using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace RedisExcel
{
    /// <summary>
    /// Manages Pub/Sub subscriptions with per-channel reference counting.
    ///
    /// A single StackExchange.Redis handler per (host, channel, pattern) broadcasts to
    /// every registered listener. Disconnecting a topic removes only its own listener;
    /// the channel is unsubscribed when the last listener leaves. On connection restore,
    /// all channels of the host are automatically re-subscribed.
    ///
    /// Hardened lifecycle: network calls (Subscribe/Unsubscribe) are serialized per
    /// channel through a dedicated lock but always run outside the listener lock, so
    /// message fan-out never blocks behind socket I/O. Subscribe racing with Dispose()
    /// either throws ObjectDisposedException or fully rolls back its listener, and
    /// channel entries are removed with a value-checked atomic remove so a stale
    /// removal can never evict a state concurrently re-created for the same key.
    ///
    /// Duplicate suppression: feeds republish unchanged values constantly; identical
    /// consecutive payloads are compared as raw bytes (no string decoding) and skipped
    /// when SkipRepeatedMessages is on (default). Patterns are never deduplicated
    /// because different channels can interleave messages with the same payload.
    ///
    /// Before: each topic registered its own handler and DisconnectData called
    /// Unsubscribe(channel) without a handler, tearing down every other topic's handler
    /// on the same channel and never re-subscribing them.
    /// </summary>
    public sealed class RedisSubscriptionManager
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly RedisConnectionManager _connections;
        private readonly ConcurrentDictionary<string, ChannelState> _channels =
            new ConcurrentDictionary<string, ChannelState>();
        private readonly bool _skipRepeated;
        private long _nextListenerId;
        private volatile bool _disposed;

        public RedisSubscriptionManager(RedisConnectionManager connections)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _skipRepeated = AppConfig.Current.SkipRepeatedMessages;
        }

        public int ChannelCount => _channels.Count;

        public int ListenerCount
        {
            get
            {
                int total = 0;
                foreach (var state in _channels.Values)
                    total += state.Listeners.Count;
                return total;
            }
        }

        public IDisposable Subscribe(string host, string channel, bool pattern, Action<string> onMessage)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(RedisSubscriptionManager));
            if (string.IsNullOrWhiteSpace(host)) throw new ArgumentException("host is required", nameof(host));
            if (string.IsNullOrWhiteSpace(channel)) throw new ArgumentException("channel is required", nameof(channel));
            if (onMessage == null) throw new ArgumentNullException(nameof(onMessage));

            string key = MakeKey(host, channel, pattern);
            ChannelState state;
            long id;
            while (true)
            {
                state = _channels.GetOrAdd(key, _ => new ChannelState(host, channel, pattern, _skipRepeated));
                bool needSubscribe;
                lock (state.Sync)
                {
                    if (state.Disposed)
                    {
                        Thread.Yield();
                        continue;
                    }
                    id = Interlocked.Increment(ref _nextListenerId);
                    state.Listeners[id] = onMessage;
                    needSubscribe = state.Listeners.Count == 1;
                    state.RebuildSnapshot();
                    // Network I/O is deliberately kept outside this lock (below).
                }

                if (needSubscribe)
                {
                    bool subscribed;
                    try
                    {
                        subscribed = state.TryEnsureSubscribed(_connections);
                    }
                    catch
                    {
                        // StackExchange.Redis may have registered the handler before
                        // throwing; undo everything so no zombie subscription is left.
                        lock (state.Sync)
                        {
                            state.Listeners.TryRemove(id, out _);
                            state.RebuildSnapshot();
                        }
                        ((ICollection<KeyValuePair<string, ChannelState>>)_channels)
                            .Remove(new KeyValuePair<string, ChannelState>(key, state));
                        state.ReleaseSubscription();
                        throw;
                    }

                    if (!subscribed)
                    {
                        // The state was disposed while this lock was released (manager
                        // shutdown): drop the listener and retry. The leading _disposed
                        // check throws on the next iteration.
                        lock (state.Sync)
                        {
                            state.Listeners.TryRemove(id, out _);
                            state.RebuildSnapshot();
                        }
                        continue;
                    }
                }

                break;
            }
            logger.Debug($"Subscribe: host={host}, channel={channel}, pattern={pattern}, listeners={state.Listeners.Count}");
            return new Registration(this, state, id);
        }

        public void Dispose()
        {
            _disposed = true;
            foreach (var state in _channels.Values)
            {
                lock (state.Sync)
                    state.Disposed = true;
                state.ReleaseSubscription();
            }
            _channels.Clear();
        }

        private void Remove(ChannelState state, long id)
        {
            bool lastListener;
            lock (state.Sync)
            {
                if (state.Disposed || !state.Listeners.TryRemove(id, out _))
                    return;
                state.RebuildSnapshot();
                lastListener = state.Listeners.IsEmpty;
                if (lastListener)
                    state.Disposed = true;
            }
            if (!lastListener)
                return;

            // Value-checked atomic remove: a state concurrently re-created under the
            // same registry key must never be evicted by this stale removal.
            ((ICollection<KeyValuePair<string, ChannelState>>)_channels)
                .Remove(new KeyValuePair<string, ChannelState>(
                    MakeKey(state.Host, state.Name, state.Pattern), state));
            state.ReleaseSubscription();
            logger.Debug($"Remove: host={state.Host}, channel={state.Name}, pattern={state.Pattern} unsubscribed");
        }

        /// <summary>
        /// Registry key for a channel state. The host length is length-prefixed so
        /// hosts and channels that themselves contain the \u0001 separator cannot
        /// collide. Format: {L|P}\u0001{host.Length}:{host}{channel}.
        /// </summary>
        internal static string MakeKey(string host, string channel, bool pattern)
        {
            return $"{(pattern ? 'P' : 'L')}\u0001{host.Length}:{host}{channel}";
        }

        private sealed class ChannelState
        {
            private static readonly Action<string>[] EmptyListeners = new Action<string>[0];

            public readonly string Host;
            public readonly string Name;
            public readonly bool Pattern;
            public readonly RedisChannel Channel;
            public readonly ConcurrentDictionary<long, Action<string>> Listeners =
                new ConcurrentDictionary<long, Action<string>>();
            public readonly object Sync = new object();
            public bool Disposed;

            private readonly bool _skipRepeated;
            private readonly object _serSync = new object();
            private volatile Action<string>[] _listenersSnapshot = EmptyListeners;
            private ISubscriber _subscriber;
            private RedisValue _lastMessage;
            private bool _hasLastMessage;

            public ChannelState(string host, string channel, bool pattern, bool skipRepeated)
            {
                Host = host;
                Name = channel;
                Pattern = pattern;
                _skipRepeated = skipRepeated;
                Channel = new RedisChannel(
                    channel,
                    pattern ? RedisChannel.PatternMode.Pattern : RedisChannel.PatternMode.Literal);
            }

            /// <summary>
            /// Copy-on-write snapshot of the listeners: HandleMessage reads it without
            /// locks. Called under Sync whenever a listener is added or removed.
            /// (The previous Listeners.Values enumeration allocated a list copy for
            /// every single message.)
            /// </summary>
            public void RebuildSnapshot()
            {
                var snapshot = new Action<string>[Listeners.Count];
                int index = 0;
                foreach (var kvp in Listeners)
                    snapshot[index++] = kvp.Value;
                _listenersSnapshot = snapshot;
            }

            /// <summary>
            /// Ensures exactly one StackExchange.Redis handler is registered for this
            /// channel. Serialized by _serSync; the listener lock is only touched to
            /// observe disposal. Returns false when the state was disposed before the
            /// handler could be installed.
            /// </summary>
            public bool TryEnsureSubscribed(RedisConnectionManager connections)
            {
                lock (_serSync)
                {
                    lock (Sync)
                    {
                        if (Disposed)
                            return false;
                    }
                    if (_subscriber != null)
                        return true;
                    var subscriber = connections.GetSubscriber(Host);
                    // Assign first: if Subscribe throws after registering the handler,
                    // the rollback path can still unsubscribe it.
                    _subscriber = subscriber;
                    subscriber.Subscribe(Channel, HandleMessage);
                    return true;
                }
            }

            /// <summary>
            /// Clears the stored subscriber under _serSync and unsubscribes the handler
            /// outside every lock. Nulling first makes concurrent ReleaseSubscription
            /// calls no-ops instead of double-unsubscribing the same handler.
            /// </summary>
            public void ReleaseSubscription()
            {
                ISubscriber subscriber;
                lock (_serSync)
                {
                    subscriber = _subscriber;
                    _subscriber = null;
                }
                if (subscriber == null)
                    return;
                try
                {
                    subscriber.Unsubscribe(Channel, HandleMessage);
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, $"ReleaseSubscription: host={Host}, channel={Name}");
                }
            }

            private void HandleMessage(RedisChannel channel, RedisValue message)
            {
                // Duplicate suppression: identical consecutive payloads (price feeds
                // republish unchanged values constantly) change nothing in Excel, so
                // skip the string decode and the whole fan-out. StackExchange.Redis
                // delivers messages for a channel sequentially, so no lock is needed.
                // Patterns are excluded because different channels interleave here.
                if (_skipRepeated && !Pattern && _hasLastMessage && message == _lastMessage)
                    return;
                _lastMessage = message;
                _hasLastMessage = true;

                string text = message; // implicit RedisValue -> string conversion (may be null, as in the original code)
                var listeners = _listenersSnapshot;
                for (int i = 0; i < listeners.Length; i++)
                {
                    try
                    {
                        listeners[i](text);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"HandleMessage: listener error host={Host}, channel={Name}");
                    }
                }
            }
        }

        private sealed class Registration : IDisposable
        {
            private RedisSubscriptionManager _owner;
            private readonly ChannelState _state;
            private readonly long _id;

            public Registration(RedisSubscriptionManager owner, ChannelState state, long id)
            {
                _owner = owner;
                _state = state;
                _id = id;
            }

            public void Dispose()
            {
                var owner = Interlocked.Exchange(ref _owner, null);
                if (owner != null)
                    owner.Remove(_state, _id);
            }
        }
    }
}

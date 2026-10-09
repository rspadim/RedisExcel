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
    /// Listeners may carry a caller-defined origin tag (for example "RTD" or "UDF"),
    /// used only by the origin-scoped counters (ListenerCountWithOrigin(string) /
    /// ChannelCountWithOrigin(string)); message delivery ignores it.
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

        /// <summary>Total number of channel states (literal and pattern).</summary>
        public int ChannelCount => _channels.Count;

        /// <summary>Total number of registered listeners across all channels.</summary>
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

        /// <summary>
        /// Number of channel states that have at least one listener with the given
        /// caller-defined origin tag (for example "RTD" or "UDF"). Comparison is
        /// ordinal (case-sensitive); a null origin matches untagged listeners.
        /// </summary>
        public int ChannelCountWithOrigin(string origin)
        {
            int total = 0;
            foreach (var state in _channels.Values)
            {
                if (HasListenerWithOrigin(state, origin))
                    total++;
            }
            return total;
        }

        /// <summary>
        /// Number of registered listeners with the given caller-defined origin tag
        /// (for example "RTD" or "UDF"). Comparison is ordinal (case-sensitive);
        /// a null origin matches untagged listeners.
        /// </summary>
        public int ListenerCountWithOrigin(string origin)
        {
            int total = 0;
            foreach (var state in _channels.Values)
            {
                foreach (var listener in state.Listeners.Values)
                {
                    if (listener != null && string.Equals(listener.Origin, origin, StringComparison.Ordinal))
                        total++;
                }
            }
            return total;
        }

        private static bool HasListenerWithOrigin(ChannelState state, string origin)
        {
            foreach (var listener in state.Listeners.Values)
            {
                if (listener != null && string.Equals(listener.Origin, origin, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Registers a listener for the channel. The optional <paramref name="origin"/>
        /// is a caller-defined tag (for example "RTD" or "UDF") used only for counting
        /// via <see cref="ListenerCountWithOrigin(string)"/> / <see cref="ChannelCountWithOrigin(string)"/>.
        /// </summary>
        public IDisposable Subscribe(string host, string channel, bool pattern, Action<string> onMessage, string origin = null)
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
                if (_disposed)
                    throw new ObjectDisposedException(nameof(RedisSubscriptionManager));

                state = _channels.GetOrAdd(key, _ => new ChannelState(host, channel, pattern, _skipRepeated));
                lock (state.Sync)
                {
                    if (state.Disposed)
                    {
                        Thread.Yield();
                        continue;
                    }
                    id = Interlocked.Increment(ref _nextListenerId);
                    state.Listeners[id] = new Listener(onMessage, origin);
                    state.RebuildSnapshot();
                    // Network I/O is deliberately kept outside this lock (below).
                }

                // Always ensure the handler is installed: a joiner may arrive while the
                // first subscriber's Subscribe is still failing/rolling back, so the
                // fast path (_subscriber != null) matters, not a listener count.
                bool subscribed;
                try
                {
                    subscribed = state.TryEnsureSubscribed(_connections, this);
                }
                catch
                {
                    // Roll back only this listener; other joiners on the same state
                    // keep the channel alive and retry on their own path.
                    RollbackListener(state, id);
                    throw;
                }

                if (!subscribed)
                {
                    // State or manager disposed while the listener lock was released;
                    // undo this listener and retry (the loop re-checks _disposed).
                    RollbackListener(state, id);
                    continue;
                }

                if (_disposed)
                {
                    // Dispose() completed while we were subscribing: the manager is
                    // shutting down, so fail with the documented contract.
                    RollbackListener(state, id);
                    throw new ObjectDisposedException(nameof(RedisSubscriptionManager));
                }

                break;
            }
            logger.Debug($"Subscribe: host={host}, channel={channel}, pattern={pattern}, origin={origin ?? "<null>"}, listeners={state.Listeners.Count}");
            return new Registration(this, state, id);
        }

        /// <summary>
        /// Removes a single listener after a failed/lost Subscribe. Only when this was
        /// the last listener the state is marked disposed, evicted (value-checked) and
        /// unsubscribed; otherwise concurrent joiners keep the channel alive.
        /// </summary>
        private void RollbackListener(ChannelState state, long id)
        {
            bool lastListener;
            lock (state.Sync)
            {
                if (!state.Listeners.TryRemove(id, out _))
                    return;
                state.RebuildSnapshot();
                lastListener = state.Listeners.IsEmpty;
                if (lastListener)
                    state.Disposed = true;
            }
            if (!lastListener)
                return;
            RemoveChannelEntry(MakeKey(state.Host, state.Name, state.Pattern), state);
            state.ReleaseSubscription();
        }

        /// <summary>
        /// Removes the registry entry only if it still maps to this exact state.
        /// .NET Framework has no atomic value-checked TryRemove, so a mismatched
        /// (freshly re-created) entry is simply restored instead of retried.
        /// </summary>
        private void RemoveChannelEntry(string key, ChannelState state)
        {
            if (_channels.TryRemove(key, out var removed) && !ReferenceEquals(removed, state))
                _channels.TryAdd(key, removed);
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

            RemoveChannelEntry(MakeKey(state.Host, state.Name, state.Pattern), state);
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

        /// <summary>
        /// One registered listener plus its caller-defined origin tag ("RTD", "UDF",
        /// ...). The tag is never used for delivery, only for diagnostics/counting.
        /// </summary>
        private sealed class Listener
        {
            public readonly Action<string> Callback;
            public readonly string Origin;

            public Listener(Action<string> callback, string origin)
            {
                Callback = callback;
                Origin = origin;
            }
        }

        private sealed class ChannelState
        {
            private static readonly Action<string>[] EmptyListeners = new Action<string>[0];

            public readonly string Host;
            public readonly string Name;
            public readonly bool Pattern;
            public readonly RedisChannel Channel;
            public readonly ConcurrentDictionary<long, Listener> Listeners =
                new ConcurrentDictionary<long, Listener>();
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
                    snapshot[index++] = kvp.Value.Callback;
                _listenersSnapshot = snapshot;
            }

            /// <summary>
            /// Ensures exactly one StackExchange.Redis handler is registered for this
            /// channel. Serialized by _serSync; the listener lock is only touched to
            /// observe disposal. Returns false when the state OR the owning manager
            /// was disposed before the handler could be installed. On a failed
            /// Subscribe the stored subscriber is reset and best-effort unsubscribed,
            /// so a concurrent joiner retries cleanly.
            /// </summary>
            public bool TryEnsureSubscribed(RedisConnectionManager connections, RedisSubscriptionManager manager)
            {
                lock (_serSync)
                {
                    lock (Sync)
                    {
                        if (Disposed)
                            return false;
                    }
                    // The manager sets its flag before touching any state, so this
                    // closes the window where a state is not yet marked disposed
                    // while the manager is already shutting down.
                    if (manager._disposed)
                        return false;
                    if (_subscriber != null)
                        return true;
                    var subscriber = connections.GetSubscriber(Host);
                    // Assign first: if Subscribe throws after registering the handler,
                    // the rollback path can still unsubscribe it.
                    _subscriber = subscriber;
                    try
                    {
                        subscriber.Subscribe(Channel, HandleMessage);
                    }
                    catch
                    {
                        // Reset the slot so the next joiner (or retry) attempts a fresh
                        // Subscribe instead of trusting a half-installed handler.
                        _subscriber = null;
                        try
                        {
                            subscriber.Unsubscribe(Channel, HandleMessage);
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, $"TryEnsureSubscribed rollback: host={Host}, channel={Name}");
                        }
                        throw;
                    }
                    return true;
                }
            }

            /// <summary>
            /// Clears the stored subscriber and unsubscribes it while holding _serSync,
            /// so a concurrent TryEnsureSubscribed can never install a fresh handler
            /// between the clear and the unsubscribe. The listener lock is not held.
            /// </summary>
            public void ReleaseSubscription()
            {
                lock (_serSync)
                {
                    var subscriber = _subscriber;
                    _subscriber = null;
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

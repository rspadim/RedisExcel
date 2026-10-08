using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
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
        private long _nextListenerId;

        public RedisSubscriptionManager(RedisConnectionManager connections)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
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
            if (string.IsNullOrWhiteSpace(host)) throw new ArgumentException("host is required", nameof(host));
            if (string.IsNullOrWhiteSpace(channel)) throw new ArgumentException("channel is required", nameof(channel));
            if (onMessage == null) throw new ArgumentNullException(nameof(onMessage));

            string key = MakeKey(host, channel, pattern);
            ChannelState state;
            long id;
            while (true)
            {
                state = _channels.GetOrAdd(key, _ => new ChannelState(host, channel, pattern));
                lock (state.Sync)
                {
                    if (state.Disposed)
                        continue;
                    id = Interlocked.Increment(ref _nextListenerId);
                    state.Listeners[id] = onMessage;
                    if (state.Listeners.Count == 1)
                    {
                        try
                        {
                            state.Subscribe(_connections);
                        }
                        catch
                        {
                            state.Listeners.TryRemove(id, out _);
                            throw;
                        }
                    }
                    break;
                }
            }
            logger.Debug($"Subscribe: host={host}, channel={channel}, pattern={pattern}, listeners={state.Listeners.Count}");
            return new Registration(this, state, id);
        }

        /// <summary>Re-subscribes every channel of a host after a reconnect (StackExchange.Redis does not do this on its own).</summary>
        public void ResubscribeHost(string host)
        {
            foreach (var state in _channels.Values)
            {
                if (!string.Equals(state.Host, host, StringComparison.Ordinal))
                    continue;
                try
                {
                    lock (state.Sync)
                    {
                        if (state.Disposed)
                            continue;
                        state.Unsubscribe();
                        state.Subscribe(_connections);
                    }
                    logger.Info($"ResubscribeHost: re-subscribed host={host}, channel={state.Name}, pattern={state.Pattern}");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"ResubscribeHost: failed host={host}, channel={state.Name}, pattern={state.Pattern}");
                }
            }
        }

        public void Dispose()
        {
            foreach (var state in _channels.Values)
            {
                lock (state.Sync)
                    state.Disposed = true;
                try
                {
                    state.Unsubscribe();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "Dispose: unsubscribe error");
                }
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
                lastListener = state.Listeners.IsEmpty;
                if (lastListener)
                    state.Disposed = true;
            }
            if (!lastListener)
                return;

            _channels.TryRemove(MakeKey(state.Host, state.Name, state.Pattern), out _);
            state.Unsubscribe();
            logger.Debug($"Remove: host={state.Host}, channel={state.Name}, pattern={state.Pattern} unsubscribed");
        }

        internal static string MakeKey(string host, string channel, bool pattern)
        {
            return $"{host}\u0001{(pattern ? 'P' : 'L')}\u0001{channel}";
        }

        private sealed class ChannelState
        {
            public readonly string Host;
            public readonly string Name;
            public readonly bool Pattern;
            public readonly RedisChannel Channel;
            public readonly ConcurrentDictionary<long, Action<string>> Listeners =
                new ConcurrentDictionary<long, Action<string>>();
            public readonly object Sync = new object();
            public bool Disposed;

            private ISubscriber _subscriber;

            public ChannelState(string host, string channel, bool pattern)
            {
                Host = host;
                Name = channel;
                Pattern = pattern;
                Channel = new RedisChannel(
                    channel,
                    pattern ? RedisChannel.PatternMode.Pattern : RedisChannel.PatternMode.Literal);
            }

            public void Subscribe(RedisConnectionManager connections)
            {
                _subscriber = connections.GetSubscriber(Host);
                _subscriber.Subscribe(Channel, HandleMessage);
            }

            public void Unsubscribe()
            {
                if (_subscriber == null)
                    return;
                try
                {
                    _subscriber.Unsubscribe(Channel, HandleMessage);
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, $"Unsubscribe: host={Host}, channel={Name}");
                }
            }

            private void HandleMessage(RedisChannel channel, RedisValue message)
            {
                string text = message; // implicit RedisValue -> string conversion (may be null, as in the original code)
                foreach (var listener in Listeners.Values)
                {
                    try
                    {
                        listener(text);
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

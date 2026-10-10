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
    /// Listeners may carry a caller-defined origin tag (for example "RTD" or "UDF"),
    /// used only by the origin-scoped counters (ListenerCountWithOrigin(string) /
    /// ChannelCountWithOrigin(string)); message delivery ignores it.
    ///
    /// Hardened lifecycle: network calls (Subscribe/Unsubscribe) are serialized per
    /// channel through a dedicated lock but always run outside the listener lock, so
    /// message fan-out never blocks behind socket I/O. Subscribe racing with Dispose()
    /// usually throws ObjectDisposedException or fully rolls back its listener; when
    /// Dispose() wins the final check the registration is still returned and is torn
    /// down with the manager. Channel entry removal re-checks the mapping and restores
    /// a raced fresh entry instead of evicting it (see RemoveChannelEntry).
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

        /// <summary>
        /// Raised after a listener was successfully registered (literal or
        /// pattern, first listener or joiner, any origin): (host, channel,
        /// pattern). Consumer-side dedup markers are cleared on join so a
        /// returning listener still receives the next payload when its value
        /// did not change. Handler exceptions are caught and logged; they
        /// never fail the Subscribe call.
        /// </summary>
        internal static event Action<string, string, bool> ListenerJoined;

        private readonly RedisConnectionManager _connections;
        private readonly ConcurrentDictionary<string, ChannelState> _channels =
            new ConcurrentDictionary<string, ChannelState>();
        private readonly bool _skipRepeated;
        private long _nextListenerId;
        private volatile bool _disposed;

        /// <summary>Test-only override of <see cref="ConfigRoot.SkipRepeatedMessages"/>
        /// (null = use the process configuration). The production sources are
        /// compiled directly into the smoke/unit test projects, so this seam
        /// lets those tests pin the duplicate-suppression branch
        /// deterministically instead of depending on a machine-local
        /// RedisExcel.json. Must be set before the manager is constructed
        /// (the flag is captured per instance).</summary>
#pragma warning disable 0649 // assigned only by the linked test sources
        internal static bool? SkipRepeatedOverrideForTests;
#pragma warning restore 0649

        private static bool SkipRepeatedEnabled =>
            SkipRepeatedOverrideForTests ?? AppConfig.Current.SkipRepeatedMessages;

        /// <summary>Effective duplicate-suppression flag after applying the test
        /// override; test-only, so the smoke test can branch deterministically.</summary>
        internal static bool SkipRepeatedEnabledForTests => SkipRepeatedEnabled;

        // Per-key stripes serializing the network handoff (subscribe and
        // unsubscribe) ACROSS channel-state generations for the same registry
        // key. StackExchange.Redis reuses its internal Subscription objects per
        // channel: a joiner attaching while the previous generation is being
        // torn down could skip the wire SUBSCRIBE (its internal map entry was
        // already being removed) and leave the channel permanently deaf - churn
        // stress reproduced misses with 10 s gaps. Serializing the handoff
        // closes that window.
        private static readonly StripedLocks NetworkLockStripes = new StripedLocks(64);

        private static object NetworkGate(string key)
        {
            return NetworkLockStripes.For(key);
        }

        public RedisSubscriptionManager(RedisConnectionManager connections)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _skipRepeated = SkipRepeatedEnabled;
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
        /// Whether any live channel state for the host still has listeners.
        /// Consumed by the connection manager: evicting such a multiplexer
        /// would silence those subscriptions permanently.
        /// </summary>
        public bool HasActiveSubscribers(string host)
        {
            foreach (var state in _channels.Values)
            {
                if (string.Equals(state.Host, host, StringComparison.Ordinal) && !state.Listeners.IsEmpty)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Registers a listener for the channel. The optional <paramref name="origin"/>
        /// is a caller-defined tag (for example "RTD" or "UDF") used only for counting
        /// via <see cref="ListenerCountWithOrigin(string)"/> / <see cref="ChannelCountWithOrigin(string)"/>.
        /// </summary>
        /// <remarks>
        /// A concurrent <see cref="Dispose"/> can have a third outcome beyond
        /// "throws" / "rolled back": when the final disposal check loses the race
        /// by an instant, the registration is returned and is disposed together
        /// with the manager (or when its own token is disposed); it is never left
        /// half-installed. Disposing the returned token is always safe and
        /// idempotent.
        /// </remarks>
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
            int disposedSpins = 0;
            while (true)
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(RedisSubscriptionManager));

                state = _channels.GetOrAdd(key, _ => new ChannelState(host, channel, pattern, _skipRepeated));
                lock (state.Sync)
                {
                    if (state.Disposed)
                    {
                        // The remover marks the state disposed BEFORE removing the
                        // mapping. Wait bounded for that removal; if it never comes
                        // (scheduler pause, or the remover lost a restore race),
                        // evict the stale mapping ourselves so this can never spin
                        // forever on the Excel thread.
                        if (++disposedSpins > 500)
                        {
                            RemoveChannelEntry(key, state);
                            if (_channels.TryGetValue(key, out var stillCurrent) && ReferenceEquals(stillCurrent, state))
                                throw new InvalidOperationException("a disposed channel state could not be recreated");
                        }
                        else
                        {
                            Thread.Yield();
                        }
                        continue;
                    }
                    bool wasActive = !state.Listeners.IsEmpty;
                    id = Interlocked.Increment(ref _nextListenerId);
                    state.Listeners[id] = new Listener(onMessage, origin);
                    state.RebuildSnapshot();
                    // A listener joining an already-active state never saw the
                    // payload currently held by the shared dedup marker, so clear
                    // the marker: the next publish - even an identical repeat -
                    // must be fanned out to the joiner (and everyone else) instead
                    // of being suppressed and leaving the new cell blank.
                    if (wasActive)
                        state.ResetLastMessage();
                    // Network I/O is deliberately kept outside this lock (below).
                }

                // Always ensure the handler is installed: a joiner may arrive while the
                // first subscriber's Subscribe is still failing/rolling back, so the
                // fast path (_subscriber != null) matters, not a listener count.
                bool subscribed;
                try
                {
                    // Serialize the network handoff per registry key across
                    // generations: a joiner must not attach to a dying
                    // generation's subscription object (see NetworkGate).
                    lock (NetworkGate(key))
                    {
                        subscribed = state.TryEnsureSubscribed(_connections, this);
                    }
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
            RaiseListenerJoined(host, channel, pattern);
            return new Registration(this, state, id);
        }

        /// <summary>
        /// Fans out a join notification with per-handler isolation: a failing
        /// consumer cleanup must not break an already successful Subscribe.
        /// </summary>
        private static void RaiseListenerJoined(string host, string channel, bool pattern)
        {
            var handlers = ListenerJoined;
            if (handlers == null)
                return;
            foreach (Action<string, string, bool> handler in handlers.GetInvocationList())
            {
                try
                {
                    handler(host, channel, pattern);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"ListenerJoined handler failed: host={host}, channel={channel}, pattern={pattern}");
                }
            }
        }

        /// <summary>
        /// Removes a single listener after a failed/lost Subscribe. Only when this was
        /// the last listener the state is marked disposed, removed from the registry
        /// and unsubscribed; otherwise concurrent joiners keep the channel alive.
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
            string key = MakeKey(state.Host, state.Name, state.Pattern);
            RemoveChannelEntry(key, state);
            lock (NetworkGate(key))
                state.ReleaseSubscription();
        }

        /// <summary>
        /// Removes the registry entry only if it still maps to this exact state.
        /// .NET Framework has no atomic value-checked TryRemove, so the mapping
        /// is re-checked and a raced fresh entry that got evicted is restored
        /// with a bounded retry; an entry that is currently installed is never
        /// removed, so a live state stays reachable.
        /// </summary>
        private void RemoveChannelEntry(string key, ChannelState state)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                if (!_channels.TryGetValue(key, out var current))
                    return; // already gone (Dispose cleared it or a remover won)
                if (!ReferenceEquals(current, state))
                    return; // a fresh live state owns the key: leave it alone
                if (_channels.TryRemove(key, out var removed))
                {
                    if (ReferenceEquals(removed, state))
                        return; // our state was removed cleanly
                    // The mapping changed between the check and the remove:
                    // put the evicted live entry back (retry if the restore
                    // races with yet another insert).
                    if (_channels.TryAdd(key, removed))
                        return;
                }
            }
            logger.Debug($"RemoveChannelEntry: gave up restoring a raced entry for key={key}");
        }

        public void Dispose()
        {
            _disposed = true;
            foreach (var state in _channels.Values)
            {
                lock (state.Sync)
                    state.Disposed = true;
                lock (NetworkGate(MakeKey(state.Host, state.Name, state.Pattern)))
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
            lock (NetworkGate(MakeKey(state.Host, state.Name, state.Pattern)))
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
            private readonly object _dedupSync = new object();
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
            /// Clears the duplicate-suppression marker so the next payload is
            /// treated as new for the whole channel. Called when a listener joins
            /// an already-active state: the marker is shared, and the joiner never
            /// saw the suppressed payload, so an identical republish must be
            /// fanned out instead of skipped.
            /// </summary>
            public void ResetLastMessage()
            {
                // Serialized with HandleMessage through _dedupSync: the marker
                // must never be read while another thread replaces it (a torn
                // RedisValue read used to make the comparison throw and leave
                // the channel deaf). A reset racing a message either wins (the
                // payload is fanned out again) or loses (the marker already
                // reflects the delivered payload) - both safe.
                lock (_dedupSync)
                {
                    _hasLastMessage = false;
                    _lastMessage = RedisValue.Null;
                }
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
                            // Rollback unsubscribe failed: the handler may still be
                            // registered on the shared subscriber, so this state can
                            // never be trusted again. Poison it and drop the registry
                            // mapping; the next Subscribe builds a fresh state
                            // instead of risking a duplicate handler registration.
                            logger.Error(ex, $"TryEnsureSubscribed rollback: host={Host}, channel={Name}");
                            Disposed = true;
                            manager.RemoveChannelEntry(MakeKey(Host, Name, Pattern), this);
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
                        // Fire-and-forget teardown (literal and pattern channels
                        // use the same path): the handler is removed client-side
                        // immediately, while the command is sent asynchronously,
                        // so an unreachable host cannot block the Excel main
                        // thread for up to SyncTimeout per channel.
                        subscriber.Unsubscribe(Channel, HandleMessage, CommandFlags.FireAndForget);
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
                // does NOT deliver plain handlers sequentially: MessageCompletable
                // falls back to the thread pool, so several messages of one channel
                // can run HandleMessage concurrently and the marker needs the lock
                // (a torn RedisValue read once made the comparison throw inside the
                // subscriber callback - which StackExchange.Redis swallows - and the
                // channel stayed deaf afterwards). Patterns are excluded because
                // different channels interleave here. Suppression stays best-effort
                // under concurrency: two racing copies may both fan out, which is
                // harmless.
                if (_skipRepeated && !Pattern)
                {
                    lock (_dedupSync)
                    {
                        if (_hasLastMessage && message == _lastMessage)
                            return;
                        _lastMessage = message;
                        _hasLastMessage = true;
                    }
                }

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

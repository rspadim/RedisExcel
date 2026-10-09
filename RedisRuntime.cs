using NLog;
using System;

namespace RedisExcel
{
    /// <summary>
    /// Shared add-in state. Connections and subscriptions are expensive resources:
    /// they are created once per Excel process and reused by every RTD server
    /// (Excel may instantiate more than one) and by UDF calls.
    /// Both fields are volatile and published in an order that keeps them
    /// consistent: a reader that observes a non-null <c>_connections</c> is
    /// guaranteed to observe a non-null <c>_subscriptions</c>.
    /// </summary>
    internal static class RedisRuntime
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();

        private static volatile RedisConnectionManager _connections;
        private static volatile RedisSubscriptionManager _subscriptions;

        // Set at the start of Shutdown, inside the lock, and cleared only by
        // ResetAfterAddInReload (a same-process add-in reload): after an
        // explicit Shutdown (for example AutoClose) no code path may
        // recreate the managers, no matter how many RTD timer ticks or UDF
        // calls race with the teardown.
        private static volatile bool _shutdown;

        public static RedisConnectionManager Connections
        {
            get
            {
                EnsureInitialized();
                var value = _connections;
                if (value == null)
                {
                    // Shutdown raced with this access: EnsureInitialized now
                    // throws instead of recreating the managers.
                    EnsureInitialized();
                    value = _connections;
                }
                return value ?? throw new InvalidOperationException("RedisRuntime is shutting down");
            }
        }

        public static RedisSubscriptionManager Subscriptions
        {
            get
            {
                EnsureInitialized();
                var value = _subscriptions;
                if (value == null)
                {
                    // Shutdown raced with this access: EnsureInitialized now
                    // throws instead of recreating the managers.
                    EnsureInitialized();
                    value = _subscriptions;
                }
                return value ?? throw new InvalidOperationException("RedisRuntime is shutting down");
            }
        }

        private static void EnsureInitialized()
        {
            // Checked before creation: once Shutdown has started, callers (timer
            // ticks included) must fail instead of resurrecting disposed managers.
            if (_shutdown)
                throw new InvalidOperationException("RedisRuntime is shutting down");
            if (_connections != null && _subscriptions != null)
                return;
            lock (Sync)
            {
                if (_shutdown)
                    throw new InvalidOperationException("RedisRuntime is shutting down");
                if (_connections != null && _subscriptions != null)
                    return;
                var connections = new RedisConnectionManager();
                var subscriptions = new RedisSubscriptionManager(connections);
                // StackExchange.Redis re-subscribes channels automatically after a
                // reconnect, so no explicit resubscribe wiring is required here.
                // Publish _subscriptions first: a reader that observes a non-null
                // _connections must also observe a fully initialized _subscriptions.
                _subscriptions = subscriptions;
                _connections = connections;
            }
        }

        public static void Shutdown()
        {
            lock (Sync)
            {
                // The flag is set BEFORE the managers are disposed/nulled (and
                // is cleared only by ResetAfterAddInReload), so a racing
                // EnsureInitialized either observes it and throws, or waits
                // on the lock and then observes it.
                _shutdown = true;
                try
                {
                    _subscriptions?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Shutdown: error disposing subscriptions");
                }
                try
                {
                    _connections?.Shutdown();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Shutdown: error disposing connections");
                }
                _subscriptions = null;
                _connections = null;
            }
        }

        /// <summary>
        /// Clears the shutdown tombstone so a surviving AppDomain (add-in
        /// unloaded and reloaded without an Excel restart) can lazily recreate
        /// the managers on the next access. No-op when the runtime was not
        /// shut down; thread-safe under the same lock used by
        /// <see cref="Shutdown"/>.
        /// </summary>
        public static void ResetAfterAddInReload()
        {
            lock (Sync)
            {
                if (!_shutdown)
                    return;
                // Managers were disposed and nulled by Shutdown; clear the
                // fields again for clarity, then drop the tombstone last so a
                // racing EnsureInitialized never observes a cleared flag with
                // stale managers.
                _subscriptions = null;
                _connections = null;
                _shutdown = false;
            }
        }
    }
}

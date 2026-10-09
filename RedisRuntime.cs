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

        public static RedisConnectionManager Connections
        {
            get { EnsureInitialized(); return _connections; }
        }

        public static RedisSubscriptionManager Subscriptions
        {
            get { EnsureInitialized(); return _subscriptions; }
        }

        private static void EnsureInitialized()
        {
            if (_connections != null && _subscriptions != null)
                return;
            lock (Sync)
            {
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
    }
}

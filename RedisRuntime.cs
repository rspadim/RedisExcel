using NLog;
using System;

namespace RedisExcel
{
    /// <summary>
    /// Shared add-in state. Connections and subscriptions are expensive resources:
    /// they are created once per Excel process and reused by every RTD server
    /// (Excel may instantiate more than one) and by UDF calls.
    /// </summary>
    internal static class RedisRuntime
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();

        private static RedisConnectionManager _connections;
        private static RedisSubscriptionManager _subscriptions;

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
            if (_connections != null)
                return;
            lock (Sync)
            {
                if (_connections != null)
                    return;
                var connections = new RedisConnectionManager();
                var subscriptions = new RedisSubscriptionManager(connections);
                connections.ConnectionRestored += subscriptions.ResubscribeHost;
                _connections = connections;
                _subscriptions = subscriptions;
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

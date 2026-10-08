using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Threading;

namespace RedisExcel
{
    /// <summary>Connection pool identifier, kept to preserve the meaning of the status/count functions.</summary>
    public enum RedisPool
    {
        RtdData,
        RtdSub,
        UdfData
    }

    /// <summary>
    /// Single creation/cache point for ConnectionMultiplexer.
    /// Before there were three duplicated connection implementations (RTD data, RTD sub,
    /// UDF) with different options for the same host, no Lazy protection, and event
    /// handlers registered in only one of them.
    /// </summary>
    public sealed class RedisConnectionManager
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> _rtdData =
            new ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>>();
        private readonly ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> _rtdSub =
            new ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>>();
        private readonly ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> _udfData =
            new ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>>();

        /// <summary>Raised when a subscription connection is restored (channels must be re-subscribed).</summary>
        public event Action<string> ConnectionRestored;

        public int RtdConnectionCount => _rtdData.Count + _rtdSub.Count;
        public int UdfConnectionCount => _udfData.Count;

        public IDatabase GetDatabase(string host, RedisPool pool) => GetConnection(host, pool).GetDatabase();

        public ISubscriber GetSubscriber(string host) => GetConnection(host, RedisPool.RtdSub).GetSubscriber();

        public ConnectionMultiplexer GetConnection(string host, RedisPool pool)
        {
            var dictionary = DictionaryFor(pool);
            // Lazy avoids creating duplicate connections when GetOrAdd is called concurrently.
            var lazy = dictionary.GetOrAdd(host, h => new Lazy<ConnectionMultiplexer>(
                () => Connect(h, pool), LazyThreadSafetyMode.ExecutionAndPublication));
            return lazy.Value;
        }

        private ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> DictionaryFor(RedisPool pool)
        {
            switch (pool)
            {
                case RedisPool.RtdData: return _rtdData;
                case RedisPool.RtdSub: return _rtdSub;
                default: return _udfData;
            }
        }

        private ConnectionMultiplexer Connect(string host, RedisPool pool)
        {
            var config = AppConfig.Current;
            int timeoutMs = pool == RedisPool.UdfData ? config.UDF.timeout : config.RTD.timeout;

            var options = ConfigurationOptions.Parse(host);
            options.AbortOnConnectFail = false;
            options.ConnectTimeout = timeoutMs;
            options.SyncTimeout = timeoutMs;
            options.ClientName =
                $"{ClientNamePrefix(pool)} :: {BuildInfo.Tag} :: " +
                $"{Environment.UserDomainName}\\{Environment.UserName} :: {Environment.MachineName}";

            logger.Info($"RedisConnect: opening {pool} connection to {host} (timeout={timeoutMs}ms, client={options.ClientName})");
            var mux = ConnectionMultiplexer.Connect(options);

            mux.ConnectionFailed += (sender, args) =>
            {
                logger.Info($"RedisConnect: connection lost ({pool}) host={host}, endpoint={args.EndPoint}, " +
                            $"failure={args.FailureType}, error={args.Exception?.Message}");
            };
            mux.ConnectionRestored += (sender, args) =>
            {
                logger.Info($"RedisConnect: connection restored ({pool}) host={host}");
                if (pool == RedisPool.RtdSub)
                    ConnectionRestored?.Invoke(host);
            };
            return mux;
        }

        private static string ClientNamePrefix(RedisPool pool)
        {
            return pool == RedisPool.UdfData ? "RedisUDF" : "RedisRTD";
        }

        public void Shutdown()
        {
            foreach (var dictionary in new[] { _rtdData, _rtdSub, _udfData })
            {
                foreach (var kv in dictionary)
                {
                    try
                    {
                        kv.Value.Value.Close();
                        kv.Value.Value.Dispose();
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(ex, "Shutdown: error closing connection");
                    }
                }
                dictionary.Clear();
            }
        }
    }

    internal static class BuildInfo
    {
#if GIT_TAG
        public const string Tag = GIT_TAG;
#else
        public const string Tag = "not GITHub";
#endif
    }
}

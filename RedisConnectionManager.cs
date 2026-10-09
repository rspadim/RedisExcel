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
        private readonly ConcurrentDictionary<string, IDatabase> _databases =
            new ConcurrentDictionary<string, IDatabase>();
        private readonly ConcurrentDictionary<string, ISubscriber> _subscribers =
            new ConcurrentDictionary<string, ISubscriber>();

        public int RtdConnectionCount => _rtdData.Count + _rtdSub.Count;
        public int UdfConnectionCount => _udfData.Count;

        public IDatabase GetDatabase(string host, RedisPool pool)
        {
            // Cache the lightweight wrappers: volatile worksheet functions call this
            // constantly, and ConnectionMultiplexer.GetDatabase() allocates per call.
            // If the factory throws (connect failure), GetOrAdd does not insert the
            // entry, so the next call retries instead of caching the failure.
            string key = PoolKey(host, pool);
            return _databases.GetOrAdd(key, _ => GetConnection(host, pool).GetDatabase());
        }

        public ISubscriber GetSubscriber(string host) => GetSubscriber(host, RedisPool.RtdSub);

        public ISubscriber GetSubscriber(string host, RedisPool pool)
        {
            // Same GetOrAdd semantics as GetDatabase: a throwing factory leaves the
            // cache empty, so transient connect failures are not cached.
            string key = PoolKey(host, pool);
            return _subscribers.GetOrAdd(key, _ => GetConnection(host, pool).GetSubscriber());
        }

        private static string PoolKey(string host, RedisPool pool)
        {
            return ((int)pool) + "\u0001" + host;
        }

        public ConnectionMultiplexer GetConnection(string host, RedisPool pool)
        {
            var dictionary = DictionaryFor(pool);
            // Lazy avoids creating duplicate connections when GetOrAdd is called concurrently.
            var lazy = dictionary.GetOrAdd(host, h => new Lazy<ConnectionMultiplexer>(
                () => Connect(h, pool), LazyThreadSafetyMode.ExecutionAndPublication));
            try
            {
                return lazy.Value;
            }
            catch
            {
                // Do not cache a failed connect: drop the entry so the next call retries.
                // Only remove our own entry in case another thread already replaced it.
                if (dictionary.TryGetValue(host, out var current) && ReferenceEquals(current, lazy))
                    dictionary.TryRemove(host, out _);
                throw;
            }
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
            };
            return mux;
        }

        private static string ClientNamePrefix(RedisPool pool)
        {
            return pool == RedisPool.UdfData ? "RedisUDF" : "RedisRTD";
        }

        public void Shutdown()
        {
            _databases.Clear();
            _subscribers.Clear();
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
        /// <summary>
        /// Release tag embedded by the CI via AssemblyInformationalVersion
        /// (msbuild /p:InformationalVersion=&lt;tag&gt;). "dev" for local builds.
        /// </summary>
        public static string Tag { get; } = ResolveTag();

        private static string ResolveTag()
        {
            try
            {
                var attribute = Attribute.GetCustomAttribute(
                    typeof(BuildInfo).Assembly,
                    typeof(System.Reflection.AssemblyInformationalVersionAttribute))
                    as System.Reflection.AssemblyInformationalVersionAttribute;
                var value = attribute?.InformationalVersion;
                return string.IsNullOrWhiteSpace(value) ? "dev" : value;
            }
            catch
            {
                return "dev";
            }
        }
    }
}

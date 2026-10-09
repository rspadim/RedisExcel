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

        // Set at the start of Shutdown and never cleared: a connect that races
        // with shutdown must not leave an un-disposed multiplexer behind.
        private volatile bool _shutdown;

        public int RtdConnectionCount => _rtdData.Count + _rtdSub.Count;
        public int UdfConnectionCount => _udfData.Count;

        public IDatabase GetDatabase(string host, RedisPool pool)
        {
            // Cache the lightweight wrappers: volatile worksheet functions call this
            // constantly, and ConnectionMultiplexer.GetDatabase() allocates per call.
            // If the factory throws (connect failure), GetOrAdd does not insert the
            // entry, so the next call retries instead of caching the failure.
            //
            // Shutdown fence: never hand out a wrapper bound to a multiplexer that
            // Shutdown is closing. The flag is re-checked after GetOrAdd because a
            // racing call could otherwise insert the entry after the caches were
            // cleared; that entry is dropped instead of being left behind.
            if (_shutdown)
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            string key = PoolKey(host, pool);
            var database = _databases.GetOrAdd(key, _ => GetConnection(host, pool).GetDatabase());
            if (_shutdown)
            {
                _databases.TryRemove(key, out _);
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            }
            return database;
        }

        public ISubscriber GetSubscriber(string host) => GetSubscriber(host, RedisPool.RtdSub);

        public ISubscriber GetSubscriber(string host, RedisPool pool)
        {
            // Same GetOrAdd semantics as GetDatabase: a throwing factory leaves the
            // cache empty, so transient connect failures are not cached. The same
            // shutdown fence applies: no wrapper is handed out after Shutdown, and
            // an entry that raced with the cache clear is removed again.
            if (_shutdown)
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            string key = PoolKey(host, pool);
            var subscriber = _subscribers.GetOrAdd(key, _ => GetConnection(host, pool).GetSubscriber());
            if (_shutdown)
            {
                _subscribers.TryRemove(key, out _);
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            }
            return subscriber;
        }

        private static string PoolKey(string host, RedisPool pool)
        {
            return ((int)pool) + "\u0001" + host;
        }

        public ConnectionMultiplexer GetConnection(string host, RedisPool pool)
        {
            // Shutdown fence: a cached multiplexer is already being closed by
            // Shutdown, so never hand it out once the flag is set.
            if (_shutdown)
                throw new InvalidOperationException("RedisConnectionManager is shutting down");

            var dictionary = DictionaryFor(pool);
            // Lazy avoids creating duplicate connections when GetOrAdd is called concurrently.
            var lazy = dictionary.GetOrAdd(host, h => new Lazy<ConnectionMultiplexer>(
                () => Connect(h, pool), LazyThreadSafetyMode.ExecutionAndPublication));
            ConnectionMultiplexer connection;
            try
            {
                connection = lazy.Value;
            }
            catch
            {
                if (_shutdown)
                {
                    // Shutdown won the race: Connect already refused/disposed the
                    // multiplexer, so drop the failed entry instead of replacing
                    // it - no re-insertion into a cache Shutdown just cleared.
                    dictionary.TryRemove(host, out _);
                }
                else
                {
                    // Atomically replace the failed attempt with a fresh Lazy so the next call
                    // retries. TryUpdate only swaps when the current entry is still ours, so a
                    // concurrently created entry is never removed (no leak, no double connect).
                    dictionary.TryUpdate(host,
                        new Lazy<ConnectionMultiplexer>(() => Connect(host, pool), LazyThreadSafetyMode.ExecutionAndPublication),
                        lazy);
                }
                throw;
            }
            if (_shutdown)
            {
                // Shutdown started while this connection was in flight. Keep the
                // entry so Shutdown's walk still sees and disposes it, but fail
                // the caller instead of returning a dying connection.
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            }
            return connection;
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
            if (_shutdown)
                throw new InvalidOperationException("RedisConnectionManager is shutting down");

            var config = AppConfig.Current;
            int timeoutMs = pool == RedisPool.UdfData ? config.UDF.timeout : config.RTD.timeout;

            var options = ParseOptions(host);
            options.AbortOnConnectFail = false;
            options.ConnectTimeout = timeoutMs;
            options.SyncTimeout = timeoutMs;
            options.ClientName =
                $"{ClientNamePrefix(pool)} :: {BuildInfo.Tag} :: " +
                $"{Environment.UserDomainName}\\{Environment.UserName} :: {Environment.MachineName}";

            logger.Info($"RedisConnect: opening {pool} connection to {host} (timeout={timeoutMs}ms, client={options.ClientName})");
            ConnectionMultiplexer mux;
            try
            {
                mux = ConnectionMultiplexer.Connect(options);
            }
            catch (Exception ex) when (IsInvalidHostError(ex))
            {
                // A malformed endpoint (for example a port above 65535) surfaces
                // through the endpoint/DNS layer as a localized argument error;
                // normalize it to a stable English message. Real connection
                // failures (RedisConnectionException and friends) pass through.
                throw new ArgumentException($"invalid Redis host '{host}'", ex);
            }
            if (_shutdown)
            {
                // Shutdown started while this connect was in flight: dispose the
                // fresh multiplexer immediately and fail the caller instead of
                // leaving an un-disposed connection behind.
                try
                {
                    mux.Dispose();
                }
                catch
                {
                }
                throw new InvalidOperationException("RedisConnectionManager is shutting down");
            }

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

        private static ConfigurationOptions ParseOptions(string host)
        {
            try
            {
                return ConfigurationOptions.Parse(host);
            }
            catch (Exception ex) when (IsInvalidHostError(ex))
            {
                // StackExchange.Redis surfaces a malformed endpoint with a
                // localized argument/format error; give callers a stable
                // English message instead (inner exception preserved).
                throw new ArgumentException($"invalid Redis host '{host}'", ex);
            }
        }

        private static bool IsInvalidHostError(Exception ex)
        {
            // ArgumentException covers ArgumentOutOfRangeException raised by the
            // endpoint layer (localized message); Format/Overflow cover raw
            // numeric port parse failures. Connection failures are other types.
            return ex is ArgumentException || ex is FormatException || ex is OverflowException;
        }

        private static string ClientNamePrefix(RedisPool pool)
        {
            return pool == RedisPool.UdfData ? "RedisUDF" : "RedisRTD";
        }

        public void Shutdown()
        {
            // Set before anything is closed/cleared: a racing connect either sees
            // the flag (and refuses) or completes before Shutdown starts walking
            // the caches, so no new connection can appear after the teardown.
            _shutdown = true;
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

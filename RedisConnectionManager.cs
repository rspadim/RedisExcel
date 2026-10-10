using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
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

        // One-owner gate for Shutdown: concurrent callers return immediately
        // instead of racing Close/Dispose on the same multiplexers.
        private int _shutdownStarted;

        // A spreadsheet argument can produce arbitrary host strings, and every
        // parseable-but-unreachable host keeps a disconnected multiplexer (with
        // reconnect tasks) cached forever. Past the cap, disconnected entries
        // are evicted and disposed; connected multiplexers are never evicted.
        private const int MaxCachedConnectionsPerPool = 512;

        // Short negative cache of connect failures per (host, pool): a burst of
        // N SUB topics / volatile UDF reads against a dead host used to pay
        // N x ConnectTimeout on the Excel thread; the first failure now makes
        // the following attempts within the window fail fast. The retry loops
        // (RTD backoff, next recalculation) still recover.
        private const int ConnectFailureMemoMs = 2000;
        private readonly ConcurrentDictionary<string, long> _recentConnectFailures =
            new ConcurrentDictionary<string, long>();

        // Live-consumer protection for eviction: hosts with active subscribers
        // are never evicted (a disconnected walk must not orphan a channel).
        private Func<string, bool> _hostProtection;

        internal void SetEvictionProtection(Func<string, bool> isHostProtected)
        {
            _hostProtection = isHostProtected;
        }

        /// <summary>Multiplexers actually created in the RTD pools; the same
        /// number as <see cref="LiveRtdConnectionCount"/> without the shutdown
        /// fence. Entries whose Lazy never ran hold nothing and are not
        /// counted.</summary>
        public int RtdConnectionCount => CountCreated(_rtdData) + CountCreated(_rtdSub);

        /// <summary>Multiplexers actually created in the UDF pool; see
        /// <see cref="RtdConnectionCount"/>.</summary>
        public int UdfConnectionCount => CountCreated(_udfData);

        /// <summary>
        /// Number of multiplexers actually created across all pools (RtdData,
        /// RtdSub, UdfData); zero while the shutdown fence is set. A cached
        /// entry whose Lazy never ran (offline host) holds nothing and is not
        /// counted, like the legacy <see cref="RtdConnectionCount"/> /
        /// <see cref="UdfConnectionCount"/> accessors.
        /// </summary>
        public int LiveConnectionCount()
        {
            // The fence stands in for a per-entry IsDisposed check
            // (ConnectionMultiplexer.IsDisposed is internal in
            // StackExchange.Redis 2.8): while it is set, multiplexers are
            // being disposed by Shutdown's walk or by the shutdown races in
            // GetConnection/Connect, so report zero instead of counting one
            // that is being torn down.
            return _shutdown ? 0 : CountCreated(_rtdData) + CountCreated(_rtdSub) + CountCreated(_udfData);
        }

        /// <summary>
        /// Live multiplexers for the RTD pools only (RtdData + RtdSub), the
        /// value behind RedisRTDConnectionCount. Zero while the shutdown
        /// fence is set, like <see cref="LiveConnectionCount"/>.
        /// </summary>
        public int LiveRtdConnectionCount() => _shutdown ? 0 : CountCreated(_rtdData) + CountCreated(_rtdSub);

        /// <summary>
        /// Live multiplexers for the UDF pool only (UdfData). Zero while the
        /// shutdown fence is set, like <see cref="LiveConnectionCount"/>.
        /// </summary>
        public int LiveUdfConnectionCount() => _shutdown ? 0 : CountCreated(_udfData);

        private static int CountCreated(ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> dictionary)
        {
            int total = 0;
            foreach (var kv in dictionary)
            {
                if (kv.Value.IsValueCreated)
                    total++;
            }
            return total;
        }

        private void EvictDisconnectedConnections(ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> dictionary, string keepHost, RedisPool pool)
        {
            foreach (var kv in dictionary)
            {
                if (dictionary.Count <= MaxCachedConnectionsPerPool)
                    break;
                if (string.Equals(kv.Key, keepHost, StringComparison.Ordinal))
                    continue;
                // Never evict a host with live subscribers: the subscription
                // manager holds handlers on that multiplexer, and disposing it
                // would silence the channel permanently (its fast paths trust
                // the stored subscriber/wrapper). Consumer-less entries still
                // keep the cap bounded.
                var protection = _hostProtection;
                if (protection != null && protection(kv.Key))
                    continue;
                if (!kv.Value.IsValueCreated)
                {
                    // A never-materialized entry holds nothing (also the state of
                    // a failed factory on .NET Framework): reclaim it so failed
                    // host entries cannot sit past the cap forever. No wrappers
                    // can exist for a multiplexer that never materialized.
                    RemoveConnectionEntry(dictionary, kv.Key, kv.Value);
                    continue;
                }
                ConnectionMultiplexer mux;
                try
                {
                    mux = kv.Value.Value; // IsValueCreated: cannot block on a connect
                }
                catch
                {
                    continue; // failed placeholder: the retry path replaces it
                }
                if (mux.IsConnected)
                    continue;
                // Conditional remove (value identity): never removes an entry a
                // concurrent caller just replaced with a fresh Lazy.
                if (RemoveConnectionEntry(dictionary, kv.Key, kv.Value))
                {
                    DisposeConnection(mux,
                        $"GetConnection: error disposing evicted connection host={kv.Key}",
                        closeFirst: true, allowCommandsToComplete: false);
                    // Invalidate the cached wrappers too: they are bound to the
                    // disposed multiplexer, and leaving them would make every
                    // later call on this host throw ObjectDisposedException
                    // forever (GetDatabase/GetSubscriber would never rebuild).
                    _databases.TryRemove(PoolKey(kv.Key, pool), out _);
                    _subscribers.TryRemove(PoolKey(kv.Key, pool), out _);
                    logger.Info($"GetConnection: evicted disconnected connection host={kv.Key} (pool over {MaxCachedConnectionsPerPool} entries)");
                }
            }
        }

        private void DisposeConnection(ConnectionMultiplexer connection, string logMessage, bool closeFirst, bool allowCommandsToComplete)
        {
            // Shared teardown: log a failure instead of surfacing it. The
            // closeFirst/allowCommandsToComplete pair mirrors each call site's
            // current sequence (eviction and the shutdown races close(false);
            // the post-connect path disposes only).
            try
            {
                if (closeFirst)
                    connection.Close(allowCommandsToComplete);
                connection.Dispose();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, logMessage);
            }
        }

        private static InvalidOperationException ShuttingDown()
        {
            return new InvalidOperationException("RedisConnectionManager is shutting down");
        }

        private static bool RemoveConnectionEntry(ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> dictionary, string host, Lazy<ConnectionMultiplexer> lazy)
        {
            return ((ICollection<KeyValuePair<string, Lazy<ConnectionMultiplexer>>>)dictionary)
                .Remove(new KeyValuePair<string, Lazy<ConnectionMultiplexer>>(host, lazy));
        }

        public IDatabase GetDatabase(string host, RedisPool pool) =>
            GetCachedWrapper(_databases, host, pool, mux => mux.GetDatabase());

        public ISubscriber GetSubscriber(string host) => GetSubscriber(host, RedisPool.RtdSub);

        public ISubscriber GetSubscriber(string host, RedisPool pool) =>
            GetCachedWrapper(_subscribers, host, pool, mux => mux.GetSubscriber());

        /// <summary>
        /// Reads/creates a cached lightweight wrapper (IDatabase or
        /// ISubscriber): volatile worksheet functions call these constantly,
        /// and ConnectionMultiplexer.GetDatabase()/GetSubscriber() allocate
        /// per call. If the factory throws (connect failure), GetOrAdd does
        /// not insert the entry, so the next call retries instead of caching
        /// the failure.
        /// Shutdown fence: never hand out a wrapper bound to a multiplexer
        /// that Shutdown is closing. The flag is re-checked after GetOrAdd
        /// because a racing call could otherwise insert the entry after the
        /// caches were cleared; that entry is dropped instead of being left
        /// behind.
        /// </summary>
        private T GetCachedWrapper<T>(ConcurrentDictionary<string, T> cache, string host, RedisPool pool, Func<ConnectionMultiplexer, T> wrapper)
        {
            if (_shutdown)
                throw ShuttingDown();
            string key = PoolKey(host, pool);
            var value = cache.GetOrAdd(key, _ => wrapper(GetConnection(host, pool)));
            if (_shutdown)
            {
                cache.TryRemove(key, out _);
                throw ShuttingDown();
            }
            return value;
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
                throw ShuttingDown();

            var dictionary = DictionaryFor(pool);
            // Lazy avoids creating duplicate connections when GetOrAdd is called concurrently.
            var lazy = dictionary.GetOrAdd(host, h => new Lazy<ConnectionMultiplexer>(
                () => Connect(h, pool), LazyThreadSafetyMode.ExecutionAndPublication));
            ConnectionMultiplexer connection;
            try
            {
                connection = lazy.Value;
            }
            catch (ArgumentException)
            {
                // Non-transient host error: ParseOptions (and Connect's
                // wrapper) normalizes a malformed endpoint to ArgumentException.
                // The same host fails identically on every retry, so drop the
                // entry instead of replacing it - no dead placeholder
                // accumulates and the pool counters stay clean.
                dictionary.TryRemove(host, out _);
                throw;
            }
            catch
            {
                // Record the failure so a burst of calls within the memo window
                // fails fast instead of repeating the connect stall.
                _recentConnectFailures[PoolKey(host, pool)] = DateTime.UtcNow.Ticks;
                if (_shutdown)
                {
                    // Shutdown won the race: Connect already refused/disposed the
                    // multiplexer, so drop the failed entry instead of replacing
                    // it - no re-insertion into a cache Shutdown just cleared.
                    dictionary.TryRemove(host, out _);
                }
                else
                {
                    // Genuine (transient) connection failure: atomically replace the
                    // failed attempt with a fresh Lazy so the next call retries.
                    // TryUpdate only swaps when the current entry is still ours, so a
                    // concurrently created entry is never removed (no leak, no double connect).
                    dictionary.TryUpdate(host,
                        new Lazy<ConnectionMultiplexer>(() => Connect(host, pool), LazyThreadSafetyMode.ExecutionAndPublication),
                        lazy);
                }
                throw;
            }
            if (_shutdown)
            {
                // Shutdown started while this connection was in flight. The
                // walk may have run before the Lazy published its value (and
                // then skipped/cleared the entry), or may already have closed
                // it: disposing here is idempotent and guarantees the mux can
                // never leak, while still failing the caller instead of
                // returning a dying connection. Also drop the entry so the
                // pool never keeps a "host -> disposed mux" placeholder
                // (mirrors the GetDatabase/GetSubscriber shutdown fences).
                DisposeConnection(connection, "GetConnection: error disposing connection created during shutdown",
                    closeFirst: true, allowCommandsToComplete: true);
                dictionary.TryRemove(host, out _);
                throw ShuttingDown();
            }
            if (dictionary.Count > MaxCachedConnectionsPerPool)
                EvictDisconnectedConnections(dictionary, host, pool);
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
                throw ShuttingDown();

            string memoKey = PoolKey(host, pool);
            if (_recentConnectFailures.TryGetValue(memoKey, out var failedAtTicks))
            {
                long elapsedMs = (DateTime.UtcNow.Ticks - failedAtTicks) / TimeSpan.TicksPerMillisecond;
                if (elapsedMs >= 0 && elapsedMs < ConnectFailureMemoMs)
                    throw new RedisConnectionException(ConnectionFailureType.UnableToConnect,
                        $"connect to '{host}' skipped (recent failure {elapsedMs}ms ago)");
            }

            var config = AppConfig.Current;
            int timeoutMs = pool == RedisPool.UdfData ? config.UDF.timeout : config.RTD.timeout;

            var options = ParseOptions(host);
            options.AbortOnConnectFail = false;
            // The default ConnectRetry (3) multiplies the first-connect stall:
            // a dead host would hold the Excel thread for attempts x timeout.
            options.ConnectRetry = 1;
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
                DisposeConnection(mux, "Connect: error disposing connection created during shutdown",
                    closeFirst: false, allowCommandsToComplete: false);
                throw ShuttingDown();
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
            _recentConnectFailures.TryRemove(memoKey, out _);
            return mux;
        }

        internal static ConfigurationOptions ParseOptions(string host)
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
            if (Interlocked.Exchange(ref _shutdownStarted, 1) != 0)
                return; // another caller already owns the teardown
            // Set before anything is closed/cleared: a racing connect either sees
            // the flag (and refuses) or completes before Shutdown starts walking
            // the caches, so no new connection can appear after the teardown.
            _shutdown = true;
            _databases.Clear();
            _subscribers.Clear();
            _recentConnectFailures.Clear();
            foreach (var dictionary in new[] { _rtdData, _rtdSub, _udfData })
            {
                foreach (var kv in dictionary)
                {
                    var lazy = kv.Value;
                    if (!lazy.IsValueCreated)
                    {
                        // Nothing was ever created for this entry: forcing
                        // Lazy.Value here could block for seconds behind an
                        // in-flight connect to an unreachable host. The racing
                        // connect sees _shutdown when it completes and disposes
                        // its own multiplexer (post-create re-check in Connect),
                        // so there is nothing left to close for this entry.
                        continue;
                    }
                    try
                    {
                        // IsValueCreated: the value is materialized, so this
                        // read cannot block behind a connect. Close(false):
                        // never wait for in-flight commands (the no-arg Close
                        // blocked the Excel thread up to SyncTimeout per
                        // multiplexer, contradicting the non-blocking shutdown).
                        kv.Value.Value.Close(false);
                        kv.Value.Value.Dispose();
                        // Drop the host entry together with its multiplexer:
                        // no "host -> disposed mux" entry may be left behind,
                        // even for an entry a racing connect published between
                        // the walk and the Clear below (mirrors
                        // GetDatabase/GetSubscriber).
                        dictionary.TryRemove(kv.Key, out _);
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

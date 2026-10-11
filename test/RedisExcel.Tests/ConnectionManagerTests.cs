using System;
using System.Collections.Concurrent;
using System.Reflection;
using System.Threading;
using StackExchange.Redis;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Offline RedisConnectionManager lifecycle tests. No Redis server is
    /// involved: every path exercised here throws, returns, or probes a dead
    /// local endpoint (127.0.0.1:1, which refuses immediately) without a
    /// server behind it. Shares the runtime-singleton collection with
    /// RedisRuntimeReloadTests: the reload tests manipulate the process-wide
    /// managers other tests run against.
    /// </summary>
    [Collection("runtime-singleton")]
    public class ConnectionManagerTests
    {
        [Fact]
        public void Shutdown_IsIdempotent()
        {
            var manager = new RedisConnectionManager();

            manager.Shutdown();
            manager.Shutdown();

            Assert.Equal(0, manager.RtdConnectionCount);
            Assert.Equal(0, manager.UdfConnectionCount);
            Assert.Equal(0, manager.LiveConnectionCount());
            Assert.Equal(0, manager.LiveRtdConnectionCount());
            Assert.Equal(0, manager.LiveUdfConnectionCount());
        }

        [Fact]
        public void GetDatabase_AfterShutdown_ThrowsShuttingDown()
        {
            var manager = new RedisConnectionManager();
            manager.Shutdown();

            var ex = Assert.Throws<InvalidOperationException>(
                () => manager.GetDatabase("localhost:6379", RedisPool.RtdData));

            Assert.Equal("RedisConnectionManager is shutting down", ex.Message);
        }

        [Fact]
        public void GetSubscriber_AfterShutdown_ThrowsShuttingDown()
        {
            var manager = new RedisConnectionManager();
            manager.Shutdown();

            var ex = Assert.Throws<InvalidOperationException>(
                () => manager.GetSubscriber("localhost:6379", RedisPool.RtdSub));

            Assert.Equal("RedisConnectionManager is shutting down", ex.Message);
        }

        [Fact]
        public void LiveConnectionCount_IsZeroBeforeAnyConnect()
        {
            var manager = new RedisConnectionManager();

            // A cached Lazy that never ran holds no multiplexer: 0, and no
            // connect is attempted for the pools that were never touched.
            Assert.Equal(0, manager.LiveConnectionCount());
        }

        [Fact]
        public void LiveRtdConnectionCount_IsZeroBeforeAnyConnect()
        {
            var manager = new RedisConnectionManager();

            // RtdData and RtdSub hold no created multiplexer: 0.
            Assert.Equal(0, manager.LiveRtdConnectionCount());
        }

        [Fact]
        public void LiveUdfConnectionCount_IsZeroBeforeAnyConnect()
        {
            var manager = new RedisConnectionManager();

            // UdfData holds no created multiplexer: 0.
            Assert.Equal(0, manager.LiveUdfConnectionCount());
        }

        [Fact]
        public void LiveConnectionCount_MaterializedPool_HonorsTheShutdownFence()
        {
            var manager = new RedisConnectionManager();

            // A real (disconnected) multiplexer materialized in the UDF pool:
            // AbortOnConnectFail=false keeps the connect from throwing, so the
            // entry is created and counted.
            manager.GetConnection(DeadLocalHost, RedisPool.UdfData);

            Assert.Equal(1, manager.LiveConnectionCount());
            Assert.Equal(1, manager.LiveUdfConnectionCount());
            Assert.Equal(0, manager.LiveRtdConnectionCount());

            SetShutdown(manager, true);

            // The fence stands in for the per-entry IsDisposed check: while it is
            // set the counters report zero even though the pool still holds the
            // multiplexer. A suppressed `_shutdown ? 0 : ...` must go red here.
            Assert.Equal(0, manager.LiveConnectionCount());
            Assert.Equal(0, manager.LiveUdfConnectionCount());
            Assert.Equal(0, manager.LiveRtdConnectionCount());

            manager.Shutdown(); // close the materialized multiplexer (no leak)
        }

        // ---------------------------------------------------------------
        // Post-create shutdown fence in Connect. The fence is only reachable
        // on the private Connect method offline (the public path needs a real
        // in-flight connect), so it is invoked through reflection.
        // ---------------------------------------------------------------

        [Fact]
        public void Connect_ShutdownFence_PreSet_RefusesWithoutConnecting()
        {
            var manager = new RedisConnectionManager();
            SetShutdown(manager, true);

            var ex = Assert.Throws<TargetInvocationException>(
                () => ConnectMethod.Invoke(manager, new object[] { DeadLocalHost, RedisPool.UdfData }));

            var inner = Assert.IsType<InvalidOperationException>(ex.InnerException);
            Assert.Equal("RedisConnectionManager is shutting down", inner.Message);
        }

        [Fact]
        public void Connect_ShutdownSetWhileInFlight_ThrowsInsteadOfReturningTheConnection()
        {
            var manager = new RedisConnectionManager();
            Exception thrown = null;
            var started = new ManualResetEventSlim(false);
            var thread = new Thread(() =>
            {
                started.Set();
                try
                {
                    ConnectMethod.Invoke(manager, new object[] { InFlightHost, RedisPool.UdfData });
                }
                catch (Exception ex)
                {
                    thrown = ex;
                }
            });

            thread.Start();
            Assert.True(started.Wait(5000));
            // Connect is invoked DIRECTLY here (not through GetConnection), so
            // there is no pool entry to observe: the only signal is time. The
            // host uses a dedicated connectTimeout=2000ms, so a 200ms wait is a
            // 10x margin inside the blocking Connect (this unroutable TEST-NET
            // address never answers). Trip the post-create fence while it is in
            // flight, so the in-flight connect must dispose its multiplexer and
            // fail the caller.
            Thread.Sleep(200);
            SetShutdown(manager, true);
            Assert.True(thread.Join(15000), "the in-flight connect never completed after the shutdown fence was set");

            Assert.NotNull(thrown); // a deleted fence returns the connection instead of throwing
            var inner = Assert.IsType<InvalidOperationException>(
                ((TargetInvocationException)thrown).InnerException);
            Assert.Equal("RedisConnectionManager is shutting down", inner.Message);
            Assert.Equal(0, manager.UdfConnectionCount); // no entry published by the refused connect
        }

        [Fact]
        public void ParseOptions_MalformedPort_ThrowsInvalidHost()
        {
            var ex = Assert.Throws<ArgumentException>(
                () => RedisConnectionManager.ParseOptions("localhost:99999"));

            Assert.Equal("invalid Redis host 'localhost:99999'", ex.Message);
            Assert.NotNull(ex.InnerException); // original endpoint error preserved
        }

        [Fact]
        public void GetConnection_MalformedHost_ThrowsAndDropsCacheEntry()
        {
            var manager = new RedisConnectionManager();

            // Deterministic host error: the failed Lazy is removed, not kept
            // as a retry placeholder, so the pool counters stay clean.
            var ex = Assert.Throws<ArgumentException>(
                () => manager.GetConnection("localhost:99999", RedisPool.UdfData));

            Assert.Contains("invalid Redis host 'localhost:99999'", ex.Message);
            Assert.Equal(0, manager.UdfConnectionCount);
            Assert.Equal(0, manager.LiveConnectionCount());
        }

        // ---------------------------------------------------------------
        // Eviction and the connect-failure memo. Reflection seeds the private
        // pool/wrapper caches and invokes the private eviction walk directly,
        // so the tests never need hundreds of real connections.
        // ---------------------------------------------------------------

        [Fact]
        public void Eviction_DoesNotEvictProtectedHosts()
        {
            var manager = new RedisConnectionManager();
            var pool = UdfPool(manager);
            int cap = MaxCachedConnectionsPerPool;

            // Every entry is evictable (its Lazy never ran): the only reason
            // the pool can stay over the cap after the walk is the protection
            // veto. Deterministic regardless of enumeration order: with the
            // veto broken, at least one entry disappears.
            for (int i = 0; i <= cap; i++)
                pool["evict.protected:" + i] = NeverMaterializingLazy();

            manager.SetEvictionProtection(host => true);
            InvokeEviction(manager, pool, keepHost: null, RedisPool.UdfData);

            Assert.Equal(cap + 1, pool.Count);
            foreach (var kv in pool)
                Assert.False(kv.Value.IsValueCreated, "a protected entry was evicted or materialized");
        }

        [Fact]
        public void Eviction_ReclaimsNeverCreatedEntries()
        {
            var manager = new RedisConnectionManager();
            var pool = UdfPool(manager);
            int cap = MaxCachedConnectionsPerPool;

            int factoryCalls = 0;
            for (int i = 0; i <= cap; i++)
            {
                pool["evict.never:" + i] = new Lazy<ConnectionMultiplexer>(
                    () =>
                    {
                        Interlocked.Increment(ref factoryCalls);
                        throw new InvalidOperationException("the Lazy factory must never run");
                    },
                    LazyThreadSafetyMode.ExecutionAndPublication);
            }

            InvokeEviction(manager, pool, keepHost: null, RedisPool.UdfData);

            // Exactly one entry over the cap: one never-created entry is
            // reclaimed and no factory is ever forced.
            Assert.Equal(cap, pool.Count);
            Assert.Equal(0, Volatile.Read(ref factoryCalls));
        }

        [Fact]
        public void Eviction_RemovesWrapperCacheEntriesOfTheEvictedHost()
        {
            var manager = new RedisConnectionManager();
            var pool = UdfPool(manager);
            var databases = Databases(manager);
            var subscribers = Subscribers(manager);
            int cap = MaxCachedConnectionsPerPool;

            string target = "evict.target:" + Guid.NewGuid().ToString("N");
            // A dead local endpoint yields a multiplexer that connects never
            // and fails fast (refused), which is what eviction targets.
            ConnectionMultiplexer mux =
                ConnectionMultiplexer.Connect("127.0.0.1:1,abortConnect=False,connectTimeout=100,connectRetry=0");
            try
            {
                Assert.False(mux.IsConnected, "the dead test endpoint unexpectedly accepted a connection");

                pool[target] = new Lazy<ConnectionMultiplexer>(() => mux, LazyThreadSafetyMode.ExecutionAndPublication);
                // Force the Lazy like GetConnection does: eviction only reaches
                // the disposal/wrapper-invalidation branch for materialized
                // multiplexers.
                Assert.Same(mux, pool[target].Value);
                for (int i = 0; i < cap; i++)
                    pool["evict.filler:" + i] = NeverMaterializingLazy();

                // Seed the wrapper caches exactly like GetDatabase/GetSubscriber
                // would for this host/pool.
                string poolKey = PoolKey(target, RedisPool.UdfData);
                databases[poolKey] = mux.GetDatabase();
                subscribers[poolKey] = mux.GetSubscriber();

                // The fillers are protected so the target is the only removable
                // entry: ConcurrentDictionary enumeration order is unspecified,
                // and the walk must visit the target deterministically.
                manager.SetEvictionProtection(host => !string.Equals(host, target, StringComparison.Ordinal));
                InvokeEviction(manager, pool, keepHost: null, RedisPool.UdfData);

                Assert.False(pool.ContainsKey(target));
                Assert.False(databases.ContainsKey(poolKey),
                    "the evicted host's cached database wrapper survived");
                Assert.False(subscribers.ContainsKey(poolKey),
                    "the evicted host's cached subscriber wrapper survived");
                Assert.Equal(cap, pool.Count);
            }
            finally
            {
                try
                {
                    mux.Dispose();
                }
                catch
                {
                    // already disposed by the eviction
                }
            }
        }

        [Fact]
        public void GetConnection_RecentConnectFailureMemo_FailsFastWithoutRetrying()
        {
            var manager = new RedisConnectionManager();
            const string host = "127.0.0.1:1";
            var pool = UdfPool(manager);

            // RedisConnectionManager always configures AbortOnConnectFail=false,
            // so a real connect to a dead host returns a disconnected
            // multiplexer instead of throwing and would never reach the memo.
            // Inject the failure through a seeded Lazy to exercise the
            // production recording path in GetConnection's catch.
            int attempts = 0;
            pool[host] = new Lazy<ConnectionMultiplexer>(
                () =>
                {
                    Interlocked.Increment(ref attempts);
                    throw new RedisConnectionException(
                        ConnectionFailureType.UnableToConnect, "seeded connect failure");
                },
                LazyThreadSafetyMode.ExecutionAndPublication);

            var first = Assert.Throws<RedisConnectionException>(
                () => manager.GetConnection(host, RedisPool.UdfData));
            Assert.Contains("seeded connect failure", first.Message);
            Assert.Equal(1, Volatile.Read(ref attempts));

            var failures = RecentConnectFailures(manager);
            Assert.NotEmpty(failures);

            // Keep the window open even if the runner paused between the two
            // calls; the recording itself is asserted above.
            failures[PoolKey(host, RedisPool.UdfData)] = DateTime.UtcNow.Ticks;

            var second = Assert.Throws<RedisConnectionException>(
                () => manager.GetConnection(host, RedisPool.UdfData));

            // Deterministic "fails fast without retrying": the memoized window
            // short-circuits before the Lazy factory, so the seeded connect is
            // attempted exactly once. (No wall-clock bound: a paused runner must
            // not turn this into a flake.)
            Assert.Contains("skipped (recent failure", second.Message);
            Assert.Equal(1, Volatile.Read(ref attempts)); // no second real connect
        }

        /// <summary>Reads the private pool-cap constant (reflection keeps the
        /// test in sync when it changes).</summary>
        private static int MaxCachedConnectionsPerPool =>
            (int)typeof(RedisConnectionManager)
                .GetField("MaxCachedConnectionsPerPool", BindingFlags.Static | BindingFlags.NonPublic)
                .GetRawConstantValue();

        // A refused local port: the connect fails fast without leaving the
        // process, so no server is needed for the materialized-pool tests.
        private const string DeadLocalHost = "127.0.0.1:1,connectTimeout=500,connectRetry=0";
        // RFC 5737 TEST-NET-1 address: connect parks until the 1000ms timeout.
        private const string UnreachableHost = "10.255.255.1:6379,connectTimeout=1000,connectRetry=1";
        // A dedicated, longer connect timeout for the in-flight shutdown-fence
        // test: it blocks on Connect for ~2s, so a short wait is safely inside.
        private const string InFlightHost = "10.255.255.1:6379,connectTimeout=2000,connectRetry=0";

        private static MethodInfo ConnectMethod =>
            typeof(RedisConnectionManager)
                .GetMethod("Connect", BindingFlags.Instance | BindingFlags.NonPublic);

        private static void SetShutdown(RedisConnectionManager manager, bool value)
            => typeof(RedisConnectionManager)
                .GetField("_shutdown", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(manager, value);

        private static ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> UdfPool(RedisConnectionManager manager)
            => (ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>>)ReadField(manager, "_udfData");

        private static ConcurrentDictionary<string, IDatabase> Databases(RedisConnectionManager manager)
            => (ConcurrentDictionary<string, IDatabase>)ReadField(manager, "_databases");

        private static ConcurrentDictionary<string, ISubscriber> Subscribers(RedisConnectionManager manager)
            => (ConcurrentDictionary<string, ISubscriber>)ReadField(manager, "_subscribers");

        private static ConcurrentDictionary<string, long> RecentConnectFailures(RedisConnectionManager manager)
            => (ConcurrentDictionary<string, long>)ReadField(manager, "_recentConnectFailures");

        private static object ReadField(RedisConnectionManager manager, string name)
            => typeof(RedisConnectionManager)
                .GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(manager);

        private static string PoolKey(string host, RedisPool pool)
            => (string)typeof(RedisConnectionManager)
                .GetMethod("PoolKey", BindingFlags.Static | BindingFlags.NonPublic)
                .Invoke(null, new object[] { host, pool });

        private static void InvokeEviction(
            RedisConnectionManager manager,
            ConcurrentDictionary<string, Lazy<ConnectionMultiplexer>> pool,
            string keepHost,
            RedisPool redisPool)
        {
            typeof(RedisConnectionManager)
                .GetMethod("EvictDisconnectedConnections", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(manager, new object[] { pool, keepHost, redisPool });
        }

        private static Lazy<ConnectionMultiplexer> NeverMaterializingLazy()
        {
            return new Lazy<ConnectionMultiplexer>(
                () => throw new InvalidOperationException("the Lazy factory must never run"),
                LazyThreadSafetyMode.ExecutionAndPublication);
        }
    }

    /// <summary>
    /// Offline RedisRuntime reload tests: the add-in can be unloaded/reloaded
    /// in the same process, so the shutdown tombstone must be clearable.
    /// Shares the runtime-singleton collection with ConnectionManagerTests:
    /// Shutdown/Reset touch the process-wide managers, which no other test in
    /// this collection may observe mid-swap.
    /// </summary>
    [Collection("runtime-singleton")]
    public class RedisRuntimeReloadTests
    {
        [Fact]
        public void ResetAfterAddInReload_AfterShutdown_RecreatesManagers()
        {
            RedisRuntime.Shutdown();
            RedisRuntime.ResetAfterAddInReload();

            Assert.NotNull(RedisRuntime.Connections);
            Assert.NotNull(RedisRuntime.Subscriptions);

            RedisRuntime.Shutdown(); // leave a clean state for other tests
        }

        [Fact]
        public void ResetAfterAddInReload_IsNoOpWhileRuntimeIsLive()
        {
            RedisRuntime.ResetAfterAddInReload(); // known state, even if a previous test shut down
            var connections = RedisRuntime.Connections;

            RedisRuntime.ResetAfterAddInReload();

            Assert.Same(connections, RedisRuntime.Connections);
            RedisRuntime.Shutdown();
        }
    }
}

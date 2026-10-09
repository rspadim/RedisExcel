using System;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Offline RedisConnectionManager lifecycle tests. No Redis server is
    /// involved: every path exercised here throws or returns before I/O
    /// (shutdown fence, never-created cache entries, malformed endpoints).
    /// </summary>
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
    }

    /// <summary>
    /// Offline RedisRuntime reload tests: the add-in can be unloaded/reloaded
    /// in the same process, so the shutdown tombstone must be clearable.
    /// </summary>
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

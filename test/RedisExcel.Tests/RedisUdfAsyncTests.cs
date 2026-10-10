using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Unit tests for <see cref="RedisUdfAsync"/>.
    ///
    /// The Run tests cover the synchronous path only (AsyncWrites = false) and
    /// pin the mode through the test seam, so the result cannot depend on a
    /// RedisExcel.json found on the machine. The queue tests drive the
    /// dispatcher's per-host FIFO queue directly; they never call
    /// ExcelAsyncUtil (which needs an Excel host) and are fully offline.
    /// </summary>
    public class RedisUdfAsyncTests
    {
        /// <summary>Runs the test with AsyncWrites forced to false.</summary>
        private static void WithSyncMode(Action test)
        {
            RedisUdfAsync.AsyncWritesOverrideForTests = false;
            try
            {
                test();
            }
            finally
            {
                RedisUdfAsync.AsyncWritesOverrideForTests = null;
            }
        }

        /// <summary>Runs the test with AsyncWrites forced to true.</summary>
        private static void WithAsyncMode(Action test)
        {
            RedisUdfAsync.AsyncWritesOverrideForTests = true;
            try
            {
                test();
            }
            finally
            {
                RedisUdfAsync.AsyncWritesOverrideForTests = null;
            }
        }

        private static string UniqueHost() => "test.host:" + Guid.NewGuid().ToString("N");

        private static async Task<object> WaitCompleted(Task<object> task, int timeoutMs = 5000)
        {
            Task completed = await Task.WhenAny(task, Task.Delay(timeoutMs));
            Assert.True(completed == task, "Queued work item did not complete in time.");
            return task;
        }

        /// <summary>Raises <paramref name="stored"/> to <paramref name="candidate"/> when larger.</summary>
        private static void InterlockedMax(ref int stored, int candidate)
        {
            int seen = Volatile.Read(ref stored);
            while (candidate > seen)
            {
                int prior = Interlocked.CompareExchange(ref stored, candidate, seen);
                if (prior == seen)
                    break;
                seen = prior;
            }
        }

        // ---------------------------------------------------------------
        // Sync path: Run must be a pure passthrough with AsyncWrites false.
        // ---------------------------------------------------------------

        [Fact]
        public void Run_SyncMode_ReturnsWorkResult()
        {
            WithSyncMode(() =>
            {
                object result = RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "key", "value" }, () => (object)42);

                Assert.Equal(42, result);
            });
        }

        [Fact]
        public void Run_SyncMode_RunsWorkOnCallerThread()
        {
            WithSyncMode(() =>
            {
                int callerThread = Thread.CurrentThread.ManagedThreadId;
                int workThread = -1;

                RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "arg" }, () =>
                {
                    workThread = Thread.CurrentThread.ManagedThreadId;
                    return null;
                });

                Assert.Equal(callerThread, workThread);
            });
        }

        [Fact]
        public void Run_SyncMode_RunsWorkExactlyOnce()
        {
            WithSyncMode(() =>
            {
                int calls = 0;

                RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "arg" }, () => { calls++; return null; });

                Assert.Equal(1, calls);
            });
        }

        [Fact]
        public void Run_SyncMode_PropagatesExceptions()
        {
            WithSyncMode(() =>
            {
                var ex = Assert.Throws<InvalidOperationException>(() =>
                {
                    RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "arg" }, () => throw new InvalidOperationException("boom"));
                });

                Assert.Equal("boom", ex.Message);
            });
        }

        [Fact]
        public void Run_SyncMode_PassesThroughReferenceResults()
        {
            WithSyncMode(() =>
            {
                var matrix = new object[,] { { 1, 2 }, { 3, 4 } };

                object result = RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "arg" }, () => matrix);

                Assert.Same(matrix, result);
            });
        }

        [Fact]
        public void Run_SyncMode_IgnoresTheHostArgument()
        {
            WithSyncMode(() =>
            {
                // The host is only used for queueing in async mode; the sync
                // path must not resolve or validate it in any way.
                object result = RedisUdfAsync.Run("RedisUDFSet", new object(), new object[] { "arg" }, () => (object)"ok");

                Assert.Equal("ok", result);
            });
        }

        [Fact]
        public void Run_SyncMode_AcceptsNullIdentityArgs()
        {
            WithSyncMode(() =>
            {
                // The sync path must not dereference the identity arguments.
                object result = RedisUdfAsync.Run("RedisUDFSet", null, null, () => (object)"ok");

                Assert.Equal("ok", result);
            });
        }

        // ---------------------------------------------------------------
        // Async mode offline: the caller refusal, the invalid-host fallback
        // and the sync passthrough still running on the calling thread.
        // ---------------------------------------------------------------

        [Fact]
        public void Run_AsyncMode_WithoutCaller_RefusesWork()
        {
            WithAsyncMode(() =>
            {
                int calls = 0;

                // Offline there is no Excel host, so XlCall.Excel(xlfCaller)
                // throws and CallerIdentityForDedup returns null: the async
                // dispatch must refuse the write instead of sharing an identity.
                object result = RedisUdfAsync.Run(
                    "RedisUDFSet",
                    null,
                    new object[] { "key", "value" },
                    () => { Interlocked.Increment(ref calls); return (object)42; });

                Assert.Equal("Error: async write needs a worksheet caller (AsyncWrites)", result);
                Assert.Equal(0, Volatile.Read(ref calls));
            });
        }

        [Fact]
        public void Run_AsyncMode_InvalidHost_FallsBackToSync()
        {
            WithAsyncMode(() =>
            {
                // A non-text host argument is rejected by ResolveHost, so
                // ResolveHostForDispatch yields null and Run must fall back to
                // the synchronous path (the core body surfaces the host error).
                Assert.Null(RedisUDF.ResolveHostForDispatch(new object()));

                int callerThread = Thread.CurrentThread.ManagedThreadId;
                int workThread = -1;

                object result = RedisUdfAsync.Run("RedisUDFSet", new object(), new object[] { "arg" }, () =>
                {
                    workThread = Thread.CurrentThread.ManagedThreadId;
                    return (object)42;
                });

                Assert.Equal(callerThread, workThread);
                Assert.Equal(42, result);
            });
        }

        [Fact]
        public void Run_SyncMode_UsesTheCallingThread()
        {
            WithSyncMode(() =>
            {
                int threadId = -1;
                int workThread = -1;
                object result = null;

                using (var completed = new ManualResetEventSlim(false))
                {
                    var thread = new Thread(() =>
                    {
                        threadId = Thread.CurrentThread.ManagedThreadId;
                        result = RedisUdfAsync.Run("RedisUDFSet", null, new object[] { "arg" }, () =>
                        {
                            workThread = Thread.CurrentThread.ManagedThreadId;
                            return (object)"ok";
                        });
                        completed.Set();
                    });
                    thread.IsBackground = true;
                    thread.Start();

                    Assert.True(completed.Wait(5000), "The dedicated thread never completed.");
                    Assert.True(thread.Join(5000), "The dedicated thread did not finish.");
                }

                Assert.Equal(threadId, workThread);
                Assert.Equal("ok", result);
            });
        }

        // ---------------------------------------------------------------
        // Queue: per-host FIFO, hosts independent, failures isolated,
        // idle queues removed.
        // ---------------------------------------------------------------

        [Fact]
        public async Task Enqueue_SameHost_PreservesFifoOrder()
        {
            string host = UniqueHost();
            var order = new List<int>();
            var tasks = new List<Task<object>>();

            for (int i = 0; i < 32; i++)
            {
                int captured = i;
                tasks.Add(RedisUdfAsync.Enqueue(host, () =>
                {
                    lock (order)
                        order.Add(captured);
                    return (object)captured;
                }));
            }

            foreach (var task in tasks)
                await WaitCompleted(task);

            Assert.Equal(Enumerable.Range(0, 32), order);
        }

        [Fact]
        public async Task Enqueue_SameHost_SerializesItems()
        {
            string host = UniqueHost();
            using (var firstStarted = new ManualResetEventSlim(false))
            using (var releaseFirst = new ManualResetEventSlim(false))
            using (var secondStarted = new ManualResetEventSlim(false))
            {
                Task<object> first = RedisUdfAsync.Enqueue(host, () =>
                {
                    firstStarted.Set();
                    releaseFirst.Wait(5000);
                    return (object)"first";
                });
                Assert.True(firstStarted.Wait(5000), "The first item never started.");

                Task<object> second = RedisUdfAsync.Enqueue(host, () =>
                {
                    secondStarted.Set();
                    return (object)"second";
                });

                // The second item must stay queued behind the blocked first.
                Assert.False(secondStarted.Wait(200), "The second item started while the first was still running.");

                releaseFirst.Set();
                await WaitCompleted(first);
                await WaitCompleted(second);

                Assert.Equal("first", await first);
                Assert.Equal("second", await second);
                Assert.True(secondStarted.IsSet);
            }
        }

        [Fact]
        public async Task Enqueue_DifferentHosts_RunConcurrently()
        {
            using (var firstStarted = new ManualResetEventSlim(false))
            using (var releaseFirst = new ManualResetEventSlim(false))
            {
                Task<object> first = RedisUdfAsync.Enqueue(UniqueHost(), () =>
                {
                    firstStarted.Set();
                    releaseFirst.Wait(5000);
                    return (object)"first";
                });
                Assert.True(firstStarted.Wait(5000), "The first host's item never started.");

                // A different host is not blocked by the first host's queue.
                Task<object> second = RedisUdfAsync.Enqueue(UniqueHost(), () => (object)"second");
                await WaitCompleted(second);

                releaseFirst.Set();
                await WaitCompleted(first);

                Assert.Equal("first", await first);
                Assert.Equal("second", await second);
            }
        }

        [Fact]
        public async Task Enqueue_WrapsExceptionsAsErrorText()
        {
            Task<object> task = RedisUdfAsync.Enqueue(
                UniqueHost(),
                () => throw new InvalidOperationException("boom"));

            await WaitCompleted(task);

            Assert.Equal("Error: boom", await task);
        }

        [Fact]
        public async Task Enqueue_FailedItemDoesNotBlockFollowers()
        {
            string host = UniqueHost();

            Task<object> failed = RedisUdfAsync.Enqueue(host, () => throw new InvalidOperationException("first"));
            Task<object> follower = RedisUdfAsync.Enqueue(host, () => (object)"second");

            await WaitCompleted(failed);
            await WaitCompleted(follower);

            Assert.Equal("Error: first", await failed);
            Assert.Equal("second", await follower);
        }

        [Fact]
        public async Task Enqueue_RemovesIdleQueue()
        {
            string host = UniqueHost();

            Task<object> task = RedisUdfAsync.Enqueue(host, () => (object)1);
            await WaitCompleted(task);

            // The release continuation runs after the item completes.
            Assert.True(
                SpinWait.SpinUntil(() => !RedisUdfAsync.HasQueueForTests(host), 5000),
                "The idle per-host queue was not removed.");
        }

        [Fact]
        public async Task Enqueue_SameHost_NeverOverlaps()
        {
            string host = UniqueHost();
            int active = 0;
            int maxActive = 0;
            var order = new ConcurrentQueue<int>();
            var tasks = new List<Task<object>>();

            for (int i = 0; i < 100; i++)
            {
                int captured = i;
                tasks.Add(RedisUdfAsync.Enqueue(host, () =>
                {
                    InterlockedMax(ref maxActive, Interlocked.Increment(ref active));
                    order.Enqueue(captured);
                    Thread.Sleep(1);
                    Interlocked.Decrement(ref active);
                    return (object)captured;
                }));
            }

            foreach (var task in tasks)
                await WaitCompleted(task, 30000);

            Assert.Equal(1, Volatile.Read(ref maxActive));
            Assert.Equal(100, order.Count);
            Assert.Equal(Enumerable.Range(0, 100), order);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Offline tests for <see cref="RedisWriteObservable"/>, the Excel-DNA
    /// observable used by the AsyncWrites dispatch. Subscribe is exercised
    /// directly (there is no Excel host and no Redis): it must enqueue on the
    /// calling thread, deliver exactly one terminal result and then complete,
    /// while the shared per-host queue keeps same-host writes serial and lets
    /// different hosts run independently. Dispose is a no-op by contract, so a
    /// detached subscription still completes.
    /// </summary>
    public class RedisUdfObserveTests
    {
        private static string UniqueHost() => "test.observe.host:" + Guid.NewGuid().ToString("N");

        /// <summary>
        /// IExcelObserver test double recording OnNext values, OnError
        /// exceptions and the completion count, with a wait handle.
        /// </summary>
        private sealed class RecordingObserver : IExcelObserver
        {
            private readonly object _gate = new object();
            private readonly List<object> _values = new List<object>();
            private readonly List<Exception> _errors = new List<Exception>();
            private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
            private int _completedCount;

            public void OnNext(object value)
            {
                lock (_gate)
                    _values.Add(value);
            }

            public void OnError(Exception exception)
            {
                lock (_gate)
                    _errors.Add(exception);
            }

            public void OnCompleted()
            {
                lock (_gate)
                    _completedCount++;
                _completed.Set();
            }

            public object[] Values
            {
                get { lock (_gate) return _values.ToArray(); }
            }

            public Exception[] Errors
            {
                get { lock (_gate) return _errors.ToArray(); }
            }

            public int CompletedCount
            {
                get { lock (_gate) return _completedCount; }
            }

            public bool WaitCompleted(int timeoutMs = 5000) => _completed.Wait(timeoutMs);
        }

        /// <summary>
        /// IExcelObserver test double whose OnNext always throws: pins that the
        /// delivery path swallows an observer failure and still completes.
        /// </summary>
        private sealed class OnNextThrowingObserver : IExcelObserver
        {
            private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
            private int _nextCalls;
            private int _completedCount;

            public void OnNext(object value)
            {
                Interlocked.Increment(ref _nextCalls);
                throw new InvalidOperationException("observer rejected the value");
            }

            public void OnError(Exception exception)
            {
            }

            public void OnCompleted()
            {
                Interlocked.Increment(ref _completedCount);
                _completed.Set();
            }

            public int OnNextCalls => Volatile.Read(ref _nextCalls);

            public int CompletedCount => Volatile.Read(ref _completedCount);

            public bool WaitCompleted(int timeoutMs = 5000) => _completed.Wait(timeoutMs);
        }

        [Fact]
        public void Subscribe_DeliversResultOnceThenCompletes()
        {
            var observer = new RecordingObserver();
            var observable = new RedisWriteObservable(UniqueHost(), () => (object)42);

            using (observable.Subscribe(observer))
            {
                Assert.True(observer.WaitCompleted(), "The observable never completed.");
            }

            Assert.Equal(new object[] { 42 }, observer.Values);
            Assert.Equal(1, observer.CompletedCount);
            Assert.Empty(observer.Errors);
        }

        [Fact]
        public void Subscribe_WorkException_DeliversErrorTextThenCompletes()
        {
            var observer = new RecordingObserver();
            var observable = new RedisWriteObservable(
                UniqueHost(),
                () => throw new InvalidOperationException("boom"));

            using (observable.Subscribe(observer))
            {
                Assert.True(observer.WaitCompleted(), "The observable never completed.");
            }

            Assert.Equal(new object[] { "Error: boom" }, observer.Values);
            Assert.Equal(1, observer.CompletedCount);
            Assert.Empty(observer.Errors);
        }

        [Fact]
        public void Subscribe_SameHost_RunsSeriallyInSubmissionOrder()
        {
            string host = UniqueHost();
            var order = new List<string>();
            var firstStarted = new ManualResetEventSlim(false);
            var releaseFirst = new ManualResetEventSlim(false);
            var secondStarted = new ManualResetEventSlim(false);

            var first = new RedisWriteObservable(host, () =>
            {
                lock (order)
                    order.Add("first-start");
                firstStarted.Set();
                releaseFirst.Wait(5000);
                lock (order)
                    order.Add("first-end");
                return (object)"first";
            });
            var second = new RedisWriteObservable(host, () =>
            {
                lock (order)
                    order.Add("second");
                secondStarted.Set();
                return (object)"second";
            });

            var firstObserver = new RecordingObserver();
            var secondObserver = new RecordingObserver();

            using (first.Subscribe(firstObserver))
            {
                Assert.True(firstStarted.Wait(5000), "The first write never started.");

                // Second is submitted while the first is still running: the
                // shared host queue must hold it back until the first ends.
                using (second.Subscribe(secondObserver))
                {
                    Assert.False(
                        secondStarted.Wait(200),
                        "The second write started while the first was still running.");

                    releaseFirst.Set();

                    Assert.True(firstObserver.WaitCompleted(), "The first observable never completed.");
                    Assert.True(secondObserver.WaitCompleted(), "The second observable never completed.");
                }
            }

            Assert.Equal(new[] { "first-start", "first-end", "second" }, order);
            Assert.Equal(new object[] { "first" }, firstObserver.Values);
            Assert.Equal(new object[] { "second" }, secondObserver.Values);
        }

        [Fact]
        public void Subscribe_DifferentHosts_CanInterleave()
        {
            var firstStarted = new ManualResetEventSlim(false);
            var releaseFirst = new ManualResetEventSlim(false);

            var first = new RedisWriteObservable(UniqueHost(), () =>
            {
                firstStarted.Set();
                releaseFirst.Wait(5000);
                return (object)"first";
            });
            var second = new RedisWriteObservable(UniqueHost(), () => (object)"second");

            var firstObserver = new RecordingObserver();
            var secondObserver = new RecordingObserver();

            using (first.Subscribe(firstObserver))
            using (second.Subscribe(secondObserver))
            {
                Assert.True(firstStarted.Wait(5000), "The first host's write never started.");

                // A different host is not blocked by the first host's queue.
                Assert.True(secondObserver.WaitCompleted(), "The second observable never completed.");

                releaseFirst.Set();
                Assert.True(firstObserver.WaitCompleted(), "The first observable never completed.");
            }

            Assert.Equal(new object[] { "first" }, firstObserver.Values);
            Assert.Equal(new object[] { "second" }, secondObserver.Values);
        }

        [Fact]
        public void Dispose_IsNoOpAndWriteStillCompletes()
        {
            var observer = new RecordingObserver();
            using (var release = new ManualResetEventSlim(false))
            {
                var observable = new RedisWriteObservable(UniqueHost(), () =>
                {
                    release.Wait(5000);
                    return (object)"done";
                });

                IDisposable subscription = observable.Subscribe(observer);

                // Detaching the subscription (Excel disconnecting the topic)
                // must not abort a queued or executing write.
                subscription.Dispose();

                release.Set();
                Assert.True(observer.WaitCompleted(), "The disposed observable never completed.");
            }

            Assert.Equal(new object[] { "done" }, observer.Values);
            Assert.Equal(1, observer.CompletedCount);
            Assert.Empty(observer.Errors);
        }

        [Fact]
        public void Subscribe_Twice_EnqueuesOnlyOnce()
        {
            string host = UniqueHost();
            int runs = 0;
            var observable = new RedisWriteObservable(host, () =>
            {
                Interlocked.Increment(ref runs);
                return (object)1;
            });

            var first = new RecordingObserver();
            var second = new RecordingObserver();

            using (observable.Subscribe(first))
            {
                Assert.True(first.WaitCompleted(), "The first subscription never completed.");
            }

            // The one-shot guard: a duplicate Subscribe returns a disposable
            // (Excel-DNA always expects one) and never enqueues again, but the
            // second observer still receives the queued result exactly once.
            IDisposable duplicate = observable.Subscribe(second);
            Assert.NotNull(duplicate);
            duplicate.Dispose();

            Assert.True(second.WaitCompleted(), "The duplicate subscription never completed.");

            Assert.True(
                SpinWait.SpinUntil(() => !RedisUdfAsync.HasQueueForTests(host), 5000),
                "The idle per-host queue was not removed.");

            Assert.Equal(1, Volatile.Read(ref runs));
            Assert.Equal(1, second.CompletedCount);
            Assert.Equal(new object[] { 1 }, second.Values);
            Assert.Empty(second.Errors);
        }

        [Fact]
        public async Task Subscribe_ConcurrentDuplicates_EnqueueOnceAndDeliverToAll()
        {
            string host = UniqueHost();
            int runs = 0;
            var observable = new RedisWriteObservable(host, () =>
            {
                Interlocked.Increment(ref runs);
                return (object)7;
            });

            const int subscribers = 8;
            var observers = new RecordingObserver[subscribers];
            var ready = new ManualResetEventSlim(false);
            var tasks = new List<Task>();

            for (int i = 0; i < subscribers; i++)
            {
                int index = i;
                observers[index] = new RecordingObserver();
                tasks.Add(Task.Run(() =>
                {
                    ready.Wait();
                    observable.Subscribe(observers[index]).Dispose();
                }));
            }

            ready.Set();
            await Task.WhenAll(tasks);

            // The Interlocked guard must admit exactly one enqueue; every
            // racing subscriber still receives the same single result.
            Assert.True(
                SpinWait.SpinUntil(() => !RedisUdfAsync.HasQueueForTests(host), 5000),
                "The idle per-host queue was not removed.");

            foreach (RecordingObserver observer in observers)
            {
                Assert.True(observer.WaitCompleted(), "A racing subscription never completed.");
                Assert.Equal(1, observer.CompletedCount);
                Assert.Equal(new object[] { 7 }, observer.Values);
                Assert.Empty(observer.Errors);
            }

            Assert.Equal(1, Volatile.Read(ref runs));
        }

        [Fact]
        public void Subscribe_OnNextThrows_StillCompletes()
        {
            var observer = new OnNextThrowingObserver();
            int runs = 0;
            var observable = new RedisWriteObservable(UniqueHost(), () =>
            {
                Interlocked.Increment(ref runs);
                return (object)42;
            });

            // Subscribe itself must not surface an observer failure (delivery
            // happens on the queue continuation).
            IDisposable subscription = observable.Subscribe(observer);
            Assert.NotNull(subscription);

            Assert.True(observer.WaitCompleted(), "The observable never completed after OnNext threw.");

            subscription.Dispose();

            Assert.Equal(1, observer.OnNextCalls);
            Assert.Equal(1, observer.CompletedCount);
            Assert.Equal(1, Volatile.Read(ref runs));
        }

        [Fact]
        public void Subscribe_DisposingAQueuedWriteIsNoOp()
        {
            string host = UniqueHost();
            int secondRuns = 0;
            var firstStarted = new ManualResetEventSlim(false);
            var releaseFirst = new ManualResetEventSlim(false);

            var first = new RedisWriteObservable(host, () =>
            {
                firstStarted.Set();
                releaseFirst.Wait(5000);
                return (object)"first";
            });
            var second = new RedisWriteObservable(host, () =>
            {
                Interlocked.Increment(ref secondRuns);
                return (object)"second";
            });

            var firstObserver = new RecordingObserver();
            var secondObserver = new RecordingObserver();

            using (first.Subscribe(firstObserver))
            {
                Assert.True(firstStarted.Wait(5000), "The first write never started.");

                IDisposable secondSubscription = second.Subscribe(secondObserver);
                Assert.True(RedisUdfAsync.HasQueueForTests(host), "The second write was not queued.");

                // Excel detaching the queued topic must not cancel the write.
                secondSubscription.Dispose();

                releaseFirst.Set();

                Assert.True(firstObserver.WaitCompleted(), "The first observable never completed.");
                Assert.True(secondObserver.WaitCompleted(), "The disposed observable never completed.");
            }

            Assert.Equal(1, Volatile.Read(ref secondRuns));
            Assert.Equal(new object[] { "second" }, secondObserver.Values);
            Assert.Equal(1, secondObserver.CompletedCount);
            Assert.Empty(secondObserver.Errors);
        }

        [Fact]
        public void Subscribe_EnqueuesBeforeReturning()
        {
            string host = UniqueHost();
            var observer = new RecordingObserver();

            using (var release = new ManualResetEventSlim(false))
            {
                var observable = new RedisWriteObservable(host, () =>
                {
                    release.Wait(5000);
                    return (object)"done";
                });

                // Subscribe enqueues synchronously before it returns; the
                // blocked item keeps the queue alive, so this is deterministic.
                using (observable.Subscribe(observer))
                {
                    Assert.True(
                        RedisUdfAsync.HasQueueForTests(host),
                        "Subscribe returned before enqueuing the write.");

                    release.Set();
                    Assert.True(observer.WaitCompleted(), "The observable never completed.");
                }
            }

            Assert.Equal(new object[] { "done" }, observer.Values);
            Assert.Equal(1, observer.CompletedCount);
            Assert.Empty(observer.Errors);
        }
    }
}

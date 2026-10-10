using ExcelDna.Integration;
using NLog;
using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace RedisExcel
{
    /// <summary>
    /// Async dispatch for UDF work, gated by the AsyncWrites configuration key.
    ///
    /// When AsyncWrites is false this is a pure synchronous passthrough: the
    /// caller's thread runs the work and sees its value or exception exactly as
    /// before. When it is true the work is handed to a per-host serial queue and
    /// Excel-DNA's async machinery makes the calculation return Excel's pending
    /// marker (#N/A) immediately; the cell is filled with the real value (or the
    /// usual "Error: ..." text) when the queued item finishes.
    ///
    /// Ordering: every work item submitted for a host is appended to that
    /// host's FIFO queue, so writes to one host execute in submission order.
    /// Submission happens in RedisWriteObservable.Subscribe, which Excel-DNA
    /// invokes synchronously on the Excel thread during the internal RTD
    /// ConnectData of the Observe registration, so the queue order is the
    /// formula evaluation order. Different hosts use different queues and can
    /// run concurrently. Idle queues are removed from the registry as soon as
    /// their last item completes.
    ///
    /// Identity caveats: the calling cell reference is part of the identity, so
    /// inserting/moving rows or columns (coordinates change) makes the next
    /// evaluation a new call and re-issues the write. While the internal RTD
    /// topic stays connected, repeated evaluations with unchanged arguments
    /// return the cached completed value, so the write is not re-issued (this
    /// includes F9 on volatile write cells under AsyncWrites). Calls without a
    /// worksheet caller are refused with an Error cell because the identity
    /// would be shared or unstable.
    /// </summary>
    internal static class RedisUdfAsync
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // ExcelAsyncUtil identifies an async call by the (name, parameters)
        // pair. The prefix keeps these keys apart from other add-ins/wrappers
        // that happen to use the same UDF function names.
        private const string AsyncNamePrefix = "RedisUdfAsync:";

        private sealed class HostQueue
        {
            public readonly object Gate = new object();

            // Tail of the chained task sequence: the previous item, or a
            // completed sentinel for a brand-new queue.
            public Task<object> Tail = Task.FromResult<object>(null);
        }

        private static readonly ConcurrentDictionary<string, HostQueue> _queues =
            new ConcurrentDictionary<string, HostQueue>(StringComparer.Ordinal);

        /// <summary>
        /// Test-only override of <see cref="AppConfig.AsyncWrites"/> (null =
        /// use the process configuration). The offline unit tests pin this to
        /// false so they exercise the sync path without touching Excel async.
        /// </summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static bool? AsyncWritesOverrideForTests;
#pragma warning restore 0649

        private static bool AsyncWritesEnabled => AsyncWritesOverrideForTests ?? AppConfig.AsyncWrites;

        /// <summary>
        /// Frozen entry point used by the RedisUDF write functions. In sync mode
        /// (AsyncWrites false) it returns <c>work()</c> directly on the caller
        /// thread - the exact pre-async behavior. In async mode it returns the
        /// object Excel expects for a pending async call.
        /// </summary>
        /// <param name="identityArgs">The UDF's own arguments, in their declared
        /// order; Excel-DNA uses them as the async call identity. They must be
        /// value-equal on the re-call that follows the completed async call, so
        /// the cached result is returned instead of running the write again.
        /// They must also be values Excel-DNA supports (null, scalars, strings,
        /// DateTime, numerics, enums, ExcelReference/ExcelError/ExcelEmpty/
        /// ExcelMissing and arrays of those). May be null (treated as empty).</param>
        internal static object Run(string functionName, object optionalHost, object[] identityArgs, Func<object> work)
        {
            // Sync mode: pure passthrough, the caller's thread runs the work.
            if (!AsyncWritesEnabled)
                return work();

            // Host resolution is used for queueing only. On an invalid host
            // (null) fall back to the sync path so the core body can surface
            // its usual "Error: ..." text on the Excel thread.
            string host = RedisUDF.ResolveHostForDispatch(optionalHost);
            if (host == null)
                return work();

            try
            {
                // The async name/parameters identify the registered call for
                // Excel. The name is stable per write function; the parameters
                // combine the calling cell reference (structurally equal on
                // the completed re-call, distinct across cells), the resolved
                // host and the UDF's own arguments (identityArgs) so a changed
                // argument is never served by an unrelated pending call.
                // Excel-DNA matches (name, parameters) by value and returns the
                // cached result on the re-call instead of registering again.
                object caller = CallerIdentityForDedup();
                if (caller == null)
                {
                    // Without a worksheet cell reference the identity is shared
                    // across callers (or unstable between the call and the
                    // delivery re-call), which could skip or duplicate the
                    // write - fail loudly instead. Rare: a call outside
                    // worksheet evaluation (e.g. Application.Run) or a
                    // transient xlfCaller failure.
                    logger.Error($"RedisUdfAsync.Run: no worksheet caller for {functionName}; async write refused");
                    return "Error: async write needs a worksheet caller (AsyncWrites)";
                }
                object parameters = new object[] { caller, host, identityArgs ?? new object[0] };
                string asyncName = AsyncNamePrefix + (functionName ?? string.Empty);

                // Observe overload with a custom one-shot observable: Excel-DNA
                // creates the observable at registration and invokes its
                // Subscribe synchronously on the Excel thread during the
                // internal RTD ConnectData, so the enqueue there is the single
                // side effect per registered call (the delivery re-call returns
                // the cached value from Excel-DNA's state lookup and never
                // re-subscribes). The queue continuation delivers
                // OnNext/OnCompleted, so no thread-pool thread waits while an
                // item is queued. The ExcelAsyncHandle overload (deferred
                // SetResult) crashed Excel (0xc0000409) under COM automation,
                // so it stays unused.
                object asyncResult = ExcelAsyncUtil.Observe(asyncName, parameters, () => new RedisWriteObservable(host, work));
                if (asyncResult == null)
                {
                    // Excel-DNA returns null when its internal RTD registration
                    // failed; Subscribe was not called, so failing here cannot
                    // duplicate the write.
                    logger.Error($"RedisUdfAsync.Run: async RTD registration failed for {functionName}");
                    return "Error: async dispatch failed (RTD registration)";
                }
                return asyncResult;
            }
            catch (Exception ex)
            {
                // The work has not started (RedisWriteObservable.Subscribe owns
                // the enqueue), so this cannot duplicate a write; surface the
                // failure instead of silently retrying on the Excel thread.
                logger.Error(ex, $"RedisUdfAsync.Run: async dispatch failed for {functionName}");
                return "Error: async dispatch failed: " + ex.Message;
            }
        }

        /// <summary>
        /// Calling cell reference used as part of the async call identity: the
        /// same cell produces a structurally equal ExcelReference on the
        /// completed re-call, while different cells never share an identity.
        /// Returns null outside a worksheet call (or on a transient xlfCaller
        /// failure); Run refuses the async dispatch in that case because the
        /// identity could be shared or unstable.
        /// </summary>
        private static object CallerIdentityForDedup()
        {
            try
            {
                object caller = XlCall.Excel(XlCall.xlfCaller);
                if (caller is ExcelReference)
                    return caller;
            }
            catch
            {
                // xlfCaller is a live C API call and can transiently fail.
            }
            return null;
        }

        /// <summary>
        /// Appends a work item to the host's FIFO queue and returns the task
        /// that completes with the item's result. The continuation chain is
        /// built under the queue's gate, so two submissions for one host can
        /// never race; a stale queue instance (removed while idle) is detected
        /// after taking the gate and the item retries on the live instance,
        /// which guarantees a host never runs items on two queues at once.
        /// </summary>
        internal static Task<object> Enqueue(string host, Func<object> work)
        {
            while (true)
            {
                HostQueue queue = _queues.GetOrAdd(host, _ => new HostQueue());
                lock (queue.Gate)
                {
                    if (!_queues.TryGetValue(host, out var current) || !ReferenceEquals(current, queue))
                        continue; // idle-removal race: retry against the live queue

                    Task<object> next = queue.Tail.ContinueWith(
                        _ => ExecuteSafely(work),
                        CancellationToken.None,
                        TaskContinuationOptions.None,
                        TaskScheduler.Default);
                    queue.Tail = next;

                    // Drop the queue once its last item completed and no newer
                    // item claimed it (ReferenceEquals check under the gate).
                    next.ContinueWith(
                        _ => ReleaseQueue(host, queue, next),
                        CancellationToken.None,
                        TaskContinuationOptions.None,
                        TaskScheduler.Default);

                    return next;
                }
            }
        }

        private static void ReleaseQueue(string host, HostQueue queue, Task<object> completed)
        {
            lock (queue.Gate)
            {
                // Remove only when this is still the dictionary's queue for the
                // host and nothing newer claimed it (both checks under the
                // gate; a racing Enqueue retries against the live instance).
                if (ReferenceEquals(queue.Tail, completed)
                    && _queues.TryGetValue(host, out var current)
                    && ReferenceEquals(current, queue))
                {
                    _queues.TryRemove(host, out _);
                }
            }
        }

        /// <summary>
        /// Runs one queued work item. Exceptions must never surface through the
        /// async handle: they are logged and converted to the same "Error: ..."
        /// text the synchronous UDF bodies return.
        /// </summary>
        private static object ExecuteSafely(Func<object> work)
        {
            try
            {
                return work();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUdfAsync: queued work item failed");
                return "Error: " + ex.Message;
            }
        }

        /// <summary>Whether a queue currently exists for the host; test-only.</summary>
        internal static bool HasQueueForTests(string host) => _queues.ContainsKey(host);
    }

    /// <summary>
    /// One-shot Excel-DNA observable for an async write. Subscribe runs
    /// synchronously on the Excel thread during the internal RTD ConnectData
    /// (inside the ExcelAsyncUtil.Observe registration), so the per-host
    /// enqueue happens there and the per-host FIFO order is the formula
    /// evaluation order. The queue continuation delivers the result (or the
    /// usual "Error: ..." text) and completion, so no thread-pool thread waits
    /// while the item is queued. Excel cancelling a calculation or detaching
    /// the topic only disposes this subscription; the returned disposable is a
    /// no-op, so a queued or executing write still runs to completion in
    /// per-host order (same semantics as the previous thread-pool delegate).
    /// </summary>
    internal sealed class RedisWriteObservable : IExcelObservable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // Every Subscribe returns this shared instance: Dispose must be a no-op
        // and the object carries no per-subscription state.
        private static readonly IDisposable NoOpDisposable = new NoOpDisposableImpl();

        private sealed class NoOpDisposableImpl : IDisposable
        {
            public void Dispose()
            {
            }
        }

        private readonly string _host;
        private readonly Func<object> _work;

        internal RedisWriteObservable(string host, Func<object> work)
        {
            _host = host;
            _work = work;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            try
            {
                // Enqueue synchronously HERE: Excel-DNA calls Subscribe during
                // the RTD ConnectData on the Excel thread, so same-host writes
                // enter the FIFO in formula evaluation order. The continuation
                // runs only when the item reaches the head of the queue.
                RedisUdfAsync.Enqueue(_host, _work).ContinueWith(
                    OnQueuedCompleted,
                    observer,
                    CancellationToken.None,
                    TaskContinuationOptions.None,
                    TaskScheduler.Default);
            }
            catch (Exception ex)
            {
                // Enqueue converts work failures to "Error: ..." task results;
                // this guards truly unexpected failures. The write did not
                // start, so reporting the failure cannot duplicate it.
                logger.Error(ex, "RedisWriteObservable: enqueue failed");
                DeliverAndComplete(observer, "Error: " + ex.Message);
            }

            // No-op: a queued/executing write completes even if Excel detaches
            // the topic before the result arrives.
            return NoOpDisposable;
        }

        private static void OnQueuedCompleted(Task<object> queued, object state)
        {
            DeliverAndComplete((IExcelObserver)state, ResultForExcel(queued));
        }

        /// <summary>
        /// Maps the queued task's outcome to the value Excel should show.
        /// ExecuteSafely already converts a work failure into "Error: ..."
        /// text; cancelled/faulted tasks are handled defensively.
        /// </summary>
        private static object ResultForExcel(Task<object> queued)
        {
            if (queued.Status == TaskStatus.RanToCompletion)
                return queued.Result;

            if (queued.IsCanceled)
                return "Error: async write was cancelled";

            Exception error = queued.Exception != null ? queued.Exception.GetBaseException() : null;
            return "Error: " + (error != null ? error.Message : "async write failed");
        }

        /// <summary>
        /// Delivers one value and then completes, never throwing into the queue
        /// continuation: Excel may already have detached the topic, and the
        /// continuation must still clear the queue slot. Failures are logged;
        /// completion is attempted even when the delivery failed.
        /// </summary>
        private static void DeliverAndComplete(IExcelObserver observer, object result)
        {
            try
            {
                observer.OnNext(result);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisWriteObservable: delivering the async write result failed");
            }

            try
            {
                observer.OnCompleted();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisWriteObservable: completing the async write observer failed");
            }
        }
    }
}

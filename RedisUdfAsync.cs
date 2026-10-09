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
    /// host's FIFO queue, so writes to one host execute in submission order
    /// (submission happens on the Excel calculation thread, i.e. in formula
    /// evaluation order). Different hosts use different queues and can run
    /// concurrently. Idle queues are removed from the registry as soon as their
    /// last item completes.
    /// </summary>
    internal static class RedisUdfAsync
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // ExcelAsyncUtil identifies an async call by the (name, parameters)
        // pair. The prefix keeps these keys apart from other add-ins/wrappers
        // that happen to use the same UDF function names.
        private const string AsyncNamePrefix = "RedisUdfAsync:";

        // Identity element used when xlfCaller returns no cell reference.
        private const string CallerIdentityMissing = "RedisUdfAsync:no-caller";

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
                object parameters = new object[]
                {
                    CallerIdentityForDedup(),
                    host,
                    identityArgs ?? new object[0]
                };
                string asyncName = AsyncNamePrefix + (functionName ?? string.Empty);

                // Classic ExcelFunc overload: the delegate runs on a
                // thread-pool thread. The enqueue lives INSIDE the delegate
                // (RunQueued): Excel re-evaluates the formula to deliver the
                // result, so the body must stay side-effect free or the write
                // would run twice. The ExcelAsyncHandle overload (deferred
                // SetResult) crashed Excel (0xc0000409) under COM automation,
                // so the battle-tested overload is used instead.
                return ExcelAsyncUtil.Run(asyncName, parameters, () => RunQueued(host, work));
            }
            catch (Exception ex)
            {
                // The work has not started (the delegate owns the enqueue), so
                // this cannot duplicate a write; surface the failure instead of
                // silently retrying on the Excel thread.
                logger.Error(ex, $"RedisUdfAsync.Run: async dispatch failed for {functionName}");
                return "Error: async dispatch failed: " + ex.Message;
            }
        }

        /// <summary>
        /// Calling cell reference used as part of the async call identity: the
        /// same cell produces a structurally equal ExcelReference on the
        /// completed re-call, while different cells never share an identity.
        /// Falls back to a constant sentinel outside a worksheet call (rare; a
        /// null element would merge unrelated no-caller cells).
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
            return CallerIdentityMissing;
        }

        /// <summary>
        /// ExcelFunc body: invoked by Excel-DNA exactly once per registered
        /// async call (single-subscription guard). The enqueue is HERE, not in
        /// Run: Excel re-evaluates the formula to deliver the result and only
        /// the identity lookup in Run is deduplicated, so a body-level enqueue
        /// would run the write a second time on that re-call. Excel cancelling
        /// a calculation only detaches the cell from the call; the queued write
        /// itself is not aborted - like an already-running synchronous write it
        /// runs to completion in per-host order.
        /// </summary>
        private static object RunQueued(string host, Func<object> work)
        {
            try
            {
                return WaitQueued(Enqueue(host, work));
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUdfAsync: queued work failed");
                return "Error: " + ex.Message;
            }
        }

        /// <summary>
        /// Waits for a queued item and returns its result (or the usual
        /// "Error: ..." text) for Excel-DNA to deliver to the cell.
        /// </summary>
        private static object WaitQueued(Task<object> queued)
        {
            try
            {
                return queued.GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUdfAsync: queued work failed");
                return "Error: " + ex.Message;
            }
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
}

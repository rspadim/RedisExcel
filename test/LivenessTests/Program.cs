using ExcelDna.Integration;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using RedisExcel.TestSupport;

namespace RedisExcel.LivenessTests
{
    /// <summary>
    /// Adversarial liveness (stall) suite for the RedisExcel production sources.
    ///
    /// Drives the REAL sources (RedisConnectionManager, RedisSubscriptionManager,
    /// RedisUdfAsync queue + RedisWriteObservable, RedisUDF publish path, TickGate)
    /// against a Redis server (default 127.0.0.1:6399; optionally a disposable
    /// redis:7-alpine container) while attacking liveness:
    ///   W) warmup/connect: ping, first live delivery, pool warmup.
    ///   A0) subscribe/unsubscribe churn hot path without the publish flood.
    ///   A) ThreadPool starvation + heavy subscribe/unsubscribe/publish churn
    ///      + an async-write backlog that must drain after release.
    ///   B) lock contention on every public entry point (subscribe/dispose,
    ///      counters, connections, UDF publishes, queue, observable, TickGate).
    ///   C) CLIENT KILL storms + (with --container, unless --skip-restart) a real
    ///      server restart mid-stream.
    ///   D1) async queue bursts (same-host FIFO, no overlap, idle removal).
    ///   D2) observable burst under ThreadPool starvation.
    ///   D3) duplicate-subscribe race + failing observer isolation.
    ///   E) dedup marker / (re)join: an identical republish must reach a joiner.
    ///
    /// Heartbeats: a dedicated process thread and a dedicated watchdog thread keep
    /// measuring while pool threads and locks are attacked; the delivery probe
    /// measures message flow and resume latency. The run fails (exit 1) when a
    /// heartbeat exceeds the stuck bound or delivery/queue never resume.
    /// RedisTimeoutException during the flood is tolerated ONLY if the entry
    /// points recover afterwards (the dedicated recovery checks).
    ///
    /// Bounds: stuck 5s, resume 30s, queue drain 20s.
    ///
    /// Modes: the default phases above, plus two 24x7-survival modes:
    ///   --matrix: scripted failure matrix (8 faults) against a managed
    ///             container under continuous traffic; per fault it asserts the
    ///             heartbeat bound, delivery resume, re-subscription (PUBSUB
    ///             NUMSUB) and monotonic counters, then resource stability.
    ///   --soak [minutes]: mixed traffic with a random matrix fault every
    ///             20-40s (default 5 minutes), same per-fault bounds, then
    ///             resource stability + GC [SUMMARY] statistics.
    ///
    /// Usage:
    ///   dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False" [--container &lt;name&gt;] [--skip-restart] [--quick]
    ///   dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False" --container &lt;name&gt; --matrix
    ///   dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False" --container &lt;name&gt; --soak 5
    /// </summary>
    internal static class Program
    {
        private static string Host = "127.0.0.1:6399";
        private static string ContainerName;   // --container <name>
        private static bool ContainerManaged;  // container started by this run
        private static string ContainerSkipReason;
        private static bool SkipRestart;       // --skip-restart
        private static bool Quick;             // --quick
        private static bool MatrixMode;        // --matrix
        private static double SoakMinutes = -1; // --soak [minutes] (default 5)

        private const string LiveChannel = "liveness:live";

        // Matrix/soak bookkeeping: everything they must explicitly dispose or
        // verify at the end (resource-stability checks).
        private static ConnectionMultiplexer AdminMux; // short-timeout direct mux for CLIENT/PUBSUB
        private static readonly List<string> TrackedChannels = new List<string>();
        private static readonly List<IDisposable> TrackedTokens = new List<IDisposable>();
        private static readonly List<string> TrackedQueueHosts = new List<string>();
        private static readonly Random Rng = new Random();
        private static long FaultsRun;
        private static long FaultOutageCount; // delivery count at the fault's outage point

        private const double StuckBoundMs = 5000;        // "no thread stuck > 5 seconds"
        private const double ResumeBoundMs = 30000;      // kill storm / restart resume
        private const double QueueDrainBoundMs = 20000;  // queued async items after release

        private static readonly Stopwatch Clock = Stopwatch.StartNew();
        private static readonly List<string> Failures = new List<string>();
        private static readonly ConcurrentDictionary<string, Latency> Lats = new ConcurrentDictionary<string, Latency>();
        private static readonly object WorstGate = new object();
        private static readonly Dictionary<string, double> Worst = new Dictionary<string, double>();

        private static RedisConnectionManager C;
        private static RedisSubscriptionManager S;
        private static DeliveryProbe Delivery;
        private static ProcessHeartbeat Proc;
        private static PoolHeartbeat Pool;
        private static WatchdogProbe Wake;
        private static long JoinedEvents;

        private static int Main(string[] args)
        {
            // Needed for the soak [SUMMARY] allocation/survival statistics.
            try { AppDomain.MonitoringIsEnabled = true; } catch { }
            try
            {
                ParseArgs(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex.Message);
                Console.WriteLine(Usage);
                return 1;
            }

            Console.WriteLine("[LIVENESS] start host=" + Host
                + " quick=" + Quick
                + " skipRestart=" + SkipRestart
                + " container=" + (ContainerName ?? "-"));
            Console.WriteLine("[LIVENESS] mode=" + (MatrixMode ? "matrix" : (SoakMinutes > 0 ? "soak" : "phases"))
                + (SoakMinutes > 0 ? " minutes=" + SoakMinutes.ToString("0.##", Inv) : ""));
            RedisSubscriptionManager.ListenerJoined += (h, c, p) => Interlocked.Increment(ref JoinedEvents);
            try
            {
                StartContainerIfRequested();
                if (MatrixMode)
                    RunMatrix();
                else if (SoakMinutes > 0)
                    RunSoak();
                else
                    RunAll();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex);
                Failures.Add("fatal: " + ex.Message);
            }
            finally
            {
                StopContainerIfManaged();
            }

            Console.WriteLine();
            Console.WriteLine("[WORST] procHeartbeatMs=" + Fmt(GetWorst("procHeartbeatMs")));
            Console.WriteLine("[WORST] watchdogMs=" + Fmt(GetWorst("watchdogMs")));
            Console.WriteLine("[WORST] deliveryGapMs=" + Fmt(GetWorst("deliveryGapMs")));
            Console.WriteLine("[WORST] deliveryResumeMs=" + Fmt(GetWorst("deliveryResumeMs")));
            Console.WriteLine("[WORST] queueDrainMs=" + Fmt(GetWorst("queueDrainMs")));
            Console.WriteLine("[WORST] poolHeartbeatGapMs=" + Fmt(GetWorst("poolHeartbeatGapMs")));
            Console.WriteLine("[SUMMARY] listenerJoinedEvents=" + Interlocked.Read(ref JoinedEvents));
            Console.WriteLine("[SUMMARY] failed=" + Failures.Count + " elapsedSec=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));
            foreach (var f in Failures) Console.WriteLine("[FAILED] " + f);
            return Failures.Count == 0 ? 0 : 1;
        }

        private const string Usage =
            "Usage: dotnet run --project test\\LivenessTests -c Release -- [host] [options]\n" +
            "  host              Redis endpoint/config, e.g. \"127.0.0.1:6399,abortConnect=False\"\n" +
            "                    (default: " + "127.0.0.1:6399" + ")\n" +
            "  --container <n>   manage a disposable redis:7-alpine container named <n>, mapped\n" +
            "                    to the host port; phase C restarts it, the run removes it at the end\n" +
            "  --skip-restart    do not restart the server in phase C (the CLIENT KILL storm still\n" +
            "                    runs); the restart sub-phase prints SKIP\n" +
            "  --quick           shortened attack windows for a fast sanity run\n" +
            "  --matrix          scripted failure matrix (continuous traffic + 8 faults +\n" +
            "                    resource stability); docker faults require --container,\n" +
            "                    destructive command faults require a loopback host\n" +
            "  --soak [minutes]  mixed traffic with a random matrix fault every 20-40s,\n" +
            "                    resource stability + GC summary at the end (default 5)";

        private static void ParseArgs(string[] args)
        {
            var positional = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (string.Equals(a, "--quick", StringComparison.OrdinalIgnoreCase))
                    Quick = true;
                else if (string.Equals(a, "--skip-restart", StringComparison.OrdinalIgnoreCase))
                    SkipRestart = true;
                else if (string.Equals(a, "--container", StringComparison.OrdinalIgnoreCase))
                {
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("--container requires a container name");
                    ContainerName = args[++i];
                }
                else if (string.Equals(a, "--matrix", StringComparison.OrdinalIgnoreCase))
                    MatrixMode = true;
                else if (string.Equals(a, "--soak", StringComparison.OrdinalIgnoreCase))
                {
                    double minutes = 5;
                    if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                    {
                        if (!double.TryParse(args[i + 1], NumberStyles.Float, Inv, out minutes))
                            throw new ArgumentException("--soak expects a number of minutes");
                        i++;
                    }
                    if (minutes <= 0)
                        throw new ArgumentException("--soak minutes must be > 0");
                    SoakMinutes = minutes;
                }
                else if (a.StartsWith("--", StringComparison.Ordinal))
                    throw new ArgumentException("unknown option '" + a + "'");
                else
                    positional.Add(a);
            }
            if (positional.Count > 0) Host = positional[0];
            if (positional.Count > 1)
                throw new ArgumentException("unexpected extra argument '" + positional[1] + "'");
            if (MatrixMode && SoakMinutes > 0)
                throw new ArgumentException("--matrix and --soak are mutually exclusive");
        }

        private static void RunAll()
        {
            StartProbes();

            Warmup();
            PhaseA0();
            PhaseA();
            PhaseB();
            PhaseC();
            PhaseD1();
            PhaseD2();
            PhaseD3();
            PhaseE();
            Cleanup();
        }

        // ------------------------------------------------------------ warmup

        private static void Warmup()
        {
            PhaseBanner("W", "warmup/connect");
            var wd = ConnectionMultiplexer.Connect(DirectConfig);
            var db = wd.GetDatabase();
            long up = WaitUntil(() => { try { db.Ping(); return true; } catch { return false; } }, 30000);
            Check("W.server-up", up >= 0, "Ping within 30s (host=" + Host + ")");
            Delivery = new DeliveryProbe();
            Delivery.Start(C, S, Host, "liveness:live");
            long prev = Delivery.Count;
            double first = Delivery.ResumeAfter(prev, ResumeBoundMs);
            Check("W.first-delivery", first >= 0, "first live delivery ms=" + Fmt(first));
            Metric("W", "firstDeliveryMs", first);
            Metric("W", "channelCount", S.ChannelCount);
            Metric("W", "listenerCount", S.ListenerCount);
            wd.Dispose();
            WarmPools();
            DumpLatencies("W");
        }

        /// <summary>Materializes every pool/subscriber used by later phases so
        /// first-connect cost is never attributed to an attack phase.</summary>
        private static void WarmPools()
        {
            try { C.GetDatabase(Host, RedisPool.RtdData).Ping(); } catch { }
            try { C.GetDatabase(Host, RedisPool.UdfData).Ping(); } catch { }
            try { C.GetSubscriber(Host, RedisPool.RtdSub).Publish(new RedisChannel("liveness:warm", RedisChannel.PatternMode.Literal), "warm-rtd"); } catch { }
            try { C.GetSubscriber(Host, RedisPool.UdfData).Publish(new RedisChannel("liveness:warm", RedisChannel.PatternMode.Literal), "warm-udf"); } catch { }
        }

        // ------------------------------------- A0: churn baseline hot path

        private static void PhaseA0()
        {
            long start = PhaseBanner("A0", "churn baseline hot path (no publish flood)");
            var leftovers = new ConcurrentBag<IDisposable>();
            var churn0 = new Worker("A0-churn", 4, (id, n) =>
            {
                string ch = "liveness:a0:churn:" + ((n / 2) % 16);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, ch, false, NoopMessage, "A");
                if ((n % 7) == 0) leftovers.Add(token);
                else token.Dispose();
                GetLat("A0.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            });
            churn0.Start();
            Thread.Sleep(Quick ? 1500 : 4000);
            churn0.StopAndJoin();
            foreach (var t in leftovers) { try { t.Dispose(); } catch { } }

            Metric("A0", "churnOps", churn0.Ops);
            Metric("A0", "churnErrors", churn0.Errors);
            Metric("A0", "churnTimeouts", churn0.Timeouts);
            Metric("A0", "durationSec", MsSince(start) / 1000.0);
            Check("A0.churn-baseline-hot-path", churn0.Ops >= (Quick ? 700 : 2000) && churn0.Errors == 0,
                "ops=" + churn0.Ops + " errors=" + churn0.Errors);
            DumpLatencies("A0");
        }

        // ------------------------------------------- A: ThreadPool starvation

        private static void PhaseA()
        {
            long start = PhaseBanner("A", "ThreadPool starvation + subscribe/unsubscribe/publish churn + async backlog");

            var release = StarvePool(out int minW, out int minIo, out int maxW, out int maxIo, out bool setMin, out bool setMax, 400);
            Check("A.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            if (!(setMin && setMax))
            {
                RestorePool(minW, minIo, maxW, maxIo, release);
                return;
            }

            // Clear the running max gap right before the attack so the gap
            // measured after release is specific to THIS starvation window.
            Pool.Reset();
            long poolTicksAtStarveStart = Pool.Ticks;
            int attackMs = Quick ? 2500 : 6000;

            // Backlog enqueued while starved: its continuations need pool threads.
            int backlog = Quick ? 60 : 150;
            string qhost = Host + "|queueA";
            var drained = new CountdownEvent(backlog);
            for (int i = 0; i < backlog; i++)
                RedisUdfAsync.Enqueue(qhost, () => { drained.Signal(); return (object)null; });

            long prevDeliveries = Delivery.Count;
            var leftovers = new ConcurrentBag<IDisposable>();
            var churn = new Worker("A-churn", 4, (id, n) =>
            {
                string ch = "liveness:a:churn:" + ((n / 2) % 16);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, ch, false, NoopMessage, "A");
                if ((n % 4) == 0) leftovers.Add(token);
                else token.Dispose();
                GetLat("A.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            });
            var pubs = new Worker("A-publish", 2, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var sub = C.GetSubscriber(Host, RedisPool.RtdSub);
                sub.Publish(new RedisChannel("liveness:a:churn:" + ((n / 2) % 16), RedisChannel.PatternMode.Literal),
                    "A-" + id + "-" + n, CommandFlags.FireAndForget);
                GetLat("A.getsubscriber-publish").Record(Stopwatch.GetTimestamp() - t0);
            });

            churn.Start();
            pubs.Start();
            Thread.Sleep(attackMs);
            churn.StopAndJoin();
            pubs.StopAndJoin();

            double poolGapDuring = Pool.MaxGapMs;
            double deliveryGapDuring = Delivery.MaxGapMs;
            double watchdogDuring = Wake.MaxIterMs;
            long poolTicksDuringAttack = Pool.Ticks - poolTicksAtStarveStart;
            int queuePendingAtAttackEnd = drained.CurrentCount;

            // Release the pool: sleepers end, continuations and heartbeats resume.
            RestorePool(minW, minIo, maxW, maxIo, release);
            WaitUntil(() => Pool.Ticks > poolTicksAtStarveStart, 3000); // let the first post-release tick record the gap
            double poolGapAfterRelease = Pool.MaxGapMs;

            long drainStart = Stopwatch.GetTimestamp();
            bool drainedOk = drained.Wait((int)QueueDrainBoundMs);
            double drainMs = MsSince(drainStart);
            double resumeMs = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);

            Metric("A", "churnOps", churn.Ops);
            Metric("A", "churnErrors", churn.Errors);
            Metric("A", "churnTimeouts", churn.Timeouts);
            Metric("A", "churnFirstError", churn.FirstError ?? "-");
            Metric("A", "publishOps", pubs.Ops);
            Metric("A", "publishErrors", pubs.Errors);
            Metric("A", "poolTicksDuringAttack", poolTicksDuringAttack);
            Metric("A", "poolHeartbeatGapAfterMs", poolGapAfterRelease);
            Metric("A", "deliveryGapDuringMs", deliveryGapDuring);
            Metric("A", "watchdogMaxMs", watchdogDuring);
            Metric("A", "queuePendingAtAttackEnd", queuePendingAtAttackEnd);
            Metric("A", "deliveryResumeMs", resumeMs);
            Metric("A", "queueDrainMs", drainMs);

            // The pool heartbeat's continuation only runs on a pool thread; while
            // the six blocked sleepers hold the clamped (4,4) pool it cannot
            // tick. The first gap measured AFTER release therefore spans the
            // whole attack window: requiring it to reach ~the window is positive
            // proof the pool was actually starved (a counter stuck at 0 could
            // also mean the heartbeat thread died).
            Check("A.attack-bit", poolTicksDuringAttack == 0 && queuePendingAtAttackEnd > 0
                    && poolGapAfterRelease >= attackMs - 500,
                "pool ticks during attack=" + poolTicksDuringAttack + " queued items stuck=" + queuePendingAtAttackEnd + " " + poolGapAfterRelease.ToString("0.0", Inv) + "ms gap measured after release (window " + attackMs + "ms)");
            Check("A.dedicated-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            Check("A.entrypoints-not-stuck", watchdogDuring < StuckBoundMs, "watchdog max iteration=" + Fmt(watchdogDuring) + "ms");
            Check("A.delivery-flows-during-pool-starvation", deliveryGapDuring >= 0 && deliveryGapDuring < StuckBoundMs,
                "delivery max gap during attack=" + Fmt(deliveryGapDuring) + "ms (inline fan-out on the SER reader thread keeps flowing)");
            Check("A.queue-drains-after-release", drainedOk, "drainMs=" + Fmt(drainMs) + " queue removed=" + (WaitUntil(() => !RedisUdfAsync.HasQueueForTests(qhost), 5000) >= 0));
            Check("A.delivery-resumes-after-release", resumeMs >= 0 && resumeMs <= ResumeBoundMs, "resumeMs=" + Fmt(resumeMs));
            Check("A.queue-removed-when-idle", WaitUntil(() => !RedisUdfAsync.HasQueueForTests(qhost), 5000) >= 0, "HasQueueForTests false");
            Check("A.no-unexpected-errors", churn.Errors - churn.Timeouts == 0,
                "churn errors=" + churn.Errors + " (all RedisTimeoutException? " + churn.Timeouts + ") first=" + (churn.FirstError ?? "-"));

            foreach (var t in leftovers) { try { t.Dispose(); } catch { } }
            NoteWorst("poolHeartbeatGapMs", poolGapAfterRelease);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("deliveryGapMs", deliveryGapDuring);
            NoteWorst("watchdogMs", watchdogDuring);
            NoteWorst("deliveryResumeMs", resumeMs);
            NoteWorst("queueDrainMs", drainMs);
            Metric("A", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("A");
        }

        // ------------------------------------------------ B: lock contention

        private static void PhaseB()
        {
            long start = PhaseBanner("B", "lock contention on every public entry point");
            string prevSyncWrite = RedisUDF.SyncWriteOverrideForTests;
            RedisUDF.SyncWriteOverrideForTests = "sync"; // reply-awaiting publishes: strongest contention
            long prevDeliveries = Delivery.Count;
            long udfErrors = 0;
            long udfTimeouts = 0;
            string firstUdfError = null;
            int latestCount = 4, pubCount = 4;
            // The runner can be a slow shared box: scale the stress workers down in
            // --quick so the artificial contention cannot starve the dedicated
            // heartbeat/thread pool and trip a bound for the wrong reason. The
            // explicit fault phases and the bounds themselves are unchanged (the
            // full run keeps the original counts).
            int contentionWorkers = Quick ? 1 : 2;
            var pubChannels = Enumerable.Range(0, pubCount).Select(i => "liveness:b:pub:" + i).ToArray();
            var latestChannels = Enumerable.Range(0, latestCount).Select(i => "liveness:b:latest:" + i).ToArray();
            var leftovers = new ConcurrentBag<IDisposable>();

            var workers = new List<Worker>();
            workers.Add(new Worker("B-subscribe", contentionWorkers, (id, n) =>
            {
                string ch = "liveness:b:churn:" + ((n / 2) % 24);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, ch, false, NoopMessage, "B");
                if ((n % 5) == 0) leftovers.Add(token);
                else token.Dispose();
                GetLat("B.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            }));
            workers.Add(new Worker("B-reads", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                int a = S.ChannelCount, b = S.ListenerCount, c = S.ChannelCountWithOrigin("B"), d = S.ListenerCountWithOrigin("RTD");
                int e = C.LiveConnectionCount(), f = C.RtdConnectionCount, g = C.UdfConnectionCount;
                int h = C.LiveRtdConnectionCount(), k = C.LiveUdfConnectionCount();
                GetLat("B.counts-livecounts").Record(Stopwatch.GetTimestamp() - t0);
                GC.KeepAlive(a + b + c + d + e + f + g + h + k);
            }));
            workers.Add(new Worker("B-publish", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var sub = C.GetSubscriber(Host, RedisPool.RtdSub);
                sub.Publish(new RedisChannel("liveness:b:churn:" + ((n / 2) % 24), RedisChannel.PatternMode.Literal),
                    "B-" + n, CommandFlags.FireAndForget);
                GetLat("B.getsubscriber-publish").Record(Stopwatch.GetTimestamp() - t0);
            }));
            workers.Add(new Worker("B-db", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                C.GetDatabase(Host, RedisPool.RtdData).Ping();
                C.GetDatabase(Host, RedisPool.UdfData).Ping();
                GetLat("B.getdatabase-ping").Record(Stopwatch.GetTimestamp() - t0);
            }));
            workers.Add(new Worker("B-conn", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var mux = C.GetConnection(Host, RedisPool.RtdSub);
                var sub = C.GetSubscriber(Host, RedisPool.RtdSub);
                GetLat("B.getconnection-getsubscriber").Record(Stopwatch.GetTimestamp() - t0);
                GC.KeepAlive(mux);
                GC.KeepAlive(sub);
            }));
            workers.Add(new Worker("B-udf-publish", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object res = RedisUDF.RedisUDFChannelPublish(pubChannels[n % pubCount], "B-" + n, Host);
                GetLat("B.udf.publish").Record(Stopwatch.GetTimestamp() - t0);
                if (res is string s && s.StartsWith("Error:", StringComparison.Ordinal))
                {
                    if (IsTimeout(s)) Interlocked.Increment(ref udfTimeouts);
                    else
                    {
                        Interlocked.Increment(ref udfErrors);
                        if (Volatile.Read(ref firstUdfError) == null) Volatile.Write(ref firstUdfError, s);
                    }
                }
            }));
            workers.Add(new Worker("B-udf-pif", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object res = RedisUDF.RedisUDFChannelPublishIfChanged(pubChannels[n % pubCount], "pif-" + n, Host);
                GetLat("B.udf.publishIfChanged").Record(Stopwatch.GetTimestamp() - t0);
                if (res is string s2 && s2.StartsWith("Error:", StringComparison.Ordinal))
                {
                    if (IsTimeout(s2)) Interlocked.Increment(ref udfTimeouts);
                    else
                    {
                        Interlocked.Increment(ref udfErrors);
                        if (Volatile.Read(ref firstUdfError) == null) Volatile.Write(ref firstUdfError, s2);
                    }
                }
            }));
            workers.Add(new Worker("B-udf-latest", 1, (id, n) =>
            {
                string ch = latestChannels[n % latestCount];
                long t0 = Stopwatch.GetTimestamp();
                string latest = RedisUDF.RedisUDFChannelLatest(ch, Host);
                GetLat("B.udf.latest").Record(Stopwatch.GetTimestamp() - t0);
                if (latest != null && latest.StartsWith("Error:", StringComparison.Ordinal))
                {
                    if (IsTimeout(latest)) Interlocked.Increment(ref udfTimeouts);
                    else
                    {
                        Interlocked.Increment(ref udfErrors);
                        if (Volatile.Read(ref firstUdfError) == null) Volatile.Write(ref firstUdfError, latest);
                    }
                }
                long t1 = Stopwatch.GetTimestamp();
                RedisUDF.RedisUDFChannelUnsubscribe(ch, Host);
                GetLat("B.udf.unsubscribe").Record(Stopwatch.GetTimestamp() - t1);
            }));
            workers.Add(new Worker("B-queue", contentionWorkers, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                Task<object> t = RedisUdfAsync.Enqueue(Host + "|queueB", () => { Thread.SpinWait(1000); return (object)n; });
                if (!t.Wait(5000)) Interlocked.Increment(ref udfErrors);
                GetLat("B.enqueue-wait").Record(Stopwatch.GetTimestamp() - t0);
            }));
            workers.Add(new Worker("B-observable", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                new RedisWriteObservable(Host + "|queueB2", () => (object)"ok").Subscribe(new RecordingObserver());
                GetLat("B.observable-subscribe").Record(Stopwatch.GetTimestamp() - t0);
            }));
            workers.Add(new Worker("B-tickgate", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                for (int k = 0; k < 20; k++)
                {
                    if (!Gate.TryEnter()) throw new InvalidOperationException("tick gate is busy");
                    Gate.Exit();
                }
                GetLat("B.tickgate").Record(Stopwatch.GetTimestamp() - t0);
            }));

            foreach (var w in workers) w.Start();
            Thread.Sleep(Quick ? 3000 : 8000);
            foreach (var w in workers) w.StopAndJoin();
            foreach (var w in workers)
                Console.WriteLine("[METRIC] phase=B worker=" + w.Name + " ops=" + w.Ops + " errors=" + w.Errors
                    + " timeouts=" + w.Timeouts + " firstError=" + (w.FirstError ?? "-"));

            long ops = workers.Sum(w => w.Ops);
            long errors = workers.Sum(w => w.Errors);
            long timeouts = workers.Sum(w => w.Timeouts);
            double resume = Delivery.ResumeAfter(prevDeliveries, 5000);

            // Post-flood recovery: a fresh subscription must succeed again once
            // the flood stops feeding the subscription connection's inbound pipe.
            bool recovered = false;
            string recoverError = null;
            string recoverChannel = "liveness:b:recover:" + Guid.NewGuid().ToString("N");
            int recoveredMessages = 0;
            WaitUntil(() =>
            {
                try
                {
                    var token = S.Subscribe(Host, recoverChannel, false, m => Interlocked.Increment(ref recoveredMessages), "B");
                    try
                    {
                        var sub = C.GetSubscriber(Host, RedisPool.RtdSub);
                        sub.Publish(new RedisChannel(recoverChannel, RedisChannel.PatternMode.Literal), "recover");
                        if (WaitUntil(() => Volatile.Read(ref recoveredMessages) > 0, 5000) < 0)
                            throw new TimeoutException("recovery delivery did not arrive");
                        recovered = true;
                        return true;
                    }
                    finally { token.Dispose(); }
                }
                catch (Exception ex)
                {
                    recoverError = ex.GetType().Name + ": " + ex.Message;
                    return false;
                }
            }, ResumeBoundMs);

            Metric("B", "ops", ops);
            Metric("B", "workerErrors", errors);
            Metric("B", "workerTimeouts", timeouts);
            Metric("B", "udfErrors", udfErrors);
            Metric("B", "udfTimeouts", udfTimeouts);
            Metric("B", "firstUdfError", firstUdfError ?? "-");
            Metric("B", "deliveryGapMs", Delivery.MaxGapMs);
            Metric("B", "deliveryResumeMs", resume);
            Metric("B", "watchdogMaxMs", Wake.MaxIterMs);
            Metric("B", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("B", "recovered", recovered);
            Metric("B", "recoverError", recoverError ?? "-");

            Check("B.dedicated-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            Check("B.entrypoints-not-stuck", Wake.MaxIterMs < StuckBoundMs, "watchdog max iteration=" + Fmt(Wake.MaxIterMs) + "ms");
            Check("B.delivery-stays-live", Delivery.MaxGapMs < StuckBoundMs, "delivery max gap=" + Fmt(Delivery.MaxGapMs) + "ms");
            // Under the intentional flood + starvation a synchronous UDF call may
            // legitimately hit the client timeout: that is transient, the same
            // class as the worker timeouts (recovery is asserted separately).
            // Only non-timeout errors fail the run, and the transient timeouts
            // stay bounded so a pathological collapse still fails.
            Check("B.no-unexpected-errors", errors - timeouts == 0 && udfErrors == 0 && udfTimeouts <= 1000,
                "errors=" + errors + " timeouts=" + timeouts + " udfErrors=" + udfErrors + " udfTimeouts=" + udfTimeouts
                + " firstUdfError=" + (firstUdfError ?? "-"));
            Check("B.subscribe-recovers-after-flood", recovered, "recovery subscribe succeeded once the flood stopped" + (recoverError == null ? "" : " lastError=" + recoverError));

            foreach (var t in leftovers) { try { t.Dispose(); } catch { } }
            foreach (var ch in latestChannels) { try { RedisUDF.RedisUDFChannelUnsubscribe(ch, Host); } catch { } }
            RedisUDF.SyncWriteOverrideForTests = prevSyncWrite;
            Metric("B", "durationSec", MsSince(start) / 1000.0);
            if (Quick)
                Metric("B", "quick-summary",
                    "workerErrors=" + errors + " workerTimeouts=" + timeouts + " udfErrors=" + udfErrors
                    + " udfTimeouts=" + udfTimeouts + " deliveryGapMs=" + Fmt(Delivery.MaxGapMs)
                    + " procHeartbeatMs=" + Fmt(Proc.MaxGapMs) + " watchdogMs=" + Fmt(Wake.MaxIterMs));
            NoteWorst("deliveryGapMs", Delivery.MaxGapMs);
            NoteWorst("watchdogMs", Wake.MaxIterMs);
            NoteWorst("deliveryResumeMs", resume);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            DumpLatencies("B");
        }

        // ------------------------------------- C: kills + server restart

        private static void PhaseC()
        {
            long start = PhaseBanner("C", "CLIENT KILL storm" + (ContainerManaged && !SkipRestart ? " + server restart mid-stream" : " (restart skipped)"));
            var wd = ConnectionMultiplexer.Connect(DirectConfig);
            var db = wd.GetDatabase();

            if (IsLocalEndpoint())
            {
                long prevDeliveries = Delivery.Count;
                long clientsKilled = 0;
                int killErrors = 0;
                int rounds = Quick ? 4 : 12;
                for (int i = 0; i < rounds; i++)
                {
                    try
                    {
                        RedisResult r = db.Execute("CLIENT", "KILL", "TYPE", "pubsub");
                        long killed;
                        if (long.TryParse(r.ToString(), NumberStyles.Integer, Inv, out killed))
                            clientsKilled += killed;
                    }
                    catch { killErrors++; }
                    Thread.Sleep(300);
                }

                Metric("C", "killRounds", rounds);
                Metric("C", "clientsKilled", clientsKilled);
                Metric("C", "killErrors", killErrors);
                double killResume = CheckDeliveryRecovered(prevDeliveries,
                    "killDeliveryResumeMs", "killReadersAfterMs",
                    "C.kill-storm-delivery-resumes", "C.kill-storm-resubscribed");
                Metric("C", "deliveryGapMs", Delivery.MaxGapMs);
                Metric("C", "procHeartbeatMs", Proc.MaxGapMs);
                Check("C.kill-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
                NoteWorst("deliveryGapMs", Delivery.MaxGapMs);
                NoteWorst("deliveryResumeMs", killResume);
            }
            else
            {
                Console.WriteLine("[SKIP] phase=C sub=kill-storm reason=non-local host " + Host);
            }

            // --- server restart ---
            if (ContainerManaged && !SkipRestart)
            {
                long prevDeliveries2 = Delivery.Count;
                double restartWallMs = RestartContainer();
                long serverUpMs = WaitUntil(() => { try { db.Ping(); return true; } catch { return false; } }, ResumeBoundMs);

                Metric("C", "restartWallMs", restartWallMs);
                Metric("C", "serverUpAfterMs", serverUpMs);
                Check("C.restart-server-up", serverUpMs >= 0, "ping after restart ms=" + Fmt(serverUpMs));
                double restartResume = CheckDeliveryRecovered(prevDeliveries2,
                    "restartDeliveryResumeMs", "restartReaders",
                    "C.restart-delivery-resumes", "C.restart-resubscribed");
                Metric("C", "deliveryGapKillAndRestartMs", Delivery.MaxGapMs);
                Metric("C", "publishErrorsTotal", Delivery.PublishErrors);
                Check("C.restart-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
                Check("C.delivery-gap-across-kills-and-restart", Delivery.MaxGapMs < ResumeBoundMs,
                    "max inter-arrival gap=" + Fmt(Delivery.MaxGapMs) + "ms across the kills + restart");
                NoteWorst("deliveryResumeMs", restartResume);
                NoteWorst("deliveryGapMs", Delivery.MaxGapMs);
            }
            else
            {
                string reason = SkipRestart
                    ? "--skip-restart"
                    : (ContainerSkipReason != null ? ContainerSkipReason
                        : (ContainerName == null ? "run with --container to manage a disposable server" : "container not managed"));
                Console.WriteLine("[SKIP] phase=C sub=server-restart reason=" + reason);
            }

            wd.Dispose();
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("C", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("C");
        }

        // ------------------------------------------- D1: queue burst FIFO

        private static void PhaseD1()
        {
            long start = PhaseBanner("D1", "async queue burst: same-host FIFO + no overlap + idle-queue removal");
            string[] hosts = { Host + "|d1a", Host + "|d1b" };
            const int producersPerHost = 2;
            int itemsPerProducer = Quick ? 60 : 150;
            int total = hosts.Length * producersPerHost * itemsPerProducer;
            var done = new CountdownEvent(total);
            var active = new int[hosts.Length];
            var maxActive = new int[hosts.Length];
            var lastSeq = new Dictionary<string, int>();
            object seqGate = new object();
            int violations = 0;
            var producers = new List<Thread>();

            for (int h = 0; h < hosts.Length; h++)
            {
                int hi = h;
                string host = hosts[h];
                for (int p = 0; p < producersPerHost; p++)
                {
                    int pi = p;
                    var t = new Thread(() =>
                    {
                        for (int i = 0; i < itemsPerProducer; i++)
                        {
                            int item = i;
                            RedisUdfAsync.Enqueue(host, () =>
                            {
                                int a = Interlocked.Increment(ref active[hi]);
                                UpdateMax(ref maxActive[hi], a);
                                lock (seqGate)
                                {
                                    string key = hi + "|" + pi;
                                    int prev;
                                    if (lastSeq.TryGetValue(key, out prev) && prev + 1 != item) violations++;
                                    lastSeq[key] = item;
                                }
                                Thread.SpinWait(2000);
                                Interlocked.Decrement(ref active[hi]);
                                done.Signal();
                                return (object)item;
                            });
                        }
                    });
                    t.IsBackground = true;
                    t.Name = "D1-producer-" + h + "-" + p;
                    producers.Add(t);
                }
            }

            foreach (var t in producers) t.Start();
            bool drainedOk = done.Wait((int)QueueDrainBoundMs);
            foreach (var t in producers) t.Join(10000);
            double drainMs = MsSince(start);
            bool overlap = maxActive.Any(m => m > 1);
            bool queuesGone = WaitUntil(() => !RedisUdfAsync.HasQueueForTests(hosts[0]) && !RedisUdfAsync.HasQueueForTests(hosts[1]), 5000) >= 0;

            Metric("D1", "items", total);
            Metric("D1", "drainMs", drainMs);
            Metric("D1", "fifoViolations", violations);
            Metric("D1", "maxConcurrentPerHost", string.Join(",", maxActive));
            Check("D1.all-executed", drainedOk, "completed " + (total - done.CurrentCount) + "/" + total);
            Check("D1.fifo-per-producer", violations == 0, "violations=" + violations);
            Check("D1.no-overlap-per-host", !overlap, "maxConcurrentPerHost=" + string.Join(",", maxActive));
            Check("D1.queues-removed", queuesGone, "idle queue entries removed");
            NoteWorst("queueDrainMs", drainMs);
            DumpLatencies("D1");
        }

        // ------------------------------ D2: observable burst + starvation

        private static void PhaseD2()
        {
            long start = PhaseBanner("D2", "observable burst under ThreadPool starvation");
            string host = Host + "|d2";
            int count = Quick ? 40 : 60;

            // Starve first, then subscribe: the queued work items can only run
            // when the pool is released. Unlike phase A, D2 never bails out on
            // a failed guard: the gate assertions below must still run.
            var release = StarvePool(out int minW, out int minIo, out int maxW, out int maxIo, out bool setMin, out bool setMax, 300);

            var observers = new RecordingObserver[count];
            for (int i = 0; i < count; i++)
            {
                int tag = i;
                observers[i] = new RecordingObserver();
                new RedisWriteObservable(host, () => (object)tag).Subscribe(observers[i]);
            }

            Thread.Sleep(Quick ? 800 : 1500);
            int completedWhileStarved = CountCompleted(observers);
            long starveTicks = Pool.Ticks;
            RestorePool(minW, minIo, maxW, maxIo, release);
            WaitUntil(() => Pool.Ticks > starveTicks, 3000);

            long drainT0 = Stopwatch.GetTimestamp();
            long drainedAt = WaitUntil(() => CountCompleted(observers) == count, QueueDrainBoundMs);
            double drainMs = drainedAt < 0 ? -1 : MsSince(drainT0);
            int deliveredWrong = 0;
            for (int i = 0; i < count; i++)
            {
                object[] values = observers[i].Values;
                if (observers[i].CompletedCount != 1 || values.Length != 1 || !(values[0] is int v) || v != i)
                    deliveredWrong++;
            }
            bool queueGone = WaitUntil(() => !RedisUdfAsync.HasQueueForTests(host), 5000) >= 0;

            Metric("D2", "items", count);
            Metric("D2", "completedWhileStarved", completedWhileStarved);
            Metric("D2", "drainMs", drainMs);
            Metric("D2", "wrongDeliveries", deliveredWrong);
            Metric("D2", "poolHeartbeatGapMs", Pool.MaxGapMs);
            Metric("D2", "procHeartbeatMs", Proc.MaxGapMs);
            Check("D2.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            Check("D2.work-stalled-during-starvation", completedWhileStarved == 0, "completed while starved=" + completedWhileStarved + "/" + count);
            Check("D2.drain-after-release", drainedAt >= 0, "completed=" + CountCompleted(observers) + "/" + count);
            Check("D2.exactly-once-each", deliveredWrong == 0, "wrong deliveries=" + deliveredWrong);
            Check("D2.queue-removed", queueGone, "idle queue entry removed");
            Check("D2.heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            NoteWorst("poolHeartbeatGapMs", Pool.MaxGapMs);
            NoteWorst("queueDrainMs", drainMs);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("D2", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("D2");
        }

        // --------------------- D3: duplicate subscribe + failing observer

        private static void PhaseD3()
        {
            PhaseBanner("D3", "observable duplicate-subscribe race + failing observer isolation");
            int races = Quick ? 8 : 20;
            int raceFailures = 0;
            for (int r = 0; r < races; r++)
            {
                string host = Host + "|d3r" + r;
                int runs = 0;
                int round = r;
                var observable = new RedisWriteObservable(host, () => { Interlocked.Increment(ref runs); return (object)round; });
                var obs = new RecordingObserver[4];
                var barrier = new Barrier(4);
                var threads = new Thread[4];
                for (int j = 0; j < 4; j++)
                {
                    int jj = j;
                    obs[jj] = new RecordingObserver();
                    threads[jj] = new Thread(() => { barrier.SignalAndWait(); observable.Subscribe(obs[jj]); });
                    threads[jj].IsBackground = true;
                    threads[jj].Start();
                }
                foreach (var t in threads) t.Join(10000);
                if (WaitUntil(() => CountCompleted(obs) == 4, 5000) < 0) raceFailures++;
                if (runs != 1) raceFailures++;
                foreach (var o in obs)
                    if (o.CompletedCount != 1 || o.Values.Length != 1) raceFailures++;
            }
            Check("D3.duplicate-subscribe-once", raceFailures == 0, "failures=" + raceFailures + "/" + (races * 6));

            string hostF = Host + "|d3fail";
            var throwing = new OnNextThrowingObserver();
            new RedisWriteObservable(hostF, () => (object)"first").Subscribe(throwing);
            bool throwingCompleted = throwing.WaitCompleted(5000);
            var after = new RecordingObserver();
            new RedisWriteObservable(hostF, () => (object)"second").Subscribe(after);
            bool afterCompleted = after.WaitCompleted(5000);
            Metric("D3", "postThrowValue", after.Values.Length > 0 ? after.Values[0] : "<none>");
            Check("D3.throwing-observer-completes", throwingCompleted && throwing.CompletedCount == 1, "completed=" + throwing.CompletedCount);
            Check("D3.queue-survives-throwing-observer", afterCompleted && after.CompletedCount == 1, "next item delivered=" + afterCompleted);
            DumpLatencies("D3");
        }

        // ----------------------------- E: dedup marker + listener rejoin

        private static void PhaseE()
        {
            PhaseBanner("E", "dedup marker: identical republish must reach a joiner");
            string ch = "liveness:e:" + Guid.NewGuid().ToString("N");
            int a = 0;
            int b = 0;
            IDisposable tokenA = S.Subscribe(Host, ch, false, msg => Interlocked.Increment(ref a), "E");
            var sub = C.GetSubscriber(Host, RedisPool.RtdSub);
            long readers1 = sub.Publish(new RedisChannel(ch, RedisChannel.PatternMode.Literal), "SAME-PAYLOAD", CommandFlags.None);
            long a1 = WaitUntil(() => Volatile.Read(ref a) >= 1, 5000);
            IDisposable tokenB = S.Subscribe(Host, ch, false, msg => Interlocked.Increment(ref b), "E");
            long readers2 = sub.Publish(new RedisChannel(ch, RedisChannel.PatternMode.Literal), "SAME-PAYLOAD", CommandFlags.None);
            long b1 = WaitUntil(() => Volatile.Read(ref b) >= 1, 5000);
            long a2 = WaitUntil(() => Volatile.Read(ref a) >= 2, 5000);

            Metric("E", "readersFirst", readers1);
            Metric("E", "readersSecond", readers2);
            Metric("E", "listenerA", a);
            Metric("E", "listenerB", b);
            Check("E.first-delivery", a1 >= 0, "listener A never got the first payload");
            Check("E.identical-republish-reaches-joiner", b1 >= 0 && a2 >= 0, "a=" + a + " b=" + b + " (identical payload must not be dedup-suppressed for the joiner)");
            tokenB.Dispose();
            tokenA.Dispose();
            DumpLatencies("E");
        }

        // ======================================================== matrix/soak

        /// <summary>Starts the dedicated heartbeat threads and the real managers
        /// (shared by the default phases and the matrix/soak modes).</summary>
        private static void StartProbes()
        {
            Proc = new ProcessHeartbeat("proc-heartbeat", 20);
            Proc.Start();
            Pool = new PoolHeartbeat();
            Pool.Start();
            C = new RedisConnectionManager();
            S = new RedisSubscriptionManager(C);
            Wake = new WatchdogProbe(C, S);
            Wake.Start();
        }

        private const int PauseHoldMs = 5000;   // "half-open, hold ~5s"
        private const int FlapCycles = 3;       // "3x flapping stop/start loop"
        private const int KillStormRounds = 3;  // "3x CLIENT KILL TYPE pubsub"

        private static bool DockerFaultsEnabled { get { return ContainerManaged; } }
        private static bool CommandFaultsEnabled { get { return IsLocalEndpoint(); } }

        private static string DockerFaultSkipReason
        {
            get
            {
                if (ContainerSkipReason != null) return ContainerSkipReason;
                return ContainerName == null
                    ? "run with --container <name> to manage a disposable server"
                    : "container not managed";
            }
        }

        private static string CommandFaultSkipReason
        {
            get { return "destructive command faults require a loopback host (" + Host + ")"; }
        }

        /// <summary>
        /// Scripted failure matrix: every fault runs under continuous traffic
        /// (publisher + delivery counter + dedicated process heartbeat); after
        /// each fault the run asserts a bounded process heartbeat, delivery
        /// resume, automatic re-subscription (PUBSUB NUMSUB) and monotonic
        /// counters, then finishes with resource-stability checks (zero runtime
        /// listeners, NUMSUB back to 0, no idle queues).
        /// </summary>
        private static void RunMatrix()
        {
            StartProbes();
            Warmup();
            TrackChannel(LiveChannel);
            long faultsBefore = Interlocked.Read(ref FaultsRun);
            long matrixStart = PhaseBanner("M", "scripted failure matrix container=" + (ContainerManaged ? ContainerName : "-"));

            if (DockerFaultsEnabled)
                RunFault("M1", "docker stop + docker start", LiveChannel, () => InjectStopStart("M1"));
            else
                SkipFault("M1", "docker stop + docker start", DockerFaultSkipReason);

            if (DockerFaultsEnabled)
                RunFault("M2", "docker restart", LiveChannel, () => InjectRestart("M2"));
            else
                SkipFault("M2", "docker restart", DockerFaultSkipReason);

            if (DockerFaultsEnabled)
                RunFault("M3", "docker pause + unpause (half-open, hold " + PauseHoldMs + "ms)", LiveChannel, () => InjectPauseUnpause("M3"));
            else
                SkipFault("M3", "docker pause + unpause", DockerFaultSkipReason);

            if (CommandFaultsEnabled)
                RunFault("M4", KillStormRounds + "x CLIENT KILL TYPE pubsub", LiveChannel, () => InjectKillStorm("M4"));
            else
                SkipFault("M4", "CLIENT KILL storm", CommandFaultSkipReason);

            if (CommandFaultsEnabled)
                RunFault("M5", "CLIENT PAUSE 3000 + UNPAUSE", LiveChannel, () => InjectClientPause("M5"));
            else
                SkipFault("M5", "CLIENT PAUSE 3000 + UNPAUSE", CommandFaultSkipReason);

            if (DockerFaultsEnabled)
                RunFault("M6", FlapCycles + "x flapping stop/start", LiveChannel, () => InjectFlapping("M6"));
            else
                SkipFault("M6", "flapping stop/start", DockerFaultSkipReason);

            RunFault("M7", "reload under traffic (Shutdown + ResetAfterAddInReload x2)", LiveChannel, () => InjectReload("M7"));

            if (DockerFaultsEnabled || CommandFaultsEnabled)
                RunFault("M8", "write path during fault (sync error + fire-and-forget + async backlog)", LiveChannel, () => InjectWritePath("M8"));
            else
                SkipFault("M8", "write path during fault", "needs a managed container or a loopback host");

            Console.WriteLine("[SUMMARY] matrixFaultsRan=" + (Interlocked.Read(ref FaultsRun) - faultsBefore)
                + " matrixSeconds=" + (MsSince(matrixStart) / 1000.0).ToString("0.0", Inv));
            FinishMode("M9", "matrix");
        }

        /// <summary>
        /// Soak mode: continuous mixed traffic (delivery publishes, UDF reads,
        /// writes, channel publishes, async queue items) with a random fault
        /// from the matrix every 20-40s. Every fault is asserted with the same
        /// heartbeat/resume/monotonic bounds as --matrix; the run ends with
        /// resource stability and GC statistics. Any failure exits 1.
        /// </summary>
        private static void RunSoak()
        {
            StartProbes();
            Warmup();
            TrackChannel(LiveChannel);
            PhaseBanner("SOAK", "soak mixed traffic minutes=" + SoakMinutes.ToString("0.##", Inv));
            string prevOverride = RedisUDF.SyncWriteOverrideForTests;
            RedisUDF.SyncWriteOverrideForTests = "fireforget";
            var choices = BuildFaultChoices();
            if (choices.Count == 0)
                Failures.Add("soak: no injectable fault available (needs --container or a loopback host)");

            long readErrors = 0, writeErrors = 0, publishErrors = 0, enqueueTimeouts = 0;
            // "Continuous" traffic is paced to an Excel-like rate (the reply-less
            // paths would otherwise run flat-out and allocate tens of GB, which
            // says nothing about fault survival and risks OOM on slower hosts).
            var read = new Worker("SOAK-read", 2, (id, n) =>
            {
                string r = RedisUDF.RedisUDFGet("liveness:soak:key", Host);
                if (r != null && r.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref readErrors);
                Thread.Sleep(2);
            });
            var write = new Worker("SOAK-write", 2, (id, n) =>
            {
                string r = Convert.ToString(RedisUDF.RedisUDFSet("liveness:soak:key", "soak-" + n, Host), Inv);
                if (r != null && r.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref writeErrors);
                Thread.Sleep(5);
            });
            var publish = new Worker("SOAK-publish", 1, (id, n) =>
            {
                string r = Convert.ToString(RedisUDF.RedisUDFChannelPublish("liveness:soak:ch", "soak-" + n, Host), Inv);
                if (r != null && r.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref publishErrors);
                Thread.Sleep(5);
            });
            string soakQueueHost = Host + "|soak";
            TrackQueueHost(soakQueueHost);
            var queue = new Worker("SOAK-enqueue", 1, (id, n) =>
            {
                Task<object> t = RedisUdfAsync.Enqueue(soakQueueHost, () => (object)RedisUDF.RedisUDFGet("liveness:soak:key", Host));
                if (!t.Wait(10000)) Interlocked.Increment(ref enqueueTimeouts);
            });
            var workers = new List<Worker> { read, write, publish, queue };
            foreach (var w in workers) w.Start();

            long soakStart = Stopwatch.GetTimestamp();
            double durationMs = SoakMinutes * 60000.0;
            int faults = 0;
            string lastFaultKey = null;
            while (MsSince(soakStart) < durationMs)
            {
                double remaining = durationMs - MsSince(soakStart);
                if (remaining <= 0) break;
                int delayMs = Rng.Next(20000, 40001); // "random fault every 20-40s"
                Thread.Sleep((int)Math.Min(delayMs, remaining));
                if (MsSince(soakStart) >= durationMs) break;
                FaultChoice choice = PickSoakFault(choices, lastFaultKey);
                if (choice == null) break;
                lastFaultKey = choice.Key;
                string id = "SOAK-F" + (++faults);
                RunFault(id, choice.Desc, LiveChannel, () => choice.Inject(id));
            }

            foreach (var w in workers) w.StopAndJoin();
            Metric("SOAK", "faults", faults);
            Metric("SOAK", "elapsedSec", MsSince(soakStart) / 1000.0);
            Metric("SOAK", "readOps", read.Ops);
            Metric("SOAK", "writeOps", write.Ops);
            Metric("SOAK", "publishOps", publish.Ops);
            Metric("SOAK", "enqueueOps", queue.Ops);
            Metric("SOAK", "readErrorResults", readErrors);
            Metric("SOAK", "writeErrorResults", writeErrors);
            Metric("SOAK", "publishErrorResults", publishErrors);
            Metric("SOAK", "enqueueTimeouts", enqueueTimeouts);
            Check("SOAK.traffic-ran", read.Ops + write.Ops + publish.Ops + queue.Ops > 0,
                "read=" + read.Ops + " write=" + write.Ops + " publish=" + publish.Ops + " enqueue=" + queue.Ops);

            // Post-soak recovery: a sync write + read round-trip and a fresh
            // runtime subscription end-to-end once the traffic stopped.
            RedisUDF.SyncWriteOverrideForTests = "sync";
            string finalKey = "liveness:soak:final:" + Guid.NewGuid().ToString("N");
            long finalSet = WaitUntil(() => string.Equals(
                Convert.ToString(RedisUDF.RedisUDFSet(finalKey, "final-ok", Host), Inv), "OK", StringComparison.Ordinal), 15000);
            string finalGet = RedisUDF.RedisUDFGet(finalKey, Host);
            RedisUDF.SyncWriteOverrideForTests = prevOverride;
            Check("SOAK.final-write", finalSet >= 0, "sync RedisUDFSet returned OK ms=" + Fmt(finalSet));
            Check("SOAK.final-read", finalGet == "final-ok", "RedisUDFGet=" + Shorten(finalGet ?? "", 160));

            string chFresh = "liveness:soak:final:" + Guid.NewGuid().ToString("N");
            TrackChannel(chFresh);
            int fresh = 0;
            IDisposable freshToken = RedisRuntime.Subscriptions.Subscribe(Host, chFresh, false,
                _ => Interlocked.Increment(ref fresh), "SOAK");
            TrackToken(freshToken);
            long freshOk = WaitUntil(() =>
            {
                if (Volatile.Read(ref fresh) > 0) return true;
                PublishDirect(chFresh, "final");
                return Volatile.Read(ref fresh) > 0;
            }, ResumeBoundMs);
            Check("SOAK.final-subscription", freshOk >= 0, "fresh runtime subscription delivered within " + Fmt(ResumeBoundMs) + "ms");

            Console.WriteLine("[SUMMARY] soakFaults=" + faults + " soakMinutes=" + SoakMinutes.ToString("0.##", Inv));
            FinishMode("SOAK-end", "soak");
        }

        /// <summary>
        /// Runs one fault: continuous traffic and heartbeats were started by the
        /// mode; this resets the per-fault worst trackers, injects, then asserts
        /// heartbeat bound, delivery resume, re-subscription and monotonic
        /// counters. A fault never aborts the run: failures are collected and
        /// exit 1 at the end.
        /// </summary>
        private static void RunFault(string id, string desc, string channel, Action inject)
        {
            Interlocked.Increment(ref FaultsRun);
            long faultStart = PhaseBanner(id, desc);
            int listenersBefore = S.ListenerCount;
            int channelsBefore = S.ChannelCount;
            long joinedBefore = Interlocked.Read(ref JoinedEvents);
            long deliveriesBefore = Delivery.Count;
            long publishedBefore = Delivery.Published;
            Interlocked.Exchange(ref FaultOutageCount, deliveriesBefore);

            try { inject(); }
            catch (Exception ex)
            {
                Failures.Add(id + " :: fault injection threw " + ex.GetType().Name + ": " + ex.Message);
            }
            double injectMs = MsSince(faultStart);

            // A delivery strictly after the outage mark proves the fault window
            // really ended (never just pre-outage in-flight messages). The
            // injectors mark the outage once it is fully in effect.
            long resumePrev = Interlocked.Read(ref FaultOutageCount);
            if (resumePrev < deliveriesBefore) resumePrev = deliveriesBefore;
            double resume = Delivery.ResumeAfter(resumePrev, ResumeBoundMs);
            long readersWait = WaitForNumSub(channel, 1, ResumeBoundMs, out long readers);
            double heartbeat = Proc.MaxGapMs;
            double gap = Delivery.MaxGapMs;
            double watchdog = Wake.MaxIterMs;
            bool monotonic = Delivery.Count >= deliveriesBefore
                && Delivery.Published >= publishedBefore
                && S.ListenerCount >= listenersBefore
                && S.ChannelCount >= channelsBefore
                && Interlocked.Read(ref JoinedEvents) >= joinedBefore;

            Metric(id, "injectWallMs", injectMs);
            Metric(id, "heartbeatMs", heartbeat);
            Metric(id, "watchdogMaxMs", watchdog);
            Metric(id, "deliveryGapMs", gap);
            Metric(id, "deliveryResumeMs", resume);
            Metric(id, "readersAfter", readers);
            Metric(id, "readersWaitMs", readersWait);
            Metric(id, "poolHeartbeatGapMs", Pool.MaxGapMs);
            Metric(id, "deliveries", Delivery.Count);
            Metric(id, "publishes", Delivery.Published);
            Metric(id, "publishErrors", Delivery.PublishErrors);

            Check(id + ".heartbeat", heartbeat < StuckBoundMs,
                "proc heartbeat max gap=" + Fmt(heartbeat) + "ms (bound " + StuckBoundMs + "ms)");
            Check(id + ".entrypoints-not-stuck", watchdog < StuckBoundMs,
                "watchdog max iteration=" + Fmt(watchdog) + "ms");
            Check(id + ".delivery-resumes", resume >= 0 && resume <= ResumeBoundMs,
                "delivery resume after fault=" + Fmt(resume) + "ms (bound " + ResumeBoundMs + "ms)");
            Check(id + ".resubscribed", readers >= 1,
                "PUBSUB NUMSUB " + channel + "=" + readers + " (expected >= 1)");
            Check(id + ".monotonic", monotonic,
                "deliveries=" + Delivery.Count + " (>= " + deliveriesBefore + ") listeners=" + S.ListenerCount
                + " (>= " + listenersBefore + ") channels=" + S.ChannelCount + " (>= " + channelsBefore + ")");

            NoteWorst("procHeartbeatMs", heartbeat);
            NoteWorst("watchdogMs", watchdog);
            NoteWorst("deliveryGapMs", gap);
            NoteWorst("deliveryResumeMs", resume);
            NoteWorst("poolHeartbeatGapMs", Pool.MaxGapMs);
        }

        private static void SkipFault(string id, string desc, string reason)
        {
            Console.WriteLine("[SKIP] phase=" + id + " desc=" + desc + " reason=" + reason);
        }

        // ------------------------------------------------- fault injectors

        /// <summary>docker CLI step that must succeed; a failure is recorded as an
        /// assertion failure of the current fault instead of aborting the mode.</summary>
        private static bool DockerStep(string checkName, string arguments, int timeoutMs)
        {
            DockerResult r = RunDocker(arguments, timeoutMs);
            bool ok = r.ExitCode == 0;
            Check(checkName, ok, "docker " + arguments
                + (ok ? "" : " exit=" + r.ExitCode + " " + Shorten((r.StdErr ?? "").Trim(), 200)));
            return ok;
        }

        private static bool PingServer()
        {
            try { AdminDb().Ping(); return true; }
            catch { return false; }
        }

        private static void WaitServerDown(double timeoutMs)
        {
            WaitUntil(() => !PingServer(), timeoutMs);
        }

        private static void WaitServerUp(double timeoutMs)
        {
            WaitUntil(() => PingServer(), timeoutMs);
        }

        /// <summary>Records the delivery count at the moment the current fault's
        /// outage is fully in effect; the runner only accepts a delivery after
        /// this mark as proof of resume (pre-outage in-flight messages must not
        /// satisfy the check).</summary>
        private static void MarkOutage()
        {
            Interlocked.Exchange(ref FaultOutageCount, Delivery.Count);
        }

        // Fault 1: docker stop (dead server, full reconnect) then start.
        private static void InjectStopStart(string id)
        {
            if (!DockerStep(id + ".docker-stop", "stop " + ContainerName, 120000)) return;
            WaitServerDown(10000);
            MarkOutage();
            Thread.Sleep(1000); // keep the outage visible to the delivery probe
            if (!DockerStep(id + ".docker-start", "start " + ContainerName, 120000)) return;
            WaitServerUp(ResumeBoundMs);
        }

        // Fault 2: docker restart (stop + start in one command).
        private static void InjectRestart(string id)
        {
            MarkOutage();
            DockerStep(id + ".docker-restart", "restart " + ContainerName, 120000);
            WaitServerUp(ResumeBoundMs);
        }

        // Fault 3: docker pause/unpause: half-open server (connections stay
        // established, commands hang) held for ~5s.
        private static void InjectPauseUnpause(string id)
        {
            if (!DockerStep(id + ".docker-pause", "pause " + ContainerName, 60000)) return;
            MarkOutage();
            Thread.Sleep(PauseHoldMs);
            DockerStep(id + ".docker-unpause", "unpause " + ContainerName, 60000);
            WaitServerUp(ResumeBoundMs);
        }

        // Fault 4: 3x CLIENT KILL TYPE pubsub (kills the subscriber connections).
        private static void InjectKillStorm(string id)
        {
            long killed = 0;
            int errors = 0;
            for (int i = 0; i < KillStormRounds; i++)
            {
                try
                {
                    RedisResult r = AdminDb().Execute("CLIENT", "KILL", "TYPE", "pubsub");
                    long n;
                    if (long.TryParse(Convert.ToString(r, Inv), NumberStyles.Integer, Inv, out n)) killed += n;
                }
                catch { errors++; }
                Thread.Sleep(500);
            }
            MarkOutage();
            Metric(id, "clientsKilled", killed);
            Metric(id, "killErrors", errors);
        }

        // Fault 5: CLIENT PAUSE 3000 (all commands delayed) + CLIENT UNPAUSE.
        // The unpause is fire-and-forget: a fully paused server only processes
        // it when the pause window expires anyway.
        private static void InjectClientPause(string id)
        {
            bool issued = true;
            try { AdminDb().Execute("CLIENT", "PAUSE", "3000"); }
            catch (Exception ex) { issued = false; Failures.Add(id + " :: CLIENT PAUSE failed: " + ex.Message); }
            Check(id + ".client-pause-issued", issued, "CLIENT PAUSE 3000");
            MarkOutage();
            Thread.Sleep(1000);
            try { AdminDb().Execute("CLIENT", new object[] { "UNPAUSE" }, CommandFlags.FireAndForget); } catch { }
            WaitServerUp(ResumeBoundMs);
        }

        // Fault 6: 3x flapping stop/start: repeated short outages under traffic.
        private static void InjectFlapping(string id)
        {
            for (int i = 0; i < FlapCycles; i++)
            {
                if (!DockerStep(id + ".flap-stop-" + (i + 1), "stop " + ContainerName, 120000)) return;
                WaitServerDown(10000);
                MarkOutage();
                Thread.Sleep(1000);
                if (!DockerStep(id + ".flap-start-" + (i + 1), "start " + ContainerName, 120000)) return;
                WaitServerUp(ResumeBoundMs);
                Thread.Sleep(500);
            }
        }

        // Fault 7: same-process add-in reload under traffic: RedisRuntime is
        // shut down (the add-in unload), then ResetAfterAddInReload is called on
        // both RedisRuntime and RedisUDF exactly like AddIn.AutoOpen does on a
        // reload; new subscriptions and UDF entry points must work afterwards.
        private static void InjectReload(string id)
        {
            string suffix = Guid.NewGuid().ToString("N");
            string chOld = "liveness:matrix:reload:old:" + suffix;
            string chNew = "liveness:matrix:reload:new:" + suffix;
            string chUdf = "liveness:matrix:reload:udf:" + suffix;
            TrackChannel(chOld);
            TrackChannel(chNew);
            TrackChannel(chUdf);

            int oldGot = 0;
            IDisposable oldToken = RedisRuntime.Subscriptions.Subscribe(Host, chOld, false,
                _ => Interlocked.Increment(ref oldGot), "MATRIX");
            TrackToken(oldToken);
            PublishDirect(chOld, "pre-reload");
            long pre = WaitUntil(() => Volatile.Read(ref oldGot) >= 1, 10000);
            Check(id + ".pre-reload-subscription-live", pre >= 0, "pre-reload subscription delivered=" + (pre >= 0));

            MarkOutage();
            RedisRuntime.Shutdown();
            RedisRuntime.ResetAfterAddInReload();
            RedisUDF.ResetAfterAddInReload();
            try { oldToken.Dispose(); } catch { }

            int newGot = 0;
            IDisposable newToken = RedisRuntime.Subscriptions.Subscribe(Host, chNew, false,
                _ => Interlocked.Increment(ref newGot), "MATRIX");
            TrackToken(newToken);
            long newOk = WaitUntil(() =>
            {
                if (Volatile.Read(ref newGot) > 0) return true;
                PublishDirect(chNew, "post-reload");
                return Volatile.Read(ref newGot) > 0;
            }, ResumeBoundMs);
            Check(id + ".new-subscription-after-reload", newOk >= 0,
                "fresh runtime subscription delivered after reload=" + (newOk >= 0) + " (wait=" + Fmt(newOk) + "ms)");

            string reloadKey = "liveness:matrix:reload:key:" + suffix;
            string setText = Convert.ToString(RedisUDF.RedisUDFSet(reloadKey, "reload-ok", Host), Inv) ?? "";
            Check(id + ".udf-write-after-reload", !setText.StartsWith("Error:", StringComparison.Ordinal),
                "RedisUDFSet=" + Shorten(setText, 160));
            string getText = RedisUDF.RedisUDFGet(reloadKey, Host);
            Check(id + ".udf-read-after-reload", getText == "reload-ok", "RedisUDFGet=" + Shorten(getText ?? "", 160));

            RedisUDF.RedisUDFChannelLatest(chUdf, Host);
            RedisUDF.RedisUDFChannelPublish(chUdf, "udf-payload", Host);
            long udfOk = WaitUntil(() => RedisUDF.RedisUDFChannelLatest(chUdf, Host) == "udf-payload", 10000);
            Check(id + ".udf-channel-after-reload", udfOk >= 0, "RedisUDFChannelLatest saw the publish=" + (udfOk >= 0));
            try { RedisUDF.RedisUDFChannelUnsubscribe(chUdf, Host); } catch { }
            try { newToken.Dispose(); } catch { }
        }

        // Fault 8: writes during an outage: the sync core write must return
        // "Error:" within the configured timeout bound and succeed after
        // recovery; the fire-and-forget write must never throw into the caller;
        // the async queue backlog must complete every task.
        private static void InjectWritePath(string id)
        {
            string prevOverride = RedisUDF.SyncWriteOverrideForTests;
            bool dockerOutage = DockerFaultsEnabled;
            string key = "liveness:matrix:wp:" + Guid.NewGuid().ToString("N").Substring(0, 12);
            string qhost = Host + "|matrix-wp";
            TrackQueueHost(qhost);
            int backlog = 15;
            var tasks = new Task<object>[backlog];
            try
            {
                if (dockerOutage)
                {
                    if (!DockerStep(id + ".docker-stop", "stop " + ContainerName, 120000)) return;
                    WaitServerDown(10000);
                    MarkOutage();
                }
                else
                {
                    try { AdminDb().Execute("CLIENT", "PAUSE", "3000"); }
                    catch (Exception ex) { Failures.Add(id + " :: CLIENT PAUSE failed: " + ex.Message); }
                    MarkOutage();
                    Thread.Sleep(250);
                }

                RedisUDF.SyncWriteOverrideForTests = "sync";
                long t0 = Stopwatch.GetTimestamp();
                string syncText = Convert.ToString(RedisUDF.RedisUDFSet(key, "during-fault", Host), Inv) ?? "";
                double syncMs = MsSince(t0);
                Check(id + ".sync-write-error-during-fault", syncText.StartsWith("Error:", StringComparison.Ordinal),
                    "RedisUDFSet(sync)=" + Shorten(syncText, 160));
                Check(id + ".sync-write-bounded", syncMs <= WriteOutageBoundMs(),
                    "ms=" + Fmt(syncMs) + " bound=" + Fmt(WriteOutageBoundMs()) + "ms");
                Metric(id, "syncWriteDuringFaultMs", syncMs);

                RedisUDF.SyncWriteOverrideForTests = "fireforget";
                bool ffThrew = false;
                string ffText = null;
                try { ffText = Convert.ToString(RedisUDF.RedisUDFSet(key, "ff-during-fault", Host), Inv); }
                catch (Exception ex) { ffThrew = true; ffText = ex.GetType().Name + ": " + ex.Message; }
                Check(id + ".fireforget-never-throws", !ffThrew, "result=" + Shorten(ffText ?? "", 160));

                RedisUDF.SyncWriteOverrideForTests = "sync";
                for (int i = 0; i < backlog; i++)
                {
                    int n = i;
                    tasks[n] = RedisUdfAsync.Enqueue(qhost, () => RedisUDF.RedisUDFSet(key + ":" + n, "backlog-" + n, Host));
                }
            }
            finally
            {
                if (dockerOutage)
                {
                    DockerStep(id + ".docker-start", "start " + ContainerName, 120000);
                    WaitServerUp(ResumeBoundMs);
                }
                else
                {
                    try { AdminDb().Execute("CLIENT", new object[] { "UNPAUSE" }, CommandFlags.FireAndForget); } catch { }
                    WaitServerUp(ResumeBoundMs);
                }
                RedisUDF.SyncWriteOverrideForTests = prevOverride;
            }

            int enqueued = 0;
            foreach (var t in tasks) if (t != null) enqueued++;
            long drainStart = Stopwatch.GetTimestamp();
            bool drained = enqueued == backlog && Task.WaitAll(tasks, (int)QueueDrainBoundMs);
            double drainMs = MsSince(drainStart);
            int completed = 0;
            foreach (var t in tasks) if (t != null && t.IsCompleted) completed++;
            Metric(id, "queueDrainMs", drainMs);
            Metric(id, "queueCompleted", completed);
            Check(id + ".enqueue-backlog-completes", drained && completed == backlog,
                "completed=" + completed + "/" + backlog + " drained=" + drained + " drainMs=" + Fmt(drainMs));
            NoteWorst("queueDrainMs", drainMs);

            RedisUDF.SyncWriteOverrideForTests = "sync";
            try
            {
                string probeKey = key + ":after";
                long ok = WaitUntil(() => string.Equals(
                    Convert.ToString(RedisUDF.RedisUDFSet(probeKey, "after-recovery", Host), Inv), "OK", StringComparison.Ordinal), 15000);
                Check(id + ".sync-write-succeeds-after-recovery", ok >= 0, "RedisUDFSet(sync)=OK ms=" + Fmt(ok));
            }
            finally
            {
                RedisUDF.SyncWriteOverrideForTests = prevOverride;
            }
        }

        /// <summary>Time budget for a sync write during an outage: one connect
        /// attempt (ConnectTimeout) plus the command sync timeout (both are the
        /// configured pool timeout) plus slack.</summary>
        private static double WriteOutageBoundMs()
        {
            return RedisUDF.ResponseTimeoutMs() + 4000;
        }

        // ------------------------------------------------- soak fault set

        private sealed class FaultChoice
        {
            public string Key;
            public string Desc;
            public Action<string> Inject;
        }

        private static List<FaultChoice> BuildFaultChoices()
        {
            var choices = new List<FaultChoice>();
            if (DockerFaultsEnabled)
            {
                choices.Add(new FaultChoice { Key = "stop-start", Desc = "docker stop + docker start", Inject = InjectStopStart });
                choices.Add(new FaultChoice { Key = "restart", Desc = "docker restart", Inject = InjectRestart });
                choices.Add(new FaultChoice { Key = "pause", Desc = "docker pause + unpause (half-open, hold " + PauseHoldMs + "ms)", Inject = InjectPauseUnpause });
                choices.Add(new FaultChoice { Key = "flap", Desc = FlapCycles + "x flapping stop/start", Inject = InjectFlapping });
            }
            if (CommandFaultsEnabled)
            {
                choices.Add(new FaultChoice { Key = "kill", Desc = KillStormRounds + "x CLIENT KILL TYPE pubsub", Inject = InjectKillStorm });
                choices.Add(new FaultChoice { Key = "client-pause", Desc = "CLIENT PAUSE 3000 + UNPAUSE", Inject = InjectClientPause });
            }
            choices.Add(new FaultChoice { Key = "reload", Desc = "reload under traffic (Shutdown + ResetAfterAddInReload x2)", Inject = InjectReload });
            if (DockerFaultsEnabled || CommandFaultsEnabled)
            {
                choices.Add(new FaultChoice
                {
                    Key = "write-path",
                    Desc = "write path during fault (sync error + fire-and-forget + async backlog)",
                    Inject = InjectWritePath
                });
            }
            return choices;
        }

        private static FaultChoice PickSoakFault(List<FaultChoice> choices, string lastKey)
        {
            if (choices.Count == 0) return null;
            for (int attempt = 0; attempt < 10; attempt++)
            {
                FaultChoice choice = choices[Rng.Next(choices.Count)];
                if (choices.Count == 1 || choice.Key != lastKey) return choice;
            }
            return choices[0];
        }

        // ------------------------------------------------- end-of-mode checks

        /// <summary>
        /// Resource stability: all subscriber tokens are disposed by the caller
        /// first; then zero runtime listeners, PUBSUB NUMSUB back to 0 for every
        /// tracked channel and no idle async queues are required before the
        /// managers are shut down.
        /// </summary>
        private static void FinishMode(string phase, string label)
        {
            Console.WriteLine();
            Console.WriteLine("[PHASE] id=" + phase + " desc=resource stability (listeners, NUMSUB, queues, shutdown)");
            try { Delivery.Dispose(); } catch { }
            DisposeTrackedTokens();

            long noLeaks = WaitUntil(() => S.ChannelCount == 0 && S.ListenerCount == 0, 10000);
            Metric(phase, "channelCount", S.ChannelCount);
            Metric(phase, "listenerCount", S.ListenerCount);
            Check(phase + ".no-channel-leaks", noLeaks >= 0,
                "ChannelCount=" + S.ChannelCount + " ListenerCount=" + S.ListenerCount);

            ResourceStabilityChecks(phase);
            DumpGcSummary();

            Console.WriteLine("[SUMMARY] " + label + "Deliveries=" + Delivery.Count
                + " " + label + "Publishes=" + Delivery.Published
                + " publishErrors=" + Delivery.PublishErrors);
            Console.WriteLine("[SUMMARY] listenerJoinedEvents=" + Interlocked.Read(ref JoinedEvents));

            Wake.Stop();
            Pool.Stop();
            Proc.Stop();
            C.Shutdown();
            Metric(phase, "liveConnectionsAfterShutdown", C.LiveConnectionCount());
            Check(phase + ".shutdown-pools", C.LiveConnectionCount() == 0,
                "LiveConnectionCount after Shutdown=" + C.LiveConnectionCount());
            try { RedisRuntime.Shutdown(); } catch { }
            DisposeAdmin();
            DumpLatencies(phase);
        }

        private static void ResourceStabilityChecks(string phase)
        {
            // A throwing RedisRuntime is a FAILURE, not a clean state: returning
            // true here would let a genuinely broken runtime (e.g. "shutting
            // down" InvalidOperationException) pass the leak check silently.
            // Capture and surface the exception text so the cause is diagnosable.
            Exception runtimeError = null;
            long runtimeZero = WaitUntil(() =>
            {
                try
                {
                    runtimeError = null;
                    return RedisRuntime.Subscriptions.ListenerCount == 0;
                }
                catch (Exception ex)
                {
                    runtimeError = ex;
                    return false;
                }
            }, 10000);
            int runtimeListeners = -1;
            string runtimeReadError = null;
            try { runtimeListeners = RedisRuntime.Subscriptions.ListenerCount; }
            catch (Exception ex) { runtimeReadError = ex.GetType().Name + ": " + ex.Message; }
            string runtimeDetail = runtimeError != null
                ? " runtime threw " + runtimeError.GetType().Name + ": " + runtimeError.Message
                : (runtimeReadError != null ? " runtime threw " + runtimeReadError : "");
            Check(phase + ".runtime-listeners-zero", runtimeZero >= 0 && runtimeError == null,
                "RedisRuntime.Subscriptions.ListenerCount=" + runtimeListeners
                + " (runtimeAccess=" + (runtimeError == null && runtimeReadError == null ? "ok" : "threw") + ")" + runtimeDetail);

            foreach (string channel in TrackedChannels.ToArray())
            {
                long zero = WaitForNumSub(channel, 0, 10000, out long last);
                Check(phase + ".numsub-zero[" + channel + "]", zero >= 0,
                    "PUBSUB NUMSUB " + channel + "=" + last + " (expected 0)");
            }

            foreach (string queueHost in TrackedQueueHosts.ToArray())
            {
                long gone = WaitUntil(() => !RedisUdfAsync.HasQueueForTests(queueHost), 10000);
                Check(phase + ".queue-idle[" + queueHost + "]", gone >= 0,
                    "HasQueueForTests(" + queueHost + ")=" + RedisUdfAsync.HasQueueForTests(queueHost));
            }
        }

        /// <summary>GC/resource statistics for the soak/end of the run.</summary>
        private static void DumpGcSummary()
        {
            // Settle the heap first so the live/survived numbers are meaningful.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long allocated = 0, survived = 0;
            try
            {
                allocated = AppDomain.CurrentDomain.MonitoringTotalAllocatedMemorySize;
                survived = AppDomain.CurrentDomain.MonitoringSurvivedMemorySize;
            }
            catch { }
            long workingSet = 0, peakWorkingSet = 0;
            try
            {
                using (var p = Process.GetCurrentProcess())
                {
                    workingSet = p.WorkingSet64;
                    peakWorkingSet = p.PeakWorkingSet64;
                }
            }
            catch { }
            Console.WriteLine("[SUMMARY] gcGen0=" + GC.CollectionCount(0)
                + " gcGen1=" + GC.CollectionCount(1)
                + " gcGen2=" + GC.CollectionCount(2));
            Console.WriteLine("[SUMMARY] allocatedMB=" + (allocated / 1048576.0).ToString("0.0", Inv)
                + " survivedMB=" + (survived / 1048576.0).ToString("0.0", Inv)
                + " liveHeapMB=" + (GC.GetTotalMemory(false) / 1048576.0).ToString("0.0", Inv)
                + " workingSetMB=" + (workingSet / 1048576.0).ToString("0.0", Inv)
                + " peakWorkingSetMB=" + (peakWorkingSet / 1048576.0).ToString("0.0", Inv));
        }

        private static void TrackChannel(string channel)
        {
            if (!TrackedChannels.Contains(channel)) TrackedChannels.Add(channel);
        }

        private static void TrackToken(IDisposable token)
        {
            if (token != null) TrackedTokens.Add(token);
        }

        private static void TrackQueueHost(string host)
        {
            if (!TrackedQueueHosts.Contains(host)) TrackedQueueHosts.Add(host);
        }

        private static void DisposeTrackedTokens()
        {
            foreach (var token in TrackedTokens)
            {
                try { token.Dispose(); } catch { }
            }
            TrackedTokens.Clear();
        }

        private static void DisposeAdmin()
        {
            try { if (AdminMux != null) AdminMux.Dispose(); } catch { }
            AdminMux = null;
        }

        /// <summary>Direct short-timeout multiplexer used only for admin commands
        /// (CLIENT KILL/PAUSE, PUBSUB NUMSUB, PUBLISH probes). Kept separate from
        /// the managers on purpose: faults must not depend on the code under
        /// test to be injected or measured.</summary>
        private static IDatabase AdminDb()
        {
            if (AdminMux == null)
            {
                ConfigurationOptions options = ConfigurationOptions.Parse(DirectConfig);
                options.AbortOnConnectFail = false;
                options.ConnectTimeout = 1000;
                options.SyncTimeout = 1000;
                options.ConnectRetry = 1;
                AdminMux = ConnectionMultiplexer.Connect(options);
            }
            return AdminMux.GetDatabase();
        }

        private static void PublishDirect(string channel, string message)
        {
            try { AdminDb().Execute("PUBLISH", channel, message); } catch { }
        }

        /// <summary>PUBSUB NUMSUB for one channel (>= 0) or -1 when unreadable.</summary>
        private static long NumSub(string channel)
        {
            RedisResult result = AdminDb().Execute("PUBSUB", "NUMSUB", channel);
            if (result.IsNull) return -1;
            var array = (RedisResult[])result;
            if (array.Length < 2) return -1;
            return (long)array[1];
        }

        /// <summary>Polls PUBSUB NUMSUB until the channel has at least
        /// <paramref name="expected"/> subscribers (or exactly 0 when expected
        /// is 0). Returns the elapsed ms or -1 on timeout; <paramref name="last"/>
        /// receives the last observed count.</summary>
        private static long WaitForNumSub(string channel, long expected, double timeoutMs, out long last)
        {
            long start = Stopwatch.GetTimestamp();
            long deadline = start + MsToTicks(timeoutMs);
            last = -1;
            while (Stopwatch.GetTimestamp() < deadline)
            {
                try { last = NumSub(channel); } catch { last = -1; }
                if (expected <= 0 ? last == 0 : last >= expected)
                    return (long)MsSince(start);
                Thread.Sleep(100);
            }
            return -1;
        }

        // ------------------------------------------------------------ cleanup

        private static void Cleanup()
        {
            PhaseBanner("cleanup", "drain, leak checks, shutdown");
            try { Delivery.Dispose(); } catch { }
            Wake.Stop();
            Pool.Stop();
            Proc.Stop();
            for (int i = 0; i < 4; i++)
            {
                try { RedisUDF.RedisUDFChannelUnsubscribe("liveness:b:pub:" + i, Host); } catch { }
                try { RedisUDF.RedisUDFChannelUnsubscribe("liveness:b:latest:" + i, Host); } catch { }
            }

            long noLeaks = WaitUntil(() => S.ChannelCount == 0 && S.ListenerCount == 0, 10000);
            Metric("cleanup", "channelCount", S.ChannelCount);
            Metric("cleanup", "listenerCount", S.ListenerCount);
            Metric("cleanup", "queueD1a", RedisUdfAsync.HasQueueForTests(Host + "|d1a"));
            Metric("cleanup", "queueB", RedisUdfAsync.HasQueueForTests(Host + "|queueB"));
            Check("cleanup.no-channel-leaks", noLeaks >= 0, "ChannelCount=" + S.ChannelCount + " ListenerCount=" + S.ListenerCount);

            C.Shutdown();
            Metric("cleanup", "liveConnectionsAfterShutdown", C.LiveConnectionCount());
            Check("cleanup.shutdown-pools", C.LiveConnectionCount() == 0, "LiveConnectionCount after Shutdown=" + C.LiveConnectionCount());
            DumpLatencies("cleanup");

            // Final hygiene: also tear down the process-wide runtime the UDF
            // entry points use (its own manager), so no reconnect threads keep
            // the test process alive after the checks above.
            try { RedisRuntime.Shutdown(); } catch { }
        }

        // ------------------------------------------------------------ container

        private static void StartContainerIfRequested()
        {
            if (ContainerName == null)
                return;
            if (!IsLocalEndpoint())
            {
                ContainerSkipReason = "--container requires a local host";
                Console.WriteLine("[SKIP] phase=container sub=start reason=" + ContainerSkipReason);
                return;
            }
            string version = DockerVersion();
            if (version == null)
            {
                ContainerSkipReason = "docker unavailable";
                Console.WriteLine("[SKIP] phase=container sub=start reason=" + ContainerSkipReason);
                return;
            }
            int port = ResolvePort(Host);
            Console.WriteLine("[PHASE] id=container desc=start container=" + ContainerName + " image=redis:7-alpine port=" + port);
            RunDocker("rm -f " + ContainerName, 60000); // previous disposable instance, if any
            DockerResult run = RunDocker("run -d --name " + ContainerName + " -p " + port + ":6379 redis:7-alpine", 300000);
            if (run.ExitCode != 0)
            {
                ContainerSkipReason = "docker run failed: " + Shorten(run.StdErr.Trim().Length > 0 ? run.StdErr.Trim() : run.StdOut.Trim(), 200);
                Console.WriteLine("[SKIP] phase=container sub=start reason=" + ContainerSkipReason);
                return;
            }
            ContainerManaged = true;
            Metric("container", "port", port);
        }

        private static void StopContainerIfManaged()
        {
            if (!ContainerManaged)
                return;
            Console.WriteLine("[PHASE] id=container desc=stop container=" + ContainerName);
            DockerResult stop = RunDocker("rm -f " + ContainerName, 60000);
            Metric("container", "removed", stop.ExitCode == 0);
        }

        private static double RestartContainer()
        {
            Console.WriteLine("[PHASE] id=restart desc=docker restart " + ContainerName);
            return RunDockerTimed("restart " + ContainerName, 120000);
        }

        private static string DockerVersion()
        {
            DockerResult r = RunDocker("version --format {{.Server.Version}}", 30000);
            return r.ExitCode == 0 ? r.StdOut.Trim() : null;
        }

        private sealed class DockerResult
        {
            public int ExitCode;
            public string StdOut = "";
            public string StdErr = "";
        }

        private static DockerResult RunDocker(string arguments, int timeoutMs)
        {
            var result = new DockerResult();
            try
            {
                var psi = new ProcessStartInfo("docker", arguments)
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    result.StdOut = p.StandardOutput.ReadToEnd();
                    result.StdErr = p.StandardError.ReadToEnd();
                    if (!p.WaitForExit(timeoutMs))
                    {
                        try { p.Kill(); } catch { }
                        result.ExitCode = -1;
                        result.StdErr = "docker timed out after " + timeoutMs + "ms";
                        return result;
                    }
                    result.ExitCode = p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                result.ExitCode = -1;
                result.StdErr = ex.Message;
            }
            return result;
        }

        private static double RunDockerTimed(string arguments, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            DockerResult r = RunDocker(arguments, timeoutMs);
            if (r.ExitCode != 0)
                Console.WriteLine("[WARN] docker " + arguments + " failed (exit " + r.ExitCode + "): " + Shorten(r.StdErr.Trim(), 200));
            return sw.Elapsed.TotalMilliseconds;
        }

        /// <summary>Connection string for direct (out-of-manager) multiplexers:
        /// respects an explicit abortConnect in the host argument, defaults it
        /// to abortConnect=False so a stopped server never aborts the connect.</summary>
        private static string DirectConfig
        {
            get
            {
                return Host.IndexOf("abortConnect", StringComparison.OrdinalIgnoreCase) >= 0
                    ? Host
                    : Host + ",abortConnect=False";
            }
        }

        /// <summary>True when every configured endpoint is loopback (localhost,
        /// 127.0.0.1, ::1). Only then the destructive CLIENT KILL storm and
        /// container management are allowed.</summary>
        private static bool IsLocalEndpoint()
        {
            try
            {
                ConfigurationOptions options = ConfigurationOptions.Parse(Host);
                bool any = false;
                foreach (EndPoint endpoint in options.EndPoints)
                {
                    any = true;
                    var dns = endpoint as DnsEndPoint;
                    if (dns != null)
                    {
                        if (!IsLocalHostName(dns.Host))
                            return false;
                        continue;
                    }
                    var ip = endpoint as IPEndPoint;
                    if (ip != null)
                    {
                        if (!IPAddress.IsLoopback(ip.Address))
                            return false;
                        continue;
                    }
                    return false;
                }
                return any;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsLocalHostName(string host)
        {
            if (string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase))
                return true;
            IPAddress address;
            return IPAddress.TryParse(host, out address) && IPAddress.IsLoopback(address);
        }

        /// <summary>Port of the first configured endpoint (6379 fallback).</summary>
        private static int ResolvePort(string host)
        {
            try
            {
                ConfigurationOptions options = ConfigurationOptions.Parse(host);
                foreach (EndPoint endpoint in options.EndPoints)
                {
                    var dns = endpoint as DnsEndPoint;
                    if (dns != null)
                        return dns.Port;
                    var ip = endpoint as IPEndPoint;
                    if (ip != null)
                        return ip.Port;
                }
            }
            catch { }
            return 6379;
        }

        private static string Shorten(string text, int max)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= max)
                return text ?? "";
            return text.Substring(0, max) + "...";
        }

        // ------------------------------------------------------------ helpers

        private static long PhaseBanner(string id, string desc)
        {
            Console.WriteLine();
            Console.WriteLine("[PHASE] id=" + id + " desc=" + desc + " t=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));
            foreach (var l in Lats.Values) l.Reset();
            if (Proc != null) Proc.Reset();
            if (Pool != null) Pool.Reset();
            if (Delivery != null) Delivery.ResetMax();
            if (Wake != null) Wake.ResetMax();
            return Stopwatch.GetTimestamp();
        }

        private static void Metric(string phase, string name, object value)
        {
            string text = value is double d ? Fmt(d) : Convert.ToString(value, Inv);
            Console.WriteLine("[METRIC] phase=" + phase + " name=" + name + " value=" + text);
        }

        private static bool IsTimeout(string text)
        {
            if (text == null) return false;
            // StackExchange.Redis 2.x RedisTimeoutException messages:
            // "Timeout performing {command} (...)", "Timeout awaiting response
            // (...)" and "Timeout before awaiting for tasks (...)" (all surface
            // here with the product's "Error: " prefix), plus the send-backlog
            // variant "The message timed out in the backlog ...". Match that
            // family precisely instead of only the backlog wording: otherwise a
            // transient client timeout under load is counted as a hard UDF error
            // and a slow runner fails an otherwise healthy run.
            return text.IndexOf("Error: Timeout ", StringComparison.Ordinal) >= 0
                || text.IndexOf("The message timed out", StringComparison.Ordinal) >= 0;
        }

        private static void Check(string name, bool ok, string detail)
        {
            Console.WriteLine("[ASSERT] name=" + name + " passed=" + ok + " detail=" + detail);
            if (!ok) Failures.Add(name + " :: " + detail);
        }

        private static void DumpLatencies(string phase)
        {
            foreach (var kv in Lats.OrderBy(k => k.Key))
            {
                if (kv.Value.Count == 0) continue;
                Console.WriteLine("[METRIC] phase=" + phase + " lat=" + kv.Key + " maxMs=" + Fmt(kv.Value.MaxMs) + " n=" + kv.Value.Count);
            }
        }

        private static void NoteWorst(string name, double value)
        {
            lock (WorstGate)
            {
                double cur;
                if (!Worst.TryGetValue(name, out cur) || value > cur) Worst[name] = value;
            }
        }

        private static double GetWorst(string name)
        {
            lock (WorstGate)
            {
                double cur;
                return Worst.TryGetValue(name, out cur) ? cur : -1;
            }
        }

        private static Latency GetLat(string name) { return Lats.GetOrAdd(name, _ => new Latency()); }

        private static readonly TickGate Gate = new TickGate();

        private static void NoopMessage(string message) { }

        private static void UpdateMax(ref int target, int value)
        {
            int cur;
            while ((cur = Volatile.Read(ref target)) < value)
            {
                if (Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }

        private static void RecordMax(ref long target, long value)
        {
            long cur;
            while ((cur = Interlocked.Read(ref target)) < value)
            {
                if (Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }

        private static long WaitUntil(Func<bool> probe, double timeoutMs)
        {
            long start = Stopwatch.GetTimestamp();
            long deadline = start + MsToTicks(timeoutMs);
            while (Stopwatch.GetTimestamp() < deadline)
            {
                try { if (probe()) return (long)MsSince(start); } catch { }
                Thread.Sleep(20);
            }
            return -1;
        }

        private static int CountCompleted(RecordingObserver[] observers)
        {
            int n = 0;
            for (int i = 0; i < observers.Length; i++)
                if (observers[i].CompletedCount == 1) n++;
            return n;
        }

        private static void GetPool(out int minW, out int minIo, out int maxW, out int maxIo)
        {
            ThreadPool.GetMinThreads(out minW, out minIo);
            ThreadPool.GetMaxThreads(out maxW, out maxIo);
        }

        /// <summary>
        /// Starves the pool for an attack window: saves the current bounds,
        /// clamps min/max to (4,4) and blocks six queued sleepers on the
        /// returned event (released by <see cref="RestorePool"/>). The saved
        /// bounds are returned for the restore.
        /// </summary>
        private static ManualResetEventSlim StarvePool(
            out int minW, out int minIo, out int maxW, out int maxIo,
            out bool setMin, out bool setMax, int settleMs)
        {
            GetPool(out minW, out minIo, out maxW, out maxIo);
            setMin = ThreadPool.SetMinThreads(4, 4);
            setMax = ThreadPool.SetMaxThreads(4, 4);
            var release = new ManualResetEventSlim(false);
            for (int i = 0; i < 6; i++)
                ThreadPool.QueueUserWorkItem(_ => release.Wait(120000));
            Thread.Sleep(settleMs);
            return release;
        }

        /// <summary>Ends a <see cref="StarvePool"/> window: releases the sleepers
        /// and restores the saved pool bounds.</summary>
        private static void RestorePool(int minW, int minIo, int maxW, int maxIo, ManualResetEventSlim release)
        {
            release.Set();
            ThreadPool.SetMaxThreads(maxW, maxIo);
            ThreadPool.SetMinThreads(minW, minIo);
        }

        /// <summary>
        /// Waits for the delivery probe to resume past the outage and for the
        /// live channel to be re-subscribed, emitting the two metrics and the
        /// two checks of a phase C recovery pair. Returns the resume time.
        /// </summary>
        private static double CheckDeliveryRecovered(
            long prevDeliveries, string resumeMetric, string readersMetric,
            string resumeCheck, string resubscribedCheck)
        {
            double resume = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);
            long readers = Delivery.WaitForReaders(LiveChannel, ResumeBoundMs);
            Metric("C", resumeMetric, resume);
            Metric("C", readersMetric, readers);
            Check(resumeCheck, resume >= 0 && resume <= ResumeBoundMs, "resumeMs=" + Fmt(resume));
            Check(resubscribedCheck, readers > 0, "publish readers=" + readers);
            return resume;
        }

        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private static double TicksToMs(long ticks) { return ticks * 1000.0 / Stopwatch.Frequency; }
        private static long MsToTicks(double ms) { return (long)(ms * Stopwatch.Frequency / 1000.0); }
        private static double MsSince(long startTicks) { return TicksToMs(Stopwatch.GetTimestamp() - startTicks); }

        private static string Fmt(double ms)
        {
            return ms < 0 ? "timeout" : ms.ToString("0.0", Inv);
        }

        // ------------------------------------------------------------ monitors

        private sealed class Latency
        {
            private long _count;
            private long _maxTicks;

            public void Record(long ticks)
            {
                Interlocked.Increment(ref _count);
                RecordMax(ref _maxTicks, ticks);
            }

            public void Reset()
            {
                Interlocked.Exchange(ref _count, 0);
                Interlocked.Exchange(ref _maxTicks, 0);
            }

            public long Count { get { return Interlocked.Read(ref _count); } }
            public double MaxMs { get { return TicksToMs(Interlocked.Read(ref _maxTicks)); } }
        }

        private sealed class Worker
        {
            private readonly Thread[] _threads;
            private volatile bool _stop;
            private long _ops;
            private long _errors;
            private long _timeouts;
            private string _firstError;

            public readonly string Name;

            public Worker(string name, int count, Action<int, long> body)
            {
                Name = name;
                _threads = new Thread[count];
                for (int i = 0; i < count; i++)
                {
                    int id = i;
                    var t = new Thread(() =>
                    {
                        long n = 0;
                        while (!_stop)
                        {
                            try { body(id, n); Interlocked.Increment(ref _ops); }
                            catch (Exception ex)
                            {
                                Interlocked.Increment(ref _errors);
                                if (ex is RedisTimeoutException) Interlocked.Increment(ref _timeouts);
                                if (Volatile.Read(ref _firstError) == null)
                                    Volatile.Write(ref _firstError, ex.GetType().Name + ": " + ex.Message);
                            }
                            n++;
                        }
                    });
                    t.IsBackground = true;
                    t.Name = name + "-" + i;
                    _threads[i] = t;
                }
            }

            public long Ops { get { return Interlocked.Read(ref _ops); } }
            public long Errors { get { return Interlocked.Read(ref _errors); } }
            public long Timeouts { get { return Interlocked.Read(ref _timeouts); } }
            public string FirstError { get { return Volatile.Read(ref _firstError); } }
            public void Start() { foreach (var t in _threads) t.Start(); }

            public void StopAndJoin(int timeoutMs = 15000)
            {
                _stop = true;
                foreach (var t in _threads) t.Join(timeoutMs);
            }
        }

        /// <summary>Dedicated-thread heartbeat: a long gap means the process
        /// itself stalled (not the pool).</summary>
        private sealed class ProcessHeartbeat
        {
            private readonly string _name;
            private readonly int _intervalMs;
            private Thread _thread;
            private volatile bool _stop;
            private long _maxGapTicks;
            private long _ticks;

            public ProcessHeartbeat(string name, int intervalMs) { _name = name; _intervalMs = intervalMs; }

            public void Start()
            {
                _thread = new Thread(() =>
                {
                    long last = Stopwatch.GetTimestamp();
                    while (!_stop)
                    {
                        Thread.Sleep(_intervalMs);
                        long now = Stopwatch.GetTimestamp();
                        RecordMax(ref _maxGapTicks, now - last);
                        last = now;
                        Interlocked.Increment(ref _ticks);
                    }
                });
                _thread.IsBackground = true;
                _thread.Name = _name;
                _thread.Start();
            }

            public void Stop() { _stop = true; }
            public void Reset() { Interlocked.Exchange(ref _maxGapTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public long Ticks { get { return Interlocked.Read(ref _ticks); } }
        }

        /// <summary>ThreadPool heartbeat: self-rescheduling continuation; its gap
        /// measures how long the pool could not run work (starvation probe).</summary>
        private sealed class PoolHeartbeat
        {
            private volatile bool _stop;
            private long _maxGapTicks;
            private long _lastTicks;
            private long _ticks;

            public void Start()
            {
                Interlocked.Exchange(ref _lastTicks, Stopwatch.GetTimestamp());
                Schedule();
            }

            public void Stop() { _stop = true; }
            public void Reset() { Interlocked.Exchange(ref _maxGapTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public long Ticks { get { return Interlocked.Read(ref _ticks); } }

            private void Schedule()
            {
                Task.Delay(25).ContinueWith(_ =>
                {
                    if (_stop) return;
                    long now = Stopwatch.GetTimestamp();
                    RecordMax(ref _maxGapTicks, now - Interlocked.Read(ref _lastTicks));
                    Interlocked.Exchange(ref _lastTicks, now);
                    Interlocked.Increment(ref _ticks);
                    Schedule();
                }, TaskScheduler.Default);
            }
        }

        /// <summary>Dedicated watchdog calling real public entry points; the max
        /// iteration time is the "is any entry point stuck behind a lock" probe.</summary>
        private sealed class WatchdogProbe
        {
            private readonly RedisConnectionManager _connections;
            private readonly RedisSubscriptionManager _subscriptions;
            private readonly TickGate _gate = new TickGate();
            private Thread _thread;
            private volatile bool _stop;
            private long _maxIterTicks;
            private long _errors;

            public WatchdogProbe(RedisConnectionManager connections, RedisSubscriptionManager subscriptions)
            {
                _connections = connections;
                _subscriptions = subscriptions;
            }

            public void Start()
            {
                _thread = new Thread(Loop);
                _thread.IsBackground = true;
                _thread.Name = "watchdog-probe";
                _thread.Start();
            }

            public void Stop() { _stop = true; }
            public void ResetMax() { Interlocked.Exchange(ref _maxIterTicks, 0); }
            public double MaxIterMs { get { return TicksToMs(Interlocked.Read(ref _maxIterTicks)); } }
            public long Errors { get { return Interlocked.Read(ref _errors); } }

            private void Loop()
            {
                while (!_stop)
                {
                    long t0 = Stopwatch.GetTimestamp();
                    try
                    {
                        int a = _subscriptions.ChannelCount;
                        int b = _subscriptions.ListenerCount;
                        int c = _subscriptions.ChannelCountWithOrigin("B");
                        int d = _connections.LiveConnectionCount();
                        int e = _connections.RtdConnectionCount;
                        if (!_gate.TryEnter()) throw new InvalidOperationException("tick gate unexpectedly busy");
                        _gate.Exit();
                        GC.KeepAlive(a + b + c + d + e);
                    }
                    catch { Interlocked.Increment(ref _errors); }
                    RecordMax(ref _maxIterTicks, Stopwatch.GetTimestamp() - t0);
                    Thread.Sleep(20);
                }
            }
        }

        /// <summary>Live delivery probe: a real subscription on "liveness:live"
        /// fed by a dedicated publishing thread; records inter-arrival gaps,
        /// latency and the resume time after an attack window.</summary>
        private sealed class DeliveryProbe
        {
            private readonly object _gate = new object();
            private RedisConnectionManager _connections;
            private string _host;
            private string _channel;
            private IDisposable _token;
            private volatile bool _stop;
            private Thread _publisher;
            private long _count;
            private long _lastArrivalTicks;
            private long _maxGapTicks;
            private long _maxLatencyTicks;
            private long _published;
            private long _publishErrors;
            private long _lastPublishTicks;

            public void Start(RedisConnectionManager connections, RedisSubscriptionManager subscriptions, string host, string channel)
            {
                _connections = connections;
                _host = host;
                _channel = channel;
                Interlocked.Exchange(ref _lastArrivalTicks, Stopwatch.GetTimestamp());
                _token = subscriptions.Subscribe(host, channel, false, OnMessage, "LIVENESS");
                _publisher = new Thread(PublishLoop) { IsBackground = true, Name = "delivery-publisher" };
                _publisher.Start();
            }

            public long Count { get { return Interlocked.Read(ref _count); } }
            public long Published { get { return Interlocked.Read(ref _published); } }
            public long PublishErrors { get { return Interlocked.Read(ref _publishErrors); } }
            public void ResetMax() { Interlocked.Exchange(ref _maxGapTicks, 0); Interlocked.Exchange(ref _maxLatencyTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public double MaxLatencyMs { get { return TicksToMs(Interlocked.Read(ref _maxLatencyTicks)); } }

            public double ResumeAfter(long previousCount, double timeoutMs)
            {
                long start = Stopwatch.GetTimestamp();
                long deadline = start + MsToTicks(timeoutMs);
                while (Stopwatch.GetTimestamp() < deadline)
                {
                    if (Interlocked.Read(ref _count) > previousCount) return MsSince(start);
                    Thread.Sleep(5);
                }
                return -1;
            }

            public long WaitForReaders(string channel, double timeoutMs)
            {
                long deadline = Stopwatch.GetTimestamp() + MsToTicks(timeoutMs);
                long probe = 0;
                while (Stopwatch.GetTimestamp() < deadline)
                {
                    try
                    {
                        var sub = _connections.GetSubscriber(_host, RedisPool.RtdSub);
                        long readers = sub.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "probe-" + Interlocked.Increment(ref probe));
                        if (readers > 0) return readers;
                    }
                    catch { }
                    Thread.Sleep(100);
                }
                return -1;
            }

            public void Dispose()
            {
                _stop = true;
                try { if (_token != null) _token.Dispose(); } catch { }
                if (_publisher != null) _publisher.Join(5000);
            }

            private void OnMessage(string message)
            {
                long now = Stopwatch.GetTimestamp();
                lock (_gate)
                {
                    if (Interlocked.Read(ref _count) > 0)
                        RecordMax(ref _maxGapTicks, now - _lastArrivalTicks);
                    _lastArrivalTicks = now;
                    Interlocked.Increment(ref _count);
                    long pub = Interlocked.Read(ref _lastPublishTicks);
                    if (pub > 0) RecordMax(ref _maxLatencyTicks, now - pub);
                }
            }

            private void PublishLoop()
            {
                var sub = _connections.GetSubscriber(_host, RedisPool.RtdSub);
                long n = 0;
                while (!_stop)
                {
                    try
                    {
                        Interlocked.Exchange(ref _lastPublishTicks, Stopwatch.GetTimestamp());
                        sub.Publish(new RedisChannel(_channel, RedisChannel.PatternMode.Literal), "hb-" + Interlocked.Increment(ref n));
                        Interlocked.Increment(ref _published);
                    }
                    catch { Interlocked.Increment(ref _publishErrors); }
                    Thread.Sleep(25);
                }
            }
        }

        // ------------------------------------------------------------ observers
    }
}

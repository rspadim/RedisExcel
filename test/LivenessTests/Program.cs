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
    /// Usage:
    ///   dotnet run --project test\LivenessTests -c Release -- "127.0.0.1:6399,abortConnect=False" [--container &lt;name&gt;] [--skip-restart] [--quick]
    /// </summary>
    internal static class Program
    {
        private static string Host = "127.0.0.1:6399";
        private static string ContainerName;   // --container <name>
        private static bool ContainerManaged;  // container started by this run
        private static string ContainerSkipReason;
        private static bool SkipRestart;       // --skip-restart
        private static bool Quick;             // --quick

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
            RedisSubscriptionManager.ListenerJoined += (h, c, p) => Interlocked.Increment(ref JoinedEvents);
            try
            {
                StartContainerIfRequested();
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
            "  --quick           shortened attack windows for a fast sanity run";

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
                else if (a.StartsWith("--", StringComparison.Ordinal))
                    throw new ArgumentException("unknown option '" + a + "'");
                else
                    positional.Add(a);
            }
            if (positional.Count > 0) Host = positional[0];
            if (positional.Count > 1)
                throw new ArgumentException("unexpected extra argument '" + positional[1] + "'");
        }

        private static void RunAll()
        {
            Proc = new ProcessHeartbeat("proc-heartbeat", 20);
            Proc.Start();
            Pool = new PoolHeartbeat();
            Pool.Start();
            C = new RedisConnectionManager();
            S = new RedisSubscriptionManager(C);
            Wake = new WatchdogProbe(C, S);
            Wake.Start();

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

            int minW, minIo, maxW, maxIo;
            GetPool(out minW, out minIo, out maxW, out maxIo);
            bool setMin = ThreadPool.SetMinThreads(4, 4);
            bool setMax = ThreadPool.SetMaxThreads(4, 4);
            Check("A.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            if (!(setMin && setMax))
            {
                ThreadPool.SetMaxThreads(maxW, maxIo);
                ThreadPool.SetMinThreads(minW, minIo);
                return;
            }

            var release = new ManualResetEventSlim(false);
            for (int i = 0; i < 6; i++)
                ThreadPool.QueueUserWorkItem(_ => release.Wait(120000));
            Thread.Sleep(400); // let the starvation take hold
            long poolTicksAtStarveStart = Pool.Ticks;

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
            Thread.Sleep(Quick ? 2500 : 6000);
            churn.StopAndJoin();
            pubs.StopAndJoin();

            double poolGapDuring = Pool.MaxGapMs;
            double deliveryGapDuring = Delivery.MaxGapMs;
            double watchdogDuring = Wake.MaxIterMs;
            long poolTicksDuringAttack = Pool.Ticks - poolTicksAtStarveStart;
            int queuePendingAtAttackEnd = drained.CurrentCount;

            // Release the pool: sleepers end, continuations and heartbeats resume.
            release.Set();
            ThreadPool.SetMaxThreads(maxW, maxIo);
            ThreadPool.SetMinThreads(minW, minIo);
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

            Check("A.attack-bit", poolTicksDuringAttack == 0 && queuePendingAtAttackEnd > 0,
                "pool ticks during attack=" + poolTicksDuringAttack + " queued items stuck=" + queuePendingAtAttackEnd + " " + poolGapAfterRelease.ToString("0.0", Inv) + "ms gap measured after release");
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
            int latestCount = 4, pubCount = 4;
            var pubChannels = Enumerable.Range(0, pubCount).Select(i => "liveness:b:pub:" + i).ToArray();
            var latestChannels = Enumerable.Range(0, latestCount).Select(i => "liveness:b:latest:" + i).ToArray();
            var leftovers = new ConcurrentBag<IDisposable>();

            var workers = new List<Worker>();
            workers.Add(new Worker("B-subscribe", 2, (id, n) =>
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
                if (res is string s && s.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref udfErrors);
            }));
            workers.Add(new Worker("B-udf-pif", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object res = RedisUDF.RedisUDFChannelPublishIfChanged(pubChannels[n % pubCount], "pif-" + n, Host);
                GetLat("B.udf.publishIfChanged").Record(Stopwatch.GetTimestamp() - t0);
                if (res is string s && s.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref udfErrors);
            }));
            workers.Add(new Worker("B-udf-latest", 1, (id, n) =>
            {
                string ch = latestChannels[n % latestCount];
                long t0 = Stopwatch.GetTimestamp();
                string latest = RedisUDF.RedisUDFChannelLatest(ch, Host);
                GetLat("B.udf.latest").Record(Stopwatch.GetTimestamp() - t0);
                if (latest != null && latest.StartsWith("Error:", StringComparison.Ordinal)) Interlocked.Increment(ref udfErrors);
                long t1 = Stopwatch.GetTimestamp();
                RedisUDF.RedisUDFChannelUnsubscribe(ch, Host);
                GetLat("B.udf.unsubscribe").Record(Stopwatch.GetTimestamp() - t1);
            }));
            workers.Add(new Worker("B-queue", 2, (id, n) =>
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
            Metric("B", "deliveryGapMs", Delivery.MaxGapMs);
            Metric("B", "deliveryResumeMs", resume);
            Metric("B", "watchdogMaxMs", Wake.MaxIterMs);
            Metric("B", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("B", "recovered", recovered);
            Metric("B", "recoverError", recoverError ?? "-");

            Check("B.dedicated-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            Check("B.entrypoints-not-stuck", Wake.MaxIterMs < StuckBoundMs, "watchdog max iteration=" + Fmt(Wake.MaxIterMs) + "ms");
            Check("B.delivery-stays-live", Delivery.MaxGapMs < StuckBoundMs, "delivery max gap=" + Fmt(Delivery.MaxGapMs) + "ms");
            Check("B.no-unexpected-errors", errors - timeouts == 0 && udfErrors == 0,
                "errors=" + errors + " timeouts=" + timeouts + " udfErrors=" + udfErrors);
            Check("B.subscribe-recovers-after-flood", recovered, "recovery subscribe succeeded once the flood stopped" + (recoverError == null ? "" : " lastError=" + recoverError));

            foreach (var t in leftovers) { try { t.Dispose(); } catch { } }
            foreach (var ch in latestChannels) { try { RedisUDF.RedisUDFChannelUnsubscribe(ch, Host); } catch { } }
            RedisUDF.SyncWriteOverrideForTests = prevSyncWrite;
            Metric("B", "durationSec", MsSince(start) / 1000.0);
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

                double killResume = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);
                long readers = Delivery.WaitForReaders("liveness:live", ResumeBoundMs);
                Metric("C", "killRounds", rounds);
                Metric("C", "clientsKilled", clientsKilled);
                Metric("C", "killErrors", killErrors);
                Metric("C", "killDeliveryResumeMs", killResume);
                Metric("C", "killReadersAfterMs", readers);
                Metric("C", "deliveryGapMs", Delivery.MaxGapMs);
                Metric("C", "procHeartbeatMs", Proc.MaxGapMs);
                Check("C.kill-storm-delivery-resumes", killResume >= 0 && killResume <= ResumeBoundMs, "resumeMs=" + Fmt(killResume));
                Check("C.kill-storm-resubscribed", readers > 0, "publish readers=" + readers);
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
                double restartResume = Delivery.ResumeAfter(prevDeliveries2, ResumeBoundMs);
                long readers2 = Delivery.WaitForReaders("liveness:live", ResumeBoundMs);

                Metric("C", "restartWallMs", restartWallMs);
                Metric("C", "serverUpAfterMs", serverUpMs);
                Metric("C", "restartDeliveryResumeMs", restartResume);
                Metric("C", "restartReaders", readers2);
                Metric("C", "deliveryGapKillAndRestartMs", Delivery.MaxGapMs);
                Metric("C", "publishErrorsTotal", Delivery.PublishErrors);
                Check("C.restart-server-up", serverUpMs >= 0, "ping after restart ms=" + Fmt(serverUpMs));
                Check("C.restart-delivery-resumes", restartResume >= 0 && restartResume <= ResumeBoundMs, "resumeMs=" + Fmt(restartResume));
                Check("C.restart-resubscribed", readers2 > 0, "publish readers=" + readers2);
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
            // when the pool is released.
            int minW, minIo, maxW, maxIo;
            GetPool(out minW, out minIo, out maxW, out maxIo);
            bool setMin = ThreadPool.SetMinThreads(4, 4);
            bool setMax = ThreadPool.SetMaxThreads(4, 4);
            var release = new ManualResetEventSlim(false);
            for (int i = 0; i < 6; i++)
                ThreadPool.QueueUserWorkItem(_ => release.Wait(120000));
            Thread.Sleep(300);

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
            release.Set();
            ThreadPool.SetMaxThreads(maxW, maxIo);
            ThreadPool.SetMinThreads(minW, minIo);
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
            var throwing = new ThrowingObserver();
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

        private sealed class RecordingObserver : IExcelObserver
        {
            private readonly object _gate = new object();
            private readonly List<object> _values = new List<object>();
            private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
            private int _completedCount;

            public void OnNext(object value) { lock (_gate) _values.Add(value); }
            public void OnError(Exception exception) { }
            public void OnCompleted()
            {
                lock (_gate) _completedCount++;
                _completed.Set();
            }

            public object[] Values { get { lock (_gate) return _values.ToArray(); } }
            public int CompletedCount { get { lock (_gate) return _completedCount; } }
            public bool WaitCompleted(int timeoutMs) { return _completed.Wait(timeoutMs); }
        }

        private sealed class ThrowingObserver : IExcelObserver
        {
            private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
            private int _completedCount;

            public void OnNext(object value) { throw new InvalidOperationException("observer rejected the value"); }
            public void OnError(Exception exception) { }
            public void OnCompleted()
            {
                Interlocked.Increment(ref _completedCount);
                _completed.Set();
            }

            public int CompletedCount { get { return Volatile.Read(ref _completedCount); } }
            public bool WaitCompleted(int timeoutMs) { return _completed.Wait(timeoutMs); }
        }
    }
}

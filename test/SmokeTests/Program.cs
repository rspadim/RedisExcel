using RedisExcel;
using StackExchange.Redis;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Behavior smoke test for RedisConnectionManager / RedisSubscriptionManager
/// (no Excel required). Requires a running Redis server; pass the connection
/// string as the first argument, or let it use the default below.
///
/// Run: dotnet run --project test\SmokeTests -c Release -- "127.0.0.1:6379,abortConnect=False"
/// </summary>
internal static class Program
{
    private const string DefaultHost = "127.0.0.1:6379,abortConnect=False";

    // Dedicated Redis for the concurrent dedup regression: a throwaway
    // container on 6396, so the burst never touches the smoke's main server.
    private const string ConcurrentEndpoint = "127.0.0.1:6396";
    private const string ConcurrentContainerName = "rs-smoke-2";

    private static int _failures;
    private static readonly object Sync = new object();

    private static void Check(bool condition, string label)
    {
        Console.WriteLine((condition ? "PASS " : "FAIL ") + label);
        if (!condition) Interlocked.Increment(ref _failures);
    }

    private static bool WaitUntil(Func<bool> condition, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            if (condition()) return true;
            Thread.Sleep(50);
        }
        return condition();
    }

    private static int Main(string[] args)
    {
        string host = args.Length > 0 ? args[0] : DefaultHost;
        string channel = "smoke:" + Guid.NewGuid().ToString("N");

        Console.WriteLine($"Redis host: {host}");
        Console.WriteLine($"Channel:    {channel}");

        var connections = new RedisConnectionManager();
        var subscriptions = new RedisSubscriptionManager(connections);

        var db1 = connections.GetDatabase(host, RedisPool.UdfData);
        var db2 = connections.GetDatabase(host, RedisPool.UdfData);
        Check(ReferenceEquals(db1, db2), "IDatabase wrapper is cached per host/pool");
        var sub1 = connections.GetSubscriber(host);
        var sub2 = connections.GetSubscriber(host);
        Check(ReferenceEquals(sub1, sub2), "ISubscriber wrapper is cached per host/pool");

        var receivedA = new List<string>();
        var receivedB = new List<string>();

        var tokenA = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedA.Add(m); });
        var tokenB = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedB.Add(m); });

        var pubConn = connections.GetConnection(host, RedisPool.UdfData);
        var publisher = pubConn.GetSubscriber();
        var server = pubConn.GetServer(pubConn.GetEndPoints().First());

        Func<long> numsub = () =>
        {
            var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", channel);
            return arr.Length >= 2 ? (long)arr[1] : 0;
        };

        Check(WaitUntil(() => numsub() == 1, 5000), "server sees exactly one physical subscription for the channel");
        Check(subscriptions.ChannelCount == 1 && subscriptions.ListenerCount == 2, "counters: 1 channel, 2 listeners");

        long readers = publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "msg1");
        Check(WaitUntil(() => { lock (Sync) return receivedA.Count == 1 && receivedB.Count == 1; }, 5000), "both listeners received msg1");
        Check(readers == 1, "publish reports 1 physical reader (local broadcast to 2 listeners)");

        // Regression (v1.1.0): disconnecting one topic must not tear down the other listener.
        tokenA.Dispose();
        Check(WaitUntil(() => subscriptions.ListenerCount == 1, 5000), "listener A removed");
        Check(WaitUntil(() => numsub() == 1, 5000), "channel stays subscribed while B is still active");

        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "msg2");
        Check(WaitUntil(() => { lock (Sync) return receivedB.Count == 2 && receivedB[1] == "msg2"; }, 5000), "B keeps receiving after A left");

        // The last listener unsubscribes the channel.
        tokenB.Dispose();
        Check(WaitUntil(() => numsub() == 0, 5000), "channel unsubscribed after the last listener left");
        Check(subscriptions.ChannelCount == 0 && subscriptions.ListenerCount == 0, "counters reset to zero");

        // Re-subscribing the same channel must work (no stuck state).
        var receivedC = new List<string>();
        var tokenC = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedC.Add(m); });
        Check(WaitUntil(() => numsub() == 1, 5000), "re-subscription registered on the server");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "msg3");
        Check(WaitUntil(() => { lock (Sync) return receivedC.Count == 1 && receivedC[0] == "msg3"; }, 5000), "new listener received msg3");
        tokenC.Dispose();

        // Duplicate suppression: identical consecutive payloads are skipped when
        // SkipRepeatedMessages is on (default). A RedisExcel.json in the user
        // profile may disable it, so adapt the expectations to the loaded config.
        bool skipRepeated = AppConfig.Current.SkipRepeatedMessages;
        var receivedD = new List<string>();
        var tokenD = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedD.Add(m); });
        Check(WaitUntil(() => numsub() == 1, 5000), "channel active for the duplicate test");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup2");
        if (skipRepeated)
        {
            Check(WaitUntil(() => { lock (Sync) return receivedD.Contains("dup2"); }, 5000), "changed payload delivered");
            Check(WaitUntil(() => { lock (Sync) return receivedD.Count == 2; }, 1000), "identical repeated payload skipped");
        }
        else
        {
            Check(WaitUntil(() => { lock (Sync) return receivedD.Contains("dup2"); }, 5000),
                "changed payload delivered (dedup disabled by config)");
            Check(WaitUntil(() => { lock (Sync) return receivedD.Count == 3; }, 1000),
                "identical repeated payloads delivered (dedup disabled by config)");
        }
        tokenD.Dispose();

        // Pattern subscriptions are never deduplicated, and pattern/literal
        // refcounts are independent (separate channel states). Unique channels
        // per run avoid interference from leftovers of an aborted prior run.
        string patBase = "smoke:pat:" + Guid.NewGuid().ToString("N");
        var receivedPat = new List<string>();
        var receivedLit = new List<string>();
        var patToken = subscriptions.Subscribe(host, patBase + ":*", pattern: true,
            onMessage: m => { lock (Sync) receivedPat.Add(m); });
        var litToken = subscriptions.Subscribe(host, patBase + ":other", pattern: false,
            onMessage: m => { lock (Sync) receivedLit.Add(m); });

        Func<string, long> numsubOf = ch =>
        {
            var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", ch);
            return arr.Length >= 2 ? (long)arr[1] : 0;
        };

        // Pattern subscriptions do not show up in PUBSUB NUMSUB; their delivery
        // assertions below prove them, while the literal side is waited on here.
        Check(WaitUntil(() => numsubOf(patBase + ":other") == 1, 5000),
            "literal channel active on the server");
        Check(subscriptions.ChannelCount == 2 && subscriptions.ListenerCount == 2,
            "pattern and literal subscriptions tracked independently");

        publisher.Publish(new RedisChannel(patBase + ":one", RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(patBase + ":one", RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(patBase + ":other", RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(patBase + ":one", RedisChannel.PatternMode.Literal), "dup2");

        // patBase + ":*" also matches patBase + ":other", so the pattern
        // listener must see all four messages (2x "dup" @ :one + 1x "dup" @
        // :other + "dup2"). If identical repeated payloads were suppressed for
        // patterns it would stop at two ("dup" + "dup2").
        Check(WaitUntil(() => { lock (Sync) return receivedPat.Count == 4; }, 5000),
            "pattern listener received every matching message (3x \"dup\" + \"dup2\", no deduplication)");
        bool patternDuplicatesDelivered;
        lock (Sync)
        {
            patternDuplicatesDelivered =
                receivedPat.Count(m => m == "dup") == 3 && receivedPat.Count(m => m == "dup2") == 1;
        }
        Check(patternDuplicatesDelivered, "identical repeated payloads are not deduplicated for patterns");
        Check(WaitUntil(() => { lock (Sync) return receivedLit.Count == 1 && receivedLit[0] == "dup"; }, 5000),
            "literal listener received only the message published to its own channel");

        // Dropping the pattern registration must not touch the literal channel.
        patToken.Dispose();
        Check(WaitUntil(() => subscriptions.ListenerCount == 1, 5000),
            "pattern listener removed, literal listener kept");
        Check(subscriptions.ChannelCount == 1, "only the literal channel state remains");
        publisher.Publish(new RedisChannel(patBase + ":other", RedisChannel.PatternMode.Literal), "after-pat");
        Check(WaitUntil(() => { lock (Sync) return receivedLit.Count == 2 && receivedLit[1] == "after-pat"; }, 5000),
            "literal channel still receives after the pattern left");
        Check(WaitUntil(() => numsubOf(patBase + ":other") == 1, 5000),
            "literal channel still subscribed on the server");
        litToken.Dispose();
        Check(WaitUntil(() => numsubOf(patBase + ":other") == 0, 5000),
            "literal channel unsubscribed after its last listener left");
        Check(subscriptions.ChannelCount == 0 && subscriptions.ListenerCount == 0,
            "all pattern/literal channels released");

        // Late-joiner dedup regression (v1.2.6): the duplicate-suppression marker
        // lives on the shared channel state, so a listener joining a channel that
        // already had listeners must still receive a republished identical payload
        // (previously the marker suppressed it and the new cell stayed blank).
        string lateJoinChannel = "smoke:late:" + Guid.NewGuid().ToString("N");
        var receivedLateA = new List<string>();
        var receivedLateB = new List<string>();
        var lateTokenA = subscriptions.Subscribe(host, lateJoinChannel, pattern: false,
            onMessage: m => { lock (Sync) receivedLateA.Add(m); });
        Check(WaitUntil(() => numsubOf(lateJoinChannel) == 1, 5000),
            "late-join channel active on the server");
        publisher.Publish(new RedisChannel(lateJoinChannel, RedisChannel.PatternMode.Literal), "same");
        Check(WaitUntil(() => { lock (Sync) return receivedLateA.Count == 1; }, 5000),
            "late-join baseline: existing listener received the initial payload");

        var lateTokenB = subscriptions.Subscribe(host, lateJoinChannel, pattern: false,
            onMessage: m => { lock (Sync) receivedLateB.Add(m); });
        publisher.Publish(new RedisChannel(lateJoinChannel, RedisChannel.PatternMode.Literal), "same");
        Check(WaitUntil(() => { lock (Sync) return receivedLateB.Count >= 1 && receivedLateB[0] == "same"; }, 5000),
            "late joiner received the republished identical payload");

        lateTokenA.Dispose();
        lateTokenB.Dispose();
        Check(WaitUntil(() => numsubOf(lateJoinChannel) == 0, 5000),
            "late-join channel unsubscribed after both listeners left");

        // Disposing the same token twice is a safe no-op.
        string doubleDisposeChannel = "smoke:double:" + Guid.NewGuid().ToString("N");
        int channelsBefore = subscriptions.ChannelCount;
        int listenersBefore = subscriptions.ListenerCount;
        var doubleDisposeToken = subscriptions.Subscribe(host, doubleDisposeChannel, pattern: false, onMessage: m => { });
        Check(WaitUntil(() => subscriptions.ListenerCount == listenersBefore + 1, 5000),
            "temporary listener registered for the double-dispose test");
        doubleDisposeToken.Dispose();
        Check(subscriptions.ChannelCount == channelsBefore && subscriptions.ListenerCount == listenersBefore,
            "counters back to baseline after the first Dispose");
        try
        {
            doubleDisposeToken.Dispose();
            Check(true, "second Dispose does not throw");
        }
        catch (Exception ex)
        {
            Check(false, "second Dispose does not throw (" + ex.GetType().Name + ")");
        }
        Check(subscriptions.ChannelCount == channelsBefore && subscriptions.ListenerCount == listenersBefore,
            "counters unchanged by the second Dispose");

        // Invalid arguments are rejected before any Redis call.
        try
        {
            subscriptions.Subscribe("", "c", false, m => { });
            Check(false, "blank host accepted");
        }
        catch (ArgumentException)
        {
            Check(true, "blank host rejected");
        }
        try
        {
            subscriptions.Subscribe(host, "", false, m => { });
            Check(false, "blank channel accepted");
        }
        catch (ArgumentException)
        {
            Check(true, "blank channel rejected");
        }

        // Origin tags: caller-defined labels ("RTD"/"UDF" in the add-in) scope the
        // counters to one consumer; the parameterless counters stay totals.
        string originChannelA = "smoke:originA:" + Guid.NewGuid().ToString("N");
        string originChannelB = "smoke:originB:" + Guid.NewGuid().ToString("N");
        int totalListenersBefore = subscriptions.ListenerCount;
        int totalChannelsBefore = subscriptions.ChannelCount;

        var originTokenA = subscriptions.Subscribe(host, originChannelA, pattern: false,
            onMessage: m => { }, origin: "smokeA");
        var originTokenB = subscriptions.Subscribe(host, originChannelB, pattern: false,
            onMessage: m => { }, origin: "smokeB");

        Check(WaitUntil(() =>
                subscriptions.ListenerCountWithOrigin("smokeA") == 1 && subscriptions.ListenerCountWithOrigin("smokeB") == 1, 5000),
            "origin listener counts track each tag");
        Check(subscriptions.ChannelCountWithOrigin("smokeA") == 1 && subscriptions.ChannelCountWithOrigin("smokeB") == 1,
            "origin channel counts track channels with that tag");
        Check(subscriptions.ListenerCount == totalListenersBefore + 2 && subscriptions.ChannelCount == totalChannelsBefore + 2,
            "parameterless counters reflect exactly the two origin subscriptions");

        originTokenA.Dispose();
        originTokenB.Dispose();

        Check(WaitUntil(() =>
                subscriptions.ListenerCountWithOrigin("smokeA") == 0 && subscriptions.ListenerCountWithOrigin("smokeB") == 0, 5000),
            "origin listener counts drop after disposal");
        Check(WaitUntil(() =>
                subscriptions.ChannelCountWithOrigin("smokeA") == 0 && subscriptions.ChannelCountWithOrigin("smokeB") == 0, 5000),
            "origin channel counts drop after disposal");
        Check(subscriptions.ListenerCount == totalListenersBefore && subscriptions.ChannelCount == totalChannelsBefore,
            "parameterless counters back to baseline after the origin test");

        // Liveness hardening regressions (v1.4.x working tree): the per-key
        // NetworkGate that serializes subscribe/unsubscribe across channel-state
        // generations, the ChannelLatest lifecycle epochs + per-listener sync,
        // and the pattern-join publish-marker purge (ClearForListener). They use
        // the process-wide RedisRuntime (what RedisUDF uses) and the local
        // manager created above, and release every listener they create.
        RunConcurrentChurnTest(subscriptions, publisher, server, host);
        RedisUdfAsync.AsyncWritesOverrideForTests = false; // exercise the synchronous UDF path
        try
        {
            RunChannelLatestRaceTest(publisher, host);
            RunGhostStaleTest(pubConn, publisher, host);
            RunPatternJoinOverClearTest(subscriptions, publisher, server, host);
        }
        finally
        {
            RedisUdfAsync.AsyncWritesOverrideForTests = null;
            RedisUDF.SyncWriteOverrideForTests = null;
        }

        // Concurrency regression for the duplicate-suppression marker, kept last
        // because it hammers a dedicated Redis (container on 6396); it creates
        // and tears down its own listener, so the counters above are unaffected.
        RunConcurrentDedupDeliveryTest(subscriptions, host, skipRepeated);

        Console.WriteLine(_failures == 0 ? "ALL PASS" : _failures + " FAILURE(S)");
        return _failures == 0 ? 0 : 1;
    }

    /// <summary>
    /// Concurrency regression for the duplicate-suppression marker: the messages
    /// of one literal channel are delivered by StackExchange.Redis through the
    /// thread pool, so several HandleMessage calls can run at the same time. The
    /// shared "last message" marker used to be read while another thread replaced
    /// it; a torn RedisValue read then threw inside the subscriber callback (which
    /// StackExchange.Redis swallows), leaving the channel permanently deaf -
    /// deliveries stalled while publishers kept publishing. The product now guards
    /// the marker with a lock. This section hammers it with 4 x 50,000 distinct
    /// payloads (200,000 deliveries required within a bounded wait) and then
    /// probes the suppression path with 50,000 identical payloads (one delivery
    /// expected, a small allowance for races), finishing with a different payload
    /// that must still arrive.
    ///
    /// Publishes go through a direct ConnectionMultiplexer/ISubscriber against the
    /// dedicated container on 127.0.0.1:6396: started here when a Docker Linux
    /// daemon is available, reused when something already answers there, and
    /// falling back to the smoke's main host when Docker cannot provide a Linux
    /// container (so a Docker-less CI still exercises the concurrency path).
    /// </summary>
    private static void RunConcurrentDedupDeliveryTest(RedisSubscriptionManager subscriptions, string fallbackHost, bool skipRepeated)
    {
        const int publisherCount = 4;
        const int messagesPerPublisher = 50000;
        const int constantProbeCount = 50000;
        const long totalDistinct = (long)publisherCount * messagesPerPublisher;

        string channel = "smoke:dedup:" + Guid.NewGuid().ToString("N");
        string target = ConcurrentEndpoint;
        bool startedContainer = false;
        IDisposable token = null;
        ConnectionMultiplexer mux = null;

        Console.WriteLine("concurrent dedup: channel " + channel);
        try
        {
            if (IsRedisReachable(ConcurrentEndpoint))
            {
                Console.WriteLine("concurrent dedup: using the Redis already answering on " + ConcurrentEndpoint);
            }
            else
            {
                string dockerDetail = null;
                if (IsLinuxDockerAvailable(out dockerDetail))
                {
                    // docker run -d --rm --name rs-smoke-2 -p 6396:6379 redis:7-alpine
                    RunTool("docker", "rm -f " + ConcurrentContainerName, 30000, out _, out _);
                    int runExit;
                    string runOutput;
                    bool launched = RunTool("docker",
                        "run -d --rm --name " + ConcurrentContainerName + " -p 6396:6379 redis:7-alpine",
                        180000, out runExit, out runOutput);
                    startedContainer = launched && runExit == 0;
                    if (!startedContainer)
                    {
                        Check(false, "concurrent dedup: docker run " + ConcurrentContainerName +
                            " failed (" + Summarize(runOutput) + ")");
                        return;
                    }
                    if (!WaitForRedis(ConcurrentEndpoint, 30000))
                    {
                        Check(false, "concurrent dedup: container " + ConcurrentContainerName +
                            " did not answer PING on " + ConcurrentEndpoint + " within 30 s");
                        return;
                    }
                    Console.WriteLine("concurrent dedup: started container " + ConcurrentContainerName +
                        " on " + ConcurrentEndpoint + " (" + dockerDetail + ")");
                }
                else
                {
                    target = fallbackHost;
                    Console.WriteLine("concurrent dedup: Docker Linux daemon unavailable (" +
                        Summarize(dockerDetail) + "); using the main host " + fallbackHost);
                }
            }

            long delivered = 0;
            token = subscriptions.Subscribe(target, channel, pattern: false,
                onMessage: m => Interlocked.Increment(ref delivered));

            var options = ConfigurationOptions.Parse(target);
            options.AbortOnConnectFail = false;
            options.ConnectRetry = 1;
            options.ConnectTimeout = 5000;
            options.SyncTimeout = 60000;
            mux = ConnectionMultiplexer.Connect(options);
            var publisher = mux.GetSubscriber();
            var server = mux.GetServer(mux.GetEndPoints().First());
            var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);

            Func<long> numsub = () =>
            {
                var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", channel);
                return arr.Length >= 2 ? (long)arr[1] : 0;
            };
            Check(WaitUntil(() => numsub() == 1, 10000),
                "concurrent dedup: listener subscribed on the " + target + " server");

            // Phase 1: 4 publisher threads x 50,000 globally distinct payloads.
            // The per-thread prefix keeps every payload unique across threads, so
            // consecutive-duplicate suppression can never legitimately drop one.
            var phase1 = Stopwatch.StartNew();
            var threads = new Thread[publisherCount];
            for (int t = 0; t < publisherCount; t++)
            {
                int threadIndex = t;
                threads[t] = new Thread(() =>
                {
                    for (int i = 0; i < messagesPerPublisher; i++)
                        publisher.Publish(redisChannel, threadIndex + ":" + i, CommandFlags.FireAndForget);
                });
                threads[t].IsBackground = true;
                threads[t].Start();
            }
            bool publishersJoined = true;
            foreach (var thread in threads)
            {
                if (!thread.Join(60000))
                    publishersJoined = false;
            }

            // PING on the publisher connection is an ordering barrier: Redis
            // processes one connection's commands in order, so the reply proves
            // the server accepted all 200,000 publishes; a missing delivery is
            // then a consumer-side stall, not a slow publisher.
            string barrierError = null;
            try
            {
                mux.GetDatabase().Ping();
            }
            catch (Exception ex)
            {
                barrierError = ex.GetType().Name + ": " + ex.Message;
            }

            bool phase1Complete = WaitUntil(() => Interlocked.Read(ref delivered) >= totalDistinct, 60000);
            long phase1Delivered = Interlocked.Read(ref delivered);
            long phase1Ms = phase1.ElapsedMilliseconds;
            if (!phase1Complete)
            {
                Console.WriteLine("  concurrent dedup diagnostics (delivery stalled while publishing continued):");
                Console.WriteLine($"    issued:            {totalDistinct:N0} (4 threads x {messagesPerPublisher:N0} FireAndForget publishes)");
                Console.WriteLine($"    delivered:         {phase1Delivered:N0}");
                Console.WriteLine($"    elapsed:           {phase1Ms:N0} ms (bound 60000 ms)");
                Console.WriteLine($"    publishers joined: {publishersJoined}");
                if (barrierError != null)
                    Console.WriteLine("    barrier PING:      " + barrierError);
            }
            Check(phase1Complete,
                $"concurrent dedup: all {totalDistinct:N0} distinct payloads delivered under concurrent publish " +
                $"(got {phase1Delivered:N0} in {phase1Ms:N0} ms)");

            // Phase 2: 50,000 identical payloads must collapse to a single
            // delivery; a few racing copies are tolerated (best-effort dedup).
            long constantBase = Interlocked.Read(ref delivered);
            for (int i = 0; i < constantProbeCount; i++)
                publisher.Publish(redisChannel, "dedup-constant", CommandFlags.FireAndForget);

            if (skipRepeated)
            {
                bool constantSeen = WaitUntil(() => Interlocked.Read(ref delivered) > constantBase, 15000);
                if (!constantSeen)
                {
                    Console.WriteLine("  concurrent dedup diagnostics: the constant probe delivered nothing " +
                        "within 15 s (listener deaf?)");
                }
                Thread.Sleep(1000); // bounded window for racing duplicates to surface
                long constantDelta = Interlocked.Read(ref delivered) - constantBase;
                Check(constantSeen && constantDelta >= 1 && constantDelta <= 3,
                    $"concurrent dedup: constant-payload probe delivered once out of {constantProbeCount:N0} " +
                    $"identical payloads (got {constantDelta:N0}; <=3 allowed for racing duplicates)");
            }
            else
            {
                bool constantComplete = WaitUntil(
                    () => Interlocked.Read(ref delivered) - constantBase >= constantProbeCount, 30000);
                long constantDelta = Interlocked.Read(ref delivered) - constantBase;
                Check(constantComplete,
                    $"concurrent dedup: constant payloads all delivered (dedup disabled by config; " +
                    $"got {constantDelta:N0}/{constantProbeCount:N0})");
            }

            // Phase 3: the channel must still deliver a different payload.
            long afterBase = Interlocked.Read(ref delivered);
            publisher.Publish(redisChannel, "dedup-after");
            bool afterDelivered = WaitUntil(() => Interlocked.Read(ref delivered) > afterBase, 10000);
            long afterDelta = Interlocked.Read(ref delivered) - afterBase;
            Check(afterDelivered,
                $"concurrent dedup: stream still delivers after the constant-payload probe (got {afterDelta:N0} of 1)");
        }
        finally
        {
            if (token != null)
            {
                try { token.Dispose(); }
                catch (Exception ex) { Console.WriteLine("concurrent dedup: listener dispose failed: " + ex.Message); }
            }
            if (mux != null)
            {
                try { mux.Dispose(); }
                catch { }
            }
            if (startedContainer)
            {
                int stopExit;
                string stopOutput;
                bool stopped = RunTool("docker", "stop " + ConcurrentContainerName, 30000, out stopExit, out stopOutput);
                if (stopped && stopExit == 0)
                    Console.WriteLine("concurrent dedup: container " + ConcurrentContainerName + " stopped (--rm removes it)");
                else
                    Console.WriteLine("concurrent dedup: warning: could not stop container " + ConcurrentContainerName +
                        ": " + Summarize(stopOutput));
                RunTool("docker", "rm -f " + ConcurrentContainerName, 30000, out _, out _);
            }
        }
    }

    /// <summary>True when a PING round-trips to the endpoint (short timeouts).</summary>
    private static bool IsRedisReachable(string endpoint)
    {
        try
        {
            var options = ConfigurationOptions.Parse(endpoint);
            options.AbortOnConnectFail = false;
            options.ConnectRetry = 1;
            options.ConnectTimeout = 1500;
            options.SyncTimeout = 3000;
            using (var probe = ConnectionMultiplexer.Connect(options))
            {
                probe.GetDatabase().Ping();
                return true;
            }
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Bounded readiness wait for a Redis endpoint.</summary>
    private static bool WaitForRedis(string endpoint, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            if (IsRedisReachable(endpoint))
                return true;
            Thread.Sleep(500);
        }
        return IsRedisReachable(endpoint);
    }

    /// <summary>
    /// Docker can only provide the redis:7-alpine container when the CLI can
    /// reach a Linux daemon (Docker Desktop/WSL2 or a Linux host). A Windows
    /// container daemon or a missing/unreachable daemon reports unavailable
    /// and the caller falls back to the smoke's main host.
    /// </summary>
    private static bool IsLinuxDockerAvailable(out string detail)
    {
        detail = string.Empty;
        int exitCode;
        string output;
        if (!RunTool("docker", "version --format {{.Server.Os}}", 20000, out exitCode, out output) || exitCode != 0)
        {
            detail = output;
            return false;
        }
        string serverOs = output.Trim();
        detail = "docker server os=" + serverOs;
        return string.Equals(serverOs, "linux", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Runs a tool with a bounded wait, capturing merged stdout/stderr. Returns
    /// false on timeout or when the process could not be started at all.
    /// </summary>
    private static bool RunTool(string fileName, string arguments, int timeoutMs, out int exitCode, out string output)
    {
        exitCode = -1;
        output = string.Empty;
        try
        {
            var psi = new ProcessStartInfo(fileName, arguments)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using (var process = Process.Start(psi))
            {
                var stdout = new StringBuilder();
                var stderr = new StringBuilder();
                process.OutputDataReceived += (sender, args) => { if (args.Data != null) { lock (stdout) stdout.AppendLine(args.Data); } };
                process.ErrorDataReceived += (sender, args) => { if (args.Data != null) { lock (stderr) stderr.AppendLine(args.Data); } };
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                if (!process.WaitForExit(timeoutMs))
                {
                    try { process.Kill(); } catch { }
                    output = "timed out after " + timeoutMs + " ms";
                    return false;
                }
                process.WaitForExit(); // flush the async output callbacks
                lock (stdout) output = stdout.ToString();
                lock (stderr)
                {
                    string errorText = stderr.ToString();
                    if (errorText.Length > 0)
                        output = output.Length > 0 ? output + " " + errorText : errorText;
                }
                exitCode = process.ExitCode;
                return true;
            }
        }
        catch (Exception ex)
        {
            output = ex.GetType().Name + ": " + ex.Message;
            return false;
        }
    }

    /// <summary>One-line, length-capped tool output for diagnostics.</summary>
    private static string Summarize(string output)
    {
        if (string.IsNullOrWhiteSpace(output))
            return "no output";
        string text = output.Replace("\r", " ").Replace("\n", " ").Trim();
        while (text.Contains("  "))
            text = text.Replace("  ", " ");
        return text.Length <= 300 ? text : text.Substring(0, 300) + "...";
    }

    // Process-wide ChannelLatest state, reached through reflection: the manager
    // listener count proves a live subscription, the registry/cache counts also
    // catch ghost entries whose subscription is already gone.
    private static readonly FieldInfo UdfChannelListenersField =
        typeof(RedisUDF).GetField("_channelListeners", BindingFlags.NonPublic | BindingFlags.Static);
    private static readonly FieldInfo UdfLatestMessagesField =
        typeof(RedisUDF).GetField("_latestMessages", BindingFlags.NonPublic | BindingFlags.Static);

    /// <summary>UDF-origin listeners in the process-wide subscription manager.</summary>
    private static int UdfListenerCount() => RedisRuntime.Subscriptions.ListenerCountWithOrigin("UDF");

    /// <summary>Size of RedisUDF's own ChannelLatest registry (reflection).</summary>
    private static int UdfRegisteredListenerCount()
        => UdfChannelListenersField == null ? -1 : ((ICollection)UdfChannelListenersField.GetValue(null)).Count;

    /// <summary>Size of RedisUDF's cached latest messages (reflection).</summary>
    private static int UdfLatestMessageCount()
        => UdfLatestMessagesField == null ? -1 : ((ICollection)UdfLatestMessagesField.GetValue(null)).Count;

    /// <summary>Atomically raises <paramref name="target"/> to <paramref name="value"/>.</summary>
    private static void UpdateMax(ref long target, long value)
    {
        long current;
        while (value > (current = Interlocked.Read(ref target)))
        {
            if (Interlocked.CompareExchange(ref target, value, current) == current)
                return;
        }
    }

    /// <summary>
    /// Liveness regression for the per-key NetworkGate working-tree hardening:
    /// 4 threads x 300 cycles, every cycle subscribe -> publish a unique token
    /// -> wait for THIS listener -> dispose, all on one channel. Before the
    /// gate, a joiner attaching while the previous channel-state generation was
    /// being torn down could skip the wire SUBSCRIBE and leave the channel
    /// permanently deaf (about 20% of the tokens were missed with no error at
    /// all). Also samples the server: while the manager holds listeners,
    /// PUBSUB NUMSUB must be 1 within a short grace window (a listener is
    /// registered before its wire SUBSCRIBE completes, so a single fresh zero
    /// sample is not a violation by itself; a zero that outlives the window
    /// while listeners remain is the deaf-channel state).
    /// </summary>
    private static void RunConcurrentChurnTest(
        RedisSubscriptionManager subscriptions, ISubscriber publisher, IServer server, string host)
    {
        const int threadCount = 4;
        const int cyclesPerThread = 300;
        const int totalCycles = threadCount * cyclesPerThread;
        const int numsubGraceMs = 1000;
        string channel = "smoke:churn:" + Guid.NewGuid().ToString("N");
        string runId = Guid.NewGuid().ToString("N");
        var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
        Console.WriteLine("concurrent churn: channel " + channel);

        Func<long> numsub = () =>
        {
            var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", channel);
            return arr.Length >= 2 ? (long)arr[1] : 0;
        };

        int misses = 0;
        long worstDeliveryMs = 0;
        int numsubSamples = 0;
        int numsubViolations = 0;
        long worstZeroWindowMs = 0;
        bool churnDone = false;
        var start = new ManualResetEventSlim(false);
        var churnWatch = Stopwatch.StartNew();

        var workers = new Thread[threadCount];
        for (int t = 0; t < threadCount; t++)
        {
            int threadIndex = t;
            workers[t] = new Thread(() =>
            {
                start.Wait();
                for (int i = 0; i < cyclesPerThread; i++)
                {
                    string token = "churn-" + threadIndex + "-" + i + "-" + runId;
                    var delivered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    IDisposable tokenSubscription = null;
                    try
                    {
                        tokenSubscription = subscriptions.Subscribe(host, channel, pattern: false,
                            onMessage: message =>
                            {
                                if (string.Equals(message, token, StringComparison.Ordinal))
                                    delivered.TrySetResult(true);
                            });
                        long publishTicks = Stopwatch.GetTimestamp();
                        publisher.Publish(redisChannel, token); // blocking: the server accepted the publish
                        if (delivered.Task.Wait(5000))
                        {
                            long elapsedMs = (Stopwatch.GetTimestamp() - publishTicks) * 1000 / Stopwatch.Frequency;
                            UpdateMax(ref worstDeliveryMs, elapsedMs);
                        }
                        else
                        {
                            Interlocked.Increment(ref misses);
                            Console.WriteLine("  churn miss: thread " + threadIndex + " cycle " + i + " token " + token);
                        }
                    }
                    finally
                    {
                        tokenSubscription?.Dispose();
                    }
                }
            });
            workers[t].IsBackground = true;
        }

        var sampler = new Thread(() =>
        {
            while (!Volatile.Read(ref churnDone))
            {
                Thread.Sleep(10);
                if (subscriptions.ListenerCount <= 0)
                    continue;
                Interlocked.Increment(ref numsubSamples);
                if (numsub() == 1)
                    continue;
                var zero = Stopwatch.StartNew();
                bool reached = false;
                while (zero.ElapsedMilliseconds < numsubGraceMs && subscriptions.ListenerCount > 0)
                {
                    if (numsub() == 1)
                    {
                        reached = true;
                        break;
                    }
                    Thread.Sleep(5);
                }
                UpdateMax(ref worstZeroWindowMs, zero.ElapsedMilliseconds);
                if (!reached && subscriptions.ListenerCount > 0)
                {
                    Interlocked.Increment(ref numsubViolations);
                    Console.WriteLine("  churn NUMSUB violation: listeners=" + subscriptions.ListenerCount +
                        ", channel not subscribed on the server for " + zero.ElapsedMilliseconds + " ms");
                }
            }
        });
        sampler.IsBackground = true;

        for (int t = 0; t < threadCount; t++)
            workers[t].Start();
        sampler.Start();
        start.Set();
        foreach (var worker in workers)
            worker.Join();
        Volatile.Write(ref churnDone, true);
        sampler.Join();
        churnWatch.Stop();

        bool released = WaitUntil(() => subscriptions.ListenerCount == 0, 5000)
            && WaitUntil(() => numsub() == 0, 5000);
        Check(released, "churn: channel and listeners released after the churn");
        Check(misses == 0,
            $"churn: {totalCycles:N0} token deliveries while {threadCount} threads churned one channel, 0 misses (got {misses})");
        Check(worstDeliveryMs < 1000,
            $"churn: worst delivery {worstDeliveryMs} ms < 1000 ms ({totalCycles:N0} deliveries)");
        Check(numsubViolations == 0,
            $"churn: PUBSUB NUMSUB == 1 whenever listeners were live ({numsubSamples} samples, {numsubViolations} violations, " +
            $"worst zero window {worstZeroWindowMs} ms, grace {numsubGraceMs} ms)");
        Console.WriteLine($"concurrent churn: {threadCount} threads x {cyclesPerThread} cycles in {churnWatch.ElapsedMilliseconds} ms");
    }

    /// <summary>
    /// Regression for the ChannelLatest lifecycle epochs + per-listener sync
    /// (working tree): a concurrent unsubscribe used to lose the subscribe
    /// race every time - the removal found nothing to remove while the network
    /// SUBSCRIBE was in flight, and the returned listener stayed alive as a
    /// ghost entry. 400 rounds, each a racing ChannelLatest call and an
    /// Unsubscribe call (the unsubscribe runs after the listener showed up in
    /// the manager, so it deterministically overlaps the install decision);
    /// after both join, the manager and the registry must hold no UDF listener.
    /// A fresh ChannelLatest must then subscribe and receive a new publish.
    /// </summary>
    private static void RunChannelLatestRaceTest(ISubscriber publisher, string host)
    {
        const int rounds = 400;
        string channel = "smoke:latest:" + Guid.NewGuid().ToString("N");
        var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
        Console.WriteLine("channel-latest race: channel " + channel);

        int badRounds = 0;
        long worstRoundMs = 0;
        var raceWatch = Stopwatch.StartNew();
        for (int round = 0; round < rounds; round++)
        {
            // Never let a leftover from a previous round pollute the next race.
            if (UdfListenerCount() != 0 || UdfRegisteredListenerCount() != 0)
                RedisUDF.RedisUDFChannelUnsubscribe(channel, host);

            var roundWatch = Stopwatch.StartNew();
            var latest = Task.Run(() => RedisUDF.RedisUDFChannelLatest(channel, host));
            // The manager listener appears after the lifecycle check snapshotted
            // the epoch, so this Unsubscribe always races the install decision.
            bool observed = SpinWait.SpinUntil(() => UdfListenerCount() > 0, 5000);
            Thread.SpinWait(1000);
            RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
            bool joined = latest.Wait(20000);
            long roundMs = roundWatch.ElapsedMilliseconds;
            if (roundMs > worstRoundMs)
                worstRoundMs = roundMs;

            int managerCount = UdfListenerCount();
            int registryCount = UdfRegisteredListenerCount();
            if (!observed || !joined || managerCount != 0 || registryCount != 0)
            {
                badRounds++;
                Console.WriteLine("  channel-latest race: round " + round + " bad (observed=" + observed +
                    ", joined=" + joined + ", manager=" + managerCount + ", registry=" + registryCount +
                    ", result='" + (joined ? latest.Result : "<not joined>") + "')");
            }
        }
        raceWatch.Stop();

        Check(badRounds == 0,
            $"channel-latest race: 0 of {rounds} racing subscribe/unsubscribe rounds left a listener " +
            $"(manager 0, registry 0; worst round {worstRoundMs} ms, total {raceWatch.ElapsedMilliseconds} ms)");

        string fresh = RedisUDF.RedisUDFChannelLatest(channel, host);
        Check(string.Equals(fresh, "(null)", StringComparison.Ordinal),
            $"channel-latest race: a fresh ChannelLatest after the race starts empty (got '{fresh}')");
        Check(UdfListenerCount() == 1 && UdfRegisteredListenerCount() == 1,
            "channel-latest race: the fresh ChannelLatest subscribed exactly one listener");

        string probe = "post-race-" + Guid.NewGuid().ToString("N");
        publisher.Publish(redisChannel, probe);
        bool delivered = SpinWait.SpinUntil(
            () => string.Equals(RedisUDF.RedisUDFChannelLatest(channel, host), probe, StringComparison.Ordinal), 5000);
        Check(delivered, "channel-latest race: the fresh listener receives a new publish");

        RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
        Check(UdfListenerCount() == 0 && UdfRegisteredListenerCount() == 0,
            "channel-latest race: unsubscribe after the fresh read leaves no listener");
    }

    /// <summary>
    /// Ghost/stale regression (working tree): a message callback that already
    /// passed the closed check used to be able to write the shared latest-message
    /// cache after the unsubscribe had cleared it; the per-listener Sync now
    /// serializes the callback write against Close+remove. ChannelLatest churns
    /// (subscribe -> unsubscribe) while a publisher spams the channel; after the
    /// publisher stops and the final unsubscribe returns, a fresh ChannelLatest
    /// must not surface any old message (no ghost re-add) until a new publish,
    /// which the fresh subscription must deliver.
    /// </summary>
    private static void RunGhostStaleTest(ConnectionMultiplexer publisherConnection, ISubscriber publisher, string host)
    {
        const int churnRounds = 300;
        string channel = "smoke:ghost:" + Guid.NewGuid().ToString("N");
        var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
        Console.WriteLine("ghost/stale: channel " + channel);

        bool stopPublishing = false;
        long published = 0;
        var publisherThread = new Thread(() =>
        {
            long i = 0;
            int burst = 0;
            while (!Volatile.Read(ref stopPublishing))
            {
                publisher.Publish(redisChannel, "ghost-churn-" + i++, CommandFlags.FireAndForget);
                Interlocked.Increment(ref published);
                if (++burst >= 32)
                {
                    burst = 0;
                    Thread.Sleep(1);
                }
            }
        });
        publisherThread.IsBackground = true;

        var churnThread = new Thread(() =>
        {
            for (int round = 0; round < churnRounds; round++)
            {
                RedisUDF.RedisUDFChannelLatest(channel, host);
                Thread.SpinWait(500);
                RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
            }
        });
        churnThread.IsBackground = true;

        var watch = Stopwatch.StartNew();
        churnThread.Start();
        publisherThread.Start();
        churnThread.Join();
        Volatile.Write(ref stopPublishing, true);
        publisherThread.Join();
        // Ordering barrier: once PING replies on the publisher connection the
        // server processed every publish; the drain wait lets those deliveries
        // reach the current listener before the final unsubscribe.
        publisherConnection.GetDatabase().Ping();
        Thread.Sleep(750);
        watch.Stop();

        RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
        Check(WaitUntil(() => UdfListenerCount() == 0, 5000) && UdfRegisteredListenerCount() == 0,
            "ghost/stale: the final unsubscribe leaves no listener and no registry entry");
        Thread.Sleep(200); // a ghost callback would have to fire in this window
        Check(UdfLatestMessageCount() == 0,
            $"ghost/stale: no stale latest-message cache entry after the churn (got {UdfLatestMessageCount()})");

        string fresh = RedisUDF.RedisUDFChannelLatest(channel, host);
        Check(string.Equals(fresh, "(null)", StringComparison.Ordinal),
            $"ghost/stale: a fresh ChannelLatest after the churn returns no ghost message (got '{fresh}')");

        string probe = "ghost-probe-" + Guid.NewGuid().ToString("N");
        publisher.Publish(redisChannel, probe);
        bool delivered = SpinWait.SpinUntil(
            () => string.Equals(RedisUDF.RedisUDFChannelLatest(channel, host), probe, StringComparison.Ordinal), 5000);
        Check(delivered, "ghost/stale: a new publish is delivered to the fresh subscription");

        RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
        Check(WaitUntil(() => UdfListenerCount() == 0, 5000) && UdfRegisteredListenerCount() == 0,
            "ghost/stale: cleanup leaves no listener");
        Console.WriteLine($"ghost/stale: {churnRounds} churn rounds, {Interlocked.Read(ref published):N0} publishes in {watch.ElapsedMilliseconds} ms");
    }

    /// <summary>
    /// Regression for the pattern-join marker purge (working tree): the local
    /// glob matcher could under-clear (reversed ranges, unterminated classes),
    /// so ClearForListener now clears every publish-if-changed marker of that
    /// host on a pattern join. A PSUB listener on a pattern that does NOT match
    /// the channel must still invalidate the marker: the next identical
    /// PublishIfChanged is published again instead of answering "No change".
    /// A raw direct subscriber (no manager dedup) proves the republish really
    /// reached the server.
    /// </summary>
    private static void RunPatternJoinOverClearTest(
        RedisSubscriptionManager subscriptions, ISubscriber publisher, IServer server, string host)
    {
        string runId = Guid.NewGuid().ToString("N");
        string channel = "smoke:overclear:" + runId + ":real";
        string pattern = "smoke:overclear-nomatch:" + runId + ":*";
        const string payload = "overclear-payload";
        Console.WriteLine("pattern-join over-clear: channel " + channel + ", pattern " + pattern);

        var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
        int rawDelivered = 0;
        Action<RedisChannel, RedisValue> rawHandler = (_, __) => Interlocked.Increment(ref rawDelivered);
        publisher.Subscribe(redisChannel, rawHandler);

        int listenerDelivered = 0;
        var listenerToken = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: _ => Interlocked.Increment(ref listenerDelivered), origin: "smokeOverClear");

        Func<long> numsub = () =>
        {
            var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", channel);
            return arr.Length >= 2 ? (long)arr[1] : 0;
        };

        try
        {
            Check(WaitUntil(() => numsub() == 2, 5000),
                "pattern-join over-clear: literal listener and raw subscriber live on the server");
            RedisUDF.SyncWriteOverrideForTests = "sync"; // deterministic readers count, marker stored only when delivered

            string first = RedisUDF.RedisUDFChannelPublishIfChanged(channel, payload, host)?.ToString();
            Check(first != "No change" && first != null && !first.StartsWith("Error:", StringComparison.Ordinal),
                $"pattern-join over-clear: first publish delivered (result '{first}')");
            Check(WaitUntil(() => rawDelivered == 1 && listenerDelivered == 1, 5000),
                "pattern-join over-clear: the first publish reached both subscribers");

            string second = RedisUDF.RedisUDFChannelPublishIfChanged(channel, payload, host)?.ToString();
            Check(string.Equals(second, "No change", StringComparison.Ordinal),
                $"pattern-join over-clear: identical republish suppressed before any join (result '{second}')");
            Check(WaitUntil(() => rawDelivered == 1 && listenerDelivered == 1, 750),
                "pattern-join over-clear: the suppressed republish was not sent to the server");

            // A pattern that cannot match the channel; its join still clears the
            // host's markers synchronously inside Subscribe (over-clearing is
            // intentional: under-clearing starves a returning listener).
            var patternToken = subscriptions.Subscribe(host, pattern, pattern: true,
                onMessage: _ => { }, origin: "smokeOverClearPattern");

            string third = RedisUDF.RedisUDFChannelPublishIfChanged(channel, payload, host)?.ToString();
            Check(third != "No change" && third != null && !third.StartsWith("Error:", StringComparison.Ordinal),
                $"pattern-join over-clear: identical republish after the non-matching pattern join was not suppressed (result '{third}')");
            // The raw subscriber sees every publish; the literal manager listener
            // may still dedup the identical payload (its own channel-state marker
            // is untouched by a pattern join), which is orthogonal by design.
            Check(WaitUntil(() => rawDelivered == 2, 5000),
                "pattern-join over-clear: the republished payload reached the server (raw subscriber count 2)");

            patternToken.Dispose();
        }
        finally
        {
            RedisUDF.SyncWriteOverrideForTests = null;
            listenerToken.Dispose();
            publisher.Unsubscribe(redisChannel, rawHandler);
        }

        Check(WaitUntil(() => numsub() == 0, 5000), "pattern-join over-clear: channel released after cleanup");
    }
}

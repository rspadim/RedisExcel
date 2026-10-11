using RedisExcel;
using StackExchange.Redis;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
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

    /// <summary>
    /// Waits until <paramref name="count"/> stays equal to <paramref name="expected"/>
    /// for a settle window (so a value still arriving in the immediate wake of a
    /// publish cannot pass) and returns the observed stable count, or the last
    /// observed count on timeout. Replaces fixed sleeps between a publish and a
    /// delivery-count assertion.
    /// </summary>
    private static int WaitForCountStable(Func<int> count, int expected, int settleMs, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        long stableSince = -1;
        int last = -1;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            last = count();
            if (last == expected)
            {
                if (stableSince < 0) stableSince = sw.ElapsedMilliseconds;
                if (sw.ElapsedMilliseconds - stableSince >= settleMs)
                    return last;
            }
            else
            {
                stableSince = -1;
            }
            Thread.Sleep(20);
        }
        return last;
    }

    /// <summary>
    /// Waits until <paramref name="value"/> stops changing for a settle window and
    /// returns the last observed value, or the last observed value on timeout.
    /// Used after a burst of fire-and-forget publishes so a racing delivery that
    /// is still in flight cannot slip past a fixed sleep (seed with a negative
    /// value so "no delivery yet" is a valid stable observation).
    /// </summary>
    private static long WaitForLongStable(Func<long> value, int settleMs, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        long stableSince = -1;
        long last = -2;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            long current = value();
            if (current == last)
            {
                if (stableSince < 0) stableSince = sw.ElapsedMilliseconds;
                if (sw.ElapsedMilliseconds - stableSince >= settleMs)
                    return current;
            }
            else
            {
                last = current;
                stableSince = -1;
            }
            Thread.Sleep(20);
        }
        return last;
    }

    /// <summary>Compact exception description ("Type: message") for a FAIL label.</summary>
    private static string DescribeException(Exception ex)
    {
        return ex == null ? "no exception" : ex.GetType().Name + ": " + ex.Message;
    }

    /// <summary>PUBSUB NUMSUB for one channel: the subscriber count (0 when the
    /// reply is missing the count element).</summary>
    private static long NumSub(IServer server, string channel)
    {
        var arr = (RedisResult[])server.Execute("PUBSUB", "NUMSUB", channel);
        return arr.Length >= 2 ? (long)arr[1] : 0;
    }

    private static int Main(string[] args)
    {
        string host = args.Length > 0 ? args[0] : DefaultHost;
        string channel = "smoke:" + Guid.NewGuid().ToString("N");

        Console.WriteLine($"Redis host: {host}");
        Console.WriteLine($"Channel:    {channel}");

        // Force SkipRepeatedMessages deterministically (test seam): the
        // duplicate-suppression assertions below must always exercise the
        // suppression branch instead of adapting to whatever RedisExcel.json a
        // machine happens to have. The flag is captured when the manager is
        // constructed, so set the override first.
        RedisSubscriptionManager.SkipRepeatedOverrideForTests = true;
        bool skipRepeated = RedisSubscriptionManager.SkipRepeatedEnabledForTests;

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

        Func<long> numsub = () => NumSub(server, channel);

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

        // Duplicate suppression: with the test seam forcing SkipRepeatedMessages
        // on, identical consecutive payloads must be skipped (the first arrives,
        // the second is suppressed, the changed third arrives).
        Check(skipRepeated, "duplicate-suppression branch pinned on by the test seam");
        var receivedD = new List<string>();
        var tokenD = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedD.Add(m); });
        Check(WaitUntil(() => numsub() == 1, 5000), "channel active for the duplicate test");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup2");
        Check(WaitUntil(() => { lock (Sync) return receivedD.Contains("dup2"); }, 5000), "changed payload delivered");
        // Bounded settle (count stable for 250 ms) instead of a fixed sleep:
        // the assertion then reads "exactly 2", never a race against a late
        // suppressed duplicate.
        int dupCount = WaitForCountStable(() => { lock (Sync) return receivedD.Count; }, 2, 250, 5000);
        Check(dupCount == 2, $"identical repeated payload skipped (received {dupCount}, expected 2: \"dup\" + \"dup2\")");
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

        Func<string, long> numsubOf = ch => NumSub(server, ch);

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
            Check(false, ex.Message);
        }
        Check(subscriptions.ChannelCount == channelsBefore && subscriptions.ListenerCount == listenersBefore,
            "counters unchanged by the second Dispose");

        // Invalid arguments are rejected before any Redis call, with the
        // SPECIFIC ArgumentException (not merely "some exception was thrown").
        try
        {
            subscriptions.Subscribe("", "c", false, m => { });
            Check(false, "blank host accepted");
        }
        catch (Exception ex)
        {
            Check(ex is ArgumentException hostEx && hostEx.ParamName == "host"
                    && hostEx.Message.StartsWith("host is required", StringComparison.Ordinal),
                "blank host rejected with the specific ArgumentException (" + DescribeException(ex) + ")");
        }
        try
        {
            subscriptions.Subscribe(host, "", false, m => { });
            Check(false, "blank channel accepted");
        }
        catch (Exception ex)
        {
            Check(ex is ArgumentException channelEx && channelEx.ParamName == "channel"
                    && channelEx.Message.StartsWith("channel is required", StringComparison.Ordinal),
                "blank channel rejected with the specific ArgumentException (" + DescribeException(ex) + ")");
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
            RunChannelLatestConvergenceTest(publisher, server, host);
        }
        finally
        {
            RedisUdfAsync.AsyncWritesOverrideForTests = null;
            RedisUDF.SyncWriteOverrideForTests = null;
        }

        // UDF surface coverage for the functions the safety scout found untested:
        // HashGetAll (matrix shape/values), ServerTime (plausible epoch) and
        // PubSubChannelsInfo (a subscribed channel with its subscriber count).
        RunUdfFunctionsSmokeTest(connections, subscriptions, server, host);

        // Real write through a ...NonVolatile twin (value round-trip + one
        // delivery to a live channel listener), no Excel involved.
        RunNonVolatileWriteSmokeTest(subscriptions, server, host);

        // Live GitHub release check (the v1.4.1 release this repo just cut):
        // RedisUDFUpdateAvailable must stay non-blocking and agree with the real
        // latest release tag. Skipped (with a note) when the network/rate limit
        // makes the answer unavailable, so CI never turns red offline.
        RunUpdateAvailableSmokeTest();

        // Concurrency regression for the duplicate-suppression marker, kept last
        // because it hammers a dedicated Redis (throwaway container on a random
        // free port); it creates and tears down its own listener, so the counters
        // above are unaffected.
        RunConcurrentDedupDeliveryTest(subscriptions, host, publisher, server, skipRepeated);

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
    /// probes the suppression path with 50,000 identical payloads (exactly one
    /// delivery expected), finishing with a different payload that must still
    /// arrive.
    ///
    /// A throwaway redis:7-alpine container is used when a Docker Linux daemon is
    /// available: the container name is randomized and bound to a free ephemeral
    /// port (no fixed name/port that a leftover container or a parallel run could
    /// collide with). A failed `docker run`/readiness/connect is a FALLBACK to the
    /// smoke's main host, never a hard failure, so a Docker-less CI still
    /// exercises the concurrency path.
    /// </summary>
    private static void RunConcurrentDedupDeliveryTest(
        RedisSubscriptionManager subscriptions, string fallbackHost,
        ISubscriber mainPublisher, IServer mainServer, bool skipRepeated)
    {
        const int publisherCount = 4;
        const int messagesPerPublisher = 50000;
        const int constantProbeCount = 50000;
        const long totalDistinct = (long)publisherCount * messagesPerPublisher;

        string channel = "smoke:dedup:" + Guid.NewGuid().ToString("N");
        string containerName = "rs-smoke-" + Guid.NewGuid().ToString("N").Substring(0, 12);
        string endpoint = null;
        bool containerStarted = false; // docker run accepted (teardown in finally)
        bool useContainer = false;     // container connected and usable
        IDisposable token = null;
        ConnectionMultiplexer mux = null;
        ISubscriber publisher = null;
        IServer server = null;

        Console.WriteLine("concurrent dedup: channel " + channel);
        try
        {
            string dockerDetail = null;
            if (IsLinuxDockerAvailable(out dockerDetail))
            {
                int port = FindFreePort();
                if (port > 0)
                {
                    endpoint = "127.0.0.1:" + port;
                    RunTool("docker", "rm -f " + containerName, 30000, out _, out _);
                    int runExit;
                    string runOutput;
                    bool launched = RunTool("docker",
                        "run -d --rm --name " + containerName + " -p " + port + ":6379 redis:7-alpine",
                        180000, out runExit, out runOutput);
                    containerStarted = launched && runExit == 0;
                    if (containerStarted && WaitForRedis(endpoint, 30000))
                    {
                        try
                        {
                            var options = ConfigurationOptions.Parse(endpoint);
                            options.AbortOnConnectFail = false;
                            options.ConnectRetry = 1;
                            options.ConnectTimeout = 5000;
                            options.SyncTimeout = 60000;
                            mux = ConnectionMultiplexer.Connect(options);
                            publisher = mux.GetSubscriber();
                            server = mux.GetServer(mux.GetEndPoints().First());
                            useContainer = true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("concurrent dedup: container connect failed (" + ex.Message +
                                "); falling back to " + fallbackHost);
                            try { mux?.Dispose(); } catch { }
                            mux = null;
                            publisher = null;
                            server = null;
                        }
                    }
                    else
                    {
                        Console.WriteLine("concurrent dedup: container " + containerName +
                            " was not ready (" + Summarize(runOutput) + "); falling back to " + fallbackHost);
                    }
                }
                else
                {
                    Console.WriteLine("concurrent dedup: no free local port found; falling back to " + fallbackHost);
                }
            }
            else
            {
                Console.WriteLine("concurrent dedup: Docker Linux daemon unavailable (" +
                    Summarize(dockerDetail) + "); falling back to " + fallbackHost);
            }

            if (!useContainer)
            {
                publisher = mainPublisher;
                server = mainServer;
                Console.WriteLine("concurrent dedup: using the main host " + fallbackHost);
            }
            else
            {
                Console.WriteLine("concurrent dedup: using disposable container " + containerName +
                    " on " + endpoint);
            }

            string target = useContainer ? endpoint : fallbackHost;
            long delivered = 0;
            token = subscriptions.Subscribe(target, channel, pattern: false,
                onMessage: m => Interlocked.Increment(ref delivered));

            var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);

            Func<long> numsub = () => NumSub(server, channel);
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
            // delivery. Handlers run on pool threads with no ordering guarantee,
            // so a join/racing pair CAN legitimately observe a second copy of the
            // same payload once the first set the marker: the assertion keeps the
            // real expectation (exactly one) while tolerating at most a couple of
            // documented racing duplicates, and the follow-up check below proves
            // the stream is not deaf.
            Check(skipRepeated, "concurrent dedup: SkipRepeatedMessages pinned on for the constant probe");
            long constantBase = Interlocked.Read(ref delivered);
            for (int i = 0; i < constantProbeCount; i++)
                publisher.Publish(redisChannel, "dedup-constant", CommandFlags.FireAndForget);

            bool constantSeen = WaitUntil(() => Interlocked.Read(ref delivered) > constantBase, 15000);
            if (!constantSeen)
            {
                Console.WriteLine("  concurrent dedup diagnostics: the constant probe delivered nothing " +
                    "within 15 s (listener deaf?)");
            }
            // Bounded settle: the delta must stop growing before it is asserted,
            // so a burst of racing deliveries cannot slip past a fixed sleep.
            long constantDelta = WaitForLongStable(() => Interlocked.Read(ref delivered) - constantBase, 1000, 15000);
            Check(constantSeen && constantDelta >= 1 && constantDelta <= 3,
                $"concurrent dedup: constant-payload probe delivered once out of {constantProbeCount:N0} " +
                $"identical payloads (got {constantDelta:N0}; <=3 allowed for documented racing duplicates)");

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
            if (containerStarted)
            {
                int stopExit;
                string stopOutput;
                bool stopped = RunTool("docker", "stop " + containerName, 30000, out stopExit, out stopOutput);
                if (stopped && stopExit == 0)
                    Console.WriteLine("concurrent dedup: container " + containerName + " stopped (--rm removes it)");
                else
                    Console.WriteLine("concurrent dedup: warning: could not stop container " + containerName +
                        ": " + Summarize(stopOutput));
                RunTool("docker", "rm -f " + containerName, 30000, out _, out _);
            }
        }
    }

    /// <summary>A free local TCP port for the disposable container, or -1 when
    /// none could be bound (the caller then falls back to the main host).</summary>
    private static int FindFreePort()
    {
        try
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
        catch
        {
            return -1;
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

        Func<long> numsub = () => NumSub(server, channel);

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

        Func<long> numsub = () => NumSub(server, channel);

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

    /// <summary>
    /// Convergence regression for the StackExchange.Redis out-of-order delivery.
    /// SE.Redis hands every channel callback to the thread pool, so delivery
    /// order is NOT guaranteed: two messages in flight together can complete out
    /// of order. The requirement is weaker than strict ordering - a later
    /// publish must never be permanently suppressed or ignored by the client.
    /// The stream converges once the FINAL publish is the last one delivered.
    ///
    /// Phase 1 publishes a fast burst of ~500 DISTINCT payloads on a literal
    /// channel with one ChannelLatest listener (all distinct, so the identical
    /// consecutive-payload dedup can never drop one). Phase 2 then republishes
    /// "FINAL-&lt;guid&gt;" up to ~50 times over about 5 s while polling
    /// ChannelLatest until it reads that FINAL value: the FINAL value must be
    /// observed at least once, proving the channel was not left deaf and no
    /// later publish was permanently ignored. Strict ordering is deliberately
    /// NOT asserted (the library does not guarantee it).
    /// </summary>
    private static void RunChannelLatestConvergenceTest(ISubscriber publisher, IServer server, string host)
    {
        const int burstCount = 500;
        const int finalAttempts = 50;
        string runId = Guid.NewGuid().ToString("N");
        string channel = "smoke:converge:" + runId;
        var redisChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
        Console.WriteLine("convergence: channel " + channel);

        Func<long> numsub = () => NumSub(server, channel);

        // One ChannelLatest listener subscribes here and stays live for both phases.
        string initial = RedisUDF.RedisUDFChannelLatest(channel, host);
        Check(string.Equals(initial, "(null)", StringComparison.Ordinal),
            $"convergence: the fresh ChannelLatest starts empty (got '{initial}')");
        Check(WaitUntil(() => numsub() == 1, 5000),
            "convergence: ChannelLatest listener subscribed on the server");

        try
        {
            // Phase 1: a fast burst of distinct payloads. Every payload is unique,
            // so no delivery may be dropped by the consecutive-duplicate dedup.
            var burstBase = Stopwatch.StartNew();
            for (int i = 0; i < burstCount; i++)
                publisher.Publish(redisChannel, "burst-" + runId + "-" + i, CommandFlags.FireAndForget);

            // Ordering barrier: PING on the publisher connection (the same
            // multiplexer that produced `server`) proves Redis accepted every
            // publish; a missing burst payload is then a consumer problem, not
            // a slow publisher. Bounded: delivery is async.
            server.Ping();
            // Deliberately NO strict-order assertion: with out-of-order delivery
            // the last published burst payload may be overwritten by an earlier
            // one that arrived late. Any burst payload observed proves the
            // stream is delivering.
            string burstPrefix = "burst-" + runId + "-";
            bool burstDelivered = WaitUntil(
                () =>
                {
                    string latest = RedisUDF.RedisUDFChannelLatest(channel, host);
                    return latest != null && latest.StartsWith(burstPrefix, StringComparison.Ordinal);
                },
                15000);
            Check(burstDelivered,
                $"convergence: a distinct burst payload is observed (stream delivering, {burstBase.ElapsedMilliseconds} ms)");

            // Phase 2: republish FINAL repeatedly; the FINAL value must be
            // observed at least once. Strict ordering is NOT asserted.
            string finalPayload = "FINAL-" + runId;
            bool finalObserved = false;
            int finalTries = 0;
            var finalWatch = Stopwatch.StartNew();
            for (int attempt = 0; attempt < finalAttempts && !finalObserved; attempt++)
            {
                finalTries = attempt + 1;
                publisher.Publish(redisChannel, finalPayload); // blocking: the server accepted it
                if (WaitUntil(
                        () => string.Equals(RedisUDF.RedisUDFChannelLatest(channel, host), finalPayload, StringComparison.Ordinal),
                        100))
                {
                    finalObserved = true;
                }
            }
            finalWatch.Stop();
            Check(finalObserved,
                $"convergence: the FINAL payload is observed at least once ({finalTries}/{finalAttempts} tries in {finalWatch.ElapsedMilliseconds} ms)");

            // The stream is still alive after convergence: one more distinct
            // payload must be delivered, proving no permanent stall.
            string afterPayload = "after-" + runId;
            publisher.Publish(redisChannel, afterPayload);
            bool stillDelivering = WaitUntil(
                () => string.Equals(RedisUDF.RedisUDFChannelLatest(channel, host), afterPayload, StringComparison.Ordinal),
                5000);
            Check(stillDelivering, "convergence: the stream keeps delivering after the FINAL payload");
        }
        finally
        {
            RedisUDF.RedisUDFChannelUnsubscribe(channel, host);
        }

        Check(WaitUntil(() => numsub() == 0, 5000), "convergence: channel released after cleanup");
    }

    /// <summary>
    /// UDF surface coverage for the functions the safety scout found untested:
    /// RedisUDFHashGetAll (rows are field/value and HGETALL values round-trip),
    /// RedisUDFServerTime (a plausible epoch, within a minute of local UTC) and
    /// RedisUDFPubSubChannelsInfo (a subscribed channel listed with its
    /// subscriber count). Uses the smoke's main server and its own local
    /// listener, so the process-wide UDF counters are left untouched.
    /// </summary>
    private static void RunUdfFunctionsSmokeTest(
        RedisConnectionManager connections, RedisSubscriptionManager subscriptions, IServer server, string host)
    {
        string runId = Guid.NewGuid().ToString("N");
        var db = connections.GetDatabase(host, RedisPool.UdfData);

        // RedisUDFHashGetAll: a 3-field hash must come back as a 3-row,
        // 2-column (field, value) matrix with every value round-tripped.
        string hashKey = "smoke:hash:" + runId;
        try
        {
            db.HashSet(hashKey, new HashEntry[]
            {
                new HashEntry("fieldA", "value-1"),
                new HashEntry("fieldB", "value-2"),
                new HashEntry("fieldC", "42")
            });
            var matrix = RedisUDF.RedisUDFHashGetAll(hashKey, host);
            bool matrixOk = matrix != null && matrix.GetLength(0) == 3 && matrix.GetLength(1) == 2;
            var seen = new Dictionary<string, string>(StringComparer.Ordinal);
            if (matrixOk)
            {
                for (int r = 0; r < matrix.GetLength(0); r++)
                    seen[Convert.ToString(matrix[r, 0], CultureInfo.InvariantCulture)] =
                        Convert.ToString(matrix[r, 1], CultureInfo.InvariantCulture);
            }
            Check(matrixOk
                    && seen.TryGetValue("fieldA", out var a) && a == "value-1"
                    && seen.TryGetValue("fieldB", out var b) && b == "value-2"
                    && seen.TryGetValue("fieldC", out var c) && c == "42",
                "HashGetAll: 3-field hash returns a 3x2 field/value matrix with the expected values");

            // An empty/missing hash returns the single empty-cell sentinel.
            var empty = RedisUDF.RedisUDFHashGetAll("smoke:missing:" + runId, host);
            Check(empty != null && empty.GetLength(0) == 1 && empty.GetLength(1) == 1
                    && string.IsNullOrEmpty(Convert.ToString(empty[0, 0], CultureInfo.InvariantCulture)),
                "HashGetAll: a missing hash returns the empty sentinel");
        }
        finally
        {
            db.KeyDelete(hashKey);
        }

        // RedisUDFServerTime: an ISO-8601 string within 60 s of local UTC.
        object timeResult = RedisUDF.RedisUDFServerTime(host);
        bool timeParsed = DateTime.TryParse(
            Convert.ToString(timeResult, CultureInfo.InvariantCulture),
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal,
            out var serverTime);
        bool timeOk = timeParsed && Math.Abs((serverTime - DateTime.UtcNow).TotalSeconds) < 60;
        Check(timeOk, $"ServerTime: returns a plausible epoch/time ('{timeResult}')");

        // RedisUDFPubSubChannelsInfo: subscribe one listener on a unique channel;
        // PUBSUB CHANNELS + NUMSUB must list it with at least one subscriber.
        string infoChannel = "smoke:info:" + runId;
        var infoToken = subscriptions.Subscribe(host, infoChannel, pattern: false, onMessage: _ => { });
        try
        {
            Check(WaitUntil(() => NumSub(server, infoChannel) == 1, 5000),
                "PubSubChannelsInfo: listener subscribed on the server");

            var info = RedisUDF.RedisUDFPubSubChannelsInfo(host);
            bool headerOk = info != null && info.GetLength(1) == 2
                && string.Equals(Convert.ToString(info[0, 0], CultureInfo.InvariantCulture), "Channel", StringComparison.Ordinal)
                && string.Equals(Convert.ToString(info[0, 1], CultureInfo.InvariantCulture), "Subscribers", StringComparison.Ordinal);
            bool channelListed = false;
            long infoSubscribers = -1;
            if (info != null)
            {
                for (int r = 1; r < info.GetLength(0); r++)
                {
                    if (string.Equals(
                            Convert.ToString(info[r, 0], CultureInfo.InvariantCulture), infoChannel, StringComparison.Ordinal))
                    {
                        channelListed = true;
                        infoSubscribers = Convert.ToInt64(info[r, 1], CultureInfo.InvariantCulture);
                        break;
                    }
                }
            }
            Check(headerOk, "PubSubChannelsInfo: returns the Channel/Subscribers header row");
            Check(channelListed && infoSubscribers >= 1,
                $"PubSubChannelsInfo: lists the subscribed channel with its subscriber count (got {infoSubscribers})");
        }
        finally
        {
            infoToken.Dispose();
        }
        Check(WaitUntil(() => NumSub(server, infoChannel) == 0, 5000),
            "PubSubChannelsInfo: channel released after cleanup");
    }

    /// <summary>
    /// Smoke-level real write through a ...NonVolatile twin (no Excel): the twin
    /// must delegate to the base write (value round-trips through Redis) and the
    /// call must happen exactly once per invocation. Its write also reaches a
    /// live channel listener with a single delivery. The E2E workbook covers the
    /// "entry-only, recalculation does not re-run" contract; this pins the
    /// actual write value/delivery, which the offline unit tests cannot.
    /// </summary>
    private static void RunNonVolatileWriteSmokeTest(
        RedisSubscriptionManager subscriptions, IServer server, string host)
    {
        string runId = Guid.NewGuid().ToString("N");
        string key = "smoke:nv:" + runId;
        string incrKey = "smoke:nvincr:" + runId;
        string channel = "smoke:nvchan:" + runId;
        var connections = new RedisConnectionManager();
        var db = connections.GetDatabase(host, RedisPool.UdfData);
        var published = new List<string>();
        IDisposable token = null;
        // Pin the write path deterministically: the synchronous path returns the
        // real reply, so the twin value assertions do not depend on a machine's
        // SyncWrite/AsyncWrites configuration.
        string prevSync = RedisUDF.SyncWriteOverrideForTests;
        bool? prevAsync = RedisUdfAsync.AsyncWritesOverrideForTests;
        RedisUDF.SyncWriteOverrideForTests = "sync";
        RedisUdfAsync.AsyncWritesOverrideForTests = false;
        try
        {
            // SetNonVolatile: delegates to Set; the value must land in Redis.
            object setResult = RedisUDF.RedisUDFSetNonVolatile(key, "nv-value", host);
            Check(string.Equals(Convert.ToString(setResult, CultureInfo.InvariantCulture), "OK", StringComparison.Ordinal),
                $"NonVolatile twin: SetNonVolatile returns the write ack (got '{setResult}')");
            Check(string.Equals(db.StringGet(key), "nv-value", StringComparison.Ordinal),
                "NonVolatile twin: SetNonVolatile actually wrote the value to Redis");

            // IncrNonVolatile on a fresh key must be exactly one: a twin that
            // double-delegated would leave 2.
            object incrResult = RedisUDF.RedisUDFIncrNonVolatile(incrKey, host);
            Check(string.Equals(Convert.ToString(incrResult, CultureInfo.InvariantCulture), "1", StringComparison.Ordinal),
                $"NonVolatile twin: IncrNonVolatile evaluated exactly once (got '{incrResult}')");

            // ChannelPublishNonVolatile reaches a live listener with one delivery.
            token = subscriptions.Subscribe(host, channel, pattern: false,
                onMessage: m => { lock (Sync) published.Add(m); });
            Check(WaitUntil(() => NumSub(server, channel) == 1, 5000),
                "NonVolatile twin: channel listener active on the server");
            object pubResult = RedisUDF.RedisUDFChannelPublishNonVolatile(channel, "nv-msg", host);
            string pubText = Convert.ToString(pubResult, CultureInfo.InvariantCulture);
            Check(!string.IsNullOrEmpty(pubText) && !pubText.StartsWith("Error:", StringComparison.Ordinal),
                $"NonVolatile twin: ChannelPublishNonVolatile did not error (got '{pubResult}')");
            int delivered = WaitForCountStable(() => { lock (Sync) return published.Count; }, 1, 250, 5000);
            Check(delivered == 1,
                $"NonVolatile twin: ChannelPublishNonVolatile delivered exactly one message (got {delivered})");
        }
        finally
        {
            token?.Dispose();
            try { db.KeyDelete(key); } catch { }
            try { db.KeyDelete(incrKey); } catch { }
            RedisUDF.SyncWriteOverrideForTests = prevSync;
            RedisUdfAsync.AsyncWritesOverrideForTests = prevAsync;
        }
        Check(WaitUntil(() => NumSub(server, channel) == 0, 5000),
            "NonVolatile twin: channel released after cleanup");
    }

    private static readonly FieldInfo UpdateLatestTagField =
        typeof(UpdateCheck).GetField("_latestTag", BindingFlags.NonPublic | BindingFlags.Static);

    /// <summary>
    /// Live check of RedisUDFUpdateAvailable: the function must never block the
    /// caller (scheduling the refresh in the background) and must agree with the
    /// real latest GitHub release tag. RedisUDFUpdateAvailable itself is a pure
    /// status read once a tag is known, so the query is forced here through the
    /// private field; if the network or the GitHub rate limit makes that
    /// impossible, the check is SKIPPED (not failed), so an offline CI never
    /// turns red. This is the smoke-level companion to the offline unit tests
    /// that only exercise the comparison.
    /// </summary>
    private static void RunUpdateAvailableSmokeTest()
    {
        Check(UpdateLatestTagField != null, "UpdateCheck: _latestTag is reflectable");
        bool originalUpdateCheck = AppConfig.Current.UpdateCheck;
        object originalTag = UpdateLatestTagField.GetValue(null);
        try
        {
            // Seed a parseable running tag (a local "dev" build is never beaten)
            // and a known older tag, then call the function: it must return
            // immediately (non-blocking) and report the NEWER known tag.
            UpdateCheck.CurrentTagOverrideForTests = "1.4.0";
            UpdateLatestTagField.SetValue(null, "v1.4.1");

            // Disable the background refresh so only the non-blocking read path
            // runs (no network thread can overwrite the seeded tag meanwhile).
            AppConfig.Current.UpdateCheck = false;

            var sw = Stopwatch.StartNew();
            bool available = RedisUDF.RedisUDFUpdateAvailable();
            sw.Stop();
            Check(available, "UpdateAvailable: a newer known tag reports TRUE");
            Check(sw.ElapsedMilliseconds < 2000,
                $"UpdateAvailable: returns without blocking ({sw.ElapsedMilliseconds}ms)");

            AppConfig.Current.UpdateCheck = originalUpdateCheck;
            UpdateLatestTagField.SetValue(null, null);

            string latest = null;
            try
            {
                using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) })
                {
                    http.DefaultRequestHeaders.UserAgent.ParseAdd("RedisExcel-UpdateCheck");
                    http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
                    string json = http.GetStringAsync(
                        "https://api.github.com/repos/rspadim/RedisExcel/releases/latest")
                        .GetAwaiter().GetResult();
                    latest = (string)Newtonsoft.Json.Linq.JObject.Parse(json)["tag_name"];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SKIP UpdateAvailable: live GitHub query unavailable (" + ex.GetType().Name + ")");
                return;
            }

            if (string.IsNullOrEmpty(latest))
            {
                Console.WriteLine("SKIP UpdateAvailable: the latest release has no tag yet");
                return;
            }

            // With the tag known, RedisUDFUpdateAvailable is a pure comparison:
            // restore the real running tag and confirm agreement with GitHub.
            UpdateCheck.CurrentTagOverrideForTests = null;
            UpdateLatestTagField.SetValue(null, latest);
            bool saysOutdated = RedisUDF.RedisUDFUpdateAvailable();
            bool realBuildIsBehind = UpdateCheck.IsNewer(latest, UpdateCheck.CurrentTag);
            Check(saysOutdated == realBuildIsBehind,
                $"UpdateAvailable: agrees with the live latest release {latest} (reports {saysOutdated}, build {UpdateCheck.CurrentTag} behind={realBuildIsBehind})");
        }
        finally
        {
            UpdateCheck.CurrentTagOverrideForTests = null;
            UpdateLatestTagField.SetValue(null, originalTag);
            AppConfig.Current.UpdateCheck = originalUpdateCheck;
        }
    }
}

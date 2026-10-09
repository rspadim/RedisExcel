using RedisExcel;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

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

        Console.WriteLine(_failures == 0 ? "ALL PASS" : _failures + " FAILURE(S)");
        return _failures == 0 ? 0 : 1;
    }
}

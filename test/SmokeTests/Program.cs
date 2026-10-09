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
        Check(numsub() == 1, "channel stays subscribed while B is still active");

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

        // Duplicate suppression: identical consecutive payloads are skipped.
        var receivedD = new List<string>();
        var tokenD = subscriptions.Subscribe(host, channel, pattern: false,
            onMessage: m => { lock (Sync) receivedD.Add(m); });
        Check(WaitUntil(() => numsub() == 1, 5000), "channel active for the duplicate test");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup");
        publisher.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), "dup2");
        Check(WaitUntil(() => { lock (Sync) return receivedD.Contains("dup2"); }, 5000), "changed payload delivered");
        Check(WaitUntil(() => { lock (Sync) return receivedD.Count == 2; }, 2000), "identical repeated payload skipped");
        tokenD.Dispose();

        Console.WriteLine(_failures == 0 ? "ALL PASS" : _failures + " FAILURE(S)");
        return _failures == 0 ? 0 : 1;
    }
}

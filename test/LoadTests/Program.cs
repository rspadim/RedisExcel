using RedisExcel;
using StackExchange.Redis;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Load test for the subscription broadcast hot path.
///
/// Usage: dotnet run --project test\LoadTests -c Release -- [mode] [host] [seconds] [publishers] [listeners] [pattern] [channel]
///   mode       manager (default) = RedisSubscriptionManager, raw = plain SE.Redis subscriber baseline
///   host       default "127.0.0.1:6379,abortConnect=False"
///   seconds    default 10
///   publishers default 2 (internal fire-and-forget threads); 0 = listen-only,
///              use an external generator such as:
///              docker exec &lt;redis&gt; redis-benchmark -n 20000 -q -P 16 PUBLISH &lt;channel&gt; &lt;payload&gt;
///              (-t publish is a silent no-op on Redis 7.4)
///   listeners  default 1 (logical listeners registered on the same channel)
///   pattern    default false; true subscribes to the given channel as a pattern
///              (use channel "*" through the host-less listen mode to receive
///              every message of a server for a short stress test)
///   channel    default random "load:<guid>"; with pattern=true and
///              publishers=0 the default is "*"
///
/// Reports throughput, allocated bytes per received message and GC counts. The
/// "published" figure counts client-side intents: in fire-and-forget mode the
/// server may process fewer (the client exits with an unsent queue), so the
/// delivery ratio is an estimate - use `INFO commandstats` for server truth.
/// </summary>
internal static class Program
{
    private static long _received;

    private const string Usage =
        "usage: manager|raw <host:port> [seconds] [publishers] [listeners] [pattern] [channel]\n" +
        "  publishers=0 listens only (generate the load externally, e.g.\n" +
        "  redis-benchmark -n 20000 -q -P 16 PUBLISH <channel> <payload>; -t publish\n" +
        "  is a silent no-op on Redis 7.4); pattern=true subscribes as a pattern.";

    private static int Main(string[] args)
    {
        string mode = args.Length > 0 ? args[0] : "manager";
        if (mode != "manager" && mode != "raw")
        {
            Console.Error.WriteLine($"unknown mode '{mode}' (expected 'manager' or 'raw')");
            Console.Error.WriteLine(Usage);
            return 2;
        }
        string host = args.Length > 1 ? args[1] : "127.0.0.1:6379,abortConnect=False";
        int seconds = 10, publishers = 2, listeners = 1;
        bool pattern = false;
        if ((args.Length > 2 && !int.TryParse(args[2], out seconds))
            || (args.Length > 3 && !int.TryParse(args[3], out publishers))
            || (args.Length > 4 && !int.TryParse(args[4], out listeners))
            || (args.Length > 5 && !bool.TryParse(args[5], out pattern)))
        {
            Console.Error.WriteLine("invalid numeric/boolean argument");
            Console.Error.WriteLine(Usage);
            return 2;
        }
        string channelName = args.Length > 6 ? args[6] : null;

        AppDomain.MonitoringIsEnabled = true;

        string channel = channelName ?? (pattern && publishers == 0 ? "*" : "load:" + Guid.NewGuid().ToString("N"));
        var connections = new RedisConnectionManager();
        var tokens = new IDisposable[listeners];

        if (mode == "manager")
        {
            var subscriptions = new RedisSubscriptionManager(connections);
            for (int i = 0; i < listeners; i++)
            {
                tokens[i] = subscriptions.Subscribe(host, channel, pattern,
                    m => Interlocked.Increment(ref _received));
            }
        }
        else
        {
            var subscriber = connections.GetConnection(host, RedisPool.RtdSub).GetSubscriber();
            var redisChannel = new RedisChannel(channel,
                pattern ? RedisChannel.PatternMode.Pattern : RedisChannel.PatternMode.Literal);
            for (int i = 0; i < listeners; i++)
                subscriber.Subscribe(redisChannel, (ch, v) => Interlocked.Increment(ref _received));
        }

        Console.WriteLine($"mode={mode} listeners={listeners} publishers={publishers} seconds={seconds} pattern={pattern}");

        // Let the subscription settle and, in listen-only mode, confirm that
        // messages are flowing before starting the measurement window.
        if (publishers == 0)
        {
            var settle = Stopwatch.StartNew();
            while (settle.Elapsed.TotalSeconds < 5 && Interlocked.Read(ref _received) == 0)
                Thread.Sleep(100);
        }
        Thread.Sleep(500);

        int gc0 = GC.CollectionCount(0), gc1 = GC.CollectionCount(1), gc2 = GC.CollectionCount(2);
        long allocBefore = AppDomain.CurrentDomain.MonitoringTotalAllocatedMemorySize;
        long published = 0;
        var token = new CancellationTokenSource();
        var overall = Stopwatch.StartNew();
        double elapsedPublish;
        long receivedAtStop;

        if (publishers == 0)
        {
            // Listen-only mode: an external generator provides the load.
            Thread.Sleep(seconds * 1000);
            elapsedPublish = overall.Elapsed.TotalSeconds;
            receivedAtStop = Interlocked.Read(ref _received);
            Console.WriteLine("published      : (external load, listen-only mode)");
        }
        else
        {
            var pubConnection = connections.GetConnection(host, RedisPool.UdfData);
            var publisher = pubConnection.GetSubscriber();
            var publishChannel = new RedisChannel(channel, RedisChannel.PatternMode.Literal);
            var publisherTasks = new Task[publishers];
            // The manager applies the SkipRepeatedMessages dedup (consecutive
            // identical payloads are skipped for literal channels), so a
            // constant payload would be delivered only once. Rotate over a
            // small precomputed set (the global counter makes every publish
            // differ from the previous one) so the broadcast path is really
            // exercised; the small set keeps the extra allocation bounded.
            var payloadSet = new string[64];
            for (int i = 0; i < payloadSet.Length; i++)
                payloadSet[i] = "1234567890.12345#" + i;
            int payloadIndex = -1;
            for (int p = 0; p < publishers; p++)
            {
                publisherTasks[p] = Task.Run(() =>
                {
                    long local = 0;
                    while (!token.IsCancellationRequested)
                    {
                        var payload = payloadSet[Interlocked.Increment(ref payloadIndex) & (payloadSet.Length - 1)];
                        // Fire-and-forget: generate pressure without waiting a round trip.
                        publisher.Publish(publishChannel, payload, CommandFlags.FireAndForget);
                        local++;
                    }
                    Interlocked.Add(ref published, local);
                });
            }

            Thread.Sleep(seconds * 1000);
            token.Cancel();
            Task.WaitAll(publisherTasks);
            elapsedPublish = overall.Elapsed.TotalSeconds;
            receivedAtStop = Interlocked.Read(ref _received);
        }

        // Drain: wait until no new messages arrive for three consecutive 500ms
        // samples (a single quiet window can truncate a post-stop backlog;
        // max 10s).
        long last = Interlocked.Read(ref _received);
        int stable = 0;
        var drain = Stopwatch.StartNew();
        while (drain.Elapsed.TotalSeconds < 10 && stable < 3)
        {
            Thread.Sleep(500);
            long now = Interlocked.Read(ref _received);
            stable = now == last ? stable + 1 : 0;
            last = now;
        }

        long received = Interlocked.Read(ref _received);
        // Capture the allocation delta at the same point as receivedAtStop, so
        // drain traffic cannot skew the bytes-per-message figure.
        long allocAtStop = AppDomain.CurrentDomain.MonitoringTotalAllocatedMemorySize;
        long allocDelta = allocAtStop - allocBefore;

        if (publishers > 0)
        {
            Console.WriteLine($"published      : {published:N0} ({published / elapsedPublish:N0}/s)");
            Console.WriteLine($"delivery ratio : {(published == 0 ? 0 : 100.0 * receivedAtStop / published):F2}%");
        }
        Console.WriteLine($"received       : {receivedAtStop:N0} ({receivedAtStop / elapsedPublish:N0}/s in window, {received:N0} after drain)");
        Console.WriteLine($"allocated      : {allocDelta:N0} bytes ({allocDelta / Math.Max(receivedAtStop, 1):N0} bytes per received message in the window; process-wide estimate that includes publisher setup)");
        Console.WriteLine($"GC collections : gen0 +{GC.CollectionCount(0) - gc0}, gen1 +{GC.CollectionCount(1) - gc1}, gen2 +{GC.CollectionCount(2) - gc2}");

        for (int i = 0; i < tokens.Length; i++)
            tokens[i]?.Dispose();
        return 0;
    }
}

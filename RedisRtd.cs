using ExcelDna.ComInterop;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static ExcelDna.Integration.Rtd.ExcelRtdServer;

namespace RedisExcel
{
    public class AddIn : IExcelAddIn
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public void AutoOpen()
        {
            // A same-process reload (Excel re-opening the add-in without a new
            // AppDomain) must not reuse the shut-down runtime.
            RedisRuntime.ResetAfterAddInReload();
            // The previous session's ChannelLatest subscriptions died with the
            // old runtime; drop the cached listeners/messages/markers so a
            // reloaded add-in resubscribes instead of answering from stale state.
            RedisUDF.ResetAfterAddInReload();
            // Raise the ThreadPool floor a little: the async write queue and the
            // RTD timers run continuations on the pool, and a starved pool
            // (other add-ins, hosted CLR, policy) would stall cells while the
            // locks themselves stay healthy. Best effort only.
            try
            {
                ThreadPool.GetMinThreads(out int minWorkers, out int minIo);
                ThreadPool.SetMinThreads(Math.Max(minWorkers, 16), Math.Max(minIo, 16));
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "AutoOpen: could not raise the thread pool floor");
            }
            try
            {
                ComServer.DllRegisterServer();
            }
            catch (Exception ex)
            {
                // A COM registration failure (no writable hive, policy...) must
                // not abort the whole AutoOpen: UDFs keep working and =RTD
                // shows #N/D until the ProgID is registered.
                logger.Error(ex, "AutoOpen: ComServer.DllRegisterServer failed");
            }
            UpdateCheck.Start();
        }

        public void AutoClose()
        {
            try
            {
                RedisRuntime.Shutdown();
            }
            catch
            {
                // best effort; the Excel process releases the sockets on shutdown
            }
            ComServer.DllUnregisterServer();
        }
    }

    public sealed class TopicData
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly object _sync = new object();
        private readonly object _subscribeSync = new object();
        private string _lastValue;

        /// <summary>Maximum payload length embedded in ToString()/log lines.</summary>
        private const int LastValueLogLimit = 64;

        // Not dirty initially: the "(ConnectData)" placeholder returned by
        // ConnectData must stay visible until the first real value arrives
        // (poll result or subscription message); flushing the initial null
        // _lastValue on the first Excel tick would blank the cell.
        private bool _dirty = false;

        // Utc ticks of the last value pushed to Excel; with a conflation window
        // this gates the next push so the latest value is delivered at most
        // once per window (0 = never pushed yet, which always pushes).
        private long _lastPushTicks;

        // A persistently failing Excel push stays dirty and is retried every
        // tick; log the first failure at Error and the rest at Debug so a stuck
        // topic cannot flood the log (cleared by a successful push).
        private bool _updateFailureLogged;

        /// <summary>
        /// Marks the topic dirty when a value already arrived while ConnectData
        /// was running: Excel-DNA drops pushes for topics that are not active
        /// yet, so the next flush tick must re-publish the cached value once the
        /// topic is active (UpdateValue is a no-op when the value is unchanged).
        /// </summary>
        internal void MarkDirtyAfterConnect()
        {
            lock (_sync)
            {
                if (_lastValue != null)
                    _dirty = true;
            }
        }

        public TopicData(Topic topic, string type, string keyOrChannel, string field, string host)
        {
            Topic = topic;
            Type = type;
            KeyOrChannel = keyOrChannel;
            Field = field;
            Host = host;
        }

        public Topic Topic { get; }
        public string Type { get; }
        public string KeyOrChannel { get; }
        public string Field { get; }
        public string Host { get; }
        public IDisposable Subscription { get; private set; }

        /// <summary>
        /// Failed Subscribe attempts for this topic: the first failure is logged at
        /// Warning with the exception, later retries at Debug with the message
        /// only (see RedisRtd.TrySubscribe).
        /// </summary>
        public int SubscribeAttempts;

        private long _nextSubscribeAttemptUtcTicks;

        /// <summary>
        /// Earliest UTC time for the next Subscribe retry after a failure: the
        /// delay doubles from 1s up to a 30s cap (1s, 2s, 4s, ...). Reset to
        /// DateTime.MinValue on success. Backed by long ticks so the timer tick
        /// and the Excel thread can read/write it without tearing.
        /// </summary>
        public DateTime NextSubscribeAttemptUtc
        {
            get { return new DateTime(Interlocked.Read(ref _nextSubscribeAttemptUtcTicks), DateTimeKind.Utc); }
            set { Interlocked.Exchange(ref _nextSubscribeAttemptUtcTicks, value.Ticks); }
        }

        private int _subscribeGate;

        /// <summary>
        /// TRUE when the caller acquired the exclusive right to run Subscribe for
        /// this topic. Guards against ConnectData's initial attempt and the Redis
        /// tick retry subscribing the same topic twice (which would leak one
        /// registration).
        /// </summary>
        public bool TryBeginSubscribe() => Interlocked.CompareExchange(ref _subscribeGate, 1, 0) == 0;

        public void EndSubscribe() => Interlocked.Exchange(ref _subscribeGate, 0);

        /// <summary>
        /// Installs the subscription unless the topic was disconnected meanwhile.
        /// Returns false when disconnected: the caller must dispose the fresh token.
        /// </summary>
        public bool InstallSubscription(IDisposable subscription)
        {
            lock (_subscribeSync)
            {
                if (Disconnected)
                    return false;
                Subscription = subscription;
                return true;
            }
        }

        /// <summary>
        /// Marks the topic disconnected and atomically detaches (and returns) the
        /// installed subscription, so a concurrent install can never be lost.
        /// Returns null when there is no subscription (for example a SUB/PSUB
        /// topic whose initial Subscribe failed and is awaiting a retry).
        /// </summary>
        public IDisposable DetachSubscription()
        {
            lock (_subscribeSync)
            {
                Disconnected = true;
                var subscription = Subscription;
                Subscription = null;
                return subscription;
            }
        }

        /// <summary>Set when the RTD server removed this topic; callbacks and flushes must stop publishing.</summary>
        public volatile bool Disconnected;

        public string LastValue { get { lock (_sync) return _lastValue; } }

        public bool Dirty { get { lock (_sync) return _dirty; } }

        private RedisValue _lastPolledValue;
        private bool _hasLastPolledValue;

        // Two-phase commit for polled values: compare first, then store the value
        // once it was ACCEPTED for delivery (queued via UpdateOnly when coalescing,
        // or pushed by UpdateAndSendToExcel). A failed Excel push is retried thanks
        // to the restored dirty flag; the commit never suppresses it.

        /// <summary>TRUE when the polled value changed since the previous tick.</summary>
        public bool HasChangedPolledValue(RedisValue value)
        {
            lock (_sync)
            {
                if (_hasLastPolledValue && _lastPolledValue == value)
                    return false;
                return true;
            }
        }

        /// <summary>Stores the polled value as seen; call once the value was accepted for delivery.</summary>
        public void CommitPolledValue(RedisValue value)
        {
            lock (_sync)
            {
                _lastPolledValue = value;
                _hasLastPolledValue = true;
            }
        }

        private HashEntry[] _lastPolledHash;
        private bool _hasLastPolledHash;

        // Two-phase commit for polled hashes: compare first, then store the reference
        // once the value was ACCEPTED for delivery (queued via UpdateOnly when coalescing,
        // or pushed by UpdateAndSendToExcel). A failed Excel push is retried thanks to
        // the restored dirty flag; the commit never suppresses it.

        /// <summary>TRUE when the polled hash changed since the previous tick.</summary>
        public bool HasChangedPolledHash(HashEntry[] entries)
        {
            lock (_sync)
            {
                if (_hasLastPolledHash && RedisResultFormatter.HashEquals(_lastPolledHash, entries))
                    return false;
                return true;
            }
        }

        /// <summary>Stores the polled hash as seen; call once the value was accepted for delivery.</summary>
        public void CommitPolledHash(HashEntry[] entries)
        {
            lock (_sync)
            {
                _lastPolledHash = entries;
                _hasLastPolledHash = true;
            }
        }

        // NOTE (upstream risk, Excel-DNA 1.9.0-beta2/rc1): a worker thread calling
        // Topic.UpdateValue takes ExcelRtdServer._updateLock and then
        // RtdUpdateSynchronization._lockObject, while the Excel main thread inside
        // ProcessUpdateNotifications takes them in the opposite order (ABBA). The
        // probability per update is low and no freeze was ever observed in
        // practice; the known mitigation (batching UpdateValue on the main thread
        // via ExcelAsyncUtil.QueueAsMacro) would rework the whole delivery path,
        // so this stays tracked as an upstream issue instead (DESIGN-v1.4.0.md).
        public void UpdateAndSendToExcel(string data)
        {
            if (Disconnected)
                return;
            lock (_sync)
            {
                // Defense in depth: the topic may have been disconnected while
                // this call waited for the lock.
                if (Disconnected)
                    return;
                // The state update and the Excel push are serialized per topic,
                // so an immediate push and a tick flush can never interleave and
                // overwrite a newer value with an older one. A failed push stays
                // dirty so the next tick retries it.
                _lastValue = data;
                try
                {
                    Topic.UpdateValue(data);
                    // Intentionally keep the topic dirty: the push may have been
                    // dropped because Excel-DNA had not activated the topic yet
                    // (the ConnectData window), and a value arriving in that
                    // window would then never reach the cell until it changes.
                    // The next Excel tick re-pushes the CURRENT value once the
                    // topic is active (UpdateValue is a no-op when unchanged),
                    // which reconciles that window; the tick clears the flag.
                    _dirty = true;
                    _lastPushTicks = DateTime.UtcNow.Ticks;
                }
                catch
                {
                    _dirty = true;
                    throw;
                }
            }
        }

        public void UpdateOnly(string data)
        {
            if (Disconnected)
                return;
            lock (_sync)
            {
                _lastValue = data;
                _dirty = true;
            }
        }

        /// <summary>
        /// Flushes the latest value to Excel when the topic is dirty AND the
        /// conflation window (if any) elapsed; otherwise it stays dirty and a
        /// later tick pushes the then-current value (latest wins).
        /// </summary>
        public void SendToExcelIfDirty(int conflationMs)
        {
            if (Disconnected)
                return;
            Exception updateError = null;
            lock (_sync)
            {
                if (!_dirty)
                    return;
                // Defense in depth: the topic may have been disconnected while
                // this tick waited for the lock.
                if (Disconnected)
                    return;
                long nowTicks = DateTime.UtcNow.Ticks;
                if (!Conflation.IsDue(_lastPushTicks, nowTicks, conflationMs))
                    return; // inside the window: keep the newest value pending
                try
                {
                    // Push while holding the lock: a concurrent state update
                    // cannot interleave and overwrite the newer value. On
                    // failure stay dirty so the next Excel tick retries.
                    Topic.UpdateValue(_lastValue);
                    _dirty = false;
                    _lastPushTicks = nowTicks;
                    _updateFailureLogged = false;
                }
                catch (Exception ex)
                {
                    // _dirty stays true; log outside the lock so a slow log
                    // write cannot delay other producers of this topic.
                    updateError = ex;
                }
            }
            if (updateError != null)
            {
                // Throttled: a permanently failing push would otherwise log an
                // Error every Excel tick (default ~100ms).
                if (_updateFailureLogged)
                {
                    if (logger.IsDebugEnabled)
                        logger.Debug(updateError, "SendToExcelIfDirty: update still failing");
                }
                else
                {
                    _updateFailureLogged = true;
                    logger.Error(updateError, "SendToExcelIfDirty: update failed");
                }
            }
        }

        public override string ToString()
        {
            return $"TopicId={Topic.TopicId}, Type={Type}, KeyOrChannel={KeyOrChannel}, Field={Field}, host={AppConfig.MaskHost(Host)}, LastValue={TruncateForLog(LastValue)}, Dirty={Dirty}";
        }

        /// <summary>
        /// Caps the payload embedded in log lines: a subscription message can be
        /// arbitrarily large, and a full ToString() would dump it into the log.
        /// </summary>
        private static string TruncateForLog(string value)
        {
            if (value == null || value.Length <= LastValueLogLimit)
                return value;
            return value.Substring(0, LastValueLogLimit) + "...(truncated)";
        }
    }

    [ComVisible(true)]
    [ProgId("RedisRtd")]
    public class RedisRtd : ExcelRtdServer
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly ConcurrentDictionary<RedisRtd, byte> Instances = new ConcurrentDictionary<RedisRtd, byte>();
        private static readonly object InstancesSync = new object();
        private static volatile RedisRtd Instance;

        // Monotonic start ordinal per RTD server instance: when the last-started
        // server terminates, the status/default surface falls back to the
        // highest-ordinal survivor instead of an arbitrary one.
        private static long _nextStartOrdinal;
        private long _startOrdinal;

        private readonly ConcurrentDictionary<int, TopicData> _polledTopics = new ConcurrentDictionary<int, TopicData>();
        private readonly ConcurrentDictionary<int, TopicData> _subscribedTopics = new ConcurrentDictionary<int, TopicData>();

        private long _messageCount;
        private long _previousSecondCount;
        private long _messageCounterThreshold;
        private volatile bool _realTimeUpdates = true;
        private ENUMExcelUpdateStyle _excelUpdateStyle = ENUMExcelUpdateStyle.Automatic;
        private double _excelUpdateRateMs = 100;
        private double _redisUpdateRateMs = 1000;
        private bool _useGetMultiple = true;
        private bool _skipRepeatedMessages = true;
        private bool _coalesceRealtimeUpdates = true;
        private int _conflationMs;
        private string _defaultHost;

        private System.Timers.Timer _excelTimer;
        private System.Timers.Timer _redisTimer;
        private System.Timers.Timer _counterTimer;

        // Reentrancy gates: a tick is skipped while the matching timer callback is still running.
        private readonly TickGate _redisTickGate = new TickGate();
        private readonly TickGate _excelTickGate = new TickGate();
        private readonly TickGate _counterTickGate = new TickGate();

        public static long CurrentMessagesCounter()
        {
            var instance = Instance;
            // Last COMPLETED second (swapped by OnCounterTick), not the partial
            // in-flight count that grows through the current second.
            return instance == null ? 0 : Interlocked.Read(ref instance._previousSecondCount);
        }

        public static int RedisConnectionsCount() => RedisRuntime.Connections.LiveRtdConnectionCount();

        /// <summary>
        /// Number of active Redis pub/sub listeners registered by the RTD layer. RTD-only
        /// (UDF listeners are excluded; they use a separate origin).
        /// </summary>
        public static int RedisSubscriptionsCount() => RedisRuntime.Subscriptions.ListenerCountWithOrigin("RTD");

        public static int TopicsCount()
        {
            int total = 0;
            foreach (var instance in Instances.Keys)
                total += instance._polledTopics.Count + instance._subscribedTopics.Count;
            return total;
        }

        /// <summary>
        /// Number of Redis channels/patterns with at least one RTD listener. RTD-only
        /// (UDF-subscribed channels are excluded; they use a separate origin).
        /// </summary>
        public static int ChannelTopicsCount() => RedisRuntime.Subscriptions.ChannelCountWithOrigin("RTD");

        public static string DefaultHost() => Instance?._defaultHost;

        public static double ExcelUpdateRate() => Instance?._excelUpdateRateMs ?? 0;

        public static double RedisUpdateRate() => Instance?._redisUpdateRateMs ?? 0;

        public static bool IsRealTimeEnabled() => Instance?._realTimeUpdates ?? false;

        protected override bool ServerStart()
        {
            logger.Info("ServerStart: starting RTD server");
            lock (InstancesSync)
            {
                Instance = this;
                _startOrdinal = Interlocked.Increment(ref _nextStartOrdinal);
                Instances[this] = 0;
            }

            var config = AppConfig.Current.RTD;
            _defaultHost = config.host;
            _excelUpdateRateMs = config.ExcelUpdateRateMs;
            _redisUpdateRateMs = config.RedisUpdateRateMs;
            _excelUpdateStyle = config.ExcelUpdateStyle;
            _messageCounterThreshold = config.MessageCounterThreshold;
            _useGetMultiple = config.UseGetMultiple;
            _skipRepeatedMessages = AppConfig.Current.SkipRepeatedMessages;
            _coalesceRealtimeUpdates = AppConfig.Current.CoalesceRealtimeUpdates;
            // Explicit ConflationMs wins; absent falls back to the legacy
            // CoalesceRealtimeUpdates boolean (true = the Excel tick window),
            // so an unchanged config keeps its previous behaviour.
            _conflationMs = Conflation.Resolve(
                AppConfig.Current.ConflationMs, _coalesceRealtimeUpdates, (int)_excelUpdateRateMs);
            // in Realtime/Timer styles the mode is fixed; in Automatic it is recalculated from the message counter
            _realTimeUpdates = _excelUpdateStyle != ENUMExcelUpdateStyle.Timer;

            logger.Info(
                "ServerStart: config " +
                $"host={AppConfig.MaskHost(_defaultHost)}, excelRate={_excelUpdateRateMs}ms, redisRate={_redisUpdateRateMs}ms, " +
                $"style={_excelUpdateStyle}, threshold={_messageCounterThreshold}, useGetMultiple={_useGetMultiple}, " +
                $"conflationMs={_conflationMs}");

            _redisTimer = CreateTimer(_redisUpdateRateMs, "redis", OnRedisTick);
            _excelTimer = CreateTimer(_excelUpdateRateMs, "excel", OnExcelTick);
            _counterTimer = CreateTimer(1000, "counter", OnCounterTick);
            _redisTimer.Start();
            _excelTimer.Start();
            _counterTimer.Start();
            return true;
        }

        protected override void ServerTerminate()
        {
            logger.Info("ServerTerminate");

            // Mark every topic before disposing the timers and clearing the registries:
            // despite its name, Timer.Dispose does not wait for an Elapsed callback that
            // is already running, so only the Disconnected flags guarantee that such a
            // tick stops publishing and never touches the live registries after teardown
            // starts.
            foreach (var td in _polledTopics.Values.Concat(_subscribedTopics.Values))
                td.Disconnected = true;

            DisposeTimer(ref _redisTimer);
            DisposeTimer(ref _excelTimer);
            DisposeTimer(ref _counterTimer);

            foreach (var td in _subscribedTopics.Values)
            {
                try
                {
                    // DetachSubscription tolerates a null Subscription (a SUB/PSUB
                    // topic whose ConnectData Subscribe failed and is awaiting a
                    // retry) and wins any race with a retry installing a fresh token.
                    td.DetachSubscription()?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "ServerTerminate: error disposing subscription");
                }
            }
            _subscribedTopics.Clear();
            _polledTopics.Clear();

            lock (InstancesSync)
            {
                Instances.TryRemove(this, out _);
                if (ReferenceEquals(Instance, this))
                    Instance = HighestOrdinalInstance();
            }
        }

        /// <summary>
        /// Survivor with the highest start ordinal; called under InstancesSync so
        /// a terminating server never promotes an older instance by accident.
        /// </summary>
        private static RedisRtd HighestOrdinalInstance()
        {
            RedisRtd best = null;
            foreach (var candidate in Instances.Keys)
            {
                if (best == null || candidate._startOrdinal > best._startOrdinal)
                    best = candidate;
            }
            return best;
        }

        private static System.Timers.Timer CreateTimer(double intervalMs, string name, Action action)
        {
            var timer = new System.Timers.Timer(intervalMs) { AutoReset = true };
            timer.Elapsed += (sender, args) =>
            {
                // System.Timers.Timer swallows unhandled exceptions; without this catch the
                // failure would be completely silent (one of the original bug symptoms).
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Timer({name}): unhandled error");
                }
            };
            return timer;
        }

        private static void DisposeTimer(ref System.Timers.Timer timer)
        {
            var toDispose = timer;
            timer = null;
            if (toDispose == null)
                return;
            try
            {
                toDispose.Stop();
                toDispose.Dispose();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "DisposeTimer: error");
            }
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            try
            {
                if (_subscribedTopics.TryGetValue(topic.TopicId, out var existingSub))
                {
                    logger.Warn($"ConnectData: TopicId={topic.TopicId} already exists in subscriptions");
                    return existingSub.LastValue;
                }
                if (_polledTopics.TryGetValue(topic.TopicId, out var existingPoll))
                {
                    logger.Warn($"ConnectData: TopicId={topic.TopicId} already exists in polled topics");
                    return existingPoll.LastValue;
                }

                string command = topicInfo.Count > 0 ? topicInfo[0].ToUpperInvariant().Trim() : null;
                string param2 = topicInfo.Count > 1 ? topicInfo[1] : null;
                string param3 = topicInfo.Count > 2 ? topicInfo[2] : null;
                string param4 = topicInfo.Count > 3 ? topicInfo[3] : null;
                logger.Info($"ConnectData: command={command}, param2={param2}, param3={AppConfig.MaskHost(param3)}, param4={AppConfig.MaskHost(param4)}, TopicId={topic.TopicId}");

                // Redis keys/fields may be empty or whitespace; only a missing
                // (null) argument is invalid for them (StackExchange.Redis
                // rejects null for keys/fields). SUB/PSUB channels are the
                // exception: a blank channel can never subscribe, so it is
                // rejected up front below instead of being retried forever.
                TopicData td;
                switch (command)
                {
                    case "GET":
                    case "HGETALL":
                        if (param2 == null)
                            throw new Exception($"{command} requires a key as the second argument");
                        // Documented arguments: key + optional host.
                        RejectExtraArguments(command, topicInfo, 3);
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        RedisConnectionManager.ParseOptions(td.Host);
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "HGET":
                        if (param2 == null)
                            throw new Exception("HGET requires a key as the second argument");
                        if (param3 == null)
                            throw new Exception("HGET requires a field as the third argument");
                        // Documented arguments: key + field + optional host.
                        RejectExtraArguments(command, topicInfo, 4);
                        td = new TopicData(topic, command, param2, param3, AppConfig.ResolveRtdHost(param4));
                        RedisConnectionManager.ParseOptions(td.Host);
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "SUB":
                    case "PSUB":
                        // A blank/whitespace channel can never subscribe, so
                        // refuse it BEFORE the topic is registered: otherwise
                        // every Redis tick would retry it forever, with a full
                        // stack trace per attempt.
                        if (string.IsNullOrWhiteSpace(param2))
                            throw new Exception($"{command} requires a channel as the second argument");
                        // Documented arguments: channel + optional host.
                        RejectExtraArguments(command, topicInfo, 3);
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        // A malformed host can never subscribe successfully, so
                        // reject it up front instead of retrying it forever.
                        RedisConnectionManager.ParseOptions(td.Host);
                        // Register BEFORE subscribing: a transient Subscribe failure
                        // keeps the topic (Subscription == null) so OnRedisTick
                        // retries it, and the cell starts with the #ERROR text
                        // below until the first message overwrites it.
                        _subscribedTopics[topic.TopicId] = td;
                        var subscribeError = TrySubscribe(td, out _);
                        if (subscribeError != null)
                            return $"#ERROR: ConnectData: {AppConfig.MaskHost(subscribeError)}";
                        break;
                    default:
                        throw new Exception($"unknown command '{command}', expected one of [GET, HGET, HGETALL, SUB, PSUB]");
                }

                logger.Info($"ConnectData: accepted {td}");
                // A subscription message can arrive between Subscribe and Excel
                // activating the topic; Excel-DNA drops pushes for topics that are
                // not active yet, so return the value cached by that push instead
                // of the sentinel (which would overwrite it permanently in
                // non-coalesced realtime mode). LastValue is read under the topic
                // lock, so this is safe against concurrent updates. A value that
                // raced the activation also marks the topic dirty, so the first
                // flush tick re-publishes it once the topic is active
                // (UpdateValue is a no-op when the value is unchanged).
                string initialValue = td.LastValue;
                td.MarkDirtyAfterConnect();
                return initialValue ?? "(ConnectData)";
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ConnectData: error");
                return $"#ERROR: ConnectData: {AppConfig.MaskHost(ex.Message)}";
            }
        }

        /// <summary>
        /// Rejects topic arguments beyond the documented ones for the command.
        /// Excel may append trailing blank/empty tokens, so only a non-blank extra
        /// is an error; a typo such as =RTD(...,"GET","k","host","extra") must
        /// fail loudly instead of being silently ignored.
        /// </summary>
        private static void RejectExtraArguments(string command, IList<string> topicInfo, int documentedCount)
        {
            for (int i = documentedCount; i < topicInfo.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(topicInfo[i]))
                    throw new Exception($"{command} accepts at most {documentedCount - 1} argument(s); unexpected extra argument '{topicInfo[i]}' at topic position {i + 1}");
            }
        }

        protected override void DisconnectData(Topic topic)
        {
            try
            {
                if (_subscribedTopics.TryRemove(topic.TopicId, out var sub))
                {
                    // Mark first so in-flight callbacks stop publishing before the
                    // shared subscription is torn down. DetachSubscription also
                    // tolerates a null Subscription (failed initial Subscribe,
                    // retry pending) and wins any race with a retry that is
                    // installing a fresh token.
                    sub.DetachSubscription()?.Dispose();
                    logger.Info($"DisconnectData: removed subscription {sub}");
                }
                else if (_polledTopics.TryRemove(topic.TopicId, out var polled))
                {
                    // In-flight poll ticks may still hold this topic; stop them from
                    // pushing values to a disconnected Excel topic.
                    polled.Disconnected = true;
                    logger.Info($"DisconnectData: removed polled topic {polled}");
                }
                else
                {
                    // Excel calls DisconnectData even for topics whose ConnectData
                    // was rejected (Excel-DNA registers them anyway), so this is
                    // expected for malformed topics and not an error.
                    logger.Debug($"DisconnectData: unknown TopicId={topic.TopicId} (expected for topics rejected at ConnectData)");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "DisconnectData: error");
            }
        }

        private IDisposable Subscribe(TopicData td)
        {
            bool pattern = td.Type == "PSUB";
            long topicId = td.Topic.TopicId;
            var subscription = RedisRuntime.Subscriptions.Subscribe(td.Host, td.KeyOrChannel, pattern, message =>
            {
                if (td.Disconnected)
                    return;
                Interlocked.Increment(ref _messageCount);
                if (logger.IsTraceEnabled)
                    logger.Trace($"Subscribe: TopicId={topicId}, channel={td.KeyOrChannel}, message={message}");
                DeliverToTopic(td, message);
            }, "RTD");
            logger.Info($"Subscribe: subscribed {td}");
            return subscription;
        }

        /// <summary>
        /// Attempts to subscribe a registered SUB/PSUB topic (idempotent). Returns
        /// the failure message, or null on success/already-subscribed. The first
        /// failure is logged at Warning with the exception; later attempts log at
        /// Debug with the message only, so a permanently unreachable host does not
        /// flood the log. Each failure schedules the next retry (1s, 2s, 4s ...
        /// capped at 30s, see NextSubscribeAttemptUtc).
        /// </summary>
        private string TrySubscribe(TopicData td, out bool attempted)
        {
            attempted = false;
            if (td.Disconnected || td.Subscription != null)
                return null;
            // Only one thread may run Subscribe for a topic at a time: both the
            // Excel thread (ConnectData) and the Redis tick (retry) can get here.
            if (!td.TryBeginSubscribe())
                return null;
            attempted = true;
            try
            {
                if (td.Disconnected || td.Subscription != null)
                    return null;
                var subscription = Subscribe(td);
                // Subscribe succeeded: clear the retry backoff so a later
                // failure starts again at 1s instead of jumping to the 30s cap.
                td.NextSubscribeAttemptUtc = default(DateTime);
                Interlocked.Exchange(ref td.SubscribeAttempts, 0);
                if (!td.InstallSubscription(subscription))
                {
                    // DisconnectData/ServerTerminate raced with the install: the
                    // topic was removed while subscribing, so tear the fresh
                    // registration down instead of leaking it.
                    try
                    {
                        subscription.Dispose();
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(ex, $"TrySubscribe: error disposing raced subscription, TopicId={td.Topic.TopicId}");
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                int attempt = Interlocked.Increment(ref td.SubscribeAttempts);
                // 1s, 2s, 4s, 8s, 16s, then 30s for every later attempt.
                int seconds = Math.Min(30, 1 << Math.Min(attempt - 1, 5));
                td.NextSubscribeAttemptUtc = DateTime.UtcNow.AddSeconds(seconds);
                if (attempt == 1)
                    logger.Warn(ex, $"Subscribe failed, retry in {seconds}s: {td}");
                else if (logger.IsDebugEnabled)
                    logger.Debug($"Subscribe retry failed (attempt {attempt}), retry in {seconds}s: {td}: {AppConfig.MaskHost(ex.Message)}");
                return ex.Message;
            }
            finally
            {
                td.EndSubscribe();
            }
        }

        // Failed SUB/PSUB registrations are retried with backoff; cap the
        // number of retries per Redis tick so a large set of unreachable
        // topics cannot stall the tick.
        private const int PendingSubscribeRetriesPerTick = 4;

        /// <summary>
        /// Retries SUB/PSUB topics whose Subscribe failed at ConnectData
        /// (Subscription is null), at most <see cref="PendingSubscribeRetriesPerTick"/>
        /// per tick and only once each topic's backoff elapsed. A successful
        /// retry lets messages flow into the cell, which still shows the initial
        /// "#ERROR: ConnectData: ..." text until the first message overwrites it.
        /// </summary>
        private void RetryPendingSubscriptions()
        {
            var now = DateTime.UtcNow;
            int retried = 0;
            foreach (var td in _subscribedTopics.Values)
            {
                if (retried >= PendingSubscribeRetriesPerTick)
                    break;
                // TrySubscribe re-checks Disconnected/Subscription itself and
                // reports whether a real attempt happened, so only the backoff
                // gate remains here.
                if (td.NextSubscribeAttemptUtc > now)
                    continue;
                TrySubscribe(td, out bool attempted);
                if (attempted)
                {
                    retried++;
                    // A recovered subscription must stop showing the initial
                    // "#ERROR: ConnectData..." text on a quiet channel: push a
                    // neutral placeholder that the next tick publishes.
                    if (td.Subscription != null && td.LastValue == null)
                        td.UpdateOnly("(subscribed)");
                }
            }
        }

        /// <summary>
        /// Reentrancy prologue shared by the three timers: a tick that fires while
        /// the previous one still runs is skipped (never queued), with the same log
        /// text as before. Callers must release the gate in a finally.
        /// </summary>
        private static bool TryBeginTick(TickGate gate, string name)
        {
            if (gate.TryEnter())
                return true;
            if (logger.IsDebugEnabled)
                logger.Debug($"{name}: previous tick still running, skipping");
            return false;
        }

        private void OnCounterTick()
        {
            if (!TryBeginTick(_counterTickGate, "OnCounterTick"))
                return;
            try
            {
                UpdateRealtimeMode(allowReenable: true);
                // Report the last COMPLETED second instead of a partial in-flight
                // count: swap the interval counter into _previousSecondCount
                // (read by RedisRTDMessagesCounter) and reset the in-flight one.
                Interlocked.Exchange(ref _previousSecondCount, Interlocked.Exchange(ref _messageCount, 0));
            }
            finally
            {
                _counterTickGate.Exit();
            }
        }

        private void OnExcelTick()
        {
            if (!TryBeginTick(_excelTickGate, "OnExcelTick"))
                return;
            try
            {
                UpdateRealtimeMode(allowReenable: false);
                if (_polledTopics.IsEmpty && _subscribedTopics.IsEmpty)
                    return;
                if (logger.IsDebugEnabled)
                    logger.Debug($"OnExcelTick: flushing dirty topics, polled={_polledTopics.Count}, subscribed={_subscribedTopics.Count}");
                // Dirty values are normally empty in non-coalesced real-time mode
                // because updates are sent immediately; flushing every tick is
                // cheap and guarantees delivery when Automatic style re-enables
                // real-time after a burst.
                foreach (var td in _polledTopics.Values.Concat(_subscribedTopics.Values))
                    td.SendToExcelIfDirty(_conflationMs);
            }
            finally
            {
                _excelTickGate.Exit();
            }
        }

        /// <summary>
        /// In Automatic style, turns off real-time updates when the message volume exceeds
        /// the threshold. allowReenable=true (the 1s tick) also goes back to real-time
        /// when the volume drops.
        /// </summary>
        private void UpdateRealtimeMode(bool allowReenable)
        {
            if (_excelUpdateStyle != ENUMExcelUpdateStyle.Automatic)
                return;

            long count = Interlocked.Read(ref _messageCount);
            bool belowThreshold = _messageCounterThreshold <= 0 || count < _messageCounterThreshold;
            if (!allowReenable && belowThreshold)
                return;

            bool next = belowThreshold;
            if (next == _realTimeUpdates)
                return;
            _realTimeUpdates = next;
            logger.Debug($"UpdateRealtimeMode: realTimeUpdates={next}, messages={count}/{_messageCounterThreshold}");
        }

        private void OnRedisTick()
        {
            if (!TryBeginTick(_redisTickGate, "OnRedisTick"))
                return;
            try
            {
                // Failed SUB/PSUB registrations are retried even when this server
                // instance has no polled topics.
                RetryPendingSubscriptions();
                if (_polledTopics.IsEmpty)
                    return;
                var hostGroups = _polledTopics.Values.GroupBy(t => t.Host).ToList();
                if (logger.IsDebugEnabled)
                    logger.Debug($"OnRedisTick: polling {_polledTopics.Count} topic(s) across {hostGroups.Count} host(s), rate={_redisUpdateRateMs}ms");
                // Bounded parallelism so one slow/unreachable host cannot delay the
                // others in the same tick. The per-host try/catch keeps the hosts
                // isolated, and the TickGate is held until every host completed, so
                // ticks never overlap (same reentrancy behavior as before).
                Parallel.ForEach(hostGroups, new ParallelOptions { MaxDegreeOfParallelism = 4 }, group =>
                {
                    try
                    {
                        PollHost(group.Key, group.ToList());
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"OnRedisTick: host={AppConfig.MaskHost(group.Key)}");
                    }
                });
            }
            finally
            {
                _redisTickGate.Exit();
            }
        }

        private void PollHost(string host, List<TopicData> topics)
        {
            var db = RedisRuntime.Connections.GetDatabase(host, RedisPool.RtdData);

            // Defense-in-depth: a topic registered with a null key/field (which
            // ConnectData now rejects) must never reach the batch, because it
            // would make StackExchange.Redis throw for the whole host tick.
            // Empty/whitespace names are valid Redis names and are allowed.
            bool IsMalformed(TopicData td)
            {
                if (td.KeyOrChannel == null ||
                    (td.Type == "HGET" && td.Field == null))
                {
                    logger.Warn($"PollHost: skipping malformed topic, TopicId={td.Topic.TopicId}, Type={td.Type}, host={AppConfig.MaskHost(host)}");
                    return true;
                }
                return false;
            }

            // GET in batch (MGET): one round-trip for all GETs of the host.
            // Isolated so an MGET failure cannot abort HGET/HGETALL polling for the host.
            if (_useGetMultiple)
            {
                try
                {
                    var gets = topics.Where(t => t.Type == "GET" && !IsMalformed(t)).ToList();
                    if (gets.Count > 0)
                    {
                        var keys = gets.Select(t => (RedisKey)t.KeyOrChannel).ToArray();
                        if (logger.IsTraceEnabled)
                            logger.Trace($"PollHost: GETMULTI host={AppConfig.MaskHost(host)}, keys=[{string.Join(", ", gets.Select(t => t.KeyOrChannel))}]");
                        var values = db.StringGet(keys);
                        for (int i = 0; i < gets.Count; i++)
                        {
                            var td = gets[i];
                            if (_skipRepeatedMessages && !td.HasChangedPolledValue(values[i]))
                                continue;
                            // Per-item try so one failing Excel push cannot abort the whole MGET batch.
                            try
                            {
                                // IsNull, not HasValue: HasValue is also false for the empty
                                // string (RedisValue.HasValue => !IsNullOrEmpty), so an existing
                                // key holding "" must display as an empty cell; only a genuinely
                                // missing key gets the sentinel.
                                Publish(td, values[i].IsNull ? "(no value)" : values[i].ToString());
                                td.CommitPolledValue(values[i]);
                            }
                            catch (Exception ex)
                            {
                                logger.Error(ex, $"PollHost: GET key={td.KeyOrChannel}, host={AppConfig.MaskHost(host)}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: GETMULTI host={AppConfig.MaskHost(host)}");
                }
                finally
                {
                    // GET values are handled by the MGET block only; keep them out of the
                    // pipeline even when MGET failed (they are retried on the next tick).
                    topics = topics.Where(t => t.Type != "GET").ToList();
                }
            }
            if (topics.Count == 0)
                return;

            // HGET/HGETALL (and GET when UseGetMultiple=false) in a pipeline
            var batch = db.CreateBatch();
            var singleTasks = new List<KeyValuePair<TopicData, Task<RedisValue>>>();
            var hashTasks = new List<KeyValuePair<TopicData, Task<HashEntry[]>>>();
            foreach (var td in topics)
            {
                if (IsMalformed(td))
                    continue;
                if (td.Type == "HGET")
                    singleTasks.Add(new KeyValuePair<TopicData, Task<RedisValue>>(td, batch.HashGetAsync(td.KeyOrChannel, td.Field)));
                else if (td.Type == "HGETALL")
                    hashTasks.Add(new KeyValuePair<TopicData, Task<HashEntry[]>>(td, batch.HashGetAllAsync(td.KeyOrChannel)));
                else if (td.Type == "GET")
                    singleTasks.Add(new KeyValuePair<TopicData, Task<RedisValue>>(td, batch.StringGetAsync(td.KeyOrChannel)));
            }
            try
            {
                batch.Execute();
            }
            catch (Exception ex)
            {
                // The batch failed as a whole (connection lost, disposed mux...):
                // report it for this host and fall through so the per-item loops
                // below still observe/drain each task and nothing is lost
                // silently (their per-item try/catch logs the individual faults).
                logger.Error(ex, $"PollHost: batch execute failed, host={AppConfig.MaskHost(host)}");
            }
            int responseTimeoutMs = RedisUDF.ResponseTimeoutMs();
            DrainBatch(singleTasks, responseTimeoutMs, host,
                // HGET and GET (UseGetMultiple=false): IsNull distinguishes a missing
                // key/field from an existing one holding the empty string (same as MGET).
                value => value.IsNull ? "(no value)" : value.ToString(),
                (td, value) => td.HasChangedPolledValue(value),
                (td, value) => td.CommitPolledValue(value));
            DrainBatch(hashTasks, responseTimeoutMs, host,
                RedisResultFormatter.FormatHash,
                (td, entries) => td.HasChangedPolledHash(entries),
                (td, entries) => td.CommitPolledHash(entries));
        }

        /// <summary>
        /// Drains one pipelined batch with a hard response bound: an orphaned task
        /// (multiplexer disposed mid-execute) surfaces as a timeout instead of
        /// blocking the tick thread forever. The skip/publish/commit order and the
        /// per-item error isolation are unchanged; the log line uses the topic type
        /// (HGET/HGETALL/GET), matching the old per-loop texts.
        /// </summary>
        private void DrainBatch<T>(
            List<KeyValuePair<TopicData, Task<T>>> tasks, int responseTimeoutMs, string host,
            Func<T, string> format, Func<TopicData, T, bool> changed, Action<TopicData, T> commit)
        {
            foreach (var pair in tasks)
            {
                try
                {
                    if (!RedisUDF.WaitBounded(pair.Value, responseTimeoutMs))
                        throw new TimeoutException("no reply within the response timeout (batch left incomplete after a dispose?)");
                    var value = pair.Value.GetAwaiter().GetResult();
                    if (_skipRepeatedMessages && !changed(pair.Key, value))
                        continue;
                    Publish(pair.Key, format(value));
                    commit(pair.Key, value);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: {pair.Key.Type} key={pair.Key.KeyOrChannel}, host={AppConfig.MaskHost(host)}");
                }
            }
        }

        private void Publish(TopicData td, string value)
        {
            if (td.Disconnected)
                return;
            Interlocked.Increment(ref _messageCount);
            if (logger.IsTraceEnabled)
                logger.Trace($"Publish: {td.Type} host={AppConfig.MaskHost(td.Host)}, key={td.KeyOrChannel}, field={td.Field}, value={value}");
            DeliverToTopic(td, value);
        }

        /// <summary>
        /// Single owner of the realtime-vs-coalesced delivery policy: a realtime
        /// push only when coalescing is off, otherwise the value waits for the
        /// Excel tick (dirty flag). Used by the subscription callback and Publish.
        /// </summary>
        private void DeliverToTopic(TopicData td, string value)
        {
            if (_realTimeUpdates && !_coalesceRealtimeUpdates)
                td.UpdateAndSendToExcel(value);
            else
                td.UpdateOnly(value);
        }
    }

    public static class RedisRtdStatus
    {
        // Every helper reports a neutral value instead of throwing after
        // RedisRuntime.Shutdown: the runtime refuses to resurrect the managers,
        // so an unguarded access would surface InvalidOperationException in the
        // cell during Excel teardown/reload. The public Excel surface (names
        // and signatures) is unchanged.

        /// <summary>Runs the status read, replacing a teardown/reload exception
        /// with the neutral value the cell used to get from its own catch.</summary>
        private static T Guarded<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch
            {
                return fallback;
            }
        }

        [ExcelFunction(Description = "Returns the number of active Redis connections.", IsVolatile = true)]
        public static int RedisRTDConnectionCount() => Guarded(() => RedisRtd.RedisConnectionsCount(), 0);

        [ExcelFunction(Description = "Returns the number of active Redis subscriptions.", IsVolatile = true)]
        public static int RedisRTDSubscriptionCount() => Guarded(() => RedisRtd.RedisSubscriptionsCount(), 0);

        [ExcelFunction(Description = "Returns the total number of active Excel RTD topics.", IsVolatile = true)]
        public static int RedisRTDTopicCount() => Guarded(() => RedisRtd.TopicsCount(), 0);

        [ExcelFunction(Description = "Returns the number of Redis channels with subscriptions.", IsVolatile = true)]
        public static int RedisRTDChannelCount() => Guarded(() => RedisRtd.ChannelTopicsCount(), 0);

        [ExcelFunction(Description = "Returns the default Redis host address used by the RTD server.", IsVolatile = true)]
        public static string RedisRTDDefaultHost() => Guarded(() => AppConfig.MaskHost(RedisRtd.DefaultHost()) ?? "", "");

        [ExcelFunction(Description = "Returns the Excel update interval in milliseconds.", IsVolatile = true)]
        public static double RedisRTDExcelUpdateInterval() => Guarded(() => RedisRtd.ExcelUpdateRate(), 0);

        [ExcelFunction(Description = "Returns the Redis polling interval in milliseconds.", IsVolatile = true)]
        public static double RedisRTDRedisUpdateInterval() => Guarded(() => RedisRtd.RedisUpdateRate(), 0);

        [ExcelFunction(Description = "Returns TRUE if real-time updates are enabled, FALSE otherwise.", IsVolatile = true)]
        public static bool RedisRTDRealTimeUpdates() => Guarded(() => RedisRtd.IsRealTimeEnabled(), false);

        [ExcelFunction(Description = "Returns last messages/second counter", IsVolatile = true)]
        public static long RedisRTDMessagesCounter() => Guarded(() => RedisRtd.CurrentMessagesCounter(), 0);
    }
}

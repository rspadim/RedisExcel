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
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
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
        private string _lastValue;
        private bool _dirty = true;

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
        public IDisposable Subscription { get; set; }

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
                    _dirty = false;
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

        public void SendToExcelIfDirty()
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
                try
                {
                    // Push while holding the lock: a concurrent state update
                    // cannot interleave and overwrite the newer value. On
                    // failure stay dirty so the next Excel tick retries.
                    Topic.UpdateValue(_lastValue);
                    _dirty = false;
                }
                catch (Exception ex)
                {
                    // _dirty stays true; log outside the lock so a slow log
                    // write cannot delay other producers of this topic.
                    updateError = ex;
                }
            }
            if (updateError != null)
                logger.Error(updateError, "SendToExcelIfDirty: update failed");
        }

        public override string ToString()
        {
            return $"TopicId={Topic.TopicId}, Type={Type}, KeyOrChannel={KeyOrChannel}, Field={Field}, Host={Host}, LastValue={LastValue}, Dirty={Dirty}";
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

        private readonly ConcurrentDictionary<int, TopicData> _polledTopics = new ConcurrentDictionary<int, TopicData>();
        private readonly ConcurrentDictionary<int, TopicData> _subscribedTopics = new ConcurrentDictionary<int, TopicData>();

        private long _messageCount;
        private long _messageCounterThreshold;
        private volatile bool _realTimeUpdates = true;
        private ENUMExcelUpdateStyle _excelUpdateStyle = ENUMExcelUpdateStyle.Automatic;
        private double _excelUpdateRateMs = 100;
        private double _redisUpdateRateMs = 1000;
        private bool _useGetMultiple = true;
        private bool _skipRepeatedMessages = true;
        private bool _coalesceRealtimeUpdates = true;
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
            return instance == null ? 0 : Interlocked.Read(ref instance._messageCount);
        }

        public static int RedisConnectionsCount() => RedisRuntime.Connections.RtdConnectionCount;

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
            // in Realtime/Timer styles the mode is fixed; in Automatic it is recalculated from the message counter
            _realTimeUpdates = _excelUpdateStyle != ENUMExcelUpdateStyle.Timer;

            logger.Info(
                "ServerStart: config " +
                $"host={_defaultHost}, excelRate={_excelUpdateRateMs}ms, redisRate={_redisUpdateRateMs}ms, " +
                $"style={_excelUpdateStyle}, threshold={_messageCounterThreshold}, useGetMultiple={_useGetMultiple}");

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
            foreach (var td in _polledTopics.Values)
                td.Disconnected = true;
            foreach (var td in _subscribedTopics.Values)
                td.Disconnected = true;

            DisposeTimer(ref _redisTimer);
            DisposeTimer(ref _excelTimer);
            DisposeTimer(ref _counterTimer);

            foreach (var td in _subscribedTopics.Values)
            {
                try
                {
                    td.Subscription?.Dispose();
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
                    Instance = Instances.Keys.FirstOrDefault();
            }
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
                logger.Info($"ConnectData: command={command}, param2={param2}, param3={param3}, param4={param4}, TopicId={topic.TopicId}");

                // Redis names (keys/fields) may be empty or whitespace; only a
                // missing (null) argument is invalid here (StackExchange.Redis
                // rejects null for keys/fields). Empty or whitespace SUB/PSUB
                // channels are still rejected later by the subscription manager.
                TopicData td;
                switch (command)
                {
                    case "GET":
                    case "HGETALL":
                        if (param2 == null)
                            throw new Exception($"{command} requires a key as the second argument");
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "HGET":
                        if (param2 == null)
                            throw new Exception("HGET requires a key as the second argument");
                        if (param3 == null)
                            throw new Exception("HGET requires a field as the third argument");
                        td = new TopicData(topic, command, param2, param3, AppConfig.ResolveRtdHost(param4));
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "SUB":
                    case "PSUB":
                        if (param2 == null)
                            throw new Exception($"{command} requires a channel as the second argument");
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        // Register the subscription first: if it throws, the catch below
                        // returns the error and no topic is stored.
                        td.Subscription = Subscribe(td);
                        _subscribedTopics[topic.TopicId] = td;
                        break;
                    default:
                        throw new Exception($"unknown command '{command}', expected one of [GET, HGET, HGETALL, SUB, PSUB]");
                }

                logger.Info($"ConnectData: accepted {td}");
                return "(ConnectData)";
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ConnectData: error");
                return $"#ERROR: ConnectData: {ex.Message}";
            }
        }

        protected override void DisconnectData(Topic topic)
        {
            try
            {
                if (_subscribedTopics.TryRemove(topic.TopicId, out var sub))
                {
                    // Mark first so in-flight callbacks stop publishing before the
                    // shared subscription is torn down.
                    sub.Disconnected = true;
                    sub.Subscription?.Dispose();
                    sub.Subscription = null;
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
                    logger.Error($"DisconnectData: unknown TopicId={topic.TopicId}");
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
                if (_realTimeUpdates && !_coalesceRealtimeUpdates)
                    td.UpdateAndSendToExcel(message);
                else
                    td.UpdateOnly(message);
            }, "RTD");
            logger.Info($"Subscribe: subscribed {td}");
            return subscription;
        }

        private void OnCounterTick()
        {
            if (!_counterTickGate.TryEnter())
            {
                if (logger.IsDebugEnabled)
                    logger.Debug("OnCounterTick: previous tick still running, skipping");
                return;
            }
            try
            {
                UpdateRealtimeMode(allowReenable: true);
                Interlocked.Exchange(ref _messageCount, 0);
            }
            finally
            {
                _counterTickGate.Exit();
            }
        }

        private void OnExcelTick()
        {
            if (!_excelTickGate.TryEnter())
            {
                if (logger.IsDebugEnabled)
                    logger.Debug("OnExcelTick: previous tick still running, skipping");
                return;
            }
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
                foreach (var td in _polledTopics.Values)
                    td.SendToExcelIfDirty();
                foreach (var td in _subscribedTopics.Values)
                    td.SendToExcelIfDirty();
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
            if (!_redisTickGate.TryEnter())
            {
                if (logger.IsDebugEnabled)
                    logger.Debug("OnRedisTick: previous tick still running, skipping");
                return;
            }
            try
            {
                if (_polledTopics.IsEmpty)
                    return;
                if (logger.IsDebugEnabled)
                    logger.Debug($"OnRedisTick: polling {_polledTopics.Count} topic(s), rate={_redisUpdateRateMs}ms");
                foreach (var group in _polledTopics.Values.GroupBy(t => t.Host))
                {
                    try
                    {
                        PollHost(group.Key, group.ToList());
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"OnRedisTick: host={group.Key}");
                    }
                }
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
                    logger.Warn($"PollHost: skipping malformed topic, TopicId={td.Topic.TopicId}, Type={td.Type}, host={host}");
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
                            logger.Trace($"PollHost: GETMULTI host={host}, keys=[{string.Join(", ", gets.Select(t => t.KeyOrChannel))}]");
                        var values = db.StringGet(keys);
                        for (int i = 0; i < gets.Count; i++)
                        {
                            var td = gets[i];
                            if (_skipRepeatedMessages && !td.HasChangedPolledValue(values[i]))
                                continue;
                            // Per-item try so one failing Excel push cannot abort the whole MGET batch.
                            try
                            {
                                Publish(td, values[i].HasValue ? values[i].ToString() : "(no value)");
                                td.CommitPolledValue(values[i]);
                            }
                            catch (Exception ex)
                            {
                                logger.Error(ex, $"PollHost: GET key={td.KeyOrChannel}, host={host}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: GETMULTI host={host}");
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
            batch.Execute();
            foreach (var pair in singleTasks)
            {
                try
                {
                    var value = pair.Value.GetAwaiter().GetResult();
                    if (_skipRepeatedMessages && !pair.Key.HasChangedPolledValue(value))
                        continue;
                    Publish(pair.Key, value.HasValue ? value.ToString() : "(no value)");
                    pair.Key.CommitPolledValue(value);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: {pair.Key.Type} key={pair.Key.KeyOrChannel}, host={host}");
                }
            }
            foreach (var pair in hashTasks)
            {
                try
                {
                    var entries = pair.Value.GetAwaiter().GetResult();
                    if (_skipRepeatedMessages && !pair.Key.HasChangedPolledHash(entries))
                        continue;
                    Publish(pair.Key, RedisResultFormatter.FormatHash(entries));
                    pair.Key.CommitPolledHash(entries);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: HGETALL key={pair.Key.KeyOrChannel}, host={host}");
                }
            }
        }

        private void Publish(TopicData td, string value)
        {
            if (td.Disconnected)
                return;
            Interlocked.Increment(ref _messageCount);
            if (logger.IsTraceEnabled)
                logger.Trace($"Publish: {td.Type} host={td.Host}, key={td.KeyOrChannel}, field={td.Field}, value={value}");
            if (_realTimeUpdates && !_coalesceRealtimeUpdates)
                td.UpdateAndSendToExcel(value);
            else
                td.UpdateOnly(value);
        }
    }

    public static class RedisRtdStatus
    {
        [ExcelFunction(Description = "Returns the number of active Redis connections.", IsVolatile = true)]
        public static int RedisRTDConnectionCount()
        {
            return RedisRtd.RedisConnectionsCount();
        }

        [ExcelFunction(Description = "Returns the number of active Redis subscriptions.", IsVolatile = true)]
        public static int RedisRTDSubscriptionCount()
        {
            return RedisRtd.RedisSubscriptionsCount();
        }

        [ExcelFunction(Description = "Returns the total number of active Excel RTD topics.", IsVolatile = true)]
        public static int RedisRTDTopicCount()
        {
            return RedisRtd.TopicsCount();
        }

        [ExcelFunction(Description = "Returns the number of Redis channels with subscriptions.", IsVolatile = true)]
        public static int RedisRTDChannelCount()
        {
            return RedisRtd.ChannelTopicsCount();
        }

        [ExcelFunction(Description = "Returns the default Redis host address used by the RTD server.", IsVolatile = true)]
        public static string RedisRTDDefaultHost()
        {
            return RedisRtd.DefaultHost();
        }

        [ExcelFunction(Description = "Returns the Excel update interval in milliseconds.", IsVolatile = true)]
        public static double RedisRTDExcelUpdateInterval()
        {
            return RedisRtd.ExcelUpdateRate();
        }

        [ExcelFunction(Description = "Returns the Redis polling interval in milliseconds.", IsVolatile = true)]
        public static double RedisRTDRedisUpdateInterval()
        {
            return RedisRtd.RedisUpdateRate();
        }

        [ExcelFunction(Description = "Returns TRUE if real-time updates are enabled, FALSE otherwise.", IsVolatile = true)]
        public static bool RedisRTDRealTimeUpdates()
        {
            return RedisRtd.IsRealTimeEnabled();
        }

        [ExcelFunction(Description = "Returns last messages/second counter", IsVolatile = true)]
        public static long RedisRTDMessagesCounter()
        {
            return RedisRtd.CurrentMessagesCounter();
        }
    }
}

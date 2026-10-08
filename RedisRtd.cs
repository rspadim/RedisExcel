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

        public string LastValue { get { lock (_sync) return _lastValue; } }

        public bool Dirty { get { lock (_sync) return _dirty; } }

        public void UpdateAndSendToExcel(string data)
        {
            Topic.UpdateValue(data);
            lock (_sync)
            {
                _lastValue = data;
                _dirty = false;
            }
        }

        public void UpdateOnly(string data)
        {
            lock (_sync)
            {
                _lastValue = data;
                _dirty = true;
            }
        }

        public void SendToExcelIfDirty()
        {
            string value;
            lock (_sync)
            {
                if (!_dirty)
                    return;
                _dirty = false;
                value = _lastValue;
            }
            Topic.UpdateValue(value);
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
        private static RedisRtd Instance;

        private readonly ConcurrentDictionary<int, TopicData> _polledTopics = new ConcurrentDictionary<int, TopicData>();
        private readonly ConcurrentDictionary<int, TopicData> _subscribedTopics = new ConcurrentDictionary<int, TopicData>();

        private long _messageCount;
        private long _messageCounterThreshold;
        private volatile bool _realTimeUpdates = true;
        private ENUMExcelUpdateStyle _excelUpdateStyle = ENUMExcelUpdateStyle.Automatic;
        private double _excelUpdateRateMs = 100;
        private double _redisUpdateRateMs = 1000;
        private bool _useGetMultiple = true;
        private string _defaultHost;

        private System.Timers.Timer _excelTimer;
        private System.Timers.Timer _redisTimer;
        private System.Timers.Timer _counterTimer;

        public static long CurrentMessagesCounter()
        {
            var instance = Instance;
            return instance == null ? 0 : Interlocked.Read(ref instance._messageCount);
        }

        public static int RedisConnectionsCount() => RedisRuntime.Connections.RtdConnectionCount;

        public static int RedisSubscriptionsCount() => RedisRuntime.Subscriptions.ListenerCount;

        public static int TopicsCount()
        {
            int total = 0;
            foreach (var instance in Instances.Keys)
                total += instance._polledTopics.Count + instance._subscribedTopics.Count;
            return total;
        }

        public static int ChannelTopicsCount() => RedisRuntime.Subscriptions.ChannelCount;

        public static string DefaultHost() => Instance?._defaultHost;

        public static double ExcelUpdateRate() => Instance?._excelUpdateRateMs ?? 0;

        public static double RedisUpdateRate() => Instance?._redisUpdateRateMs ?? 0;

        public static bool IsRealTimeEnabled() => Instance?._realTimeUpdates ?? false;

        protected override bool ServerStart()
        {
            logger.Info("ServerStart: starting RTD server");
            Instance = this;
            Instances[this] = 0;

            var config = AppConfig.Current.RTD;
            _defaultHost = config.host;
            _excelUpdateRateMs = config.ExcelUpdateRateMs;
            _redisUpdateRateMs = config.RedisUpdateRateMs;
            _excelUpdateStyle = config.ExcelUpdateStyle;
            _messageCounterThreshold = config.MessageCounterThreshold;
            _useGetMultiple = config.UseGetMultiple;
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

            Instances.TryRemove(this, out _);
            if (ReferenceEquals(Instance, this))
                Instance = Instances.Keys.FirstOrDefault();
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

                TopicData td;
                switch (command)
                {
                    case "GET":
                    case "HGETALL":
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "HGET":
                        td = new TopicData(topic, command, param2, param3, AppConfig.ResolveRtdHost(param4));
                        _polledTopics[topic.TopicId] = td;
                        break;
                    case "SUB":
                    case "PSUB":
                        td = new TopicData(topic, command, param2, null, AppConfig.ResolveRtdHost(param3));
                        _subscribedTopics[topic.TopicId] = td;
                        Subscribe(td);
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
                    sub.Subscription?.Dispose();
                    sub.Subscription = null;
                    logger.Info($"DisconnectData: removed subscription {sub}");
                }
                else if (_polledTopics.TryRemove(topic.TopicId, out var polled))
                {
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

        private void Subscribe(TopicData td)
        {
            bool pattern = td.Type == "PSUB";
            long topicId = td.Topic.TopicId;
            td.Subscription = RedisRuntime.Subscriptions.Subscribe(td.Host, td.KeyOrChannel, pattern, message =>
            {
                Interlocked.Increment(ref _messageCount);
                if (logger.IsTraceEnabled)
                    logger.Trace($"Subscribe: TopicId={topicId}, channel={td.KeyOrChannel}, message={message}");
                if (_realTimeUpdates)
                    td.UpdateAndSendToExcel(message);
                else
                    td.UpdateOnly(message);
            });
            logger.Info($"Subscribe: subscribed {td}");
        }

        private void OnCounterTick()
        {
            UpdateRealtimeMode(allowReenable: true);
            Interlocked.Exchange(ref _messageCount, 0);
        }

        private void OnExcelTick()
        {
            UpdateRealtimeMode(allowReenable: false);
            if (_realTimeUpdates)
                return;
            if (_polledTopics.IsEmpty && _subscribedTopics.IsEmpty)
                return;
            if (logger.IsDebugEnabled)
                logger.Debug($"OnExcelTick: flushing dirty topics, polled={_polledTopics.Count}, subscribed={_subscribedTopics.Count}");
            foreach (var td in _polledTopics.Values)
                td.SendToExcelIfDirty();
            foreach (var td in _subscribedTopics.Values)
                td.SendToExcelIfDirty();
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

        private void PollHost(string host, List<TopicData> topics)
        {
            var db = RedisRuntime.Connections.GetDatabase(host, RedisPool.RtdData);

            // GET in batch (MGET): one round-trip for all GETs of the host
            if (_useGetMultiple)
            {
                var gets = topics.Where(t => t.Type == "GET").ToList();
                if (gets.Count > 0)
                {
                    var keys = gets.Select(t => (RedisKey)t.KeyOrChannel).ToArray();
                    if (logger.IsTraceEnabled)
                        logger.Trace($"PollHost: GETMULTI host={host}, keys=[{string.Join(", ", gets.Select(t => t.KeyOrChannel))}]");
                    var values = db.StringGet(keys);
                    for (int i = 0; i < gets.Count; i++)
                        Publish(gets[i], values[i].HasValue ? values[i].ToString() : "(no value)");
                }
                topics = topics.Where(t => t.Type != "GET").ToList();
            }
            if (topics.Count == 0)
                return;

            // HGET/HGETALL (and GET when UseGetMultiple=false) in a pipeline
            var batch = db.CreateBatch();
            var singleTasks = new List<KeyValuePair<TopicData, Task<RedisValue>>>();
            var hashTasks = new List<KeyValuePair<TopicData, Task<HashEntry[]>>>();
            foreach (var td in topics)
            {
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
                    Publish(pair.Key, value.HasValue ? value.ToString() : "(no value)");
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
                    Publish(pair.Key, RedisResultFormatter.FormatHash(entries));
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"PollHost: HGETALL key={pair.Key.KeyOrChannel}, host={host}");
                }
            }
        }

        private void Publish(TopicData td, string value)
        {
            Interlocked.Increment(ref _messageCount);
            if (logger.IsTraceEnabled)
                logger.Trace($"Publish: {td.Type} host={td.Host}, key={td.KeyOrChannel}, field={td.Field}, value={value}");
            if (_realTimeUpdates)
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

using NLog;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using static ExcelDna.Integration.Rtd.ExcelRtdServer;
using System;
using System.Collections.Concurrent;
using StackExchange.Redis;
using System.Linq;

namespace RedisExcel
{
    public class AddIn : IExcelAddIn {
        public void AutoOpen(){ ComServer.DllRegisterServer(); }

        public void AutoClose() { ComServer.DllUnregisterServer(); }
    }
    public class TopicData
    {
        public Topic Topic { get; set; } // Armazena o Topic para notificação
        public string Type { get; set; } = null;
        public string KeyOrChannel { get; set; } = null;
        public string Field { get; set; } = null;
        public string Host { get; set; } = null;
        public string LastValue { get; set; } = null;
        public bool Dirty { get; set; } = true;
        public string GetSubHost()
        {
            return (Type == "SUB" || Type == "PSUB" ? Host : null);
        }
        public string GetChannel()
        {
            return (Type == "SUB" || Type == "PSUB" ? Host + "::" + Type + "::" + KeyOrChannel : null);
        }
        public override string ToString()
        {
            return $"TopicId={Topic.TopicId}, Type={Type}, KeyOrChannel={KeyOrChannel}, Field={Field}, Host={Host}, LastValue={LastValue}, Dirty={Dirty}";
        }
        public void SendToExcelIfDirty()
        {
            if (this.Dirty)
                this.Topic.UpdateValue(this.LastValue);
            this.Dirty = false;
        }
        public void UpdateOnly(string data)
        {
            this.LastValue = data;
            this.Dirty = true;
        }
        public void UpdateAndSendToExcel(string data)
        {
            this.Topic.UpdateValue(data);
            this.LastValue = data;
            this.Dirty = false;
        }
        public void UpdateAndSendOnlyIfNew(string data)
        {
            if (this.LastValue == data && !this.Dirty)
                return;
            this.UpdateAndSendToExcel(data);
        }
    }

    [ComVisible(true)]
    [ProgId("RedisRtd")]
    public class RedisRtd : ExcelRtdServer
    {
        private static Dictionary<string, string> Servers = null;
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private System.Timers.Timer _timerCounter;
        private System.Timers.Timer _timerRedis;
        private System.Timers.Timer _timerUpdateExcel;
        
        public bool UseGetMultiple = true;
        private ENUMExcelUpdateStyle ExcelUpdateStyle = ENUMExcelUpdateStyle.Automatic;
        private bool RealTimeUpdates = true;
        private double ExcelUpdateRateMS = 100;
        private double RedisUpdateRateMS = 100;
        private object _MessageCounter_lock = new object();
        private long MessageCounterThreshold = 0;
        private long MessageCounter = 0;
        private string ConfigDefaultHost = null;

        private readonly Dictionary<int, TopicData> _topicsSubs = new Dictionary<int, TopicData>();
        private readonly Dictionary<int, TopicData> _topics = new Dictionary<int, TopicData>();

        private static readonly ConcurrentDictionary<string, ConnectionMultiplexer> _redisDataConnections = new ConcurrentDictionary<string, ConnectionMultiplexer>();
        private static readonly ConcurrentDictionary<string, ConnectionMultiplexer> _redisSubConnections = new ConcurrentDictionary<string, ConnectionMultiplexer>();

        private static readonly ConcurrentDictionary<string, ISubscriber> _redisSubscriptions = new ConcurrentDictionary<string, ISubscriber>();
        private static readonly ConcurrentDictionary<string, HashSet<int>> _channelTopics = new ConcurrentDictionary<string, HashSet<int>>();

        public static long CurrentMessagesCounter()
        {
            if (Instance == null)
                return 0;
            lock (Instance._MessageCounter_lock)
            {
                return Instance.MessageCounter;
            }
        }
        public static int RedisConnectionsCount() => _redisDataConnections.Count + _redisSubConnections.Count;
        public static int RedisSubscriptionsCount() => _redisSubscriptions.Count;
        public static int TopicsCount() => (Instance?._topics.Count ?? 0) + (Instance?._topicsSubs.Count ?? 0);
        public static int ChannelTopicsCount() => _channelTopics.Count;
        public static string DefaultHost() => Instance?.ConfigDefaultHost ?? null;
        public static double ExcelUpdateRate() => Instance?.ExcelUpdateRateMS ?? 0;
        public static double RedisUpdateRate() => Instance?.RedisUpdateRateMS ?? 0;
        public static bool IsRealTimeEnabled() => Instance?.RealTimeUpdates ?? false;
        private static RedisRtd Instance;

        private void LoadConfig()
        {
            var config = ConfigHelper.GetConfig();
            this.ConfigDefaultHost = config.RTD.host;
            this.ExcelUpdateRateMS = config.RTD.ExcelUpdateRateMs;
            this.RedisUpdateRateMS = config.RTD.RedisUpdateRateMs;
            this.ExcelUpdateStyle = config.RTD.ExcelUpdateStyle;
            this.MessageCounterThreshold = config.RTD.MessageCounterThreshold;
            this.UseGetMultiple = config.RTD.UseGetMultiple;
            RedisRtd.Servers = config.Servers;
        }
        private string FindServerName(string host)
        {
            if (RedisRtd.Servers == null)
                this.LoadConfig();
            if (RedisRtd.Servers != null && RedisRtd.Servers.TryGetValue(host, out var server))
                return server;
            return host;
        }
        private string GetDefaultHost()
        {
            if (this.ConfigDefaultHost == null) {
                var config = ConfigHelper.GetConfig();
                this.ConfigDefaultHost = config.RTD.host;
            }
            return this.ConfigDefaultHost;
        }
        private ConnectionMultiplexer GetOrCreateRedis(string host, string subHost = null)
        {
            host = string.IsNullOrWhiteSpace(host) ? GetDefaultHost() : host;
            host = FindServerName(host);
            bool forSubscription = subHost != null;
            string connectionKey = forSubscription ? subHost : host;
            var pool = forSubscription ? _redisSubConnections : _redisDataConnections;

            if (!pool.ContainsKey(host))
            {
                lock (pool)
                {
                    if (!pool.ContainsKey(host))
                    {
                        var config = ConfigurationOptions.Parse(host);
                        config.ConnectTimeout = 1000;
                        config.AbortOnConnectFail = false;
                        string GitTag = GitVersion.Tag;
                        config.ClientName = $"RedisRTD :: {GitTag} :: {Environment.UserDomainName}\\{Environment.UserName} :: {Environment.MachineName}";
                        if (logger.IsInfoEnabled)
                            logger.Info($"GetOrCreateRedis: Creating Redis connection to {host}, key={subHost}, ClientName={config.ClientName}");
                        pool[connectionKey] = ConnectionMultiplexer.Connect(config);
                        pool[connectionKey].ConnectionFailed += (sender, args) =>
                        {
                            if (logger.IsInfoEnabled)
                                logger.Info($"GetOrCreateRedis: LOST connection to Redis. host={host}, Endpoint={args.EndPoint}, FailureType={args.FailureType}, Exception={args.Exception?.Message}");
                        };
                        pool[connectionKey].ConnectionRestored += (sender, args) =>
                        {
                            if (logger.IsInfoEnabled)
                                logger.Info($"GetOrCreateRedis: Redis reconnect detected for host={host}");
                            ResubscribeHost(host);
                        };
                    }
                }
            }

            return pool[connectionKey];
        }
        protected override bool ServerStart()
        {
            if (logger.IsInfoEnabled)
                logger.Info("ServerStart: Starting RTD Server");
            Instance = this;
            // carrega configuracoes
            LoadConfig();
            if (logger.IsInfoEnabled)
                logger.Info(
                    "Loaded configuration:\n"+
                    $"    ConfigDefaultHost        = {this.ConfigDefaultHost}\n"+
                    $"    ExcelUpdateRateMS        = {this.ExcelUpdateRateMS}\n"+
                    $"    RedisUpdateRateMS        = {this.RedisUpdateRateMS}\n"+
                    $"    ExcelUpdateStyle         = {this.ExcelUpdateStyle}\n"+
                    $"    MessageCounterThreshold  = {this.MessageCounterThreshold}\n"+
                    $"    UseGetMultiple           = {this.UseGetMultiple}"
                );

            // inicializa timers
            _timerCounter = new System.Timers.Timer(1000);
            _timerCounter.AutoReset = true;
            _timerCounter.Elapsed += TimerElapsedMessageCounter;

            _timerUpdateExcel = new System.Timers.Timer(ExcelUpdateRateMS);
            _timerUpdateExcel.AutoReset = true;
            _timerUpdateExcel.Elapsed += TimerElapsedExcel;

            _timerRedis = new System.Timers.Timer(RedisUpdateRateMS);
            _timerRedis.AutoReset = true;
            _timerRedis.Elapsed += TimerElapsedRedis;

            // start
            _timerRedis.Start();
            _timerUpdateExcel.Start();
            _timerCounter.Start();
            return true;
        }

        protected override void ServerTerminate()
        {
            if (logger.IsInfoEnabled)
                logger.Info($"ServerTerminate");
            // matar os timers
            _timerUpdateExcel?.Stop();
            _timerRedis?.Dispose();

            // TODO: matar as conexoes do redis
        }
        private void ResubscribeHost(string host)
        {
            if (!_redisSubConnections.ContainsKey(host))
            {
                if (logger.IsErrorEnabled)
                    logger.Error($"ResubscribeHost: host not found {host}");
                return;
            }
            // remove todos subscriptions, e todos channeltopics
            var topics = _topicsSubs
                .Where(kvp => kvp.Value.Host == host)
                .Select(kvp => kvp.Value)
                .ToList();
            if (logger.IsInfoEnabled)
                logger.Info($"ResubscribeHost: Cleaning topics {topics.Count} topic(s) for host={host}");
            foreach (var td in topics)
            {
                var subHost = td.GetSubHost();
                try
                {
                    _channelTopics[subHost].Clear();
                    if (_redisSubscriptions.TryRemove(subHost, out var _oldSubscription))
                        _oldSubscription.UnsubscribeAll();
                }
                catch (Exception ex)
                {
                    if (logger.IsErrorEnabled)
                        logger.Error(ex, $"ResubscribeHost: Failed to remove subHost={subHost}");
                }
            }
            // adicionar todos subscriptions e channeltopics novamente
            if (logger.IsInfoEnabled)
                logger.Info($"ResubscribeHost: Re-subscribing {topics.Count} topic(s) for host={host}");
            foreach (var topic in topics)
            {
                var topicId = topic.Topic.TopicId;
                try
                {
                    SubscribeTopicId(topicId);
                }
                catch (Exception ex)
                {
                    if (logger.IsErrorEnabled)
                        logger.Error(ex, $"ResubscribeHost: Failed to resubscribe TopicId={topicId}, host={host}, TopidData={topic}");
                }
            }
        }
        private void SubscribeTopicId(int topicId)
        {
            if (!_topicsSubs.ContainsKey(topicId))
            {
                if (logger.IsErrorEnabled)
                    logger.Error($"Subscribe: TopicId={topicId} not found");
                return;
            }
            var td = _topicsSubs[topicId];
            string subHost = td.GetSubHost();
            string channelName = td.GetChannel();
            // obtem a conexao
            var conn = GetOrCreateRedis(td.Host, subHost);
            // cria canal
            var redisChannel = new RedisChannel(
                td.KeyOrChannel, 
                (td.Type == "SUB" ? RedisChannel.PatternMode.Literal : RedisChannel.PatternMode.Pattern)
            );
            // obtem o objeto de subscription da conexao, se nao existir cria um novo
            ISubscriber subscriber;
            if (!_redisSubscriptions.TryGetValue(subHost, out subscriber))
            {
                subscriber = conn.GetSubscriber();
                _redisSubscriptions[subHost] = subscriber;
            }
            // adiciona o topico no canal da conexao
            lock (_channelTopics)
            {
                if (!_channelTopics.ContainsKey(subHost))
                    _channelTopics[subHost] = new HashSet<int>();
                if (_channelTopics[subHost].Contains(topicId))
                {
                    // ja existe
                    if (logger.IsWarnEnabled)
                        logger.Warn($"Subscribe: TopicId={topicId} already subscribed on subHost={subHost}");
                    return;
                }
                // TODO: verificar se topicId é o melhor caminho ou td.GetChannel(), caso dois topicId gerem o mesmo GetChannel()
                _channelTopics[subHost].Add(topicId);
            }
            subscriber.Subscribe(
                redisChannel,
                (channel, message) => {
                    lock (_MessageCounter_lock)
                    {
                        MessageCounter++;
                    }
                    if (logger.IsTraceEnabled)
                        logger.Trace($"Subscribe: topicId={topicId}, channel={channel}, message={message}");
                    // TODO: verificar se topicId é o melhor caminho ou td.GetChannel(), caso dois topicId gerem o mesmo GetChannel()
                    if (_topicsSubs.TryGetValue(topicId, out TopicData topicData))
                    {
                        if (RealTimeUpdates)
                            topicData.UpdateAndSendToExcel(message);
                        else
                            topicData.UpdateOnly(message);
                    }
                }
            );
            logger.Info($"Subscribe: Subscribed TopicData={td}");
        }
        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            try
            {
                if (_topicsSubs.ContainsKey(topic.TopicId))
                {
                    if (logger.IsWarnEnabled)
                        logger.Warn($"ConnectData: TopicId already exists in Subscribe, TopicId={topic.TopicId}");
                    return _topicsSubs[topic.TopicId].LastValue;
                }
                else if (_topics.ContainsKey(topic.TopicId))
                {
                    if (logger.IsWarnEnabled)
                        logger.Warn($"ConnectData: TopicId already exists in Topics, TopicId={topic.TopicId}");
                    return _topics[topic.TopicId].LastValue;
                }
                string param1 = topicInfo.Count > 0 ? topicInfo[0].ToUpper().Trim() : null;
                string param2 = topicInfo.Count > 1 ? topicInfo[1] : null;
                string param3 = topicInfo.Count > 2 ? topicInfo[2] : null;
                string param4 = topicInfo.Count > 3 ? topicInfo[3] : null;
                if (logger.IsInfoEnabled)
                    logger.Info($"ConnectData: New data request: param1={param1}, param2={param2}, param3={param3}, param4={param4}, TopicId={topic.TopicId}");
                bool sub = false;
                switch (param1){
                    case "GET":
                    case "HGETALL":
                        _topics[topic.TopicId] = new TopicData
                        {
                            Type = param1,
                            KeyOrChannel = param2,
                            Host = FindServerName(string.IsNullOrWhiteSpace(param3) ? GetDefaultHost() : param3),
                            Topic = topic
                        };
                        break;
                    case "HGET":
                        _topics[topic.TopicId] = new TopicData
                        {
                            Type = param1,
                            KeyOrChannel = param2,
                            Field = param3,
                            Host = FindServerName(string.IsNullOrWhiteSpace(param4) ? GetDefaultHost() : param4),
                            Topic = topic
                        };
                        break;
                    case "PSUB":
                    case "SUB":
                        sub = true;
                        _topicsSubs[topic.TopicId] = new TopicData
                        {
                            Type = param1,
                            KeyOrChannel = param2,
                            Host = FindServerName(string.IsNullOrWhiteSpace(param3) ? GetDefaultHost() : param3),
                            Topic = topic
                        };
                        SubscribeTopicId(topic.TopicId);
                        break;
                    default:
                        throw new Exception($"ConnectData: Unknown parameter1 {param1}, allowed: [GET, HGET, HGETALL, SUB, PSUB]");
                }
                if (logger.IsInfoEnabled)
                    logger.Info(
                        $"ConnectData: New data request accepted: " + 
                        (sub ? _topicsSubs[topic.TopicId].ToString() : _topics[topic.TopicId].ToString())
                    );
                return "(ConnectData)";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"ConnectData: Error: {ex.Message}");
                return $"#ERROR: ConnectData: {ex.Message}";
            }
        }

        protected override void DisconnectData(Topic topic)
        {
            try
            {
                bool ret = _topics.TryGetValue(topic.TopicId, out TopicData td);
                if (!ret)
                    ret = _topicsSubs.TryGetValue(topic.TopicId, out td);
                if (ret){
                    if (logger.IsInfoEnabled)
                        logger.Info($"DisconnectData: Removing {topic.TopicId}");
                    if (td.Type == "SUB" || td.Type == "PSUB")
                    {
                        string subHost = td.GetSubHost();
                        // faz unsubscribe
                        if (_redisSubscriptions.TryGetValue(subHost, out var sub))
                        {
                            var redisChannel = new RedisChannel(
                                td.KeyOrChannel,
                                (td.Type == "SUB" ? RedisChannel.PatternMode.Literal : RedisChannel.PatternMode.Pattern)
                            );
                            sub.Unsubscribe(redisChannel);
                            if (logger.IsDebugEnabled)
                                logger.Debug($"DisconnectData: Redis Unsubscribe TopicData={td}");
                        }
                        // remove da lista de canais no subHost
                        if (_channelTopics.TryGetValue(subHost, out var list))
                        {
                            list.Remove(topic.TopicId);
                            if (list.Count == 0)
                            {
                                if (_redisSubscriptions.TryRemove(subHost, out var sub2))
                                {
                                    sub2.UnsubscribeAll();
                                    if (logger.IsDebugEnabled)
                                        logger.Debug($"DisconnectData: Redis UnsubscribeAll host={td.Host}");
                                }
                                _channelTopics.TryRemove(subHost, out _);
                            }
                        }
                        _topicsSubs.Remove(topic.TopicId);
                    }
                    else
                    {
                        _topics.Remove(topic.TopicId);
                    }
                }
                else
                {
                    if (logger.IsErrorEnabled)
                        logger.Error($"DisconnectData: Unknown TopicId={topic.TopicId}");
                }
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"DisconnectData: Error: {ex.Message}");
            }
            return;
        }
        private bool CheckAutomaticRealtime(bool goback_realtime=false)
        {
            if (ExcelUpdateStyle != ENUMExcelUpdateStyle.Automatic)
                return false;
            var oldRealtime = RealTimeUpdates;
            long counterWas = 0;
            lock (_MessageCounter_lock)
            {
                counterWas = MessageCounter;
                if (ExcelUpdateStyle == ENUMExcelUpdateStyle.Automatic)
                {
                    var curr = (MessageCounter < MessageCounterThreshold || MessageCounterThreshold <= 0);
                    if (goback_realtime || (!goback_realtime && !curr))
                        RealTimeUpdates = curr;
                }
            }
            var changed = (ExcelUpdateStyle == ENUMExcelUpdateStyle.Automatic && oldRealtime != RealTimeUpdates);
            if (changed && logger.IsDebugEnabled)
                logger.Debug($"CheckAutomaticRealtime: RealTimeUpdates changed, from {oldRealtime} to {RealTimeUpdates}, counterWas={counterWas}/{MessageCounterThreshold}");
            return changed;
        }
        private void TimerElapsedMessageCounter(object sender, System.Timers.ElapsedEventArgs e)
        {
            CheckAutomaticRealtime(true);
            lock (_MessageCounter_lock)
            {
                MessageCounter = 0;
            }
        }

        private void TimerElapsedExcel(object sender, System.Timers.ElapsedEventArgs e)
        {
            CheckAutomaticRealtime();
            if (RealTimeUpdates)
                return;
            if (logger.IsDebugEnabled)
                logger.Debug($"TimerElapsedExcel: Updating Excel Values, refresh rate={ExcelUpdateRateMS}ms, topics.Count={_topics.Count}, topicsSubs.Count={_topicsSubs.Count}");
            if (_topics.Count == 0 && _topicsSubs.Count == 0)
                return;
            // aqui precisamos pegar os topicos e mandar para o excel
            foreach (var topic in _topics.Values)
                topic.SendToExcelIfDirty();
            foreach (var topic in _topicsSubs.Values)
                topic.SendToExcelIfDirty();
        }
        private void TimerElapsedRedis(object sender, System.Timers.ElapsedEventArgs e)
        {
            CheckAutomaticRealtime();
            if (logger.IsDebugEnabled)
                logger.Debug($"TimerElapsedExcel: Fetching Redis Values, refresh rate={RedisUpdateRateMS}ms, topics.Count={_topics.Count}");
            if (_topics.Count == 0)
                return;
            // Separar por tipo para GET múltiplo
            List<TopicData> otherTopics;
            if (UseGetMultiple)
            {
                var getTopics = new Dictionary<string, List<TopicData>>(); // host -> list of GETs
                otherTopics = new List<TopicData>();
                foreach (var td in _topics.Values)
                {
                    if (td.Type == "SUB" || td.Type == "PSUB") continue;

                    if (td.Type == "GET")
                    {
                        if (!getTopics.ContainsKey(td.Host))
                            getTopics[td.Host] = new List<TopicData>();
                        getTopics[td.Host].Add(td);
                    }
                    else
                    {
                        otherTopics.Add(td);
                    }
                }
                if (logger.IsDebugEnabled)
                    logger.Debug($"TimerElapsedRedis: Using GETMULTI, otherTopics.Count={otherTopics.Count}, getTopics.Count={getTopics.Count}");

                // Executar GET múltiplo por host
                foreach (var kvp in getTopics)
                {
                    try
                    {
                        var conn = GetOrCreateRedis(kvp.Key, null);
                        var db = conn.GetDatabase();

                        var topicsList = kvp.Value;
                        var keys = topicsList.Select(t => (RedisKey)t.KeyOrChannel).ToArray();
                        if (logger.IsDebugEnabled)
                            logger.Debug($"TimerElapsedRedis: GETMULTI host={kvp.Key}, keys=[{string.Join(", ", topicsList.Select(t => t.KeyOrChannel))}]");
                        var values = db.StringGet(keys);

                        for (int i = 0; i < topicsList.Count; i++)
                        {
                            var td = topicsList[i];
                            string value = (values[i].HasValue ? values[i].ToString() : "(no value)");
                            lock (_MessageCounter_lock)
                            {
                                MessageCounter++;
                            }
                            if (logger.IsTraceEnabled)
                                logger.Trace($"TimerElapsedRedis: [MULTI] host={td.Host}, key={td.KeyOrChannel}, value={value}");
                            if (RealTimeUpdates)
                                td.UpdateAndSendToExcel(value);
                            else
                                td.UpdateOnly(value);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (logger.IsErrorEnabled)
                            logger.Error(ex, $"TimerElapsedRedis: Error during GETMULTI for host={kvp.Key}");
                    }
                }
            }
            else
            {
                otherTopics = _topics.Values.ToList();
                if (logger.IsDebugEnabled)
                    logger.Debug($"TimerElapsedRedis: Not using GETMULTI, otherTopics.Count={otherTopics.Count}");
            }

            // Executar os demais tipos (HGET, HGETALL)
            foreach (var td in otherTopics)
            {
                try
                {
                    var conn = GetOrCreateRedis(td.Host);
                    var db = conn.GetDatabase();
                    string valueStr = "";
                    switch (td.Type)
                    {
                        case "GET":
                            var valueGet = db.StringGet(td.KeyOrChannel);
                            if (valueGet.HasValue)
                                valueStr = valueGet.ToString();
                            else
                                valueStr = "(no value)";
                            break;
                        case "HGET":
                            var valueHGet = db.HashGet(td.KeyOrChannel, td.Field);
                            if (valueHGet.HasValue)
                                valueStr = valueHGet.ToString();
                            else
                                valueStr = "(no value)";
                            break;
                        case "HGETALL":
                            var valueHash = db.HashGetAll(td.KeyOrChannel);
                            if (valueHash.Length > 0)
                                valueStr = string.Join(",", valueHash.Select(x => $"\"{x.Name}\":\"{x.Value}\""));
                            else
                                valueStr = "(no value)";
                            break;
                        default:
                            continue;
                    }
                    lock (_MessageCounter_lock)
                    {
                        MessageCounter++;
                    }
                    if (logger.IsTraceEnabled)
                        logger.Trace($"TimerElapsedRedis: {td.Type} host={td.Host}, key={td.KeyOrChannel}, field={td.Field}, value={valueStr}");
                    if (RealTimeUpdates)
                        td.UpdateAndSendToExcel(valueStr);
                    else
                        td.UpdateOnly(valueStr);
                }
                catch (Exception ex)
                {
                    if (logger.IsErrorEnabled)
                        logger.Error(ex, $"TimerElapsedRedis: Error fetching data for {td.KeyOrChannel}");
                }
            }
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
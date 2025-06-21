using NLog;
using ExcelDna.Integration;
using StackExchange.Redis;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Threading.Channels;

/// <summary>
/// Exemplo de arquivo RedisExcel.config.json:
/// {
///   "host": "localhost:6379,password=senha123,defaultDatabase=1",
///   "timeout": "800"
/// }

/// </summary>
namespace RedisExcel
{
    public class RedisUDF
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static TimeSpan RedisTimeout = TimeSpan.FromMilliseconds(500);
        private static string ConfigDefaultHost = null;
        private static Dictionary<string, string> Servers = null;
        private static readonly ConcurrentDictionary<string, ConnectionMultiplexer> RedisConnections = new ConcurrentDictionary<string, ConnectionMultiplexer>();

        private static readonly ConcurrentDictionary<string, string> _channelMessages = new ConcurrentDictionary<string, string>();
        private static readonly ConcurrentDictionary<string, ConnectionMultiplexer> _channelConnections = new ConcurrentDictionary<string, ConnectionMultiplexer>();
        private static readonly ConcurrentDictionary<string, ISubscriber> _channelSubscribers = new ConcurrentDictionary<string, ISubscriber>();

        private static readonly ConcurrentDictionary<string, string> _lastPublished = new ConcurrentDictionary<string, string>();
        private static void LoadConfig()
        {
            var config = ConfigHelper.GetConfig();
            RedisUDF.ConfigDefaultHost = config.UDF.host;
            RedisUDF.RedisTimeout = TimeSpan.FromMilliseconds(config.UDF.timeout);
            RedisUDF.Servers = config.Servers;
        }
        private static string GetDefaultHost()
        {
            if (RedisUDF.ConfigDefaultHost == null)
                RedisUDF.LoadConfig();
            return RedisUDF.ConfigDefaultHost;
        }
        private static string FindServerName(string host)
        {
            if (RedisUDF.Servers == null)
                RedisUDF.LoadConfig();
            if (RedisUDF.Servers != null && RedisUDF.Servers.TryGetValue(host, out var server))
                return server;
            return host;
        }
        private static ConnectionMultiplexer GetOrCreateConnection(string host)
        {
            return RedisConnections.GetOrAdd(host, h =>
            {
                if (logger.IsInfoEnabled)
                    logger.Info($"RedisConnectUDF: Connecting to {host}");
                var options = ConfigurationOptions.Parse(h);
                options.AbortOnConnectFail = false;
                options.ConnectTimeout = (int)RedisTimeout.TotalMilliseconds;
                options.SyncTimeout = (int)RedisTimeout.TotalMilliseconds;
#if GIT_TAG
		        string GitTag = GIT_TAG;
#else
		        string GitTag = "not GITHub";
#endif
		        options.ClientName = $"RedisUDF :: {GitTag} :: {Environment.UserDomainName}\\{Environment.UserName} :: {Environment.MachineName}";
                var conn = ConnectionMultiplexer.Connect(options);
                conn.ConnectionFailed += (sender, args) =>
                {
                    if (logger.IsInfoEnabled)
                        logger.Info($"RedisConnectUDF: LOST connection to Redis. host={host}, Endpoint={args.EndPoint}, FailureType={args.FailureType}, Exception={args.Exception?.Message}");
                };
                conn.ConnectionRestored += (sender, args) =>
                {
                    if (logger.IsInfoEnabled)
                        logger.Info($"RedisConnectUDF: Redis reconnect detected for host={host}");
                };
                return conn;
            });
        }

        private static IDatabase GetDatabase(string host = null)
        {
            host = string.IsNullOrWhiteSpace(host) ? GetDefaultHost() : host;
            host = FindServerName(host);

            var connection = RedisConnections.GetOrAdd(host, h =>
            {
                var options = ConfigurationOptions.Parse(h);
                options.AbortOnConnectFail = false;
                options.ConnectTimeout = (int)RedisTimeout.TotalMilliseconds;
                options.SyncTimeout = (int)RedisTimeout.TotalMilliseconds;
                return ConnectionMultiplexer.Connect(options);
            });

            return connection.GetDatabase();
        }


        private static void StartChannelListener(string channel, string host)
        {
            if (_channelSubscribers.ContainsKey(channel))
                return;

            Task.Run(() =>
            {
                try
                {
                    var mux = ConnectionMultiplexer.Connect(host);
                    var sub = mux.GetSubscriber();
                    sub.Subscribe(new RedisChannel(channel, RedisChannel.PatternMode.Literal), (chan, msg) =>
                    {
                        _channelMessages[chan] = msg;
                    });
                    // restore de conexao
                    mux.ConnectionFailed += (sender, args) =>
                    {
                        if (logger.IsInfoEnabled)
                            logger.Info($"StartChannelListener: LOST connection to Redis. host={host}, channel={channel}, Endpoint={args.EndPoint}, FailureType={args.FailureType}, Exception={args.Exception?.Message}");
                    };
                    mux.ConnectionRestored += (sender, args) =>
                    {
                        if (logger.IsInfoEnabled)
                            logger.Info($"StartChannelListener: Redis reconnect detected for host={host}, channel={channel}");
                        if (_channelConnections.TryRemove(channel, out var mux2))
                        {
                            mux2.Close();
                            mux2.Dispose();
                        }
                        var sub2 = mux.GetSubscriber();
                        sub2.Subscribe(new RedisChannel(channel, RedisChannel.PatternMode.Literal), (chan, msg) =>
                        {
                            _channelMessages[chan] = msg;
                        });
                        _channelSubscribers[channel] = sub2;
                    };

                    _channelConnections[channel] = mux;
                    _channelSubscribers[channel] = sub;
                }
                catch (Exception ex)
                {
                    _channelMessages[channel] = "Error: " + ex.Message;
                }
            });
        }
        [ExcelFunction(Description = "Unsubscribes from a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelUnsubscribe(
            [ExcelArgument(Description = "Redis channel to unsubscribe from")] string channel
        )
        {
            try
            {
                if (_channelSubscribers.TryRemove(channel, out var sub))
                {
                    sub.Unsubscribe(new RedisChannel(channel, RedisChannel.PatternMode.Literal));
                }

                if (_channelConnections.TryRemove(channel, out var mux))
                {
                    mux.Close();
                    mux.Dispose();
                }

                _channelMessages.TryRemove(channel, out _);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelUnsubscribe: channel={channel} unsubscribed");
                return $"Channel '{channel}' unsubscribed successfully.";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFChannelUnsubscribe Error: channel={channel}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Returns the number of active Redis connections", IsVolatile = true)]
        public static object RedisUDFConnectionCount()
        {
            if (logger.IsTraceEnabled)
                logger.Trace($"RedisUDFConnectionCount: connections={RedisConnections.Count}");
            return RedisConnections.Count;
        }


        [ExcelFunction(Description = "Reads the latest Pub/Sub message from a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelLatest(
            [ExcelArgument(Description = "Redis channel to read from")] string channel,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                StartChannelListener(channel, host);
                var resposta = _channelMessages.TryGetValue(channel, out var msg) ? msg : "(null)";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelLatest: channel={channel}, msg={resposta}, host={host}");
                return resposta;
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFChannelLatest Error: channel={channel}, host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishIfChangedJSON(
            [ExcelArgument(Description = "Redis channel")] string channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublishIfChanged(channel, RedisUDFMatrixToJSON(range), optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel only if subscribers are present", IsVolatile = true)]
        public static string RedisUDFChannelPublishIfChanged(
            [ExcelArgument(Description = "Redis channel to publish to")] string channel,
            [ExcelArgument(Description = "Message content to publish")] string message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var mux = GetOrCreateConnection(host);
                var sub = mux.GetSubscriber();

                long count = sub.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), message);

                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channel}, msg={message}, readers={count}, host={host}");

                if (count > 0)
                {
                    _lastPublished[channel] = message;
                    return $"{count} readers(s)";
                }
                else
                {
                    return "No Readers";
                }
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFChannelPublishIfChanged Error: channel={channel}, host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishJSON(
            [ExcelArgument(Description = "Redis channel")] string channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublish(channel, RedisUDFMatrixToJSON(range), optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelPublish(
            [ExcelArgument(Description = "Redis channel to publish to")] string channel,
            [ExcelArgument(Description = "Message content to publish")] string message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var mux = GetOrCreateConnection(host);
                var sub = mux.GetSubscriber();

                long count = sub.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), message);
                if(logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublish: channel={channel}, msg={message}, readers={count}, host={host}");

                return $"{count} readers(s)";
            }
            catch (System.Exception ex)
            {
                if(logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFChannelPublish Error: channel={channel}, host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Returns all active Pub/Sub channels and their subscriber counts from Redis", IsVolatile = true)]
        public static object[,] RedisUDFPubSubChannelsInfo(
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var conn = GetOrCreateConnection(host);
                var endpoint = conn.GetEndPoints().First();
                var server = conn.GetServer(endpoint);

                // Executa PUBSUB CHANNELS
                var channelsResult = server.Execute("PUBSUB", "CHANNELS");
                if (channelsResult.Resp2Type != ResultType.Array)
                {
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFPubSubChannelsInfo: host={host}, channels=0");
                    return new object[,] { { "Channel", "Subscribers" } };
                }

                var channels = (RedisResult[])channelsResult;
                int count = channels.Length;
                object[,] result = new object[count + 1, 2];
                result[0, 0] = "Channel";
                result[0, 1] = "Subscribers";

                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFPubSubChannelsInfo: host={host}, channels={count}");
                for (int i = 0; i < count; i++)
                {
                    string channel = channels[i].ToString();

                    // Executa PUBSUB NUMSUB <channel>
                    result[i + 1, 0] = channel;
                    result[i + 1, 1] = "Fetching ...";
                    try
                    {
                        var numsubResult = server.Execute("PUBSUB", "NUMSUB", channel);
                        var numsubArray = (RedisResult[])numsubResult;
                        long subscribers = (numsubArray.Length >= 2) ? (long)numsubArray[1] : 0;
                        result[i + 1, 1] = subscribers;


                        if (logger.IsTraceEnabled)
                            logger.Trace($"RedisUDFPubSubChannelsInfo: host={host}, channels={count}, i={i}, channel={channel}, subscribers={subscribers}");
                    }catch (Exception ex)
                    {
                        if (logger.IsErrorEnabled)
                            logger.Error(ex, $"RedisUDFPubSubChannelsInfo Error: host={host}, fetch PUBSUB NUMSUB {channel}");
                        result[i + 1, 1] = $"Error: {ex.Message}";
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFPubSubChannelsInfo Error: host={host}");
                return new object[,] { { "Error", ex.Message } };
            }
        }


        [ExcelFunction(Description = "Gets the value of a Redis key with optional host", IsVolatile = true)]
        public static string RedisUDFGet(
            [ExcelArgument(Description = "Redis key to retrieve the value from")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                var value = db.StringGet(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGet: key={key}, value={value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFGet Error: key={key}, host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with a Matrix using JSON", IsVolatile = true)]
        public static string RedisUDFSetJSON(
            [ExcelArgument(Description = "Redis key to set the value for")] string key,
            [ExcelArgument(Description = "Value to set for the given key")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            string json = null;
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                json = RedisUDFMatrixToJSON(values);
                db.StringSet(key, json);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetJSON: key={key}, value={json}, host={host}");
                return "OK";
            }
            catch (System.Exception ex)
            {
                if(logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFSetJSON Error: key={key}, value={json}, host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Sets the value of a Redis key", IsVolatile = true)]
        public static string RedisUDFSet(
            [ExcelArgument(Description = "Redis key to set the value for")] string key,
            [ExcelArgument(Description = "Value to set for the given key")] string value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                db.StringSet(key, value);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSet: key={key}, value={value}, host={host}");
                return "OK";
            }
            catch (System.Exception ex)
            {
                if(logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFSet Error: key={key}, value={value}, host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static string RedisUDFSetKV(
            [ExcelArgument(Description = "Range with Redis keys")] object[,] keys,
            [ExcelArgument(Description = "Range with Redis values")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                int L1 = keys.GetLength(0);
                int L2 = values.GetLength(0);
                int minL = (L1 < L2 ? L1 : L2);

                var entries = new List<KeyValuePair<RedisKey, RedisValue>>();
                for (int i = 0; i < minL; i++)
                {
                    var key = keys[i, 0]?.ToString();
                    var value = values[i, 0]?.ToString();
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value));
                }

                db.StringSet(entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKV: {entries.Count} pairs sent, host={host}");
                return "OK";
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFSetKV Error: host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static string RedisUDFSetKVPair(
            [ExcelArgument(Description = "2D range with Redis key-value pairs (2 columns: key, value)")] object[,] keyValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);

                var entries = new List<KeyValuePair<RedisKey, RedisValue>>();
                for (int i = 0; i < keyValuePairs.GetLength(0); i++)
                {
                    var key = keyValuePairs[i, 0]?.ToString();
                    var value = keyValuePairs[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value));
                }

                db.StringSet(entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKVPair: {entries.Count} pairs sent, host={host}");
                return "OK";
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFSetKVPair Error: host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Gets the values of multiple Redis keys", IsVolatile = true)]
        public static object[,] RedisUDFGetMultiple(
            [ExcelArgument(Description = "Array of Redis keys to retrieve")] object[] keys,
            [ExcelArgument(Description = "If TRUE, returns two columns (key, value); if FALSE, only values")] object multipleColumnsOpt,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                bool multipleColumns = multipleColumnsOpt is bool b && b;

                var db = GetDatabase(host);
                var validKeys = keys.Select(k => k?.ToString() ?? "").Where(k => !string.IsNullOrWhiteSpace(k)).ToArray();
                if (validKeys.Length == 0)
                    return new object[,] { { "Error: No valid key", "(null)" } };

                var redisKeys = validKeys.Select(k => (RedisKey)k).ToArray();
                var values = db.StringGet(redisKeys);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGetMultiple: keys={validKeys.Length}, multipleColumns={multipleColumns}, host={host}");

                if (multipleColumns)
                {
                    object[,] result = new object[values.Length, 2];
                    for (int i = 0; i < values.Length; i++)
                    {
                        result[i, 0] = redisKeys[i].ToString();
                        result[i, 1] = values[i].HasValue ? values[i].ToString() : "(null)";
                    }
                    return result;
                }
                else
                {
                    object[,] result = new object[values.Length, 1];
                    for (int i = 0; i < values.Length; i++)
                        result[i, 0] = values[i].HasValue ? values[i].ToString() : "(null)";
                    return result;
                }
            }
            catch (System.Exception ex)
            {
                if(logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFGetMultiple Error: host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }

        [ExcelFunction(Description = "Lists Redis keys matching a pattern", IsVolatile = true)]
        public static object[,] RedisUDFKeys(
            [ExcelArgument(Description = "Pattern to match Redis keys (e.g., user:*)")] string pattern,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = optionalHost as string ?? GetDefaultHost();
                host = FindServerName(host);
                var conn = GetOrCreateConnection(host);
                var server = conn.GetServer(conn.GetEndPoints().First());

                var keys = server.Keys(pattern: pattern).Select(k => k.ToString()).ToList();

                object[,] result = new object[keys.Count, 1];
                for (int i = 0; i < keys.Count; i++)
                    result[i, 0] = keys[i];
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFKeys: pattern={pattern}, found={keys.Count}, host={host}");
                return result;
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFKeys Error: padrão={pattern}, host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }
        [ExcelFunction(Description = "Returns the TTL of a Redis key in seconds", IsVolatile = true)]
        public static object RedisUDFTTL(
            [ExcelArgument(Description = "Redis key to check TTL for")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                var ttl = db.KeyTimeToLive(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTL: key={key}, ttl={ttl}, host={host}");
                return ttl.HasValue ? ttl.Value.TotalSeconds : -1;
            }
            catch (System.Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFTTL Error: host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Returns the current time of the Redis server", IsVolatile = true)]
        public static object RedisUDFServerTime(
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = optionalHost as string ?? GetDefaultHost();
                host = FindServerName(host);
                var conn = GetOrCreateConnection(host);
                var server = conn.GetServer(conn.GetEndPoints().First());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFServerTime: {server.Time()}, host={host}");
                return server.Time().ToString("o"); // formato ISO 8601
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFServerTime Error: host={host}");
                return "Error: " + ex.Message;
            }
        }
        [ExcelFunction(Description = "Checks if one or more Redis keys exist", IsVolatile = true)]
        public static object[,] RedisUDFExistsMultiples(
            [ExcelArgument(Description = "Array of Redis keys to check for existence")] object[] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                object[,] result = new object[keys.Length, 2];

                for (int i = 0; i < keys.Length; i++)
                {
                    var key = keys[i]?.ToString();
                    bool exists = db.KeyExists(key);
                    result[i, 0] = key;
                    result[i, 1] = exists ? "1" : "0";
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExistsMultiples: {keys.Length} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFExistsMultiples Error: host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }
        [ExcelFunction(Description = "Checks if a Redis key exists", IsVolatile = true)]
        public static object RedisUDFExists(
            [ExcelArgument(Description = "Redis key to check for existence")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                var resultado = db.KeyExists(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExists: key={key}, exists={resultado}");
                return resultado ? "1" : "0";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFExists Error: host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Returns the TTL of multiple Redis keys in seconds", IsVolatile = true)]
        public static object[,] RedisUDFTTLMultiples(
            [ExcelArgument(Description = "Array of Redis keys to check TTL for")] object[] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                object[,] result = new object[keys.Length, 2];

                for (int i = 0; i < keys.Length; i++)
                {
                    var key = keys[i]?.ToString();
                    var ttl = db.KeyTimeToLive(key);
                    result[i, 0] = key;
                    result[i, 1] = ttl.HasValue ? ttl.Value.TotalSeconds.ToString("F0") : "-1";
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTLMultiples: {keys.Length} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                if(logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFTTLMultiples Error: host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }
        [ExcelFunction(Description = "Sets a field in a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashSet(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "Field name to set within the hash")] string field,
            [ExcelArgument(Description = "Value to set for the given field")] string value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                db.HashSet(hashKey, field, value);
                if(logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSet: {hashKey}[{field}] = {value}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFHashSet Error: {hashKey}, host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Sets multiple fields in a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashSetMultiple(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "2D range with field-value pairs (2 columns: field, value)")] object[,] fieldValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);

                var entries = new List<HashEntry>();
                for (int i = 0; i < fieldValuePairs.GetLength(0); i++)
                {
                    var field = fieldValuePairs[i, 0]?.ToString();
                    var value = fieldValuePairs[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(field))
                        entries.Add(new HashEntry(field, value));
                }

                db.HashSet(hashKey, entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSetMultiple: {hashKey}, fields={entries.Count}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFHashSetMultiple Error: {hashKey}, host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Gets a field from a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashGet(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "Field name to retrieve from the hash")] string field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                var value = db.HashGet(hashKey, field);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGet: {hashKey}[{field}] = {value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFHashGet Error: {hashKey}[{field}], host={host}");
                return "Error: " + ex.Message;
            }
        }

        [ExcelFunction(Description = "Gets all fields of a Redis hash", IsVolatile = true)]
        public static object[,] RedisUDFHashGetAll(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);
                var all = db.HashGetAll(hashKey);

                object[,] result = new object[all.Length, 2];
                for (int i = 0; i < all.Length; i++)
                {
                    result[i, 0] = all[i].Name.ToString();
                    result[i, 1] = all[i].Value.ToString();
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGetAll: {hashKey}, fields={all.Length}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFHashGetAll Error: {hashKey}, host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }
        [ExcelFunction(Description = "Gets the same field from multiple Redis hashes", IsVolatile = true)]
        public static object[,] RedisUDFHashGetFieldMultipleKeys(
            [ExcelArgument(Description = "Array of Redis hash keys")] object[] hashKeys,
            [ExcelArgument(Description = "Field name to retrieve from each hash")] string field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = "";
            try
            {
                host = string.IsNullOrWhiteSpace(optionalHost as string) ? GetDefaultHost() : optionalHost as string;
                host = FindServerName(host);
                var db = GetDatabase(host);

                object[,] result = new object[hashKeys.Length, 2];

                for (int i = 0; i < hashKeys.Length; i++)
                {
                    string key = hashKeys[i]?.ToString();
                    string value = db.HashGet(key, field);
                    result[i, 0] = key;
                    result[i, 1] = value ?? "";
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGetFieldMultipleKeys: field={field}, hashes={hashKeys.Length}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFHashGet Error: hashKeys.Length={hashKeys.Length}, field={field}, host={host}");
                return new object[,] { { "Error: " + ex.Message } };
            }
        }


        ////////////// HELPERS
        [ExcelFunction(Description = "Converts a 2D Excel matrix to a compact JSON array of arrays.")]
        public static string RedisUDFMatrixToJSON(
            [ExcelArgument(Description = "2D Excel range to convert")] object[,] range)
        {
            try
            {
                int rows = range.GetLength(0);
                int cols = range.GetLength(1);
                var array = new object[rows][];

                for (int i = 0; i < rows; i++)
                {
                    var row = new object[cols];
                    for (int j = 0; j < cols; j++)
                    {
                        var val = range[i, j];

                        if (val is ExcelEmpty || val is ExcelMissing || val is ExcelError || val is null)
                            row[j] = null;
                        else if (val is double d)
                            row[j] = d % 1 == 0 ? (object)(long)d : d;
                        else
                            row[j] = val;
                    }
                    array[i] = row;
                }
                var settings = new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Include,
                    Formatting = Formatting.None
                };

                return JsonConvert.SerializeObject(array, settings);
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error(ex, $"RedisUDFMatrixToJSON: error {ex.Message}");
                return $"Error: {ex}";
            }
        }

        [ExcelFunction(Description = "Converts a JSON string to a 2D Excel matrix. Accepts arrays, objects, or single values.")]
        public static object[,] RedisUDFJSONToMatrix(
            [ExcelArgument(Description = "JSON string to convert to Excel matrix")] string json,
            [ExcelArgument(Description = "Value to insert for nulls (default is empty string)")] object nullValue = null
        )
        {
            var nullV = nullValue == null ? "" : nullValue;
            if (logger.IsTraceEnabled)
                logger.Trace($"RedisUDFJSONToMatrix: converting JSON data: {json}");
            try
            {
                var token = JsonConvert.DeserializeObject<JToken>(json);

                if (token is JArray array)
                {
                    // Check if it's a matrix (array of arrays)
                    if (array.All(x => x is JArray))
                    {
                        var list = array.ToObject<List<List<object>>>();
                        int rows = list.Count;
                        int cols = list.Max(r => r.Count);
                        if (logger.IsTraceEnabled)
                            logger.Trace($"RedisUDFJSONToMatrix: JArray[{rows}, {cols}] - Array of Arrays");
                        var result = new object[rows, cols];

                        for (int i = 0; i < rows; i++)
                        {
                            for (int j = 0; j < cols; j++)
                            {
                                result[i, j] = nullV;
                                if (j < list[i].Count)
                                    result[i, j] = list[i][j] ?? nullV;
                            }
                        }
                        if (logger.IsTraceEnabled)
                            logger.Trace(
                                "RedisUDFJSONToMatrix: Output\n" + 
                                string.Join("\n", 
                                    Enumerable.Range(0, result.GetLength(0))
                                    .Select(
                                        i => string.Join(
                                            ",", 
                                            Enumerable.Range(0, result.GetLength(1)).Select(
                                                j => result[i, j].ToString()
                                            )
                                        )
                                    )
                                )
                            );

                        return result;
                    }
                    else
                    {
                        // Flat array: [1,2,3,4]
                        int n = array.Count;
                        if (logger.IsTraceEnabled)
                            logger.Trace($"RedisUDFJSONToMatrix: JArray[{n}] - Flat Array");
                        var result = new object[1, n];
                        for (int i = 0; i < n; i++)
                            if (array[i] == null)
                                result[0, i] = nullV;
                            else
                                result[0, i] = array[i];
                        if (logger.IsTraceEnabled)
                            logger.Trace(
                                "RedisUDFJSONToMatrix: Output\n" +
                                string.Join("\n",
                                    Enumerable.Range(0, result.GetLength(0))
                                    .Select(
                                        i => string.Join(
                                            ",",
                                            Enumerable.Range(0, result.GetLength(1)).Select(
                                                j => result[i, j].ToString()
                                            )
                                        )
                                    )
                                )
                            );

                        return result;
                    }
                }
                else if (token is JObject obj)
                {
                    var keys = obj.Properties().Select(p => p.Name).ToList();
                    int rows = obj.Properties().Max(p =>
                    {
                        if (p.Value is JArray ja) return ja.Count;
                        return 1;
                    });
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFJSONToMatrix: JObject [{rows+1}, {keys.Count}]");
                    var result = new object[rows + 1, keys.Count];
                    for (int j = 0; j < keys.Count; j++)
                    {
                        var key = keys[j];
                        result[0, j] = key;
                        for (int i = 0; i < keys.Count; i++)
                            result[i + 1, j] = nullV;
                        var val = obj[key];
                        if (val is JArray arr)
                        {
                            for (int i = 0; i < arr.Count; i++)
                            {
                                //result[i + 1, j] = arr[i]?.ToString();
                                if (arr[i] == null)
                                    result[i + 1, j] = nullV;
                                else
                                    result[i + 1, j] = arr[i];
                            }
                        }
                        else
                        {
                            if (val == null)
                                result[1, j] = nullV;
                            else
                                result[1, j] = val;
                        }
                    }
                    if (logger.IsTraceEnabled)
                        logger.Trace(
                            "RedisUDFJSONToMatrix: Output\n" +
                            string.Join("\n",
                                Enumerable.Range(0, result.GetLength(0))
                                .Select(
                                    i => string.Join(
                                        ",",
                                        Enumerable.Range(0, result.GetLength(1)).Select(
                                            j => result[i, j].ToString()
                                        )
                                    )
                                )
                            )
                        );
                    return result;
                }
                else
                {
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFJSONToMatrix: Scalar {token}");
                    // Scalar: "abc", 123, true, etc.
                    if (token == null)
                        return new object[,] { { nullV } };
                    return new object[,] { { token } };
                }
            }
            catch (Exception ex)
            {
                if (logger.IsErrorEnabled)
                    logger.Error($"RedisUDFJSONToMatrix: Error {ex.Message}");
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }



    }
}

using ExcelDna.Integration;
using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace RedisExcel
{
    /// <summary>
    /// User Defined Functions exposed to Excel. All of them resolve host/alias through
    /// AppConfig, use the shared RedisConnectionManager and return "Error: ..." on failure,
    /// as before. Pub/Sub logic lives in RedisSubscriptionManager (per-channel refcount).
    /// </summary>
    public static class RedisUDF
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private sealed class ChannelListener
        {
            public string Host;
            public string Channel;
            public IDisposable Token;
        }

        private static readonly ConcurrentDictionary<string, ChannelListener> _channelListeners =
            new ConcurrentDictionary<string, ChannelListener>();
        private static readonly ConcurrentDictionary<string, string> _latestMessages =
            new ConcurrentDictionary<string, string>();

        private static string ChannelKey(string host, string channel) => $"{host}\u0001{channel}";

        private static string ResolveHost(object optionalHost) => AppConfig.ResolveUdfHost(optionalHost as string);

        private static IDatabase GetDb(string host) => RedisRuntime.Connections.GetDatabase(host, RedisPool.UdfData);

        private static string Fail(string function, Exception ex, string context)
        {
            logger.Error(ex, $"{function}: {context}");
            return "Error: " + ex.Message;
        }

        private static object[,] FailMatrix(string function, Exception ex, string context)
        {
            logger.Error(ex, $"{function}: {context}");
            return new object[,] { { "Error: " + ex.Message } };
        }

        [ExcelFunction(Description = "Unsubscribes from a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelUnsubscribe(
            [ExcelArgument(Description = "Redis channel to unsubscribe from")] string channel
        )
        {
            try
            {
                foreach (var kv in _channelListeners)
                {
                    if (!string.Equals(kv.Value.Channel, channel, StringComparison.Ordinal))
                        continue;
                    if (_channelListeners.TryRemove(kv.Key, out var listener))
                    {
                        listener.Token?.Dispose();
                        _latestMessages.TryRemove(kv.Key, out _);
                    }
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelUnsubscribe: channel={channel} unsubscribed");
                return $"Channel '{channel}' unsubscribed successfully.";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelUnsubscribe", ex, $"channel={channel}");
            }
        }

        [ExcelFunction(Description = "Returns the number of active Redis connections", IsVolatile = true)]
        public static object RedisUDFConnectionCount()
        {
            int count = RedisRuntime.Connections.UdfConnectionCount;
            if (logger.IsTraceEnabled)
                logger.Trace($"RedisUDFConnectionCount: connections={count}");
            return count;
        }

        [ExcelFunction(Description = "Reads the latest Pub/Sub message from a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelLatest(
            [ExcelArgument(Description = "Redis channel to read from")] string channel,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string key = ChannelKey(host, channel);
                if (!_channelListeners.ContainsKey(key))
                {
                    var listener = new ChannelListener { Host = host, Channel = channel };
                    listener.Token = RedisRuntime.Subscriptions.Subscribe(host, channel, pattern: false,
                        onMessage: message => _latestMessages[key] = message ?? "");
                    if (!_channelListeners.TryAdd(key, listener))
                        listener.Token.Dispose(); // another thread registered first
                }
                var response = _latestMessages.TryGetValue(key, out var latest) ? latest : "(null)";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelLatest: channel={channel}, msg={response}, host={host}");
                return response;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelLatest", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishIfChangedJSON(
            [ExcelArgument(Description = "Redis channel")] string channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublishIfChanged(channel, ExcelJson.RedisUDFMatrixToJSON(range), optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel only if subscribers are present", IsVolatile = true)]
        public static string RedisUDFChannelPublishIfChanged(
            [ExcelArgument(Description = "Redis channel to publish to")] string channel,
            [ExcelArgument(Description = "Message content to publish")] string message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                long readers = subscriber.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), message);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channel}, msg={message}, readers={readers}, host={host}");
                return readers > 0 ? $"{readers} readers(s)" : "No Readers";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelPublishIfChanged", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishJSON(
            [ExcelArgument(Description = "Redis channel")] string channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublish(channel, ExcelJson.RedisUDFMatrixToJSON(range), optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelPublish(
            [ExcelArgument(Description = "Redis channel to publish to")] string channel,
            [ExcelArgument(Description = "Message content to publish")] string message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                long readers = subscriber.Publish(new RedisChannel(channel, RedisChannel.PatternMode.Literal), message);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublish: channel={channel}, msg={message}, readers={readers}, host={host}");
                return $"{readers} readers(s)";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelPublish", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns all active Pub/Sub channels and their subscriber counts from Redis", IsVolatile = true)]
        public static object[,] RedisUDFPubSubChannelsInfo(
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var conn = RedisRuntime.Connections.GetConnection(host, RedisPool.UdfData);
                var server = conn.GetServer(conn.GetEndPoints().First());

                var channelsResult = server.Execute("PUBSUB", "CHANNELS");
                if (channelsResult.Resp2Type != ResultType.Array)
                {
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFPubSubChannelsInfo: host={host}, channels=0");
                    return new object[,] { { "Channel", "Subscribers" } };
                }

                var channels = (RedisResult[])channelsResult;
                var result = new object[channels.Length + 1, 2];
                result[0, 0] = "Channel";
                result[0, 1] = "Subscribers";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFPubSubChannelsInfo: host={host}, channels={channels.Length}");

                for (int i = 0; i < channels.Length; i++)
                {
                    string channel = channels[i].ToString();
                    result[i + 1, 0] = channel;
                    try
                    {
                        var numsubResult = server.Execute("PUBSUB", "NUMSUB", channel);
                        var numsubArray = (RedisResult[])numsubResult;
                        result[i + 1, 1] = numsubArray.Length >= 2 ? (long)numsubArray[1] : 0;
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"RedisUDFPubSubChannelsInfo: host={host}, PUBSUB NUMSUB {channel}");
                        result[i + 1, 1] = $"Error: {ex.Message}";
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"RedisUDFPubSubChannelsInfo: host={host}");
                return new object[,] { { "Error", ex.Message } }; // legacy 2-column shape
            }
        }

        [ExcelFunction(Description = "Gets the value of a Redis key with optional host", IsVolatile = true)]
        public static string RedisUDFGet(
            [ExcelArgument(Description = "Redis key to retrieve the value from")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var value = GetDb(host).StringGet(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGet: key={key}, value={value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFGet", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with a Matrix using JSON", IsVolatile = true)]
        public static string RedisUDFSetJSON(
            [ExcelArgument(Description = "Redis key to set the value for")] string key,
            [ExcelArgument(Description = "Value to set for the given key")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            string json = null;
            try
            {
                host = ResolveHost(optionalHost);
                json = ExcelJson.RedisUDFMatrixToJSON(values);
                GetDb(host).StringSet(key, json);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetJSON: key={key}, value={json}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetJSON", ex, $"key={key}, value={json}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key", IsVolatile = true)]
        public static string RedisUDFSet(
            [ExcelArgument(Description = "Redis key to set the value for")] string key,
            [ExcelArgument(Description = "Value to set for the given key")] string value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                GetDb(host).StringSet(key, value);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSet: key={key}, value={value}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSet", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static string RedisUDFSetKV(
            [ExcelArgument(Description = "Range with Redis keys")] object[,] keys,
            [ExcelArgument(Description = "Range with Redis values")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                int rows = Math.Min(keys.GetLength(0), values.GetLength(0));
                var entries = new List<KeyValuePair<RedisKey, RedisValue>>();
                for (int i = 0; i < rows; i++)
                {
                    var key = keys[i, 0]?.ToString();
                    var value = values[i, 0]?.ToString();
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value));
                }
                GetDb(host).StringSet(entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKV: {entries.Count} pairs sent, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetKV", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static string RedisUDFSetKVPair(
            [ExcelArgument(Description = "2D range with Redis key-value pairs (2 columns: key, value)")] object[,] keyValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var entries = new List<KeyValuePair<RedisKey, RedisValue>>();
                for (int i = 0; i < keyValuePairs.GetLength(0); i++)
                {
                    var key = keyValuePairs[i, 0]?.ToString();
                    var value = keyValuePairs[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value));
                }
                GetDb(host).StringSet(entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKVPair: {entries.Count} pairs sent, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetKVPair", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Gets the values of multiple Redis keys", IsVolatile = true)]
        public static object[,] RedisUDFGetMultiple(
            [ExcelArgument(Description = "Array of Redis keys to retrieve")] object[] keys,
            [ExcelArgument(Description = "If TRUE, returns two columns (key, value); if FALSE, only values")] object multipleColumnsOpt,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                bool multipleColumns = multipleColumnsOpt is bool b && b;

                var validKeys = keys
                    .Select(k => k?.ToString() ?? "")
                    .Where(k => !string.IsNullOrWhiteSpace(k))
                    .ToArray();
                if (validKeys.Length == 0)
                    return new object[,] { { "Error: No valid key", "(null)" } };

                var redisKeys = validKeys.Select(k => (RedisKey)k).ToArray();
                var values = GetDb(host).StringGet(redisKeys);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGetMultiple: keys={validKeys.Length}, multipleColumns={multipleColumns}, host={host}");

                int columns = multipleColumns ? 2 : 1;
                var result = new object[values.Length, columns];
                for (int i = 0; i < values.Length; i++)
                {
                    if (multipleColumns)
                        result[i, 0] = redisKeys[i].ToString();
                    result[i, columns - 1] = values[i].HasValue ? values[i].ToString() : "(null)";
                }
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFGetMultiple", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Lists Redis keys matching a pattern", IsVolatile = true)]
        public static object[,] RedisUDFKeys(
            [ExcelArgument(Description = "Pattern to match Redis keys (e.g., user:*)")] string pattern,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var conn = RedisRuntime.Connections.GetConnection(host, RedisPool.UdfData);
                var server = conn.GetServer(conn.GetEndPoints().First());
                var keys = server.Keys(pattern: pattern).Select(k => k.ToString()).ToList();

                var result = new object[keys.Count, 1];
                for (int i = 0; i < keys.Count; i++)
                    result[i, 0] = keys[i];
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFKeys: pattern={pattern}, found={keys.Count}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFKeys", ex, $"pattern={pattern}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the TTL of a Redis key in seconds", IsVolatile = true)]
        public static object RedisUDFTTL(
            [ExcelArgument(Description = "Redis key to check TTL for")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var ttl = GetDb(host).KeyTimeToLive(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTL: key={key}, ttl={ttl}, host={host}");
                return ttl.HasValue ? ttl.Value.TotalSeconds : -1;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFTTL", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the current time of the Redis server", IsVolatile = true)]
        public static object RedisUDFServerTime(
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var conn = RedisRuntime.Connections.GetConnection(host, RedisPool.UdfData);
                var server = conn.GetServer(conn.GetEndPoints().First());
                var time = server.Time();
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFServerTime: {time}, host={host}");
                return time.ToString("o"); // ISO 8601 format
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFServerTime", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Checks if one or more Redis keys exist", IsVolatile = true)]
        public static object[,] RedisUDFExistsMultiples(
            [ExcelArgument(Description = "Array of Redis keys to check for existence")] object[] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (keys.Length == 0)
                    return new object[0, 2];
                var result = new object[keys.Length, 2];
                var keysList = new List<RedisKey>(keys.Length);
                for (int i = 0; i < keys.Length; i++)
                {
                    var key = keys[i]?.ToString();
                    result[i, 0] = key;
                    keysList.Add(key);
                }
                // One round trip for all keys instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                var tasks = keysList.Select(k => batch.KeyExistsAsync(k)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                    result[i, 1] = tasks[i].GetAwaiter().GetResult() ? "1" : "0";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExistsMultiples: {keys.Length} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFExistsMultiples", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Checks if a Redis key exists", IsVolatile = true)]
        public static object RedisUDFExists(
            [ExcelArgument(Description = "Redis key to check for existence")] string key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var exists = GetDb(host).KeyExists(key);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExists: key={key}, exists={exists}");
                return exists ? "1" : "0";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFExists", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the TTL of multiple Redis keys in seconds", IsVolatile = true)]
        public static object[,] RedisUDFTTLMultiples(
            [ExcelArgument(Description = "Array of Redis keys to check TTL for")] object[] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (keys.Length == 0)
                    return new object[0, 2];
                var result = new object[keys.Length, 2];
                var keysList = new List<RedisKey>(keys.Length);
                for (int i = 0; i < keys.Length; i++)
                {
                    var key = keys[i]?.ToString();
                    result[i, 0] = key;
                    keysList.Add(key);
                }
                // One round trip for all keys instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                var tasks = keysList.Select(k => batch.KeyTimeToLiveAsync(k)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                {
                    var ttl = tasks[i].GetAwaiter().GetResult();
                    result[i, 1] = ttl.HasValue ? ttl.Value.TotalSeconds.ToString("F0") : "-1";
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTLMultiples: {keys.Length} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFTTLMultiples", ex, $"host={host}");
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
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                GetDb(host).HashSet(hashKey, field, value);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSet: {hashKey}[{field}] = {value}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashSet", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets multiple fields in a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashSetMultiple(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "2D range with field-value pairs (2 columns: field, value)")] object[,] fieldValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var entries = new List<HashEntry>();
                for (int i = 0; i < fieldValuePairs.GetLength(0); i++)
                {
                    var field = fieldValuePairs[i, 0]?.ToString();
                    var value = fieldValuePairs[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(field))
                        entries.Add(new HashEntry(field, value));
                }
                GetDb(host).HashSet(hashKey, entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSetMultiple: {hashKey}, fields={entries.Count}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashSetMultiple", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Gets a field from a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashGet(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "Field name to retrieve from the hash")] string field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var value = GetDb(host).HashGet(hashKey, field);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGet: {hashKey}[{field}] = {value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashGet", ex, $"{hashKey}[{field}], host={host}");
            }
        }

        [ExcelFunction(Description = "Gets all fields of a Redis hash", IsVolatile = true)]
        public static object[,] RedisUDFHashGetAll(
            [ExcelArgument(Description = "Redis hash key")] string hashKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                var all = GetDb(host).HashGetAll(hashKey);
                var result = new object[all.Length, 2];
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
                return FailMatrix("RedisUDFHashGetAll", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Gets the same field from multiple Redis hashes", IsVolatile = true)]
        public static object[,] RedisUDFHashGetFieldMultipleKeys(
            [ExcelArgument(Description = "Array of Redis hash keys")] object[] hashKeys,
            [ExcelArgument(Description = "Field name to retrieve from each hash")] string field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (hashKeys.Length == 0)
                    return new object[0, 2];
                var result = new object[hashKeys.Length, 2];
                var keysList = new List<RedisKey>(hashKeys.Length);
                for (int i = 0; i < hashKeys.Length; i++)
                {
                    string key = hashKeys[i]?.ToString();
                    result[i, 0] = key;
                    keysList.Add(key);
                }
                // One round trip for all hashes instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                var tasks = keysList.Select(k => batch.HashGetAsync(k, field)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                    result[i, 1] = (string)tasks[i].GetAwaiter().GetResult() ?? "";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGetFieldMultipleKeys: field={field}, hashes={hashKeys.Length}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFHashGetFieldMultipleKeys", ex, $"field={field}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns TRUE while a newer RedisExcel release is known to exist (background check, never blocks Excel).", IsVolatile = true)]
        public static bool RedisUDFUpdateAvailable()
        {
            UpdateCheck.EnsureFresh(TimeSpan.FromHours(6));
            return UpdateCheck.IsUpdateAvailable();
        }
    }
}

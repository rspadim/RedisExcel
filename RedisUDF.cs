using ExcelDna.Integration;
using NLog;
using StackExchange.Redis;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
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
            public string Channel;
            public IDisposable Token;
            // Set before the token is disposed/unsubscribed so an in-flight
            // callback cannot publish a pre-unsubscribe message after the
            // channel entry was removed.
            private volatile bool _closed;

            public bool IsClosed => _closed;

            public void Close() => _closed = true;
        }

        private static readonly ConcurrentDictionary<string, ChannelListener> _channelListeners =
            new ConcurrentDictionary<string, ChannelListener>();
        private static readonly ConcurrentDictionary<string, string> _latestMessages =
            new ConcurrentDictionary<string, string>();

        private static string ChannelKey(string host, string channel) => $"{host}\u0001{channel}";

        private static string ResolveHost(object optionalHost)
        {
            // null / empty cell / omitted argument all mean "use the default host".
            if (optionalHost == null || optionalHost is ExcelMissing || optionalHost is ExcelEmpty)
                return AppConfig.ResolveUdfHost(null);
            if (!(optionalHost is string host))
                throw new ArgumentException("host must be a text value");
            return AppConfig.ResolveUdfHost(host);
        }

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

        /// <summary>Converts an Excel cell value to the string stored in Redis.
        /// Uses the invariant culture so numbers are written as 67000.5, not
        /// 67000,5 on comma-decimal locales. Strings pass through unchanged.</summary>
        internal static string ToRedisString(object value)
        {
            if (value is Array)
                throw new ArgumentException("A multi-cell range is not a valid scalar argument");
            if (value is ExcelError)
                throw new ArgumentException("Excel error cells are not valid arguments");
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                return null;
            if (value is string s)
                return s;
            if (value is bool b)
                return b ? "true" : "false";
            if (value is double d)
            {
                // G15 keeps common values compact ("0.1" stays "0.1"); fall back to
                // G17 only when G15 does not round-trip (e.g. double.MaxValue).
                string formatted = d.ToString("G15", CultureInfo.InvariantCulture);
                if (!(double.TryParse(formatted, NumberStyles.Float, CultureInfo.InvariantCulture, out var back) && back == d))
                    formatted = d.ToString("G17", CultureInfo.InvariantCulture);
                return formatted;
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        internal static long ToInt64Invariant(object value)
        {
            if (value is Array)
                throw new ArgumentException("A multi-cell range is not a valid numeric argument");
            if (value is ExcelError)
                throw new ArgumentException("Excel error cells are not valid numeric arguments");
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                throw new ArgumentException("numeric argument is not valid");
            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch (OverflowException)
            {
                throw new ArgumentException("numeric argument is out of range");
            }
            catch (FormatException)
            {
                throw new ArgumentException("numeric argument is not valid");
            }
        }

        [ExcelFunction(Description = "Unsubscribes from a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelUnsubscribe(
            [ExcelArgument(Description = "Redis channel to unsubscribe from")] object channel
        )
        {
            try
            {
                string channelStr = ToRedisString(channel);
                foreach (var kv in _channelListeners)
                {
                    if (!string.Equals(kv.Value.Channel, channelStr, StringComparison.Ordinal))
                        continue;
                    if (_channelListeners.TryRemove(kv.Key, out var listener))
                    {
                        // Mark closed before disposing so an in-flight callback
                        // stops writing; only then drop the latest message.
                        listener.Close();
                        listener.Token?.Dispose();
                        _latestMessages.TryRemove(kv.Key, out _);
                    }
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelUnsubscribe: channel={channelStr} unsubscribed");
                return $"Channel '{channelStr}' unsubscribed successfully.";
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
            [ExcelArgument(Description = "Redis channel to read from")] object channel,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                string key = ChannelKey(host, channelStr);
                if (!_channelListeners.ContainsKey(key))
                {
                    var listener = new ChannelListener { Channel = channelStr };
                    listener.Token = RedisRuntime.Subscriptions.Subscribe(host, channelStr, pattern: false,
                        onMessage: message =>
                        {
                            if (listener.IsClosed)
                                return;
                            _latestMessages[key] = message ?? "";
                        }, origin: "UDF");
                    if (!_channelListeners.TryAdd(key, listener))
                    {
                        // Another thread registered first: close before disposing
                        // so this losing listener never writes.
                        listener.Close();
                        listener.Token.Dispose();
                    }
                }
                var response = _latestMessages.TryGetValue(key, out var latest) ? latest : "(null)";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelLatest: channel={channelStr}, msg={response}, host={host}");
                return response;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelLatest", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishIfChangedJSON(
            [ExcelArgument(Description = "Redis channel")] object channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            string json = ExcelJson.RedisUDFMatrixToJSON(range);
            // Do not publish conversion errors; return them to Excel instead.
            if (json.StartsWith("Error:", StringComparison.Ordinal))
                return json;
            return RedisUDFChannelPublishIfChanged(channel, json, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel only if subscribers are present", IsVolatile = true)]
        public static string RedisUDFChannelPublishIfChanged(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                string messageStr = ToRedisString(message) ?? "";
                long readers = subscriber.Publish(new RedisChannel(channelStr, RedisChannel.PatternMode.Literal), messageStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channelStr}, msg={message}, readers={readers}, host={host}");
                return readers > 0 ? $"{readers} readers(s)" : "No Readers";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelPublishIfChanged", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON", IsVolatile = true)]
        public static object RedisUDFChannelPublishJSON(
            [ExcelArgument(Description = "Redis channel")] object channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            string json = ExcelJson.RedisUDFMatrixToJSON(range);
            // Do not publish conversion errors; return them to Excel instead.
            if (json.StartsWith("Error:", StringComparison.Ordinal))
                return json;
            return RedisUDFChannelPublish(channel, json, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel", IsVolatile = true)]
        public static string RedisUDFChannelPublish(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                string messageStr = ToRedisString(message) ?? "";
                long readers = subscriber.Publish(new RedisChannel(channelStr, RedisChannel.PatternMode.Literal), messageStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublish: channel={channelStr}, msg={message}, readers={readers}, host={host}");
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
            [ExcelArgument(Description = "Redis key to retrieve the value from")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var value = GetDb(host).StringGet(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGet: key={keyStr}, value={value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFGet", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the type of a Redis key", IsVolatile = true)]
        public static object RedisUDFType(
            [ExcelArgument(Description = "Redis key to check the type of")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var type = GetDb(host).KeyType(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFType: key={keyStr}, type={type}, host={host}");
                switch (type)
                {
                    case RedisType.String: return "string";
                    case RedisType.List: return "list";
                    case RedisType.Set: return "set";
                    case RedisType.SortedSet: return "zset";
                    case RedisType.Hash: return "hash";
                    case RedisType.Stream: return "stream";
                    case RedisType.Unknown: return "unknown";
                    case RedisType.None: return "none";
                    default: return "none";
                }
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFType", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Renames a Redis key", IsVolatile = true)]
        public static string RedisUDFRename(
            [ExcelArgument(Description = "Redis key to rename")] object key,
            [ExcelArgument(Description = "New name for the Redis key")] object newKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                string newKeyStr = ToRedisString(newKey);
                GetDb(host).KeyRename(keyStr, newKeyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFRename: key={keyStr}, newKey={newKeyStr}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFRename", ex, $"key={key}, newKey={newKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with a Matrix using JSON", IsVolatile = true)]
        public static string RedisUDFSetJSON(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
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
                // Surface conversion errors to Excel instead of storing them as the value.
                if (json.StartsWith("Error:", StringComparison.Ordinal))
                    return json;
                string keyStr = ToRedisString(key);
                GetDb(host).StringSet(keyStr, json);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetJSON: key={keyStr}, value={json}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetJSON", ex, $"key={key}, value={json}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key", IsVolatile = true)]
        public static string RedisUDFSet(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                string valueStr = ToRedisString(value);
                // A null RedisValue would issue DEL instead of storing an empty string.
                GetDb(host).StringSet(keyStr, valueStr ?? "");
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSet: key={keyStr}, value={value}, host={host}");
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
                    var key = ToRedisString(keys[i, 0]);
                    var value = ToRedisString(values[i, 0]);
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value ?? ""));
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
                    var key = ToRedisString(keyValuePairs[i, 0]);
                    var value = ToRedisString(keyValuePairs[i, 1]);
                    if (!string.IsNullOrWhiteSpace(key))
                        entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, value ?? ""));
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
            [ExcelArgument(Description = "Array of Redis keys to retrieve")] object[,] keys,
            [ExcelArgument(Description = "If TRUE, returns two columns (key, value); if FALSE, only values")] object multipleColumnsOpt,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (multipleColumnsOpt is ExcelError)
                    throw new ArgumentException("Excel error cells are not valid arguments");
                if (multipleColumnsOpt is Array)
                    throw new ArgumentException("A multi-cell range is not a valid scalar argument");
                bool multipleColumns = multipleColumnsOpt is bool b && b;

                int rows = keys.GetLength(0);
                int cols = keys.GetLength(1);
                var validKeys = new List<string>(rows * cols);
                // Excel passes ranges row-major; flatten them in that order so the
                // output rows follow the input order.
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var key = ToRedisString(keys[r, c]) ?? "";
                        if (!string.IsNullOrWhiteSpace(key))
                            validKeys.Add(key);
                    }
                }
                if (validKeys.Count == 0)
                    return FailMatrix("RedisUDFGetMultiple", new ArgumentException("No valid key"), "keys range contains no non-blank cells");

                var redisKeys = validKeys.Select(k => (RedisKey)k).ToArray();
                var values = GetDb(host).StringGet(redisKeys);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFGetMultiple: keys={validKeys.Count}, multipleColumns={multipleColumns}, host={host}");

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
            [ExcelArgument(Description = "Pattern to match Redis keys (e.g., user:*)")] object pattern,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost,
            [ExcelArgument(Description = "Optional page size for SCAN (default 250)")] object pageSize = null
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string patternStr = ToRedisString(pattern);
                if (string.IsNullOrEmpty(patternStr))
                    throw new ArgumentException("a key pattern is required; use \"*\" to match all keys");
                var conn = RedisRuntime.Connections.GetConnection(host, RedisPool.UdfData);
                var server = conn.GetServer(conn.GetEndPoints().First());
                List<string> keys;
                bool hasPageSize = false;
                int pageSizeInt = 0;
                if (pageSize != null && !(pageSize is ExcelMissing) && !(pageSize is ExcelEmpty))
                {
                    // Excel error cells and multi-cell ranges surface as an
                    // Error cell through the shared validator.
                    if (pageSize is ExcelError || pageSize is Array)
                        ToInt64Invariant(pageSize);

                    // Historical behavior: only integral values in int range
                    // select a page size; fractional, boolean, non-numeric and
                    // out-of-range values fall back to the default.
                    if (int.TryParse(Convert.ToString(pageSize, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out pageSizeInt)
                        && pageSizeInt > 0)
                    {
                        hasPageSize = true;
                    }
                }
                if (hasPageSize)
                    keys = server.Keys(pattern: patternStr, pageSize: pageSizeInt).Select(k => k.ToString()).ToList();
                else
                    keys = server.Keys(pattern: patternStr).Select(k => k.ToString()).ToList();

                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFKeys: pattern={patternStr}, found={keys.Count}, host={host}");
                if (keys.Count == 0)
                    return new object[,] { { "" } };
                var result = new object[keys.Count, 1];
                for (int i = 0; i < keys.Count; i++)
                    result[i, 0] = keys[i];
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFKeys", ex, $"pattern={pattern}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the TTL of a Redis key in seconds", IsVolatile = true)]
        public static object RedisUDFTTL(
            [ExcelArgument(Description = "Redis key to check TTL for")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var ttl = GetDb(host).KeyTimeToLive(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTL: key={keyStr}, ttl={ttl}, host={host}");
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
            [ExcelArgument(Description = "Array of Redis keys to check for existence")] object[,] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                int rows = keys.GetLength(0);
                int cols = keys.GetLength(1);
                int count = rows * cols;
                if (count == 0)
                    return new object[,] { { "" } };
                var result = new object[count, 2];
                var keysList = new List<RedisKey>(count);
                // Excel passes ranges row-major; flatten them in that order so the
                // output rows follow the input order.
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        // Blank cells map to an empty Redis key so row positions are preserved.
                        var key = ToRedisString(keys[r, c]) ?? "";
                        int i = r * cols + c;
                        result[i, 0] = key;
                        keysList.Add(key);
                    }
                }
                // One round trip for all keys instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                var tasks = keysList.Select(k => batch.KeyExistsAsync(k)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                    result[i, 1] = tasks[i].GetAwaiter().GetResult() ? "1" : "0";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExistsMultiples: {count} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFExistsMultiples", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Checks if a Redis key exists", IsVolatile = true)]
        public static object RedisUDFExists(
            [ExcelArgument(Description = "Redis key to check for existence")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var exists = GetDb(host).KeyExists(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExists: key={keyStr}, exists={exists}");
                return exists ? "1" : "0";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFExists", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Deletes a Redis key", IsVolatile = true)]
        public static object RedisUDFDel(
            [ExcelArgument(Description = "Redis key to delete")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                long deleted = GetDb(host).KeyDelete(keyStr) ? 1L : 0L;
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFDel: key={keyStr}, deleted={deleted}, host={host}");
                return deleted;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFDel", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with an expiry in seconds", IsVolatile = true)]
        public static string RedisUDFSetEx(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Time to live in seconds")] object ttlSeconds,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long ttl = ToInt64Invariant(ttlSeconds);
                string keyStr = ToRedisString(key);
                string valueStr = ToRedisString(value);
                // A null RedisValue would issue DEL instead of storing an empty string.
                GetDb(host).StringSet(keyStr, valueStr ?? "", TimeSpan.FromSeconds(ttl));
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetEx: key={keyStr}, value={value}, ttl={ttl}s, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetEx", ex, $"key={key}, value={value}, ttl={ttlSeconds}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets a timeout on a Redis key in seconds", IsVolatile = true)]
        public static object RedisUDFExpire(
            [ExcelArgument(Description = "Redis key to set the expiry for")] object key,
            [ExcelArgument(Description = "Time to live in seconds")] object ttlSeconds,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long ttl = ToInt64Invariant(ttlSeconds);
                if (ttl <= 0)
                    return "Error: ttl must be a positive number of seconds";
                string keyStr = ToRedisString(key);
                bool expired = GetDb(host).KeyExpire(keyStr, TimeSpan.FromSeconds(ttl));
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFExpire: key={keyStr}, ttl={ttl}s, expired={expired}, host={host}");
                return expired ? "1" : "0";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFExpire", ex, $"key={key}, ttl={ttlSeconds}, host={host}");
            }
        }

        [ExcelFunction(Description = "Increments the integer value of a Redis key by one", IsVolatile = true)]
        public static object RedisUDFIncr(
            [ExcelArgument(Description = "Redis key to increment")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                long value = GetDb(host).StringIncrement(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFIncr: key={keyStr}, value={value}, host={host}");
                return value;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFIncr", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Increments the integer value of a Redis key by a given amount", IsVolatile = true)]
        public static object RedisUDFIncrBy(
            [ExcelArgument(Description = "Redis key to increment")] object key,
            [ExcelArgument(Description = "Amount to increment by")] object increment,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long incr = ToInt64Invariant(increment);
                string keyStr = ToRedisString(key);
                long value = GetDb(host).StringIncrement(keyStr, incr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFIncrBy: key={keyStr}, increment={incr}, value={value}, host={host}");
                return value;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFIncrBy", ex, $"key={key}, increment={increment}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the TTL of multiple Redis keys in seconds", IsVolatile = true)]
        public static object[,] RedisUDFTTLMultiples(
            [ExcelArgument(Description = "Array of Redis keys to check TTL for")] object[,] keys,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                int rows = keys.GetLength(0);
                int cols = keys.GetLength(1);
                int count = rows * cols;
                if (count == 0)
                    return new object[,] { { "" } };
                var result = new object[count, 2];
                var keysList = new List<RedisKey>(count);
                // Excel passes ranges row-major; flatten them in that order so the
                // output rows follow the input order.
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        // Blank cells map to an empty Redis key so row positions are preserved.
                        var key = ToRedisString(keys[r, c]) ?? "";
                        int i = r * cols + c;
                        result[i, 0] = key;
                        keysList.Add(key);
                    }
                }
                // One round trip for all keys instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                var tasks = keysList.Select(k => batch.KeyTimeToLiveAsync(k)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                {
                    var ttl = tasks[i].GetAwaiter().GetResult();
                    result[i, 1] = ttl.HasValue ? ttl.Value.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) : "-1";
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFTTLMultiples: {count} keys, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFTTLMultiples", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Sets a field in a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashSet(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Field name to set within the hash")] object field,
            [ExcelArgument(Description = "Value to set for the given field")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = ToRedisString(hashKey);
                string fieldStr = ToRedisString(field);
                string valueStr = ToRedisString(value);
                // A null RedisValue would issue HDEL instead of storing an empty string.
                GetDb(host).HashSet(hashKeyStr, fieldStr, valueStr ?? "");
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSet: {hashKeyStr}[{fieldStr}] = {value}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashSet", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets multiple fields in a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashSetMultiple(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
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
                    var field = ToRedisString(fieldValuePairs[i, 0]);
                    var value = ToRedisString(fieldValuePairs[i, 1]);
                    if (!string.IsNullOrWhiteSpace(field))
                        entries.Add(new HashEntry(field, value ?? ""));
                }
                string hashKeyStr = ToRedisString(hashKey);
                GetDb(host).HashSet(hashKeyStr, entries.ToArray());
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSetMultiple: {hashKeyStr}, fields={entries.Count}, host={host}");
                return "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashSetMultiple", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Gets a field from a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashGet(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Field name to retrieve from the hash")] object field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = ToRedisString(hashKey);
                string fieldStr = ToRedisString(field);
                var value = GetDb(host).HashGet(hashKeyStr, fieldStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGet: {hashKeyStr}[{fieldStr}] = {value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashGet", ex, $"{hashKey}[{field}], host={host}");
            }
        }

        [ExcelFunction(Description = "Gets all fields of a Redis hash", IsVolatile = true)]
        public static object[,] RedisUDFHashGetAll(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = ToRedisString(hashKey);
                var all = GetDb(host).HashGetAll(hashKeyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGetAll: {hashKeyStr}, fields={all.Length}, host={host}");
                if (all.Length == 0)
                    return new object[,] { { "" } };
                var result = new object[all.Length, 2];
                for (int i = 0; i < all.Length; i++)
                {
                    result[i, 0] = all[i].Name.ToString();
                    result[i, 1] = all[i].Value.ToString();
                }
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFHashGetAll", ex, $"{hashKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Gets the same field from multiple Redis hashes", IsVolatile = true)]
        public static object[,] RedisUDFHashGetFieldMultipleKeys(
            [ExcelArgument(Description = "Array of Redis hash keys")] object[,] hashKeys,
            [ExcelArgument(Description = "Field name to retrieve from each hash")] object field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                int rows = hashKeys.GetLength(0);
                int cols = hashKeys.GetLength(1);
                int count = rows * cols;
                if (count == 0)
                    return new object[,] { { "" } };
                var result = new object[count, 2];
                var keysList = new List<RedisKey>(count);
                // Excel passes ranges row-major; flatten them in that order so the
                // output rows follow the input order.
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        // Blank cells map to an empty Redis key so row positions are preserved.
                        string key = ToRedisString(hashKeys[r, c]) ?? "";
                        int i = r * cols + c;
                        result[i, 0] = key;
                        keysList.Add(key);
                    }
                }
                // One round trip for all hashes instead of one command per key.
                var batch = GetDb(host).CreateBatch();
                string fieldStr = ToRedisString(field);
                var tasks = keysList.Select(k => batch.HashGetAsync(k, fieldStr)).ToArray();
                batch.Execute();
                for (int i = 0; i < tasks.Length; i++)
                    result[i, 1] = (string)tasks[i].GetAwaiter().GetResult() ?? "";
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashGetFieldMultipleKeys: field={fieldStr}, hashes={count}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFHashGetFieldMultipleKeys", ex, $"field={field}, host={host}");
            }
        }

        [ExcelFunction(Description = "Deletes a field from a Redis hash", IsVolatile = true)]
        public static object RedisUDFHashDel(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Field name to delete from the hash")] object field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = ToRedisString(hashKey);
                string fieldStr = ToRedisString(field);
                long deleted = GetDb(host).HashDelete(hashKeyStr, fieldStr) ? 1L : 0L;
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashDel: {hashKeyStr}[{fieldStr}] deleted={deleted}, host={host}");
                return deleted;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFHashDel", ex, $"{hashKey}[{field}], host={host}");
            }
        }

        [ExcelFunction(Description = "Pushes a value onto the right end of a Redis list", IsVolatile = true)]
        public static object RedisUDFListPushRight(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Value to push onto the list")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                long length = GetDb(host).ListRightPush(keyStr, ToRedisString(value) ?? "");
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPushRight: key={keyStr}, value={value}, length={length}, host={host}");
                return length;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFListPushRight", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Pushes a value onto the left end of a Redis list", IsVolatile = true)]
        public static object RedisUDFListPushLeft(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Value to push onto the list")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                long length = GetDb(host).ListLeftPush(keyStr, ToRedisString(value) ?? "");
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPushLeft: key={keyStr}, value={value}, length={length}, host={host}");
                return length;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFListPushLeft", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns a range of values from a Redis list", IsVolatile = true)]
        public static object[,] RedisUDFListRange(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Start index (0-based, negative counts from the end)")] object start,
            [ExcelArgument(Description = "Stop index (inclusive, negative counts from the end)")] object stop,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var values = GetDb(host).ListRange(keyStr, ToInt64Invariant(start), ToInt64Invariant(stop));
                if (values.Length == 0)
                {
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFListRange: key={keyStr}, empty, host={host}");
                    return new object[,] { { "" } };
                }
                var result = new object[values.Length, 1];
                for (int i = 0; i < values.Length; i++)
                    result[i, 0] = values[i].ToString();
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListRange: key={keyStr}, values={values.Length}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFListRange", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Removes and returns the last element of a Redis list", IsVolatile = true)]
        public static object RedisUDFListPopRight(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var value = GetDb(host).ListRightPop(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPopRight: key={keyStr}, value={value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFListPopRight", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Removes and returns the first element of a Redis list", IsVolatile = true)]
        public static object RedisUDFListPopLeft(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var value = GetDb(host).ListLeftPop(keyStr);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPopLeft: key={keyStr}, value={value}, host={host}");
                return value.HasValue ? value.ToString() : "";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFListPopLeft", ex, $"key={key}, host={host}");
            }
        }

        [ExcelFunction(Description = "Adds a member to a Redis set", IsVolatile = true)]
        public static object RedisUDFSetAdd(
            [ExcelArgument(Description = "Redis set key")] object key,
            [ExcelArgument(Description = "Member to add to the set")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                string valueStr = ToRedisString(value);
                long added = GetDb(host).SetAdd(keyStr, valueStr ?? "") ? 1L : 0L;
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetAdd: key={keyStr}, value={valueStr}, added={added}, host={host}");
                return added;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetAdd", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Removes a member from a Redis set", IsVolatile = true)]
        public static object RedisUDFSetRemove(
            [ExcelArgument(Description = "Redis set key")] object key,
            [ExcelArgument(Description = "Member to remove from the set")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                string valueStr = ToRedisString(value);
                long removed = GetDb(host).SetRemove(keyStr, valueStr ?? "") ? 1L : 0L;
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetRemove: key={keyStr}, value={valueStr}, removed={removed}, host={host}");
                return removed;
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetRemove", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns all members of a Redis set", IsVolatile = true)]
        public static object[,] RedisUDFSetMembers(
            [ExcelArgument(Description = "Redis set key")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = ToRedisString(key);
                var members = GetDb(host).SetMembers(keyStr);
                if (members.Length == 0)
                {
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFSetMembers: key={keyStr}, empty, host={host}");
                    return new object[,] { { "" } };
                }
                var result = new object[members.Length, 1];
                for (int i = 0; i < members.Length; i++)
                    result[i, 0] = members[i].ToString();
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetMembers: key={keyStr}, members={members.Length}, host={host}");
                return result;
            }
            catch (Exception ex)
            {
                return FailMatrix("RedisUDFSetMembers", ex, $"key={key}, host={host}");
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

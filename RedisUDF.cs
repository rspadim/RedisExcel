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
        // Last payload published per host/channel key, so PublishIfChanged can
        // skip a message that did not change since the last successful publish.
        // Bounded LRU: the marker is only remembered after a publish that had
        // readers, and concurrent check/publish/store must be atomic per key.
        private static readonly PublishDedupCache _lastPublishedMessages =
            new PublishDedupCache(AppConfig.Current.PublishDedupCacheSize);

        // Striped locks: PublishIfChanged serializes its check -> publish -> store
        // sequence per host/channel key without a lock per key. The marker clears
        // in ChannelLatest/ChannelUnsubscribe take the same stripe.
        private const int PublishLockStripeCount = 64;
        private static readonly object[] _publishLocks = CreatePublishLocks();

        private static object[] CreatePublishLocks()
        {
            var locks = new object[PublishLockStripeCount];
            for (int i = 0; i < locks.Length; i++)
                locks[i] = new object();
            return locks;
        }

        private static object PublishLock(string key)
        {
            int hash = key.GetHashCode() & 0x7FFFFFFF;
            return _publishLocks[hash % _publishLocks.Length];
        }

        static RedisUDF()
        {
            // RTD SUB/PSUB listeners consume UDF publishes too. When such a
            // listener joins (initially or after a retry), the publish dedup
            // marker must be dropped, otherwise a recalculation that did not
            // change the payload answers "No change" while the returning
            // listener never received it (the ChannelState reset only helps
            // the subscription-side fan-out).
            RedisSubscriptionManager.ListenerJoined += HandleListenerJoined;
        }

        /// <summary>Dedup marker cache, exposed internally so the listener-join
        /// clearing rules can be unit tested offline.</summary>
        internal static PublishDedupCache LastPublishedMessagesForTests => _lastPublishedMessages;

        /// <summary>
        /// Entry point for <see cref="RedisSubscriptionManager.ListenerJoined"/>:
        /// clears the publish dedup markers that a newly joined listener
        /// invalidated. Never throws into Subscribe.
        /// </summary>
        internal static void HandleListenerJoined(string host, string channel, bool pattern)
        {
            try
            {
                ClearForListener(host, channel, pattern);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"HandleListenerJoined: host={host}, channel={channel}, pattern={pattern}");
            }
        }

        /// <summary>
        /// Clears the publish dedup markers invalidated by a listener join:
        /// the exact key for a literal subscription, or every marker of the
        /// same host whose channel matches the Redis glob for a pattern.
        /// </summary>
        internal static void ClearForListener(string host, string channel, bool pattern)
        {
            if (!pattern)
            {
                ClearMarker(ChannelKey(host, channel));
                return;
            }

            // Work from a snapshot: PublishIfChanged takes the per-key stripe
            // and then the cache lock, so holding the cache lock while taking
            // a stripe (the opposite order) could deadlock the two.
            string[] keys = _lastPublishedMessages.SnapshotKeys();
            for (int i = 0; i < keys.Length; i++)
            {
                if (!TryParseChannelKey(keys[i], out var keyHost, out var keyChannel))
                    continue; // not one of our markers: leave it alone
                if (!string.Equals(keyHost, host, StringComparison.Ordinal))
                    continue;
                if (GlobMatches(channel, keyChannel))
                    ClearMarker(keys[i]);
            }
        }

        private static void ClearMarker(string key)
        {
            lock (PublishLock(key))
                _lastPublishedMessages.Remove(key);
        }

        /// <summary>
        /// Splits a <see cref="ChannelKey(string, string)"/> back into host and
        /// channel. Returns false for malformed keys, so foreign markers are
        /// skipped instead of being misattributed to a host.
        /// </summary>
        internal static bool TryParseChannelKey(string key, out string host, out string channel)
        {
            host = null;
            channel = null;
            if (string.IsNullOrEmpty(key))
                return false;
            int separator = key.IndexOf(':');
            if (separator <= 0)
                return false;
            if (!int.TryParse(key.Substring(0, separator), NumberStyles.None, CultureInfo.InvariantCulture, out int hostLength))
                return false;
            int hostStart = separator + 1;
            int channelStart = hostStart + hostLength;
            // The channel (which may be empty) is always preceded by ':'.
            if (channelStart >= key.Length || key[channelStart] != ':')
                return false;
            host = key.Substring(hostStart, hostLength);
            channel = key.Substring(channelStart + 1);
            return true;
        }

        /// <summary>
        /// Case-sensitive Redis-style glob matcher: '*' matches any (possibly
        /// empty) sequence, '?' exactly one character, '[...]' a character class
        /// with optional '^' negation and 'a-z' ranges, and '\' escapes the next
        /// character. Iterative with single-star backtracking, so pathological
        /// patterns cannot blow up like a backtracking regex.
        /// </summary>
        internal static bool GlobMatches(string pattern, string value)
        {
            if (pattern == null)
                return false;
            value = value ?? "";
            int p = 0;
            int v = 0;
            int starP = -1;
            int starV = -1;
            while (v < value.Length)
            {
                if (p < pattern.Length && pattern[p] == '*')
                {
                    // Try the shortest suffix first; the fallback below extends
                    // the star match one character at a time.
                    starP = p++;
                    starV = v;
                    continue;
                }
                if (p < pattern.Length && MatchOne(pattern, ref p, value[v]))
                {
                    v++;
                    continue;
                }
                if (starP >= 0)
                {
                    p = starP + 1;
                    v = ++starV;
                    continue;
                }
                return false;
            }
            // Trailing '*' characters may match the (now exhausted) rest.
            while (p < pattern.Length && pattern[p] == '*')
                p++;
            return p == pattern.Length;
        }

        private static bool MatchOne(string pattern, ref int p, char c)
        {
            char current = pattern[p];
            if (current == '?')
            {
                p++;
                return true;
            }
            if (current == '[')
            {
                int end = FindClassEnd(pattern, p);
                if (end < 0)
                {
                    // Unterminated class: a literal '[' (like Redis).
                    p++;
                    return c == '[';
                }
                bool matched = ClassMatches(pattern, p + 1, end, c);
                if (matched)
                    p = end + 1;
                return matched;
            }
            if (current == '\\' && p + 1 < pattern.Length)
            {
                p += 2;
                return pattern[p - 1] == c;
            }
            p++;
            return current == c;
        }

        /// <summary>Index of the ']' closing the class opened at start, or -1
        /// when the class is unterminated. A ']' directly after '[' or '[^' is
        /// a literal member, like in Redis.</summary>
        private static int FindClassEnd(string pattern, int start)
        {
            int i = start + 1;
            if (i < pattern.Length && pattern[i] == '^')
                i++;
            if (i < pattern.Length && pattern[i] == ']')
                i++;
            while (i < pattern.Length)
            {
                if (pattern[i] == '\\' && i + 1 < pattern.Length)
                {
                    i += 2;
                    continue;
                }
                if (pattern[i] == ']')
                    return i;
                i++;
            }
            return -1;
        }

        private static bool ClassMatches(string pattern, int contentStart, int classEnd, char c)
        {
            bool negate = contentStart < classEnd && pattern[contentStart] == '^';
            int i = negate ? contentStart + 1 : contentStart;
            bool matched = false;
            while (i < classEnd)
            {
                // 'a-z' range when '-' sits between two members; elsewhere the
                // dash is a literal member.
                if (i + 2 < classEnd && pattern[i + 1] == '-')
                {
                    if (c >= pattern[i] && c <= pattern[i + 2])
                        matched = true;
                    i += 3;
                    continue;
                }
                if (pattern[i] == c)
                    matched = true;
                i++;
            }
            return negate ? !matched : matched;
        }

        /// <summary>Registry key for a (host, channel) pair. The host length is
        /// length-prefixed so hosts and channels that themselves contain the
        /// separator cannot collide. Format: {host.Length}:{host}:{channel}.</summary>
        internal static string ChannelKey(string host, string channel) => $"{host.Length}:{host}:{channel}";

        /// <summary>Validates and returns a required scalar text argument (key,
        /// hash key, field). Only truly missing cells (null / ExcelEmpty /
        /// ExcelMissing) are rejected; an explicit empty string is a valid Redis
        /// name, exactly like in the scalar write functions. Runs before any
        /// Redis I/O so StackExchange.Redis' "null key" error is never surfaced.</summary>
        internal static string RequireText(object value, string what)
        {
            string text = ToRedisString(value);
            if (text == null)
                throw new ArgumentException("a " + what + " is required");
            return text;
        }

        private static string ResolveHost(object optionalHost)
        {
            // null / empty cell / omitted argument all mean "use the default host".
            if (optionalHost == null || optionalHost is ExcelMissing || optionalHost is ExcelEmpty)
                return AppConfig.ResolveUdfHost(null);
            if (!(optionalHost is string host))
                throw new ArgumentException("host must be a text value");
            return AppConfig.ResolveUdfHost(host);
        }

        /// <summary>
        /// Host resolution for the async dispatcher: resolves like
        /// <see cref="ResolveHost"/> but never throws, returning null when the
        /// host argument is invalid so the async helper can dispatch and let
        /// the Core body surface the error.
        /// </summary>
        internal static string ResolveHostForDispatch(object optionalHost)
        {
            try
            {
                return ResolveHost(optionalHost);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>Test-only override of <see cref="AppConfig.SyncWrite"/>
        /// (null = use the process configuration, pinned by unit tests).</summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static string SyncWriteOverrideForTests;
#pragma warning restore 0649

        // TimeSpan.FromSeconds overflows past ~9.22e11 seconds; validate first
        // so the cell gets a stable invariant message instead of a localized
        // runtime OverflowException.
        private static readonly long MaxTtlSeconds = (long)TimeSpan.MaxValue.TotalSeconds;

        /// <summary>
        /// Whether a write function should use CommandFlags.FireAndForget,
        /// according to the configured SyncWrite mode: "sync" never fires and
        /// forgets, "fireforget-all" always does, and the default "fireforget"
        /// does it only for reply-agnostic writes.
        /// </summary>
        internal static bool ShouldFireAndForget(bool replyDependent)
        {
            switch (SyncWriteOverrideForTests ?? AppConfig.SyncWrite)
            {
                case "sync":
                    return false;
                case "fireforget-all":
                    return true;
                default: // "fireforget": reply-agnostic writes only
                    return !replyDependent;
            }
        }

        /// <summary>
        /// Result text shown in the cell when a write was sent with
        /// CommandFlags.FireAndForget: there is no reply to report, so a marker
        /// takes its place ("OK-FireForgetAll" when a reply-dependent write was
        /// forced to fire and forget by "fireforget-all").
        /// </summary>
        internal static string FireAndForgetMarker(bool replyDependent) =>
            replyDependent ? "OK-FireForgetAll" : "OK FireForget";

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

        /// <summary>Flattens an Excel range row-major (the order Excel uses), so a
        /// cell count is a stable pairing unit independent of the range shape.</summary>
        private static List<object> FlattenRowMajor(object[,] range)
        {
            int rows = range.GetLength(0);
            int cols = range.GetLength(1);
            var cells = new List<object>(rows * cols);
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    cells.Add(range[r, c]);
            return cells;
        }

        /// <summary>Returns key/value cells as pairs. A range with exactly two
        /// columns is read row by row; otherwise a range with exactly two rows is
        /// read column by column (horizontal layout). Any other shape is an
        /// error instead of silently using only part of the range.</summary>
        private static List<KeyValuePair<object, object>> FlattenPairRange(object[,] range)
        {
            int rows = range.GetLength(0);
            int cols = range.GetLength(1);
            var pairs = new List<KeyValuePair<object, object>>();
            if (cols == 2)
            {
                for (int r = 0; r < rows; r++)
                    pairs.Add(new KeyValuePair<object, object>(range[r, 0], range[r, 1]));
            }
            else if (rows == 2)
            {
                for (int c = 0; c < cols; c++)
                    pairs.Add(new KeyValuePair<object, object>(range[0, c], range[1, c]));
            }
            else
            {
                throw new ArgumentException("expected a range with 2 columns or 2 rows");
            }
            return pairs;
        }

        /// <summary>Builds the SET batch entries from key/value cell pairs. Only
        /// truly missing cells (ToRedisString returns null: null, ExcelMissing,
        /// ExcelEmpty) are skipped; an explicit empty or whitespace-only cell is
        /// a valid Redis name, exactly like in the scalar write functions.</summary>
        internal static List<KeyValuePair<RedisKey, RedisValue>> CollectStringSetEntries(
            IEnumerable<KeyValuePair<object, object>> pairs)
        {
            var entries = new List<KeyValuePair<RedisKey, RedisValue>>();
            foreach (var pair in pairs)
            {
                string key = ToRedisString(pair.Key);
                if (key == null)
                    continue;
                entries.Add(new KeyValuePair<RedisKey, RedisValue>(key, ToRedisString(pair.Value) ?? ""));
            }
            return entries;
        }

        /// <summary>Same filtering contract as <see cref="CollectStringSetEntries"/>
        /// for HSET field/value pair ranges.</summary>
        internal static List<HashEntry> CollectHashEntries(
            IEnumerable<KeyValuePair<object, object>> pairs)
        {
            var entries = new List<HashEntry>();
            foreach (var pair in pairs)
            {
                string field = ToRedisString(pair.Key);
                if (field == null)
                    continue;
                entries.Add(new HashEntry(field, ToRedisString(pair.Value) ?? ""));
            }
            return entries;
        }

        /// <summary>Coerces an Excel-friendly numeric flag: an integral number
        /// maps 0 to false and any other value to true. Returns false for
        /// non-integral or non-numeric values so the caller can reject them.</summary>
        private static bool TryGetIntegralFlag(object value, out bool flag)
        {
            double number;
            switch (value)
            {
                case double d: number = d; break;
                case float f: number = f; break;
                case decimal m: number = (double)m; break;
                case long l: number = l; break;
                case int i: number = i; break;
                case short s: number = s; break;
                case byte b: number = b; break;
                case sbyte sb: number = sb; break;
                case ushort us: number = us; break;
                case uint ui: number = ui; break;
                case ulong ul: number = ul; break;
                default: flag = false; return false;
            }
            if (double.IsNaN(number) || double.IsInfinity(number) || number != Math.Truncate(number))
            {
                flag = false;
                return false;
            }
            flag = number != 0;
            return true;
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
            // ISO-8601 round-trip text so JSON wrappers can parse date/time
            // values (only reachable programmatically, not from a cell).
            if (value is DateTime dateTime)
                return dateTime.ToString("o", CultureInfo.InvariantCulture);
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
            // Booleans convert to 1/0 and fractions silently truncate; TTL and
            // increment arguments must be whole numbers.
            if (value is bool)
                throw new ArgumentException("numeric argument is not valid");
            if (value is double d && !double.IsNaN(d) && d != Math.Truncate(d))
                throw new ArgumentException("numeric argument is not an integer");
            if (value is decimal m && m != Math.Truncate(m))
                throw new ArgumentException("numeric argument is not an integer");
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
        public static object RedisUDFChannelUnsubscribe(
            [ExcelArgument(Description = "Redis channel to unsubscribe from")] object channel,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost = null
        )
        {
            return RedisUdfAsync.Run("RedisUDFChannelUnsubscribe", optionalHost,
                new object[] { channel, optionalHost },
                () => RedisUDFChannelUnsubscribeCore(channel, optionalHost));
        }

        [ExcelFunction(Description = "Unsubscribes from a Redis channel; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFChannelUnsubscribeNonVolatile(
            [ExcelArgument(Description = "Redis channel to unsubscribe from")] object channel,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost = null
        )
        {
            return RedisUDFChannelUnsubscribe(channel, optionalHost);
        }

        // Local listener bookkeeping only - never fire and forget: the caller
        // needs the outcome and the listeners must be removed deterministically.
        private static string RedisUDFChannelUnsubscribeCore(object channel, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                // A blank channel matches nothing; report it instead of
                // claiming success.
                if (string.IsNullOrWhiteSpace(channelStr))
                    throw new ArgumentException("a channel is required");
                // Only listeners registered for this host/channel pair are
                // removed; the same channel on another host stays subscribed.
                string key = ChannelKey(host, channelStr);
                foreach (var kv in _channelListeners)
                {
                    if (!string.Equals(kv.Key, key, StringComparison.Ordinal))
                        continue;
                    if (_channelListeners.TryRemove(kv.Key, out var listener))
                    {
                        // Mark closed before disposing so an in-flight callback
                        // stops writing; only then drop the latest message.
                        listener.Close();
                        listener.Token?.Dispose();
                        _latestMessages.TryRemove(kv.Key, out _);
                        // Drop the publish dedup marker with the listener so a
                        // rejoining listener receives the next publish even when
                        // the payload did not change.
                        lock (PublishLock(key))
                            _lastPublishedMessages.Remove(key);
                    }
                }
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelUnsubscribe: channel={channelStr}, host={host} unsubscribed");
                return $"Channel '{channelStr}' unsubscribed successfully.";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFChannelUnsubscribe", ex, $"channel={channel}, host={host}");
            }
        }

        [ExcelFunction(Description = "Returns the number of active Redis connections", IsVolatile = true)]
        public static object RedisUDFConnectionCount()
        {
            int count = RedisRuntime.Connections.LiveUdfConnectionCount();
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
                if (string.IsNullOrWhiteSpace(channelStr))
                    throw new ArgumentException("a channel is required");
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
                    if (_channelListeners.TryAdd(key, listener))
                    {
                        // A freshly registered listener never saw the currently
                        // remembered payload; drop the marker so the next publish
                        // is delivered even when the payload did not change.
                        lock (PublishLock(key))
                            _lastPublishedMessages.Remove(key);
                    }
                    else
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
            return RedisUdfAsync.Run("RedisUDFChannelPublishIfChangedJSON", optionalHost,
                new object[] { channel, range, optionalHost },
                () => RedisUDFChannelPublishIfChangedJSONCore(channel, range, optionalHost));
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFChannelPublishIfChangedJSONNonVolatile(
            [ExcelArgument(Description = "Redis channel")] object channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublishIfChangedJSON(channel, range, optionalHost);
        }

        private static object RedisUDFChannelPublishIfChangedJSONCore(object channel, object[,] range, object optionalHost)
        {
            if (range == null)
                return "Error: a range is required";
            string json = ExcelJson.RedisUDFMatrixToJSON(range);
            // Do not publish conversion errors; return them to Excel instead.
            if (json.StartsWith("Error:", StringComparison.Ordinal))
                return json;
            return RedisUDFChannelPublishIfChangedCore(channel, json, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel only if subscribers are present", IsVolatile = true)]
        public static object RedisUDFChannelPublishIfChanged(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFChannelPublishIfChanged", optionalHost,
                new object[] { channel, message, optionalHost },
                () => RedisUDFChannelPublishIfChangedCore(channel, message, optionalHost));
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel only if subscribers are present; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFChannelPublishIfChangedNonVolatile(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFChannelPublishIfChanged(channel, message, optionalHost);
        }

        private static string RedisUDFChannelPublishIfChangedCore(object channel, object message, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                if (string.IsNullOrWhiteSpace(channelStr))
                    throw new ArgumentException("a channel is required");
                string messageStr = ToRedisString(message) ?? "";
                string key = ChannelKey(host, channelStr);
                // Re-publishing an identical payload is a no-op; the last
                // successfully published payload is tracked per host/channel.
                // Check -> publish -> store is serialized per key so concurrent
                // recalculations cannot invert each other.
                lock (PublishLock(key))
                {
                    if (_lastPublishedMessages.TryGet(key, out var previous) &&
                        string.Equals(previous, messageStr, StringComparison.Ordinal))
                    {
                        if (logger.IsTraceEnabled)
                            logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channelStr}, host={host}, unchanged");
                        return "No change";
                    }
                    var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                    if (ShouldFireAndForget(replyDependent: false))
                    {
                        subscriber.Publish(new RedisChannel(channelStr, RedisChannel.PatternMode.Literal), messageStr,
                            CommandFlags.FireAndForget);
                        // Fire and forget yields no reader count; remember the
                        // payload as published anyway so repeated recalculations
                        // keep deduplicating. A listener that joins later clears
                        // the marker (ListenerJoined handling).
                        _lastPublishedMessages.Set(key, messageStr);
                        if (logger.IsTraceEnabled)
                            logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channelStr}, msg={message}, host={host}, fireAndForget=true");
                        return FireAndForgetMarker(replyDependent: false);
                    }
                    long readers = subscriber.Publish(new RedisChannel(channelStr, RedisChannel.PatternMode.Literal), messageStr);
                    // Remember the payload only when it was actually delivered;
                    // with zero readers the marker is dropped, so the volatile
                    // formula recalculates and retries the publish and a late
                    // consumer is never starved by a publish it did not see.
                    if (readers > 0)
                        _lastPublishedMessages.Set(key, messageStr);
                    else
                        _lastPublishedMessages.Remove(key);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFChannelPublishIfChanged: channel={channelStr}, msg={message}, readers={readers}, host={host}");
                    return readers > 0 ? $"{readers} readers(s)" : "No Readers";
                }
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
            return RedisUdfAsync.Run("RedisUDFChannelPublishJSON", optionalHost,
                new object[] { channel, range, optionalHost },
                () => RedisUDFChannelPublishJSONCore(channel, range, optionalHost));
        }

        [ExcelFunction(Description = "Publish an Excel matrix to a Redis channel as JSON; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFChannelPublishJSONNonVolatile(
            [ExcelArgument(Description = "Redis channel")] object channel,
            [ExcelArgument(Description = "Excel range to publish")] object[,] range,
            [ExcelArgument(Description = "Optional Redis host")] object optionalHost)
        {
            return RedisUDFChannelPublishJSON(channel, range, optionalHost);
        }

        private static object RedisUDFChannelPublishJSONCore(object channel, object[,] range, object optionalHost)
        {
            if (range == null)
                return "Error: a range is required";
            string json = ExcelJson.RedisUDFMatrixToJSON(range);
            // Do not publish conversion errors; return them to Excel instead.
            if (json.StartsWith("Error:", StringComparison.Ordinal))
                return json;
            return RedisUDFChannelPublishCore(channel, json, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel", IsVolatile = true)]
        public static object RedisUDFChannelPublish(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFChannelPublish", optionalHost,
                new object[] { channel, message, optionalHost },
                () => RedisUDFChannelPublishCore(channel, message, optionalHost));
        }

        [ExcelFunction(Description = "Publishes a message to a Redis channel; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFChannelPublishNonVolatile(
            [ExcelArgument(Description = "Redis channel to publish to")] object channel,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFChannelPublish(channel, message, optionalHost);
        }

        private static string RedisUDFChannelPublishCore(object channel, object message, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string channelStr = ToRedisString(channel);
                if (string.IsNullOrWhiteSpace(channelStr))
                    throw new ArgumentException("a channel is required");
                var subscriber = RedisRuntime.Connections.GetSubscriber(host, RedisPool.UdfData);
                string messageStr = ToRedisString(message) ?? "";
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                long readers = subscriber.Publish(new RedisChannel(channelStr, RedisChannel.PatternMode.Literal), messageStr,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFChannelPublish: channel={channelStr}, msg={message}, readers={readers}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : $"{readers} readers(s)";
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
                string keyStr = RequireText(key, "key");
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
                string keyStr = RequireText(key, "key");
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
        public static object RedisUDFRename(
            [ExcelArgument(Description = "Redis key to rename")] object key,
            [ExcelArgument(Description = "New name for the Redis key")] object newKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFRename", optionalHost, new object[] { key, newKey, optionalHost }, () => RedisUDFRenameCore(key, newKey, optionalHost));
        }

        [ExcelFunction(Description = "Renames a Redis key; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFRenameNonVolatile(
            [ExcelArgument(Description = "Redis key to rename")] object key,
            [ExcelArgument(Description = "New name for the Redis key")] object newKey,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFRename(key, newKey, optionalHost);
        }

        private static string RedisUDFRenameCore(object key, object newKey, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string newKeyStr = RequireText(newKey, "new key");
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                GetDb(host).KeyRename(keyStr, newKeyStr, When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFRename: key={keyStr}, newKey={newKeyStr}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFRename", ex, $"key={key}, newKey={newKey}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with a Matrix using JSON", IsVolatile = true)]
        public static object RedisUDFSetJSON(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFSetJSON", optionalHost, new object[] { key, values, optionalHost }, () => RedisUDFSetJSONCore(key, values, optionalHost));
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with a Matrix using JSON; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetJSONNonVolatile(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetJSON(key, values, optionalHost);
        }

        private static string RedisUDFSetJSONCore(object key, object[,] values, object optionalHost)
        {
            string host = null;
            string json = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (values == null)
                    throw new ArgumentException("a range is required");
                json = ExcelJson.RedisUDFMatrixToJSON(values);
                // Surface conversion errors to Excel instead of storing them as the value.
                if (json.StartsWith("Error:", StringComparison.Ordinal))
                    return json;
                string keyStr = RequireText(key, "key");
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                GetDb(host).StringSet(keyStr, json, null, When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetJSON: key={keyStr}, value={json}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetJSON", ex, $"key={key}, value={json}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets the value of a Redis key", IsVolatile = true)]
        public static object RedisUDFSet(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFSet", optionalHost, new object[] { key, value, optionalHost }, () => RedisUDFSetCore(key, value, optionalHost));
        }

        [ExcelFunction(Description = "Sets the value of a Redis key; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetNonVolatile(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSet(key, value, optionalHost);
        }

        private static string RedisUDFSetCore(object key, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value);
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                // A null RedisValue would issue DEL instead of storing an empty string.
                GetDb(host).StringSet(keyStr, valueStr ?? "", null, When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSet: key={keyStr}, value={value}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSet", ex, $"key={key}, value={value}, host={host}");
            }
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static object RedisUDFSetKV(
            [ExcelArgument(Description = "Range with Redis keys")] object[,] keys,
            [ExcelArgument(Description = "Range with Redis values")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFSetKV", optionalHost, new object[] { keys, values, optionalHost }, () => RedisUDFSetKVCore(keys, values, optionalHost));
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetKVNonVolatile(
            [ExcelArgument(Description = "Range with Redis keys")] object[,] keys,
            [ExcelArgument(Description = "Range with Redis values")] object[,] values,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetKV(keys, values, optionalHost);
        }

        private static string RedisUDFSetKVCore(object[,] keys, object[,] values, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (keys == null || values == null)
                    throw new ArgumentException("a range is required");
                var keyCells = FlattenRowMajor(keys);
                var valueCells = FlattenRowMajor(values);
                if (keyCells.Count == 0 || keyCells.Count != valueCells.Count)
                    throw new ArgumentException("keys and values must have the same number of cells");
                var pairs = new List<KeyValuePair<object, object>>(keyCells.Count);
                for (int i = 0; i < keyCells.Count; i++)
                    pairs.Add(new KeyValuePair<object, object>(keyCells[i], valueCells[i]));
                var entries = CollectStringSetEntries(pairs);
                if (entries.Count == 0)
                    throw new ArgumentException("no entries to write");
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                GetDb(host).StringSet(entries.ToArray(), When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKV: {entries.Count} pairs sent, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
            }
            catch (Exception ex)
            {
                return Fail("RedisUDFSetKV", ex, $"host={host}");
            }
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs", IsVolatile = true)]
        public static object RedisUDFSetKVPair(
            [ExcelArgument(Description = "2D range with Redis key-value pairs (2 columns: key, value)")] object[,] keyValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFSetKVPair", optionalHost, new object[] { keyValuePairs, optionalHost }, () => RedisUDFSetKVPairCore(keyValuePairs, optionalHost));
        }

        [ExcelFunction(Description = "Sets Redis key-value pairs; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetKVPairNonVolatile(
            [ExcelArgument(Description = "2D range with Redis key-value pairs (2 columns: key, value)")] object[,] keyValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetKVPair(keyValuePairs, optionalHost);
        }

        private static string RedisUDFSetKVPairCore(object[,] keyValuePairs, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (keyValuePairs == null)
                    throw new ArgumentException("a range is required");
                var pairs = FlattenPairRange(keyValuePairs);
                var entries = CollectStringSetEntries(pairs);
                if (entries.Count == 0)
                    throw new ArgumentException("no entries to write");
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                GetDb(host).StringSet(entries.ToArray(), When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetKVPair: {entries.Count} pairs sent, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
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
                if (keys == null)
                    throw new ArgumentException("a range is required");
                if (multipleColumnsOpt is ExcelError)
                    throw new ArgumentException("Excel error cells are not valid arguments");
                if (multipleColumnsOpt is Array)
                    throw new ArgumentException("A multi-cell range is not a valid scalar argument");
                bool multipleColumns;
                if (multipleColumnsOpt is bool boolFlag)
                    multipleColumns = boolFlag;
                else if (multipleColumnsOpt == null || multipleColumnsOpt is ExcelMissing || multipleColumnsOpt is ExcelEmpty)
                    multipleColumns = false;
                else if (TryGetIntegralFlag(multipleColumnsOpt, out bool numericFlag))
                    multipleColumns = numericFlag;
                else if (multipleColumnsOpt is string flagText)
                {
                    string trimmed = flagText.Trim();
                    if (string.Equals(trimmed, "TRUE", StringComparison.OrdinalIgnoreCase))
                        multipleColumns = true;
                    else if (string.Equals(trimmed, "FALSE", StringComparison.OrdinalIgnoreCase))
                        multipleColumns = false;
                    else
                        throw new ArgumentException("multipleColumns must be TRUE or FALSE");
                }
                else
                    throw new ArgumentException("multipleColumns must be TRUE or FALSE");

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
                string keyStr = RequireText(key, "key");
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
                if (keys == null)
                    throw new ArgumentException("a range is required");
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
                string keyStr = RequireText(key, "key");
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
            return RedisUdfAsync.Run("RedisUDFDel", optionalHost, new object[] { key, optionalHost }, () => RedisUDFDelCore(key, optionalHost));
        }

        [ExcelFunction(Description = "Deletes a Redis key; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFDelNonVolatile(
            [ExcelArgument(Description = "Redis key to delete")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFDel(key, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFDelCore(object key, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).KeyDelete(keyStr, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFDel: key={keyStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
        public static object RedisUDFSetEx(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Time to live in seconds")] object ttlSeconds,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUdfAsync.Run("RedisUDFSetEx", optionalHost, new object[] { key, value, ttlSeconds, optionalHost }, () => RedisUDFSetExCore(key, value, ttlSeconds, optionalHost));
        }

        [ExcelFunction(Description = "Sets the value of a Redis key with an expiry in seconds; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetExNonVolatile(
            [ExcelArgument(Description = "Redis key to set the value for")] object key,
            [ExcelArgument(Description = "Value to set for the given key")] object value,
            [ExcelArgument(Description = "Time to live in seconds")] object ttlSeconds,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetEx(key, value, ttlSeconds, optionalHost);
        }

        private static string RedisUDFSetExCore(object key, object value, object ttlSeconds, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long ttl = ToInt64Invariant(ttlSeconds);
                if (ttl <= 0)
                    throw new ArgumentException("ttl must be a positive number of seconds");
                if (ttl > MaxTtlSeconds)
                    throw new ArgumentException("ttl is out of range");
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value);
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                // A null RedisValue would issue DEL instead of storing an empty string.
                GetDb(host).StringSet(keyStr, valueStr ?? "", TimeSpan.FromSeconds(ttl), When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFSetEx: key={keyStr}, value={value}, ttl={ttl}s, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
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
            return RedisUdfAsync.Run("RedisUDFExpire", optionalHost, new object[] { key, ttlSeconds, optionalHost }, () => RedisUDFExpireCore(key, ttlSeconds, optionalHost));
        }

        [ExcelFunction(Description = "Sets a timeout on a Redis key in seconds; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFExpireNonVolatile(
            [ExcelArgument(Description = "Redis key to set the expiry for")] object key,
            [ExcelArgument(Description = "Time to live in seconds")] object ttlSeconds,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFExpire(key, ttlSeconds, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFExpireCore(object key, object ttlSeconds, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long ttl = ToInt64Invariant(ttlSeconds);
                if (ttl <= 0)
                    return "Error: ttl must be a positive number of seconds";
                if (ttl > MaxTtlSeconds)
                    return "Error: ttl is out of range";
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).KeyExpire(keyStr, TimeSpan.FromSeconds(ttl), CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFExpire: key={keyStr}, ttl={ttl}s, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFIncr", optionalHost, new object[] { key, optionalHost }, () => RedisUDFIncrCore(key, optionalHost));
        }

        [ExcelFunction(Description = "Increments the integer value of a Redis key by one; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFIncrNonVolatile(
            [ExcelArgument(Description = "Redis key to increment")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFIncr(key, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFIncrCore(object key, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).StringIncrement(keyStr, 1, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFIncr: key={keyStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFIncrBy", optionalHost, new object[] { key, increment, optionalHost }, () => RedisUDFIncrByCore(key, increment, optionalHost));
        }

        [ExcelFunction(Description = "Increments the integer value of a Redis key by a given amount; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFIncrByNonVolatile(
            [ExcelArgument(Description = "Redis key to increment")] object key,
            [ExcelArgument(Description = "Amount to increment by")] object increment,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFIncrBy(key, increment, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFIncrByCore(object key, object increment, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                long incr = ToInt64Invariant(increment);
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).StringIncrement(keyStr, incr, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFIncrBy: key={keyStr}, increment={incr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
                if (keys == null)
                    throw new ArgumentException("a range is required");
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
                    // Same representation as RedisUDFTTL: fractional seconds,
                    // -1 when the key is missing or has no expiry.
                    result[i, 1] = ttl.HasValue ? ToRedisString(ttl.Value.TotalSeconds) : "-1";
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
            return RedisUdfAsync.Run("RedisUDFHashSet", optionalHost, new object[] { hashKey, field, value, optionalHost }, () => RedisUDFHashSetCore(hashKey, field, value, optionalHost));
        }

        [ExcelFunction(Description = "Sets a field in a Redis hash; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFHashSetNonVolatile(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Field name to set within the hash")] object field,
            [ExcelArgument(Description = "Value to set for the given field")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFHashSet(hashKey, field, value, optionalHost);
        }

        private static string RedisUDFHashSetCore(object hashKey, object field, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = RequireText(hashKey, "hash key");
                string fieldStr = RequireText(field, "field");
                string valueStr = ToRedisString(value);
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                // A null RedisValue would issue HDEL instead of storing an empty string.
                GetDb(host).HashSet(hashKeyStr, fieldStr, valueStr ?? "", When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSet: {hashKeyStr}[{fieldStr}] = {value}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
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
            return RedisUdfAsync.Run("RedisUDFHashSetMultiple", optionalHost,
                new object[] { hashKey, fieldValuePairs, optionalHost },
                () => RedisUDFHashSetMultipleCore(hashKey, fieldValuePairs, optionalHost));
        }

        [ExcelFunction(Description = "Sets multiple fields in a Redis hash; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFHashSetMultipleNonVolatile(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "2D range with field-value pairs (2 columns: field, value)")] object[,] fieldValuePairs,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFHashSetMultiple(hashKey, fieldValuePairs, optionalHost);
        }

        private static string RedisUDFHashSetMultipleCore(object hashKey, object[,] fieldValuePairs, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                if (fieldValuePairs == null)
                    throw new ArgumentException("a range is required");
                var pairs = FlattenPairRange(fieldValuePairs);
                var entries = CollectHashEntries(pairs);
                if (entries.Count == 0)
                    throw new ArgumentException("no entries to write");
                string hashKeyStr = RequireText(hashKey, "hash key");
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                GetDb(host).HashSet(hashKeyStr, entries.ToArray(),
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFHashSetMultiple: {hashKeyStr}, fields={entries.Count}, host={host}, fireAndForget={fireAndForget}");
                return fireAndForget ? FireAndForgetMarker(replyDependent: false) : "OK";
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
                string hashKeyStr = RequireText(hashKey, "hash key");
                string fieldStr = RequireText(field, "field");
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
                string hashKeyStr = RequireText(hashKey, "hash key");
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
                if (hashKeys == null)
                    throw new ArgumentException("a range is required");
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
                string fieldStr = RequireText(field, "field");
                var batch = GetDb(host).CreateBatch();
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
            return RedisUdfAsync.Run("RedisUDFHashDel", optionalHost, new object[] { hashKey, field, optionalHost }, () => RedisUDFHashDelCore(hashKey, field, optionalHost));
        }

        [ExcelFunction(Description = "Deletes a field from a Redis hash; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFHashDelNonVolatile(
            [ExcelArgument(Description = "Redis hash key")] object hashKey,
            [ExcelArgument(Description = "Field name to delete from the hash")] object field,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFHashDel(hashKey, field, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFHashDelCore(object hashKey, object field, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string hashKeyStr = RequireText(hashKey, "hash key");
                string fieldStr = RequireText(field, "field");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).HashDelete(hashKeyStr, fieldStr, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFHashDel: {hashKeyStr}[{fieldStr}], host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFListPushRight", optionalHost, new object[] { key, value, optionalHost }, () => RedisUDFListPushRightCore(key, value, optionalHost));
        }

        [ExcelFunction(Description = "Pushes a value onto the right end of a Redis list; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFListPushRightNonVolatile(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Value to push onto the list")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFListPushRight(key, value, optionalHost);
        }

        private static object RedisUDFListPushRightCore(object key, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value) ?? "";
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                long length = GetDb(host).ListRightPush(keyStr, valueStr, When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPushRight: key={keyStr}, value={value}, length={length}, host={host}, fireAndForget={fireAndForget}");
                if (fireAndForget)
                    return FireAndForgetMarker(replyDependent: false);
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
            return RedisUdfAsync.Run("RedisUDFListPushLeft", optionalHost, new object[] { key, value, optionalHost }, () => RedisUDFListPushLeftCore(key, value, optionalHost));
        }

        [ExcelFunction(Description = "Pushes a value onto the left end of a Redis list; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFListPushLeftNonVolatile(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Value to push onto the list")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFListPushLeft(key, value, optionalHost);
        }

        private static object RedisUDFListPushLeftCore(object key, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value) ?? "";
                bool fireAndForget = ShouldFireAndForget(replyDependent: false);
                long length = GetDb(host).ListLeftPush(keyStr, valueStr, When.Always,
                    fireAndForget ? CommandFlags.FireAndForget : CommandFlags.None);
                if (logger.IsTraceEnabled)
                    logger.Trace($"RedisUDFListPushLeft: key={keyStr}, value={value}, length={length}, host={host}, fireAndForget={fireAndForget}");
                if (fireAndForget)
                    return FireAndForgetMarker(replyDependent: false);
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
                string keyStr = RequireText(key, "key");
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
            return RedisUdfAsync.Run("RedisUDFListPopRight", optionalHost, new object[] { key, optionalHost }, () => RedisUDFListPopRightCore(key, optionalHost));
        }

        [ExcelFunction(Description = "Removes and returns the last element of a Redis list; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFListPopRightNonVolatile(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFListPopRight(key, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFListPopRightCore(object key, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).ListRightPop(keyStr, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFListPopRight: key={keyStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFListPopLeft", optionalHost, new object[] { key, optionalHost }, () => RedisUDFListPopLeftCore(key, optionalHost));
        }

        [ExcelFunction(Description = "Removes and returns the first element of a Redis list; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFListPopLeftNonVolatile(
            [ExcelArgument(Description = "Redis list key")] object key,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFListPopLeft(key, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFListPopLeftCore(object key, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).ListLeftPop(keyStr, CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFListPopLeft: key={keyStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFSetAdd", optionalHost, new object[] { key, value, optionalHost }, () => RedisUDFSetAddCore(key, value, optionalHost));
        }

        [ExcelFunction(Description = "Adds a member to a Redis set; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetAddNonVolatile(
            [ExcelArgument(Description = "Redis set key")] object key,
            [ExcelArgument(Description = "Member to add to the set")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetAdd(key, value, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFSetAddCore(object key, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value);
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).SetAdd(keyStr, valueStr ?? "", CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFSetAdd: key={keyStr}, value={valueStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
            return RedisUdfAsync.Run("RedisUDFSetRemove", optionalHost, new object[] { key, value, optionalHost }, () => RedisUDFSetRemoveCore(key, value, optionalHost));
        }

        [ExcelFunction(Description = "Removes a member from a Redis set; runs once per entry/argument change (non-volatile)")]
        public static object RedisUDFSetRemoveNonVolatile(
            [ExcelArgument(Description = "Redis set key")] object key,
            [ExcelArgument(Description = "Member to remove from the set")] object value,
            [ExcelArgument(Description = "Optional Redis host (e.g., host:port)")] object optionalHost
        )
        {
            return RedisUDFSetRemove(key, value, optionalHost);
        }

        // Reply-dependent: fire and forget only in "fireforget-all".
        private static object RedisUDFSetRemoveCore(object key, object value, object optionalHost)
        {
            string host = null;
            try
            {
                host = ResolveHost(optionalHost);
                string keyStr = RequireText(key, "key");
                string valueStr = ToRedisString(value);
                if (ShouldFireAndForget(replyDependent: true))
                {
                    GetDb(host).SetRemove(keyStr, valueStr ?? "", CommandFlags.FireAndForget);
                    if (logger.IsTraceEnabled)
                        logger.Trace($"RedisUDFSetRemove: key={keyStr}, value={valueStr}, host={host}, fireAndForget=true");
                    return FireAndForgetMarker(replyDependent: true);
                }
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
                string keyStr = RequireText(key, "key");
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

    /// <summary>
    /// Bounded least-recently-used cache used by PublishIfChanged to remember the
    /// last delivered payload per host/channel without growing without limit.
    /// Entries are evicted from the cold end when the configured capacity
    /// (PublishDedupCacheSize) is exceeded. Each operation is serialized
    /// internally, so different keys (which may use different striped locks in
    /// RedisUDF) can never corrupt the shared recency list; the per-key
    /// check -> publish -> store sequence is still protected by RedisUDF's
    /// striped lock.
    /// </summary>
    internal sealed class PublishDedupCache
    {
        private sealed class Entry
        {
            public readonly string Key;
            public string Value;

            public Entry(string key, string value)
            {
                Key = key;
                Value = value;
            }
        }

        private readonly int _capacity;
        private readonly Dictionary<string, LinkedListNode<Entry>> _entries;
        private readonly LinkedList<Entry> _recent = new LinkedList<Entry>();
        private readonly object _sync = new object();

        public PublishDedupCache(int capacity)
        {
            if (capacity <= 0)
                throw new ArgumentOutOfRangeException(nameof(capacity), "capacity must be positive");
            _capacity = capacity;
            // The cap can be configured very large; only use a small initial
            // hint so a typo cannot preallocate a huge dictionary at add-in start.
            _entries = new Dictionary<string, LinkedListNode<Entry>>(Math.Min(capacity, 1024), StringComparer.Ordinal);
        }

        public int Count
        {
            get
            {
                lock (_sync)
                    return _entries.Count;
            }
        }

        /// <summary>Snapshot of the tracked keys, taken under the internal lock
        /// and returned as a plain array so callers can walk it without holding
        /// that lock: the listener-join cleanup takes per-key stripes next, and
        /// PublishIfChanged takes them in the opposite order.</summary>
        public string[] SnapshotKeys()
        {
            lock (_sync)
            {
                var keys = new string[_entries.Count];
                _entries.Keys.CopyTo(keys, 0);
                return keys;
            }
        }

        /// <summary>Looks up the value and marks the entry as most recently used.</summary>
        public bool TryGet(string key, out string value)
        {
            lock (_sync)
            {
                if (!_entries.TryGetValue(key, out var node))
                {
                    value = null;
                    return false;
                }
                _recent.Remove(node);
                _recent.AddFirst(node);
                value = node.Value.Value;
                return true;
            }
        }

        /// <summary>Adds or updates the value, evicting the least recently used
        /// entry when the cache is over capacity.</summary>
        public void Set(string key, string value)
        {
            lock (_sync)
            {
                if (_entries.TryGetValue(key, out var existing))
                {
                    existing.Value.Value = value;
                    _recent.Remove(existing);
                    _recent.AddFirst(existing);
                    return;
                }
                var entry = new Entry(key, value);
                var node = _recent.AddFirst(entry);
                _entries.Add(key, node);
                if (_entries.Count > _capacity)
                {
                    var oldest = _recent.Last;
                    _recent.RemoveLast();
                    _entries.Remove(oldest.Value.Key);
                }
            }
        }

        /// <summary>Removes the entry if present; returns whether it existed.</summary>
        public bool Remove(string key)
        {
            lock (_sync)
            {
                if (!_entries.TryGetValue(key, out var node))
                    return false;
                _entries.Remove(key);
                _recent.Remove(node);
                return true;
            }
        }
    }
}

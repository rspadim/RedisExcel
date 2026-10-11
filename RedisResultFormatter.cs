using Newtonsoft.Json;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RedisExcel
{
    /// <summary>Formatting helpers for values sent to Excel.</summary>
    internal static class RedisResultFormatter
    {
        /// <summary>
        /// HGETALL as a valid JSON object: {"field":"value",...}. A missing hash
        /// (null/empty entries) returns "{}": a Redis hash is never empty, so the
        /// empty object is unambiguous and always valid JSON. The "(no value)"
        /// sentinel is reserved for GET/HGET, which have no JSON representation.
        /// </summary>
        internal static string FormatHash(HashEntry[] entries)
        {
            if (entries == null || entries.Length == 0)
                return "{}";
            return "{" + string.Join(",", entries.Select(e =>
                $"{JsonConvert.ToString(e.Name.ToString())}:{JsonConvert.ToString(e.Value.ToString())}")) + "}";
        }
        /// <summary>
        /// TRUE when two payloads are byte-for-byte identical, for the
        /// SkipRepeatedMessages dedup. Compares the RAW BYTES, not string equality:
        ///   * RedisValue's own == compares numeric-LOOKING strings numerically
        ///     ("007" == "7"), which would swallow genuinely different values and
        ///     leave the cell stale;
        ///   * ToString() decodes UTF-8, so two DIFFERENT invalid byte sequences
        ///     ("\x80" vs "\x81") both decode to U+FFFD and would compare equal,
        ///     again leaving the cell stale on a real change.
        /// A byte comparison is correct in both cases and matches the write side
        /// (PublishIfChanged compares ordinals of the same bytes).
        /// </summary>
        internal static bool SameMessage(RedisValue a, RedisValue b)
        {
            if (a.IsNull || b.IsNull)
                return a.IsNull && b.IsNull;
            // Compare the raw bytes with the framework's optimized span compare
            // (SequenceEqual): the cast returns SE.Redis's backing array with zero
            // allocation for whole-array values (SUB/GET/HGETALL), and the framework
            // uses word-at-a-time comparison (10-30x faster than a manual byte loop
            // on large payloads).
            var sa = new ReadOnlySpan<byte>((byte[])a);
            var sb = new ReadOnlySpan<byte>((byte[])b);
            return sa.SequenceEqual(sb);
        }

        /// <summary>
        /// TRUE when two HGETALL results have the same fields and values.
        /// Order-insensitive and O(n): "b" is indexed by field NAME as text
        /// (so the distinct Redis fields "1" and "01" never collide), then every
        /// entry of "a" is looked up and its value compared ORDINALLY (so a value
        /// like "007" vs "7" is a real change, not a formatting variant).
        /// The comparison assumes unique field names, as guaranteed by a Redis
        /// hash: a duplicated name collapses to a single dictionary entry.
        /// </summary>
        internal static bool HashEquals(HashEntry[] a, HashEntry[] b)
        {
            if (ReferenceEquals(a, b))
                return true;
            if (a == null || b == null || a.Length != b.Length)
                return false;
            var valuesByName = new Dictionary<string, RedisValue>(b.Length, System.StringComparer.Ordinal);
            for (int i = 0; i < b.Length; i++)
                valuesByName[b[i].Name.ToString()] = b[i].Value;
            for (int i = 0; i < a.Length; i++)
            {
                RedisValue value;
                if (!valuesByName.TryGetValue(a[i].Name.ToString(), out value) || !SameMessage(value, a[i].Value))
                    return false;
            }
            return true;
        }
    }
}

using Newtonsoft.Json;
using StackExchange.Redis;
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
        /// TRUE when two HGETALL results have the same fields and values.
        /// Order-insensitive and O(n): "b" is indexed by field NAME as text
        /// (so the distinct Redis fields "1" and "01" never collide), then every
        /// entry of "a" is looked up and its value compared with RedisValue
        /// equality (numeric formatting differences such as "1" and "1.00" are
        /// still the same value and must not trigger an update). The comparison
        /// assumes unique field names, as guaranteed by a Redis hash: a
        /// duplicated name collapses to a single dictionary entry.
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
                if (!valuesByName.TryGetValue(a[i].Name.ToString(), out value) || value != a[i].Value)
                    return false;
            }
            return true;
        }
    }
}

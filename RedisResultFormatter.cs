using Newtonsoft.Json;
using StackExchange.Redis;
using System.Linq;

namespace RedisExcel
{
    /// <summary>Formatting helpers for values sent to Excel.</summary>
    internal static class RedisResultFormatter
    {
        /// <summary>HGETALL as a valid JSON object: {"field":"value",...}.</summary>
        internal static string FormatHash(HashEntry[] entries)
        {
            if (entries == null || entries.Length == 0)
                return "(no value)";
            return "{" + string.Join(",", entries.Select(e =>
                $"{JsonConvert.ToString(e.Name.ToString())}:{JsonConvert.ToString(e.Value.ToString())}")) + "}";
        }
        /// <summary>
        /// TRUE when two HGETALL results have the same fields and values.
        /// The comparison assumes unique field names, as guaranteed by a Redis hash:
        /// the scan is order-insensitive, so duplicated field names would not be
        /// detected (multiplicity is ignored).
        /// </summary>
        internal static bool HashEquals(HashEntry[] a, HashEntry[] b)
        {
            if (ReferenceEquals(a, b))
                return true;
            if (a == null || b == null || a.Length != b.Length)
                return false;
            // Order-insensitive: Redis may return the same hash with fields in a
            // different order after a rehash, which is not a value change. Hash
            // field names are unique, so for each entry in "a" we require the same
            // Name/Value pair to exist anywhere in "b". The O(n^2) scan avoids
            // allocating a set/dictionary on this hot polling path (hashes are small).
            for (int i = 0; i < a.Length; i++)
            {
                bool found = false;
                for (int j = 0; j < b.Length; j++)
                {
                    if (a[i].Name == b[j].Name && a[i].Value == b[j].Value)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    return false;
            }
            return true;
        }
    }
}

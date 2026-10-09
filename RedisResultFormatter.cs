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
        /// <summary>TRUE when two HGETALL results have the same fields and values.</summary>
        internal static bool HashEquals(HashEntry[] a, HashEntry[] b)
        {
            if (ReferenceEquals(a, b))
                return true;
            if (a == null || b == null || a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i].Name != b[i].Name || a[i].Value != b[i].Value)
                    return false;
            }
            return true;
        }
    }
}

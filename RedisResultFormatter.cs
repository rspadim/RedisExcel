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
    }
}

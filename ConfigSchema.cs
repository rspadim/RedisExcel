using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace RedisExcel
{
    /// <summary>A single problem found while validating RedisExcel.json.</summary>
    internal sealed class ConfigIssue
    {
        /// <summary>JSON pointer-ish path of the offending value ("" = the root).</summary>
        public string Path { get; }
        public string Message { get; }

        public ConfigIssue(string path, string message)
        {
            Path = path ?? "";
            Message = message ?? "";
        }

        public override string ToString() => (Path.Length == 0 ? "/" : Path) + ": " + Message;
    }

    /// <summary>
    /// Validates RedisExcel.json against an embedded JSON Schema and produces a
    /// cleaned copy with every invalid value removed, so one wrong value can no
    /// longer discard the whole file (the tolerant parser previously failed at
    /// the JSON reader level for bool/object/array in a numeric position, which
    /// the Error callback cannot repair, and the file collapsed to defaults).
    ///
    /// The subset of JSON Schema keywords this schema uses is supported: type,
    /// properties, additionalProperties (bool or schema), enum, minimum, maximum.
    /// Comments (// and /* */) are accepted. Keys are matched case-insensitively,
    /// mirroring Newtonsoft's default property binding.
    /// </summary>
    internal static class ConfigSchema
    {
        /// <summary>The schema for RedisExcel.json. Kept as a literal so it ships
        /// inside the XLL (no external file) and is the single source of truth for
        /// both validation and the "clean then deserialize" load path.</summary>
        internal const string SchemaJson = @"{
  ""type"": ""object"",
  ""additionalProperties"": false,
  ""properties"": {
    ""RTD"": {
      ""type"": ""object"",
      ""additionalProperties"": false,
      ""properties"": {
        ""host"": { ""type"": ""string"" },
        ""timeout"": { ""type"": ""integer"", ""minimum"": 1 },
        ""RedisUpdateRateMs"": { ""type"": ""integer"", ""minimum"": 1 },
        ""ExcelUpdateRateMs"": { ""type"": ""integer"", ""minimum"": 1 },
        ""ExcelUpdateStyle"": { ""type"": ""string"", ""enum"": [ ""Timer"", ""Realtime"", ""Automatic"" ] },
        ""MessageCounterThreshold"": { ""type"": ""integer"" },
        ""UseGetMultiple"": { ""type"": ""boolean"" }
      }
    },
    ""UDF"": {
      ""type"": ""object"",
      ""additionalProperties"": false,
      ""properties"": {
        ""host"": { ""type"": ""string"" },
        ""timeout"": { ""type"": ""integer"", ""minimum"": 1 }
      }
    },
    ""Servers"": { ""type"": ""object"", ""additionalProperties"": { ""type"": ""string"" } },
    ""UpdateCheck"": { ""type"": ""boolean"" },
    ""SkipRepeatedMessages"": { ""type"": ""boolean"" },
    ""CoalesceRealtimeUpdates"": { ""type"": ""boolean"" },
    ""ConflationMs"": { ""type"": ""integer"", ""minimum"": 0, ""maximum"": 3600000 },
    ""PublishDedupCacheSize"": { ""type"": ""integer"", ""minimum"": 1 },
    ""SyncWrite"": { ""type"": ""string"", ""enum"": [ ""sync"", ""fireforget"", ""fireforget-all"" ] },
    ""AsyncWrites"": { ""type"": ""boolean"" }
  }
}";

        private static readonly JObject Schema = (JObject)JsonConvert.DeserializeObject(SchemaJson);

        /// <summary>Parses JSON accepting // and /* */ comments (Newtonsoft discards
        /// them when loading into a JToken). Returns null when the text is not a JSON
        /// object; the caller treats that as an unusable file. Throws on malformed
        /// JSON (the caller's catch handles it).</summary>
        internal static JObject ParseWithComments(string json)
        {
            using (var reader = new JsonTextReader(new StringReader(json)))
            {
                reader.DateParseHandling = DateParseHandling.None;
                var token = JToken.ReadFrom(reader);
                return token as JObject;
            }
        }

        /// <summary>Returns a copy of <paramref name="root"/> with every schema
        /// violation removed, appending one <see cref="ConfigIssue"/> per removal.
        /// Deserializing the result cannot fail on a wrong-typed value because the
        /// reader never sees it. Returns null when the root itself is invalid.</summary>
        internal static JToken Clean(JToken root, IList<ConfigIssue> issues)
        {
            return Prune(root, Schema, "", issues);
        }

        private static JToken Prune(JToken value, JObject schema, string path, IList<ConfigIssue> issues)
        {
            string type = (string)schema["type"];
            if (!MatchesType(value, type))
            {
                issues.Add(new ConfigIssue(path, "expected " + type + ", found " + Describe(value)));
                return null;
            }

            switch (type)
            {
                case "integer":
                {
                    long n = value.Value<long>();
                    var min = schema["minimum"];
                    var max = schema["maximum"];
                    if (min != null && n < min.Value<long>())
                    {
                        issues.Add(new ConfigIssue(path, "must be >= " + min.Value<long>() + " (found " + n + ")"));
                        return null;
                    }
                    if (max != null && n > max.Value<long>())
                    {
                        issues.Add(new ConfigIssue(path, "must be <= " + max.Value<long>() + " (found " + n + ")"));
                        return null;
                    }
                    return value;
                }
                case "string":
                {
                    var allowed = schema["enum"] as JArray;
                    if (allowed != null)
                    {
                        string s = value.Value<string>();
                        foreach (var e in allowed)
                        {
                            // Case-insensitive: Newtonsoft binds enum strings and the
                            // SyncWrite modes case-insensitively.
                            if (string.Equals((string)e, s, StringComparison.OrdinalIgnoreCase))
                                return value;
                        }
                        issues.Add(new ConfigIssue(path, "must be one of [" + Join(allowed) + "], found '" + s + "'"));
                        return null;
                    }
                    return value;
                }
                case "boolean":
                    return value;
                case "object":
                {
                    var obj = (JObject)value;
                    var props = schema["properties"] as JObject;
                    var additional = schema["additionalProperties"];
                    var cleaned = new JObject();
                    foreach (var prop in obj.Properties())
                    {
                        string childPath = path + "/" + prop.Name;
                        var childSchema = FindProperty(props, prop.Name);
                        if (childSchema != null)
                        {
                            var kept = Prune(prop.Value, childSchema, childPath, issues);
                            if (kept != null) cleaned[prop.Name] = kept;
                        }
                        else if (additional is JObject additionalSchema)
                        {
                            var kept = Prune(prop.Value, additionalSchema, childPath, issues);
                            if (kept != null) cleaned[prop.Name] = kept;
                        }
                        else if (additional != null && additional.Type == JTokenType.Boolean && !(bool)additional)
                        {
                            issues.Add(new ConfigIssue(childPath, "unknown key"));
                        }
                        else
                        {
                            cleaned[prop.Name] = prop.Value;
                        }
                    }
                    return cleaned;
                }
                default:
                    return value;
            }
        }

        private static JObject FindProperty(JObject properties, string name)
        {
            if (properties == null)
                return null;
            if (properties.TryGetValue(name, out var exact) && exact is JObject exactObj)
                return exactObj;
            foreach (var prop in properties.Properties())
            {
                if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase) && prop.Value is JObject obj)
                    return obj;
            }
            return null;
        }

        private static bool MatchesType(JToken value, string type)
        {
            switch (type)
            {
                case "object": return value != null && value.Type == JTokenType.Object;
                case "string": return value != null && value.Type == JTokenType.String;
                case "integer": return value != null && value.Type == JTokenType.Integer;
                case "boolean": return value != null && value.Type == JTokenType.Boolean;
                default: return true;
            }
        }

        private static string Describe(JToken value) =>
            value == null ? "null" : value.Type.ToString().ToLowerInvariant();

        private static string Join(JArray array)
        {
            var parts = new List<string>();
            foreach (var token in array)
                parts.Add(token.ToString());
            return string.Join(", ", parts);
        }
    }
}

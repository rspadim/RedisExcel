using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NJsonSchema;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace RedisExcel
{
    /// <summary>A single problem found while validating RedisExcel.json.</summary>
    internal sealed class ConfigIssue
    {
        /// <summary>JSON path of the offending value (e.g. "#/RTD.timeout").</summary>
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
    /// Validates RedisExcel.json against an embedded JSON Schema using NJsonSchema,
    /// and produces a cleaned copy with every invalid value removed, so one wrong
    /// value can no longer discard the whole file. The previous tolerant parser
    /// failed at the JSON READER level for a bool/object/array in a numeric
    /// position (the Error callback cannot repair that), which collapsed the file
    /// to defaults - repointing the host to localhost and changing the write mode.
    ///
    /// The schema lives in RedisExcel.schema.json, embedded in the assembly as a
    /// resource (so it ships inside the XLL) and exposed verbatim to Excel through
    /// RedisJsonSchemaContents(). Comments (// and /* */) are accepted on load.
    /// </summary>
    internal static class ConfigSchema
    {
        private const string ResourceSuffix = "RedisExcel.schema.json";

        private static readonly Lazy<string> _schemaText = new Lazy<string>(ReadSchemaText);
        private static readonly Lazy<JsonSchema> _schema =
            new Lazy<JsonSchema>(() => JsonSchema.FromJsonAsync(_schemaText.Value).GetAwaiter().GetResult());

        /// <summary>The embedded schema document, verbatim (for the Excel function
        /// and for tooling).</summary>
        internal static string Text => _schemaText.Value;

        private static string ReadSchemaText()
        {
            var assembly = typeof(ConfigSchema).Assembly;
            foreach (var name in assembly.GetManifestResourceNames())
            {
                if (name.EndsWith(ResourceSuffix, StringComparison.Ordinal))
                {
                    using (var stream = assembly.GetManifestResourceStream(name))
                    using (var reader = new StreamReader(stream))
                        return reader.ReadToEnd();
                }
            }
            // Fallback: a minimal object schema. Never throw at load - the config
            // path must keep working even if the resource is somehow missing.
            return "{\"type\":\"object\"}";
        }

        /// <summary>Parses JSON accepting // and /* */ comments (Newtonsoft discards
        /// them when loading into a JToken). Returns null when the text is not a
        /// JSON object; the caller treats that as an unusable file. Throws on
        /// malformed JSON (the caller's catch handles it).</summary>
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
            var errors = _schema.Value.Validate(root);
            if (errors.Count == 0)
                return root;

            var perPath = new List<KeyValuePair<string, string>>();
            foreach (var error in errors)
            {
                perPath.Add(new KeyValuePair<string, string>(NormalizePath(error.Path), Describe(error)));
            }

            // Remove the offending values from the deepest path first, so that
            // dropping a child does not shift a parent that also has an error.
            var cleaned = root.DeepClone();
            perPath.Sort((a, b) => PathDepth(b.Key).CompareTo(PathDepth(a.Key)));
            foreach (var pair in perPath)
            {
                RemovePath(cleaned as JObject, pair.Key);
                issues?.Add(new ConfigIssue(pair.Key, pair.Value));
            }
            return cleaned;
        }

        /// <summary>NJsonSchema paths look like "#/RTD.timeout" or "#/Extra"; turn
        /// them into a slash path the remover understands ("RTD/timeout").</summary>
        private static string NormalizePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return "";
            string p = path.StartsWith("#", StringComparison.Ordinal) ? path.Substring(1) : path;
            p = p.Replace('.', '/');
            while (p.StartsWith("/", StringComparison.Ordinal))
                p = p.Substring(1);
            return p;
        }

        private static int PathDepth(string path) =>
            string.IsNullOrEmpty(path) ? 0 : path.Split('/').Length;

        private static void RemovePath(JObject root, string path)
        {
            if (root == null || string.IsNullOrEmpty(path))
                return;
            var parts = path.Split('/');
            JToken current = root;
            for (int i = 0; i < parts.Length - 1; i++)
            {
                current = FindChild(current, parts[i]);
                if (current == null)
                    return;
            }
            var last = FindProperty(current as JObject, parts[parts.Length - 1]);
            if (last != null)
                last.Remove();
        }

        private static JObject FindChild(JToken token, string name)
        {
            var obj = token as JObject;
            var prop = FindProperty(obj, name);
            return prop?.Value as JObject;
        }

        private static JProperty FindProperty(JObject obj, string name)
        {
            if (obj == null)
                return null;
            foreach (var prop in obj.Properties())
            {
                if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                    return prop;
            }
            return null;
        }

        private static string Describe(NJsonSchema.Validation.ValidationError error)
        {
            switch (error.Kind)
            {
                case NJsonSchema.Validation.ValidationErrorKind.IntegerExpected: return "expected an integer";
                case NJsonSchema.Validation.ValidationErrorKind.NumberExpected: return "expected a number";
                case NJsonSchema.Validation.ValidationErrorKind.StringExpected: return "expected a string";
                case NJsonSchema.Validation.ValidationErrorKind.BooleanExpected: return "expected true or false";
                case NJsonSchema.Validation.ValidationErrorKind.ObjectExpected: return "expected an object";
                case NJsonSchema.Validation.ValidationErrorKind.ArrayExpected: return "expected an array";
                case NJsonSchema.Validation.ValidationErrorKind.NoAdditionalPropertiesAllowed: return "unknown key";
                case NJsonSchema.Validation.ValidationErrorKind.NumberTooSmall: return "value is below the minimum";
                case NJsonSchema.Validation.ValidationErrorKind.NumberTooBig: return "value is above the maximum";
                case NJsonSchema.Validation.ValidationErrorKind.NotInEnumeration: return "not one of the allowed values";
                case NJsonSchema.Validation.ValidationErrorKind.PatternMismatch: return "not one of the allowed values";
                default: return error.Kind.ToString();
            }
        }
    }
}

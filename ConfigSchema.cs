using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NJsonSchema;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    {
                        if (stream == null)
                            break; // fall through to the restrictive fallback
                        using (var reader = new StreamReader(stream))
                            return reader.ReadToEnd();
                    }
                }
            }
            // The resource is embedded and a parity test + a build check guard it, so
            // this path should never run. If it does (a broken build), fall back to a
            // RESTRICTIVE schema (no properties, additionalProperties:false) so every
            // key is dropped and the config degrades to defaults - never to a
            // permissive schema that would let a wrong-typed value reach the binder
            // and discard the whole file (the bug this class exists to prevent).
            return "{\"type\":\"object\",\"additionalProperties\":false}";
        }

        /// <summary>TRUE when the embedded schema document was found (not the
        /// restrictive fallback). Exposed so a test can assert the resource ships.</summary>
        internal static bool HasEmbeddedSchema =>
            _schemaText.Value.IndexOf("RedisUpdateRateMs", StringComparison.Ordinal) >= 0;

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
        /// reader never sees it. Returns null when the root itself is invalid.
        ///
        /// The pruning WALKS the schema (not the error path strings): NJsonSchema's
        /// paths use a notation ("Servers['prod.eu']") that a naive split corrupts,
        /// and its property lookup is case-sensitive while the object's is not, so
        /// pruning by path could remove the wrong key. Walking the document together
        /// with its schema keeps the exact JSON property names.</summary>
        internal static JToken Clean(JToken root, IList<ConfigIssue> issues)
        {
            return Prune(root, _schema.Value, "", issues);
        }

        private static JToken Prune(JToken value, JsonSchema schema, string path, IList<ConfigIssue> issues)
        {
            // Objects are walked property by property (so an unknown or bad child
            // is dropped without touching its siblings). Every other value is
            // checked as a whole: a violation drops just that value.
            if (schema.Type.HasFlag(JsonObjectType.Object) && value is JObject obj)
            {
                var cleaned = new JObject();
                // Known model properties, keyed by their canonical schema name. Two
                // keys that differ only by case bind to the same field
                // case-insensitively; keep the exact-case spelling (else the first)
                // so a bad "rtd" cannot shadow a good "RTD" - in EITHER order.
                var modelProps = new Dictionary<string, JProperty>(StringComparer.OrdinalIgnoreCase);
                var otherProps = new List<JProperty>();
                foreach (var prop in obj.Properties())
                {
                    var childSchema = FindProperty(schema, prop.Name, out string canonical);
                    if (childSchema == null)
                    {
                        otherProps.Add(prop);
                        continue;
                    }
                    if (modelProps.TryGetValue(canonical, out var existing))
                    {
                        issues?.Add(new ConfigIssue(path + "/" + prop.Name, "duplicate key (differs only by case)"));
                        bool currentExact = string.Equals(prop.Name, canonical, StringComparison.Ordinal);
                        bool existingExact = string.Equals(existing.Name, canonical, StringComparison.Ordinal);
                        if (currentExact && !existingExact)
                            modelProps[canonical] = prop;
                        // else keep the existing one
                    }
                    else
                    {
                        modelProps[canonical] = prop;
                    }
                }

                foreach (var kv in modelProps)
                {
                    var prop = kv.Value;
                    var childSchema = FindProperty(schema, prop.Name, out _);
                    var kept = Prune(prop.Value, childSchema, path + "/" + prop.Name, issues);
                    if (kept != null)
                        cleaned[prop.Name] = kept;
                }
                foreach (var prop in otherProps)
                {
                    string childPath = path + "/" + prop.Name;
                    // Free-form keys (Servers): the backing dictionary is
                    // case-SENSITIVE and the resolver supports both spellings, so
                    // they are never deduped here.
                    if (schema.AdditionalPropertiesSchema != null)
                    {
                        var kept = Prune(prop.Value, schema.AdditionalPropertiesSchema, childPath, issues);
                        if (kept != null)
                            cleaned[prop.Name] = kept;
                    }
                    else if (schema.AllowAdditionalProperties)
                    {
                        cleaned[prop.Name] = prop.Value;
                    }
                    else
                    {
                        issues?.Add(new ConfigIssue(childPath, "unknown key"));
                    }
                }
                return cleaned;
            }

            // A value the schema cannot even inspect (e.g. a JSON integer larger
            // than Int64 makes NJsonSchema throw) is treated as invalid, so it is
            // dropped instead of reaching ToObject and discarding the whole file.
            ICollection<NJsonSchema.Validation.ValidationError> errors;
            try
            {
                errors = schema.Validate(value);
            }
            catch (Exception ex)
            {
                issues?.Add(new ConfigIssue(path, "invalid value (" + ex.GetType().Name + ")"));
                return null;
            }
            if (errors.Count > 0)
            {
                var first = errors.First();
                issues?.Add(new ConfigIssue(path, Describe(first)));
                return null;
            }
            return value;
        }

        /// <summary>Finds the schema for a property name, case-insensitively, and
        /// reports the CANONICAL schema spelling so callers can prefer the exact
        /// case when two keys differ only by case.</summary>
        private static JsonSchema FindProperty(JsonSchema schema, string name, out string canonical)
        {
            canonical = null;
            if (schema.Properties.TryGetValue(name, out var exact))
            {
                canonical = name;
                return exact;
            }
            foreach (var kv in schema.Properties)
            {
                if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase))
                {
                    canonical = kv.Key;
                    return kv.Value;
                }
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

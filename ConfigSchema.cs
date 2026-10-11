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
                // Known model properties, grouped by their CANONICAL schema name:
                // two keys that differ only by case bind to the same field
                // case-insensitively. Candidates are pruned in order and the FIRST
                // that survives is kept (the exact-case spelling wins when its
                // value is valid), so a bad exact-case duplicate cannot shadow a
                // good variant and a good value is never lost.
                var candidates = new Dictionary<string, List<JProperty>>(StringComparer.OrdinalIgnoreCase);
                var order = new List<string>();
                var otherProps = new List<JProperty>();
                foreach (var prop in obj.Properties())
                {
                    var childSchema = FindProperty(schema, prop.Name, out string canonical);
                    if (childSchema == null)
                    {
                        otherProps.Add(prop);
                        continue;
                    }
                    if (!candidates.TryGetValue(canonical, out var list))
                    {
                        list = new List<JProperty>();
                        candidates[canonical] = list;
                        order.Add(canonical);
                    }
                    list.Add(prop);
                }

                foreach (var canonical in order)
                {
                    var list = candidates[canonical];
                    // Prefer the exact-case spelling first, then the document order.
                    list.Sort((a, b) =>
                    {
                        bool ax = string.Equals(a.Name, canonical, StringComparison.Ordinal);
                        bool bx = string.Equals(b.Name, canonical, StringComparison.Ordinal);
                        if (ax != bx) return ax ? -1 : 1;
                        return 0; // stable: keeps document order otherwise
                    });
                    var childSchema = FindProperty(schema, list[0].Name, out _);
                    JToken merged = null;
                    JProperty keeper = list[0];
                    foreach (var candidate in list)
                    {
                        if (candidate != list[0])
                            issues?.Add(new ConfigIssue(path + "/" + candidate.Name, "duplicate key (differs only by case)"));
                        var pruned = Prune(candidate.Value, childSchema, path + "/" + candidate.Name, issues);
                        if (pruned == null)
                            continue;
                        if (merged == null)
                        {
                            merged = pruned;
                            keeper = candidate;
                        }
                        else if (merged is JObject first && pruned is JObject second)
                        {
                            // Both spellings bind to the same object: MERGE so neither
                            // loses a valid field (a bad field in one is already gone).
                            foreach (var prop in second.Properties())
                            {
                                if (first.Property(prop.Name) == null)
                                    first[prop.Name] = prop.Value;
                            }
                        }
                        // Scalars: the first surviving value wins.
                    }
                    if (merged != null)
                        cleaned[keeper.Name] = merged;
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

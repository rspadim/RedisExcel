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
    /// <summary>Severity of a <see cref="ConfigIssue"/>. An <see cref="Error"/>
    /// rejects the whole file at load (a wrong value on a known key); a
    /// <see cref="Warning"/> is logged and ignored (e.g. an unknown/extra key).</summary>
    internal enum ConfigSeverity
    {
        Warning = 0,
        Error = 1,
    }

    /// <summary>A single problem found while validating RedisExcel.json.</summary>
    internal sealed class ConfigIssue
    {
        /// <summary>JSON path of the offending value (e.g. "#/RTD.timeout").</summary>
        public string Path { get; }
        public string Message { get; }
        public ConfigSeverity Severity { get; }

        public ConfigIssue(string path, string message, ConfigSeverity severity = ConfigSeverity.Error)
        {
            Path = path ?? "";
            Message = message ?? "";
            Severity = severity;
        }

        public override string ToString()
        {
            // The path embeds raw JSON KEY names (e.g. a Servers key that IS a
            // connection string with password=...): mask credentials and strip
            // control characters (a key name may contain \n, which would forge a
            // log line and split the #CONFIG cell). Every derived surface (cell,
            // NLog, RedisCheckConfigFile) inherits this.
            string path = Path.Length == 0 ? "/" : Path;
            return Sanitize(path) + ": " + Sanitize(Message);
        }

        /// <summary>Masks credentials and removes characters that can break a line
        /// (so one issue can never span/forge multiple log lines or split a cell).
        /// Covers C0/C1 controls plus the Unicode line/paragraph separators and NEL,
        /// which `char.IsControl` does not flag.</summary>
        private static string Sanitize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text ?? "";
            string masked = AppConfig.MaskHost(text);
            if (masked.IndexOfAny(ControlChars) < 0)
                return masked;
            var sb = new System.Text.StringBuilder(masked.Length);
            foreach (char c in masked)
                sb.Append(char.IsControl(c) || IsLineBreaker(c) ? ' ' : c);
            return sb.ToString();
        }

        private static bool IsLineBreaker(char c) =>
            c == '\u0085' || c == '\u2028' || c == '\u2029' || c == '\u009f';

        private static readonly char[] ControlChars =
            Enumerable.Range(0, 32).Select(i => (char)i)
                .Concat(new[] { (char)127, '\u0085', '\u2028', '\u2029', '\u009f' })
                .ToArray();
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
                // Skip leading whitespace/comments BEFORE the root object: a file
                // that starts with `// note` or `/* note */` is documented as valid
                // (comments are allowed anywhere), but JToken.ReadFrom would return
                // the comment token (a non-object) and the whole config would be
                // silently discarded. Consume comments until the first real token.
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.Comment)
                        continue;
                    break;
                }
                if (reader.TokenType == JsonToken.None
                    || reader.TokenType == JsonToken.StartArray
                    || reader.TokenType == JsonToken.Integer
                    || reader.TokenType == JsonToken.Float
                    || reader.TokenType == JsonToken.String
                    || reader.TokenType == JsonToken.Boolean
                    || reader.TokenType == JsonToken.Null)
                    return null; // not an object
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
            // Validate under the INVARIANT culture: the schema patterns are declared
            // case-insensitive with (?i), and .NET's (?i) folds case using the
            // CURRENT culture. Under tr-TR the capital I folds to the dotless "ı",
            // so `"TIMER"` no longer matches `Timer`/`REALTIME`/`FIREFORGET-ALL` and
            // the loader would silently drop a valid value (reverting to defaults).
            // Pinning the culture here is what makes the schema's own
            // "case-insensitive" promise hold on every machine (audit round O).
            var saved = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture =
                    System.Globalization.CultureInfo.InvariantCulture;
                // The removed v1 keys must be ERRORS, not tolerated extra keys:
                // otherwise a pre-1.6 file loads on defaults (host -> localhost)
                // while the validator says "OK". Report them before/independently of
                // the schema walk (which would only warn about an unknown key).
                ReportLegacyKeys(root, issues);
                ReportDanglingDefault(root, issues);
                return Prune(root, _schema.Value, "", issues);
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = saved;
            }
        }

        /// <summary>The property with the given name, matched case-insensitively
        /// (exact case first) - the SAME way the schema walker and the Newtonsoft
        /// binder match. The pre-checks below must not use an ordinal lookup: the
        /// walker merges case-variant keys, so an ordinal lookup both rejected a
        /// valid config (a lowercase "servers") and skipped the lenient guards on a
        /// lowercase "default"/"rtd" (audit D3), and it missed a case-variant
        /// legacy member like "RTD.Host" that the validator flags (audit F3).</summary>
        private static JProperty FindProperty(JObject obj, string name)
        {
            var exact = obj.Property(name);
            if (exact != null)
                return exact;
            return obj.Properties().FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Every property whose name matches <paramref name="name"/>
        /// case-insensitively (exact case first). The pre-checks below must inspect
        /// ALL of them: a document can carry two case-variant sections
        /// (`rtd` and `RTD`), and checking only the first let a pre-1.6 key hide in
        /// the other (audit H1/D1).</summary>
        private static IEnumerable<JProperty> FindProperties(JObject obj, string name)
        {
            var exact = obj.Property(name);
            if (exact != null)
                yield return exact;
            foreach (var p in obj.Properties())
            {
                if (p == exact)
                    continue;
                if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))
                    yield return p;
            }
        }

        /// <summary>Flags the pre-1.6 keys that the unified shape removed, as
        /// ERRORs with the v1.6 replacement, so an old file fails closed instead of
        /// silently reverting to defaults.</summary>
        internal static void ReportLegacyKeys(JToken root, IList<ConfigIssue> issues)
        {
            if (issues == null || !(root is JObject obj))
                return;
            void Legacy(string section, string key, string replacement)
            {
                // Every case-variant spelling of the section may carry the key.
                foreach (var sectionProp in FindProperties(obj, section))
                    if (sectionProp.Value is JObject sec && FindProperty(sec, key) != null)
                        issues.Add(new ConfigIssue("/" + section + "/" + key,
                            "removed in v1.6; use " + replacement, ConfigSeverity.Error));
            }
            Legacy("RTD", "host", "\"Default\" (or a \"Servers\" name)");
            Legacy("RTD", "timeout", "the server's \"connectTimeout\"/\"syncTimeout\" (or the global \"Timeout\")");
            // The whole UDF section was removed (it only held host/timeout).
            if (FindProperties(obj, "UDF").Any())
                issues.Add(new ConfigIssue("/UDF",
                    "removed in v1.6; declare servers under \"Servers\" and pick one with \"Default\"", ConfigSeverity.Error));
            // The root-level boolean moved into RTD (any casing).
            if (FindProperties(obj, "CoalesceRealtimeUpdates").Any())
                issues.Add(new ConfigIssue("/CoalesceRealtimeUpdates",
                    "moved into \"RTD\" in v1.6", ConfigSeverity.Error));
        }

        /// <summary>Flags a <c>Default</c> whose value is a bare word that is
        /// neither a defined <c>Servers</c> name nor a recognizable connection
        /// string. It is a WARNING, not an error: a single-label host
        /// ("localhost", a Docker/K8s service name) is a perfectly valid
        /// port-less connection string, so rejecting it would lock a valid
        /// config out (audit F1). The warning still surfaces the likely typo,
        /// and the value is used as-is (never a silent localhost fallback).</summary>
        internal static void ReportDanglingDefault(JToken root, IList<ConfigIssue> issues)
        {
            if (issues == null || !(root is JObject obj))
                return;
            var defaultToken = FindProperty(obj, "Default")?.Value;
            if (defaultToken == null || defaultToken.Type != JTokenType.String)
                return;
            string name = (string)defaultToken;
            if (string.IsNullOrWhiteSpace(name))
                return;
            // A raw connection string is allowed ("host", "host:port", "host,opt=..").
            if (LooksLikeConnectionString(name))
                return;
            // A name may be defined in ANY case-variant "Servers" map (the walker
            // merges them), so scan them all before calling it dangling (audit
            // H1/D3).
            foreach (var serversProp in FindProperties(obj, "Servers"))
                if (serversProp.Value is JObject servers && FindProperty(servers, name) != null)
                    return;
            issues.Add(new ConfigIssue("/Default",
                "'" + name + "' is not defined in \"Servers\" and is not a connection string; using it as a host",
                ConfigSeverity.Warning));
        }

        /// <summary>Returns a copy of <paramref name="obj"/> with each property
        /// renamed to the canonical spelling declared by the oneOf OBJECT branch
        /// (case-insensitively). Keys the schema does not know keep their spelling
        /// so the strict branch still fails on a genuine unknown key. When two keys
        /// differ only by case, the EXACT-canonical spelling wins (agreeing with the
        /// rest of the loader's dedup, audit H1/C2) and a warning is recorded. Lets
        /// the schema's case-sensitive `required`/`additionalProperties` agree with
        /// the case-insensitive binder (audit G1).</summary>
        private static JObject NormalizeObjectKeyCase(JObject obj, JsonSchema oneOfSchema, IList<ConfigIssue> issues, string path)
        {
            var objectBranch = oneOfSchema?.OneOf?.FirstOrDefault(b => b.Type.HasFlag(JsonObjectType.Object));
            if (objectBranch?.Properties == null)
                return obj;
            // Group by canonical target so a case-only duplicate is resolved the
            // same way everywhere: the exact spelling wins, else the first seen.
            var chosen = new Dictionary<string, JProperty>(StringComparer.Ordinal);
            var order = new List<string>();
            foreach (var prop in obj.Properties())
            {
                var canonical = objectBranch.Properties
                    .FirstOrDefault(p => string.Equals(p.Key, prop.Name, StringComparison.OrdinalIgnoreCase)).Key ?? prop.Name;
                if (!chosen.TryGetValue(canonical, out var keeper))
                {
                    chosen[canonical] = prop;
                    order.Add(canonical);
                    continue;
                }
                // A duplicate that differs only by case: the discarded value must
                // still be VALIDATED (a wrong value on a known key rejects the whole
                // file). Record the duplicate, and if the discarded value is itself
                // invalid, fail closed rather than silently dropping it (audit J2;
                // the earlier version kept the valid spelling and validated only it,
                // so {"host":"ok","HOST":123} was wrongly accepted).
                issues?.Add(new ConfigIssue(path + "/" + prop.Name,
                    "duplicate key (differs only by case)", ConfigSeverity.Warning));
                if (!ValueMatchesProperty(prop.Value, objectBranch.Properties[canonical]))
                    issues?.Add(new ConfigIssue(path + "/" + prop.Name,
                        "duplicate key has an invalid value for '" + canonical + "'", ConfigSeverity.Error));
                bool keepNew = string.Equals(prop.Name, canonical, StringComparison.Ordinal)
                    && !string.Equals(keeper.Name, canonical, StringComparison.Ordinal);
                if (keepNew)
                    chosen[canonical] = prop;
            }
            var copy = new JObject();
            foreach (var canonical in order)
                copy[canonical] = chosen[canonical].Value;
            return copy;
        }

        /// <summary>True when a discarded duplicate's value would itself satisfy
        /// the schema property (so an invalid value cannot hide behind a valid
        /// case-variant spelling).</summary>
        private static bool ValueMatchesProperty(JToken value, JsonSchema propertySchema)
        {
            try
            {
                return propertySchema.Validate(value).Count == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>True when a value looks like a Redis connection string (or a
        /// bare host / host:port) rather than a Servers name.</summary>
        private static bool LooksLikeConnectionString(string value)
        {
            string s = value.Trim();
            if (s.Length == 0)
                return false;
            // A well-known single-label host is a valid port-less endpoint (avoid a
            // noisy warning for the very common Default: "localhost").
            if (s.Equals("localhost", StringComparison.OrdinalIgnoreCase))
                return true;
            // A host:port, an IPv4/IPv6 host, a dotted hostname, or any option list.
            if (s.IndexOf(',') >= 0 || s.IndexOf('=') >= 0 || s.IndexOf('/') >= 0)
                return true; // connection string with options, a path, or redis://
            if (s.IndexOf(':') >= 0)
                return true; // host:port (an alias may not contain ':')
            if (s.IndexOf('.') >= 0)
                return true; // dotted hostname (an alias may not contain '.')
            return false;    // a bare word: warn (it may be a Servers name or a service host)
        }

        private static JToken Prune(JToken value, JsonSchema schema, string path, IList<ConfigIssue> issues)
        {
            // Objects are walked property by property (so an unknown or bad child
            // is handled without touching its siblings). Every other value is
            // checked as a whole.
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
                            issues?.Add(new ConfigIssue(path + "/" + candidate.Name, "duplicate key (differs only by case)", ConfigSeverity.Warning));
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
                        // An extra/unknown key is a WARNING, not an error: we do not
                        // use it, and it must not reject an otherwise working file.
                        issues?.Add(new ConfigIssue(childPath, "unknown key (ignored)", ConfigSeverity.Warning));
                    }
                }
                return cleaned;
            }

            // A schema that is a oneOf (a Servers value: string or object): keep the
            // value when any branch validates; otherwise it is an error.
            if (schema.OneOf != null && schema.OneOf.Count > 0)
            {
                // Property names are matched case-insensitively everywhere else in
                // the loader (and by the Newtonsoft binder), but the schema's own
                // `required`/`additionalProperties` are case-SENSITIVE - so a
                // Servers object spelled `Host`/`HOST`/`ConnectTimeout` would fail
                // here while every other part accepts it (audit G1). Normalize the
                // object's keys to the canonical spelling before validating.
                JToken probe = value;
                if (value is JObject valueObj)
                    probe = NormalizeObjectKeyCase(valueObj, schema, issues, path);
                bool any;
                try
                {
                    any = schema.OneOf.Any(branch => branch.Validate(probe).Count == 0);
                }
                catch (Exception ex)
                {
                    issues?.Add(new ConfigIssue(path, "invalid value (" + ex.GetType().Name + ")"));
                    return null;
                }
                if (any)
                    return probe;
                issues?.Add(new ConfigIssue(path, "expected a connection string or an object with host/connectTimeout/syncTimeout"));
                return null;
            }

            // A value the schema cannot even inspect (e.g. a JSON integer larger
            // than Int64 makes NJsonSchema throw) is treated as invalid.
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
                // An unknown-key report from `additionalProperties:false` is a
                // warning; every other kind is a value error (fail-closed).
                var severity = first.Kind == NJsonSchema.Validation.ValidationErrorKind.NoAdditionalPropertiesAllowed
                    ? ConfigSeverity.Warning
                    : ConfigSeverity.Error;
                issues?.Add(new ConfigIssue(path, Describe(first), severity));
                return severity == ConfigSeverity.Error ? null : value;
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

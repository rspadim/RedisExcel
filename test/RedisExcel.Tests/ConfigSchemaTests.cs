using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Dedicated suite for the embedded RedisExcel.json schema (v1.6 shape). It
    /// does NOT test the loader end to end; it pins the schema itself:
    ///   * every model field has a schema entry of the matching type;
    ///   * value limits match the C# type;
    ///   * a wrong VALUE on a known key is an ERROR (fail-closed at load), an
    ///     EXTRA/unknown key is only a WARNING;
    ///   * Servers values accept a string or an object;
    ///   * a document that passes the schema ALWAYS binds to ConfigRoot.
    /// </summary>
    [Collection("write-mode-seam")]
    public class ConfigSchemaTests
    {
        // ---- helpers -----------------------------------------------------------

        private static IList<ConfigIssue> Clean(string json)
            => CleanOut(json, out _);

        private static IList<ConfigIssue> CleanOut(string json, out ConfigRoot bound)
        {
            var issues = new List<ConfigIssue>();
            var root = ConfigSchema.ParseWithComments(json);
            var cleaned = ConfigSchema.Clean(root, issues);
            bound = null;
            if (cleaned != null)
                bound = cleaned.ToObject<ConfigRoot>(); // must never throw
            return issues;
        }

        // Accepted = no ERROR (a warning like an extra key does not reject the file).
        private static bool Accepted(string json) => Clean(json).All(i => i.Severity != ConfigSeverity.Error);
        private static bool Rejected(string json) => Clean(json).Any(i => i.Severity == ConfigSeverity.Error);
        private static bool Warned(string json) => Clean(json).Any(i => i.Severity == ConfigSeverity.Warning);

        // ---- the schema document itself ---------------------------------------

        [Fact]
        public void Schema_ResourceIsEmbeddedAndParses()
        {
            string text = ConfigSchema.Text;
            Assert.False(string.IsNullOrWhiteSpace(text));
            Assert.True(ConfigSchema.HasEmbeddedSchema,
                "the embedded RedisExcel.schema.json was not found in the assembly");
            var schema = JObject.Parse(text);
            Assert.Equal("object", (string)schema["type"]);
            Assert.False((bool)schema["additionalProperties"]);
        }

        [Fact]
        public void Schema_DeclaresTheV16Shape()
        {
            var schema = JObject.Parse(ConfigSchema.Text);
            var props = (JObject)schema["properties"];
            foreach (var rootField in new[] { "Default", "Timeout", "Servers", "RTD", "UpdateCheck", "SkipRepeatedMessages", "ConflationMs", "PublishDedupCacheSize", "SyncWrite", "AsyncWrites" })
                Assert.True(props.Property(rootField) != null, "schema missing root key " + rootField);

            // The old sections host/timeout are gone; CoalesceRealtimeUpdates moved into RTD.
            Assert.Null(props.Property("UDF"));
            var rtd = (JObject)((JObject)props["RTD"])["properties"];
            Assert.True(rtd.Property("CoalesceRealtimeUpdates") != null);
            Assert.Null(rtd.Property("host"));
            Assert.Null(rtd.Property("timeout"));

            var timeout = (JObject)((JObject)props["Timeout"])["properties"];
            Assert.True(timeout.Property("connectTimeout") != null);
            Assert.True(timeout.Property("syncTimeout") != null);
        }

        // ---- ServerDef: string or object --------------------------------------

        [Fact]
        public void Servers_AcceptStringOrObjectWithAnyKeyShape()
        {
            Assert.True(Accepted("{\"Servers\":{\"prod\":\"h:6379\"}}"));
            Assert.True(Accepted("{\"Servers\":{\"prod.eu\":\"h:6379\"}}"));   // dotted alias
            Assert.True(Accepted("{\"Servers\":{\"a/b\":\"h\"}}"));
            Assert.True(Accepted("{\"Servers\":{}}"));                          // empty map
            Assert.True(Accepted("{\"Servers\":{\"p\":{\"host\":\"h\",\"connectTimeout\":8000,\"syncTimeout\":9000}}}"));
            Assert.True(Accepted("{\"Servers\":{\"p\":{\"host\":\"h\"}}}"));    // object, no timeouts
        }

        [Fact]
        public void Servers_ObjectRequiresHost()
        {
            // An object without "host" is not a valid server.
            Assert.True(Rejected("{\"Servers\":{\"p\":{\"connectTimeout\":8000}}}"));
        }

        [Fact]
        public void Servers_ObjectRejectsUnknownField()
        {
            Assert.True(Rejected("{\"Servers\":{\"p\":{\"host\":\"h\",\"nope\":1}}}"));
        }

        [Fact]
        public void Servers_ObjectCaseOnlyDuplicate_InvalidVariantStillRejects()
        {
            // Audit J2: a case-only duplicate in a Servers object must not let an
            // INVALID value hide behind a valid spelling (fail-closed).
            Assert.True(Rejected("{\"Servers\":{\"s\":{\"host\":\"localhost:6379\",\"HOST\":123}},\"Default\":\"s\"}"));
            Assert.True(Rejected("{\"Servers\":{\"s\":{\"host\":123,\"HOST\":\"localhost:6379\"}},\"Default\":\"s\"}"));
            // Two VALID spellings collapse to the exact-canonical one (warning only).
            Assert.False(Rejected("{\"Servers\":{\"s\":{\"host\":\"a:1\",\"HOST\":\"b:2\"}},\"Default\":\"s\"}"));
        }

        [Fact]
        public void Servers_ObjectKeyCaseIsInsensitive()
        {
            // The loader (and the Newtonsoft binder) match property names
            // case-insensitively everywhere else, so a Servers object spelled
            // Host/HOST/ConnectTimeout must be accepted, not rejected (audit G1).
            Assert.False(Rejected("{\"Servers\":{\"a\":{\"HOST\":\"x\"}}}"));
            Assert.False(Rejected("{\"Servers\":{\"a\":{\"Host\":\"x\"}}}"));
            Assert.False(Rejected("{\"Servers\":{\"a\":{\"host\":\"x\",\"ConnectTimeout\":5}}}"));
            Assert.False(Rejected("{\"Servers\":{\"a\":{\"Host\":\"x\",\"SyncTimeout\":5}}}"));
            // A genuine unknown field must still fail (only the KNOWN names are
            // case-folded).
            Assert.True(Rejected("{\"Servers\":{\"p\":{\"host\":\"h\",\"nope\":1}}}"));
        }

        [Fact]
        public void Servers_NonStringNonObjectValue_IsRejected()
        {
            Assert.True(Rejected("{\"Servers\":{\"bad\":123}}"));
            Assert.True(Rejected("{\"Servers\":{\"bad\":[1]}}"));
        }

        [Fact]
        public void Servers_StringAndObjectBothBind()
        {
            CleanOut("{\"Servers\":{\"a\":\"h:1\",\"b\":{\"host\":\"h:2\",\"syncTimeout\":9000}}}", out var bound);
            Assert.Equal("h:1", bound.Servers["a"].host);
            Assert.Equal("h:2", bound.Servers["b"].host);
            Assert.Equal(9000, bound.Servers["b"].syncTimeout);
        }

        [Fact]
        public void Default_IsAString()
        {
            Assert.True(Accepted("{\"Default\":\"prod\",\"Servers\":{\"prod\":\"h:6379\"}}"));
            Assert.True(Accepted("{\"Default\":\"localhost:6379\"}"));
            Assert.True(Rejected("{\"Default\":42}"));
        }

        // ---- Timeout (global) --------------------------------------------------

        [Theory]
        [InlineData("connectTimeout")]
        [InlineData("syncTimeout")]
        public void Timeout_AcceptsPositiveIntRejectsZeroOrOverflow(string key)
        {
            Assert.True(Accepted("{\"Timeout\":{\"" + key + "\":1}}"));
            Assert.True(Accepted("{\"Timeout\":{\"" + key + "\":" + int.MaxValue + "}}"));
            Assert.True(Rejected("{\"Timeout\":{\"" + key + "\":0}}"));
            Assert.True(Rejected("{\"Timeout\":{\"" + key + "\":-1}}"));
            Assert.True(Rejected("{\"Timeout\":{\"" + key + "\":" + ((long)int.MaxValue + 1) + "}}"));
        }

        [Fact]
        public void Timeout_UnknownFieldIsOnlyAWarning()
        {
            // Consistent with D4: an extra/unknown key is tolerated (warning), not
            // a value error.
            Assert.False(Rejected("{\"Timeout\":{\"nope\":1}}"));
            Assert.True(Warned("{\"Timeout\":{\"nope\":1}}"));
        }

        // ---- integer limits match the C# type ----------------------------------

        [Theory]
        [InlineData("RTD", "RedisUpdateRateMs")]
        [InlineData("RTD", "ExcelUpdateRateMs")]
        public void IntFields_AcceptAtOrAboveOneRejectZeroAndOverflow(string section, string key)
        {
            Assert.True(Accepted(IntJson(section, key, 1)), key + " 1");
            Assert.True(Accepted(IntJson(section, key, int.MaxValue)), key + " int.MaxValue");
            Assert.True(Rejected(IntJson(section, key, 0)), key + " 0");
            Assert.True(Rejected(IntJson(section, key, -1)), key + " -1");
            Assert.True(Rejected(IntJson(section, key, (long)int.MaxValue + 1)), key + " 2^31");
        }

        [Theory]
        [InlineData(1)]
        [InlineData(5000)]
        [InlineData(int.MaxValue)]
        public void PublishDedupCacheSize_AcceptsPositiveInt(int value)
        {
            Assert.True(Accepted("{\"PublishDedupCacheSize\":" + value + "}"));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(2147483648L)]
        public void PublishDedupCacheSize_RejectsNonPositiveOrOverflow(long value)
        {
            Assert.True(Rejected("{\"PublishDedupCacheSize\":" + value + "}"));
        }

        [Fact]
        public void ConflationMs_AcceptsAnyInt32AndSanitizeClamps()
        {
            Assert.True(Accepted("{\"ConflationMs\":0}"));
            Assert.True(Accepted("{\"ConflationMs\":200}"));
            Assert.True(Accepted("{\"ConflationMs\":-1}"));
            Assert.True(Accepted("{\"ConflationMs\":2147483647}"));
            Assert.True(Rejected("{\"ConflationMs\":99999999999}"));
        }

        [Fact]
        public void MessageCounterThreshold_AcceptsLongRejectsBeyond()
        {
            Assert.True(Accepted("{\"RTD\":{\"MessageCounterThreshold\":0}}"));
            Assert.True(Accepted("{\"RTD\":{\"MessageCounterThreshold\":" + long.MaxValue + "}}"));
            Assert.True(Rejected("{\"RTD\":{\"MessageCounterThreshold\":99999999999999999999}}"));
        }

        private static string IntJson(string section, string key, long value)
            => "{\"" + section + "\":{\"" + key + "\":" + value + "}}";

        // ---- boolean fields ----------------------------------------------------

        [Theory]
        [InlineData("UpdateCheck")]
        [InlineData("SkipRepeatedMessages")]
        [InlineData("AsyncWrites")]
        public void BooleanFields_AcceptTrueFalseRejectOthers(string key)
        {
            Assert.True(Accepted("{\"" + key + "\":true}"));
            Assert.True(Accepted("{\"" + key + "\":false}"));
            Assert.True(Rejected("{\"" + key + "\":\"yes\"}"));
            Assert.True(Rejected("{\"" + key + "\":1}"));
        }

        [Theory]
        [InlineData("UseGetMultiple")]
        [InlineData("CoalesceRealtimeUpdates")]
        public void RtdBooleanFields_AcceptTrueFalseRejectOthers(string key)
        {
            Assert.True(Accepted("{\"RTD\":{\"" + key + "\":true}}"));
            Assert.True(Rejected("{\"RTD\":{\"" + key + "\":\"x\"}}"));
        }

        // ---- enums / patterns --------------------------------------------------

        [Theory]
        [InlineData("SYNC")]
        [InlineData("FIREFORGET")]
        [InlineData("FIREFORGET-ALL")]
        [InlineData("FireForget-All")]
        public void SyncWrite_AcceptSpellingsCaseInsensitiveUnderTurkishCulture(string value)
        {
            using (CultureScope("tr-TR"))
                Assert.True(Accepted("{\"SyncWrite\":\"" + value + "\"}"));
        }

        [Theory]
        [InlineData("bogus")]
        [InlineData("fire")]
        [InlineData("")]
        [InlineData("syncc")]
        public void SyncWrite_RejectsUnknown(string value)
        {
            Assert.True(Rejected("{\"SyncWrite\":\"" + value + "\"}"));
        }

        [Theory]
        [InlineData("Timer")]
        [InlineData("Realtime")]
        [InlineData("AUTOMATIC")]
        [InlineData("TIMER")]
        public void ExcelUpdateStyle_AcceptsDocumentedSpellings(string value)
        {
            Assert.True(Accepted("{\"RTD\":{\"ExcelUpdateStyle\":\"" + value + "\"}}"));
        }

        [Theory]
        [InlineData("Fast")]
        [InlineData("")]
        public void ExcelUpdateStyle_RejectsUnknown(string value)
        {
            Assert.True(Rejected("{\"RTD\":{\"ExcelUpdateStyle\":\"" + value + "\"}}"));
        }

        private static IDisposable CultureScope(string name)
        {
            var culture = CultureInfo.GetCultureInfo(name);
            var saved = Thread.CurrentThread.CurrentCulture;
            var savedDefault = CultureInfo.DefaultThreadCurrentCulture;
            Thread.CurrentThread.CurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentCulture = culture;
            return new RestoreCulture(saved, savedDefault);
        }

        private sealed class RestoreCulture : IDisposable
        {
            private readonly CultureInfo _saved;
            private readonly CultureInfo _savedDefault;
            public RestoreCulture(CultureInfo saved, CultureInfo savedDefault) { _saved = saved; _savedDefault = savedDefault; }
            public void Dispose()
            {
                Thread.CurrentThread.CurrentCulture = _saved;
                CultureInfo.DefaultThreadCurrentCulture = _savedDefault;
            }
        }

        // ---- shape / container mismatches (value errors) -----------------------

        [Theory]
        [InlineData("{\"RTD\":[]}")]
        [InlineData("{\"RTD\":5}")]
        [InlineData("{\"ConflationMs\":[1,2]}")]
        [InlineData("{\"ConflationMs\":{}}")]
        [InlineData("{\"AsyncWrites\":[]}")]
        [InlineData("{\"Timeout\":[]}")]
        public void WrongContainerShapes_AreRejected(string json)
        {
            Assert.True(Rejected(json));
        }

        // ---- extra keys are warnings, not errors -------------------------------

        [Theory]
        [InlineData("{\"UDF\":{\"host\":\"x\"}}", "/UDF")]                    // the removed section
        [InlineData("{\"RTD\":{\"host\":\"x\"}}", "/RTD/host")]               // removed
        [InlineData("{\"RTD\":{\"timeout\":1000}}", "/RTD/timeout")]         // removed
        [InlineData("{\"CoalesceRealtimeUpdates\":true}", "/CoalesceRealtimeUpdates")] // moved into RTD
        // Case-variant section headers must be caught too: the walker/binder are
        // case-insensitive, so a lowercase pre-1.6 file must NOT slip through on
        // defaults (audit D3 A3).
        [InlineData("{\"rtd\":{\"host\":\"myhost:6379\"}}", "/RTD/host")]
        [InlineData("{\"udf\":{\"host\":\"myhost:6379\"}}", "/UDF")]
        // A legacy key hidden in a case-variant SIBLING section must still be an
        // error (audit H1/D1): {"rtd":{host}, "RTD":{...}}.
        [InlineData("{\"rtd\":{\"host\":\"h:6379\"},\"RTD\":{\"RedisUpdateRateMs\":500}}", "/RTD/host")]
        [InlineData("{\"RTD\":{\"RedisUpdateRateMs\":500},\"rtd\":{\"host\":\"h:6379\"}}", "/RTD/host")]
        [InlineData("{\"rtd\":{\"timeout\":1000},\"RTD\":{\"RedisUpdateRateMs\":500}}", "/RTD/timeout")]
        [InlineData("{\"coalescerealtimeupdates\":true}", "/CoalesceRealtimeUpdates")]
        public void LegacyKeys_AreErrorsNamedInTheReason(string json, string path)
        {
            var issues = Clean(json);
            Assert.Contains(issues, i => i.Severity == ConfigSeverity.Error && i.Path == path);
        }

        [Fact]
        public void DanglingDefault_FindsANameInACaseVariantServersMap()
        {
            // "b" is defined in the lowercase "servers" map (the walker merges the
            // maps), so it must NOT be reported dangling (audit H1/D3). A duplicate
            // case-variant key may still warn, but not about /Default.
            var issues = Clean("{\"Default\":\"b\",\"Servers\":{\"a\":\"h1\"},\"servers\":{\"b\":\"h2\"}}");
            Assert.DoesNotContain(issues, i => i.Path == "/Default");
        }

        [Fact]
        public void DanglingDefault_IsCaseInsensitiveAndOnlyAWarning()
        {
            // A lowercase "servers" still defines the server (the walker/binder are
            // case-insensitive), so a valid config must NOT be rejected (D3 A1).
            Assert.False(Rejected("{\"Default\":\"prod\",\"servers\":{\"prod\":\"h:1\"}}"));
            // A bare word that is not a Servers key is a WARNING (it may be a valid
            // port-less host like "localhost"), never a rejection (audit F1), and
            // the guard is case-insensitive like the walker (D3 A2).
            Assert.False(Rejected("{\"default\":\"typo\"}"));
            Assert.True(Warned("{\"default\":\"typo\"}"));
            Assert.True(Warned("{\"DEFAULT\":\"typo\"}"));
            // A real host is accepted with no warning at all.
            Assert.False(Warned("{\"Default\":\"localhost\"}"));
            Assert.False(Warned("{\"Default\":\"localhost:6379\"}"));
        }

        [Theory]
        [InlineData("{\"Unknown\":1}")]
        [InlineData("{\"RTD\":{\"Unknown\":1}}")]
        public void UnknownKeys_AreOnlyWarnings(string json)
        {
            Assert.False(Rejected(json));
            Assert.True(Warned(json));
        }

        // ---- the core guarantee: anything the schema accepts binds -------------

        [Fact]
        public void EveryAcceptedDocumentBindsToTheModel()
        {
            string[] docs =
            {
                "{}",
                "{\"Default\":\"prod\"}",
                "{\"Default\":\"localhost:6379\",\"Timeout\":{\"connectTimeout\":5000,\"syncTimeout\":5000}}",
                "{\"Servers\":{\"a\":\"h1\",\"b\":{\"host\":\"h2\",\"syncTimeout\":9000}}}",
                "{\"RTD\":{\"RedisUpdateRateMs\":1000,\"ExcelUpdateRateMs\":100,\"ExcelUpdateStyle\":\"Automatic\",\"CoalesceRealtimeUpdates\":true}}",
                "{\"RTD\":{\"RedisUpdateRateMs\":{}},\"ConflationMs\":[1],\"SyncWrite\":\"sync\"}",
                "{\"SyncWrite\":\"Bogus\",\"Default\":\"myhost\"}",
                "{\"AsyncWrites\":\"yes\",\"SyncWrite\":\"sync\"}",
            };
            foreach (var doc in docs)
            {
                var issues = CleanOut(doc, out var bound);
                Assert.NotNull(bound);
            }
        }

        [Fact]
        public void ExtraKey_DoesNotRejectButStillBindsTheRest()
        {
            CleanOut("{\"Default\":\"myhost\",\"Nope\":1,\"SyncWrite\":\"sync\"}", out var bound);
            Assert.Equal("myhost", bound.Default);
            Assert.Equal("sync", bound.SyncWrite);
        }

        // ---- the built schema types match the model types (BIDIRECTIONAL) -----

        /// <summary>Full parity both ways, so a config knob cannot silently become
        /// dead config (a model field with no schema entry binds to its default
        /// and is only logged as an unknown key) and the schema cannot grow an
        /// orphan property. Driven by reflection over the CLR model, so a new
        /// field fails the test until the schema declares it.</summary>
        [Fact]
        public void SchemaAndModelAgreeOnEveryProperty_BothDirections()
        {
            var schema = JObject.Parse(ConfigSchema.Text);
            var props = (JObject)schema["properties"];

            AssertSectionParity(props, typeof(ConfigRoot), "root");
            AssertSectionParity((JObject)((JObject)props["RTD"])["properties"], typeof(RTDConfig), "RTD");
            AssertSectionParity((JObject)((JObject)props["Timeout"])["properties"], typeof(TimeoutConfig), "Timeout");

            // A Servers value is a string or an object; the object branch is ServerDef.
            var oneOf = (JArray)((JObject)props["Servers"])["additionalProperties"]["oneOf"];
            var serverObj = (JObject)oneOf.First(t => (string)t["type"] == "object");
            AssertSectionParity((JObject)serverObj["properties"], typeof(ServerDef), "ServerDef");
            // host is the only required field (matches ServerDef.host).
            Assert.Equal(new[] { "host" }, ((JArray)serverObj["required"]).Select(t => (string)t).ToArray());
        }

        private static void AssertSectionParity(JObject sectionProps, Type model, string label)
        {
            var expected = ModelSchemaNames(model);
            var declared = sectionProps.Properties().Select(p => p.Name).ToList();
            foreach (var kv in expected)
            {
                Assert.True(sectionProps.Property(kv.Key) != null,
                    label + " schema is missing the model property '" + kv.Key + "'");
                Assert.Equal(kv.Value, (string)((JObject)sectionProps[kv.Key])["type"]);
            }
            foreach (var name in declared)
                Assert.True(expected.ContainsKey(name),
                    label + " schema declares '" + name + "', which has no model field (orphan)");
        }

        private static Dictionary<string, string> ModelSchemaNames(Type type)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var p in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
            {
                if (p.GetCustomAttributes(typeof(Newtonsoft.Json.JsonIgnoreAttribute), true).Length > 0)
                    continue; // [JsonIgnore]: not part of the file (Rejected/RejectedReason)
                map[p.Name] = SchemaTypeName(p.PropertyType);
            }
            return map;
        }

        private static string SchemaTypeName(Type t)
        {
            if (t == typeof(bool)) return "boolean";
            if (t == typeof(int) || t == typeof(long) || t == typeof(int?) || t == typeof(long?)) return "integer";
            if (t.IsEnum) return "string";
            if (t == typeof(string)) return "string";
            // Any other member is a nested object (RTDConfig, TimeoutConfig,
            // Dictionary<string, ServerDef>), validated by its own section check.
            return "object";
        }

        /// <summary>The embedded schema (the one the add-in actually validates
        /// with) must be byte-identical to the on-disk <c>RedisExcel.schema.json</c>,
        /// so a stale re-embed cannot pass every other test.</summary>
        [Fact]
        public void EmbeddedSchemaBytesMatchTheOnDiskFile()
        {
            string repoRoot = FindRepoRoot();
            string onDisk = System.IO.File.ReadAllText(System.IO.Path.Combine(repoRoot, "RedisExcel.schema.json"));
            Assert.Equal(Normalize(onDisk), Normalize(ConfigSchema.Text));
        }

        private static string Normalize(string s) =>
            s.Replace("\r\n", "\n").Replace("\r", "\n").Trim('\uFEFF', '\n', ' ', '\t');

        private static string FindRepoRoot()
        {
            var dir = new System.IO.DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                if (System.IO.File.Exists(System.IO.Path.Combine(dir.FullName, "RedisExcel.sln")))
                    return dir.FullName;
                dir = dir.Parent;
            }
            throw new InvalidOperationException("RedisExcel.sln not found above " + AppContext.BaseDirectory);
        }
    }
}

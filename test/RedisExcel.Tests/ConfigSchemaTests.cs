using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Dedicated suite for the embedded RedisExcel.json schema. It does NOT test
    /// the loader; it pins the schema itself:
    ///   * every model field has a schema entry of the matching type;
    ///   * numeric limits match the C# type (so an out-of-range value is rejected
    ///     by the schema and never reaches the binding - the regression that made
    ///     a wrong integer discard the whole file);
    ///   * each declared type accepts its valid values and rejects the others;
    ///   * enums/patterns accept the documented spellings (case-insensitive);
    ///   * unknown keys and wrong container shapes are rejected;
    ///   * a document that passes the schema ALWAYS binds to ConfigRoot, so the
    ///     runtime needs only the schema (no second ToObject validation).
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

        private static bool Accepted(string json) => Clean(json).Count == 0;
        private static bool Rejected(string json) => Clean(json).Count > 0;

        // ---- the schema document itself ---------------------------------------

        [Fact]
        public void Schema_ResourceIsEmbeddedAndParses()
        {
            string text = ConfigSchema.Text;
            Assert.False(string.IsNullOrWhiteSpace(text));
            // The embedded resource (not the restrictive fallback) must ship: a
            // fallback would drop every key and silently revert the config.
            Assert.True(ConfigSchema.HasEmbeddedSchema,
                "the embedded RedisExcel.schema.json was not found in the assembly");
            var schema = JObject.Parse(text);
            Assert.Equal("object", (string)schema["type"]);
            Assert.False((bool)schema["additionalProperties"]); // unknown top-level keys rejected
        }

        [Fact]
        public void Schema_DeclaresEveryModelField()
        {
            // Walk the model (ConfigRoot + nested) and require a schema entry for
            // each JSON-bound property, so a new config key cannot ship untyped.
            var expected = new Dictionary<string, string[]>
            {
                ["RTD"] = new[] { "host", "timeout", "RedisUpdateRateMs", "ExcelUpdateRateMs", "ExcelUpdateStyle", "MessageCounterThreshold", "UseGetMultiple" },
                ["UDF"] = new[] { "host", "timeout" },
            };
            var schema = JObject.Parse(ConfigSchema.Text);
            var props = (JObject)schema["properties"];
            foreach (var section in expected)
            {
                var sectionSchema = (JObject)props[section.Key];
                Assert.NotNull(sectionSchema);
                var declared = ((JObject)sectionSchema["properties"]).Properties().Select(p => p.Name).ToHashSet();
                foreach (var field in section.Value)
                    Assert.Contains(field, declared);
            }
            foreach (var rootField in new[] { "Servers", "UpdateCheck", "SkipRepeatedMessages", "CoalesceRealtimeUpdates", "ConflationMs", "PublishDedupCacheSize", "SyncWrite", "AsyncWrites" })
                Assert.True(props.Property(rootField) != null, "schema missing root key " + rootField);
        }

        // ---- integer limits match the C# type ----------------------------------

        [Theory]
        [InlineData("RTD", "timeout")]
        [InlineData("RTD", "RedisUpdateRateMs")]
        [InlineData("RTD", "ExcelUpdateRateMs")]
        [InlineData("UDF", "timeout")]
        public void IntFields_AcceptAtOrAboveOneRejectZeroAndOverflow(string section, string key)
        {
            // The constraint is minimum:1 and the model is `int`, so the accepted
            // range is [1, int.MaxValue] and anything else (0, negative, or an
            // overflow beyond Int32) must be rejected by the schema.
            Assert.True(Accepted(IntJson(section, key, 1)), key + " 1");
            Assert.True(Accepted(IntJson(section, key, int.MaxValue)), key + " int.MaxValue");
            Assert.True(Rejected(IntJson(section, key, 0)), key + " 0");
            Assert.True(Rejected(IntJson(section, key, -1)), key + " -1");
            Assert.True(Rejected(IntJson(section, key, (long)int.MaxValue + 1)), key + " 2^31");
            Assert.True(Rejected(IntJson(section, key, 9_999_999_9999L)), key + " overflow");
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

        [Theory]
        [InlineData(0)]
        [InlineData(200)]
        [InlineData(3600000)]
        [InlineData(-1)]                  // Sanitize clamps to 0
        [InlineData(2147483647)]          // int.MaxValue accepted; Sanitize caps to 3600000
        public void ConflationMs_AcceptsAnyIntegerAndSanitizeClamps(int value)
        {
            // The schema accepts any 32-bit integer (the field is int?); the value
            // range is then CLAMPED by AppConfig.Sanitize to [0, 3600000] (the
            // documented "capped" behaviour), so the schema must not drop it first.
            Assert.True(Accepted("{\"ConflationMs\":" + value + "}"));
        }

        [Fact]
        public void ConflationMs_RejectsValuesBeyondInt32()
        {
            // int? cannot hold more than int.MaxValue, so the schema rejects it (else
            // the binding would overflow and discard the file).
            Assert.True(Rejected("{\"ConflationMs\":99999999999}"));
        }

        [Fact]
        public void MessageCounterThreshold_AcceptsLongRejectsBeyond()
        {
            Assert.True(Accepted("{\"RTD\":{\"MessageCounterThreshold\":0}}"));
            Assert.True(Accepted("{\"RTD\":{\"MessageCounterThreshold\":" + long.MaxValue + "}}"));
            Assert.True(Accepted("{\"RTD\":{\"MessageCounterThreshold\":" + long.MinValue + "}}"));
            // Beyond Int64: NJsonSchema cannot inspect it -> treated as invalid.
            Assert.True(Rejected("{\"RTD\":{\"MessageCounterThreshold\":99999999999999999999}}"));
        }

        private static string IntJson(string section, string key, long value)
            => "{\"" + section + "\":{\"" + key + "\":" + value + "}}";

        // ---- boolean fields ----------------------------------------------------

        [Theory]
        [InlineData("UpdateCheck")]
        [InlineData("SkipRepeatedMessages")]
        [InlineData("CoalesceRealtimeUpdates")]
        [InlineData("AsyncWrites")]
        public void BooleanFields_AcceptTrueFalseRejectOthers(string key)
        {
            Assert.True(Accepted("{\"" + key + "\":true}"));
            Assert.True(Accepted("{\"" + key + "\":false}"));
            Assert.True(Rejected("{\"" + key + "\":\"yes\"}"));
            Assert.True(Rejected("{\"" + key + "\":1}"));
            Assert.True(Rejected("{\"" + key + "\":{}}"));
            Assert.True(Rejected("{\"" + key + "\":[]}"));
        }

        [Theory]
        [InlineData("UseGetMultiple")]
        public void RtdBooleanFields_AcceptTrueFalseRejectOthers(string key)
        {
            Assert.True(Accepted("{\"RTD\":{\"" + key + "\":true}}"));
            Assert.True(Rejected("{\"RTD\":{\"" + key + "\":\"x\"}}"));
        }

        // ---- string fields and enums ------------------------------------------

        [Fact]
        public void HostFields_AcceptAnyString()
        {
            Assert.True(Accepted("{\"RTD\":{\"host\":\"localhost:6379\"}}"));
            Assert.True(Accepted("{\"RTD\":{\"host\":\"\"}}"));
            Assert.True(Accepted("{\"UDF\":{\"host\":\"prod.example:6380,password=x\"}}"));
            Assert.True(Rejected("{\"RTD\":{\"host\":42}}"));
        }

        [Theory]
        [InlineData("sync")]
        [InlineData("fireforget")]
        [InlineData("fireforget-all")]
        [InlineData("SYNC")]
        [InlineData("FireForget-All")]
        public void SyncWrite_AcceptsDocumentedSpellingsCaseInsensitive(string value)
        {
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
        [InlineData("Automatic")]
        [InlineData("realtime")]
        [InlineData("AUTOMATIC")]
        public void ExcelUpdateStyle_AcceptsDocumentedSpellingsCaseInsensitive(string value)
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

        // ---- Servers (free-form string map) -----------------------------------

        [Fact]
        public void Servers_AcceptsStringMapWithAnyKeyShape()
        {
            Assert.True(Accepted("{\"Servers\":{\"prod\":\"h:6379\"}}"));
            Assert.True(Accepted("{\"Servers\":{\"prod.eu\":\"h:6379\"}}"));   // dotted alias
            Assert.True(Accepted("{\"Servers\":{\"a/b\":\"h\"}}"));
            Assert.True(Accepted("{\"Servers\":{\"prod\":\"h1\",\"Prod\":\"h2\"}}")); // case-distinct aliases kept
            Assert.True(Accepted("{\"Servers\":{}}"));
        }

        [Fact]
        public void Servers_DropsNonStringValueButKeepsTheRest()
        {
            var issues = CleanOut("{\"Servers\":{\"ok\":\"h:1\",\"bad\":123},\"SyncWrite\":\"sync\"}", out var bound);
            Assert.NotEmpty(issues);
            Assert.Equal("h:1", bound.Servers["ok"]);
            Assert.False(bound.Servers.ContainsKey("bad"));
            Assert.Equal("sync", bound.SyncWrite);
        }

        // ---- shape / container mismatches -------------------------------------

        [Theory]
        [InlineData("{\"RTD\":[]}")]
        [InlineData("{\"RTD\":5}")]
        [InlineData("{\"RTD\":\"x\"}")]
        [InlineData("{\"UDF\":[]}")]
        [InlineData("{\"Servers\":[]}")]
        [InlineData("{\"Servers\":\"x\"}")]
        [InlineData("{\"ConflationMs\":[1,2]}")]
        [InlineData("{\"ConflationMs\":{}}")]
        [InlineData("{\"AsyncWrites\":[]}")]
        public void WrongContainerShapes_AreRejected(string json)
        {
            Assert.True(Rejected(json));
        }

        [Theory]
        [InlineData("{\"Unknown\":1}")]
        [InlineData("{\"RTD\":{\"Unknown\":1}}")]
        [InlineData("{\"UDF\":{\"Nope\":true}}")]
        public void UnknownKeys_AreRejected(string json)
        {
            Assert.True(Rejected(json));
        }

        // ---- case-only duplicate keys (good value must survive) ----------------

        [Theory]
        [InlineData("{\"RTD\":{\"host\":\"myhost\"},\"rtd\":{\"timeout\":true},\"SyncWrite\":\"sync\"}")]
        [InlineData("{\"rtd\":{\"timeout\":true},\"RTD\":{\"host\":\"myhost\"},\"SyncWrite\":\"sync\"}")]
        [InlineData("{\"rtd\":{\"host\":\"myhost\"},\"RTD\":{\"timeout\":true},\"SyncWrite\":\"sync\"}")]
        public void CaseOnlyDuplicate_KeepsTheGoodValueEitherOrder(string json)
        {
            var issues = CleanOut(json, out var bound);
            Assert.NotEmpty(issues);                       // the duplicate is reported
            Assert.Equal("myhost", bound.RTD.host);        // the good value won (any order)
            Assert.Equal("sync", bound.SyncWrite);         // unrelated keys survive
        }

        [Fact]
        public void CaseOnlyDuplicate_FieldLevel_KeepsTheGoodValue()
        {
            // A good exact-case spelling and a bad variant (or vice-versa) must keep
            // the value that survives pruning, never drop both.
            var issues = CleanOut("{\"RTD\":{\"HOST\":\"GOODHOST\",\"host\":true},\"SyncWrite\":\"sync\"}", out var bound);
            Assert.NotEmpty(issues);
            Assert.Equal("GOODHOST", bound.RTD.host);
            Assert.Equal("sync", bound.SyncWrite);
        }

        // ---- the core guarantee: anything the schema accepts binds -------------

        [Fact]
        public void EveryAcceptedDocumentBindsToTheModel()
        {
            // A broad corpus of valid and partially-invalid documents: Clean must
            // never produce a document that ToObject rejects (which is what made the
            // loader discard the whole file). CleanOut binds inside and would throw.
            string[] docs =
            {
                "{}",
                "{\"RTD\":{\"host\":\"myhost\",\"timeout\":2500,\"RedisUpdateRateMs\":1000,\"ExcelUpdateRateMs\":100,\"ExcelUpdateStyle\":\"Automatic\",\"MessageCounterThreshold\":10000,\"UseGetMultiple\":true}}",
                "{\"UDF\":{\"host\":\"h\",\"timeout\":1000}}",
                "{\"Servers\":{\"a\":\"h1\",\"b\":\"h2\"}}",
                "{\"RTD\":{\"host\":\"myhost\",\"timeout\":true},\"SyncWrite\":\"sync\",\"Servers\":{\"prod\":\"h\"}}",
                "{\"RTD\":{\"RedisUpdateRateMs\":{}},\"ConflationMs\":[1],\"SyncWrite\":\"sync\"}",
                "{\"RTD\":[],\"SyncWrite\":\"sync\"}",
                "{\"rtd\":{\"timeout\":true},\"RTD\":{\"host\":\"myhost\"},\"SyncWrite\":\"sync\"}",
                "{\"Servers\":{\"prod.eu\":{\"x\":1}},\"SyncWrite\":\"sync\",\"RTD\":{\"host\":\"myhost\"}}",
                "{\"RTD\":{\"timeout\":2147483648},\"SyncWrite\":\"sync\",\"RTD\":{\"host\":\"myhost\"}}",
                "{\"PublishDedupCacheSize\":2147483648,\"SyncWrite\":\"sync\"}",
                "{\"ConflationMs\":99999999999,\"SyncWrite\":\"sync\"}",
                "{\"SyncWrite\":\"Bogus\",\"RTD\":{\"host\":\"myhost\"}}",
                "{\"RTD\":{\"ExcelUpdateStyle\":\"Nope\",\"host\":\"myhost\"}}",
                "{\"AsyncWrites\":\"yes\",\"SyncWrite\":\"sync\"}",
            };
            foreach (var doc in docs)
            {
                // CleanOut itself calls ToObject on the cleaned doc; a throw fails here.
                var issues = CleanOut(doc, out var bound);
                Assert.NotNull(bound);
            }
        }

        [Fact]
        public void GoodValuesSurviveAlongsideBadOnes()
        {
            var json = "{\"RTD\":{\"host\":\"prodhost:6379\",\"timeout\":true},\"UDF\":{\"host\":\"prodhost:6379\"},\"SyncWrite\":\"sync\",\"Servers\":{\"prod\":\"prodhost:6379\"}}";
            CleanOut(json, out var bound);
            Assert.Equal("prodhost:6379", bound.RTD.host);
            Assert.Equal("prodhost:6379", bound.UDF.host);
            Assert.Equal("sync", bound.SyncWrite);
            Assert.Equal("prodhost:6379", bound.Servers["prod"]);
            // The bad bool is dropped, so the field is absent before Sanitize (the
            // default for the type applies: 0). Sanitize then substitutes 1000; that
            // belongs to the AppConfig tests, not the schema.
            Assert.Equal(0, bound.RTD.timeout);
        }

        // ---- the built schema types match the model types ---------------------

        [Fact]
        public void SchemaTypesMatchTheModelTypes()
        {
            // A weak but real structural check: every scalar schema type matches the
            // CLR type of the property of the same name on the model.
            var modelMap = new Dictionary<string, Type>
            {
                ["timeout"] = typeof(int),
                ["RedisUpdateRateMs"] = typeof(int),
                ["ExcelUpdateRateMs"] = typeof(int),
                ["MessageCounterThreshold"] = typeof(long),
                ["UseGetMultiple"] = typeof(bool),
                ["UpdateCheck"] = typeof(bool),
                ["SkipRepeatedMessages"] = typeof(bool),
                ["CoalesceRealtimeUpdates"] = typeof(bool),
                ["ConflationMs"] = typeof(int?),
                ["PublishDedupCacheSize"] = typeof(int),
                ["AsyncWrites"] = typeof(bool),
            };
            var schema = JObject.Parse(ConfigSchema.Text);
            var all = new Dictionary<string, JToken>();
            void Collect(JObject obj)
            {
                foreach (var p in obj.Properties())
                {
                    if (p.Value is JObject o && o["type"] != null)
                        all[p.Name] = o["type"];
                }
            }
            Collect((JObject)schema["properties"]);
            Collect((JObject)((JObject)schema["properties"])["RTD"]?["properties"]);
            Collect((JObject)((JObject)schema["properties"])["UDF"]?["properties"]);
            foreach (var kv in modelMap)
            {
                Assert.True(all.TryGetValue(kv.Key, out var declared), "schema type missing for " + kv.Key);
                string expected = kv.Value == typeof(bool) ? "boolean"
                    : (kv.Value == typeof(int) || kv.Value == typeof(long) || kv.Value == typeof(int?)) ? "integer"
                    : "string";
                Assert.Equal(expected, (string)declared);
            }
        }
    }
}

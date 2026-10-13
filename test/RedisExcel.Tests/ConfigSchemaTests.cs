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
        [InlineData("{\"UDF\":{\"host\":\"x\"}}")]   // the removed section is now an ERROR (a legacy key)
        public void LegacyKeys_AreErrors(string json)
        {
            Assert.True(Rejected(json));
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

        // ---- the built schema types match the model types ---------------------

        [Fact]
        public void SchemaTypesMatchTheModelTypes()
        {
            var modelMap = new Dictionary<string, Type>
            {
                ["RedisUpdateRateMs"] = typeof(int),
                ["ExcelUpdateRateMs"] = typeof(int),
                ["MessageCounterThreshold"] = typeof(long),
                ["UseGetMultiple"] = typeof(bool),
                ["CoalesceRealtimeUpdates"] = typeof(bool),
                ["UpdateCheck"] = typeof(bool),
                ["SkipRepeatedMessages"] = typeof(bool),
                ["ConflationMs"] = typeof(int?),
                ["PublishDedupCacheSize"] = typeof(int),
                ["AsyncWrites"] = typeof(bool),
                ["connectTimeout"] = typeof(int?),
                ["syncTimeout"] = typeof(int?),
            };
            var schema = JObject.Parse(ConfigSchema.Text);
            var all = new Dictionary<string, JToken>();
            void Collect(JObject obj)
            {
                if (obj == null) return;
                foreach (var p in obj.Properties())
                    if (p.Value is JObject o && o["type"] != null)
                        all[p.Name] = o["type"];
            }
            var props = (JObject)schema["properties"];
            Collect(props);
            Collect((JObject)((JObject)props["RTD"])["properties"]);
            Collect((JObject)((JObject)props["Timeout"])["properties"]);
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

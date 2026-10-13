using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    [Collection("write-mode-seam")]
    public class AppConfigTests
    {
        [Fact]
        public void CandidatePaths_PutTheAddInFolderFirst()
        {
            // A user who keeps the XLL in a dedicated folder (e.g. C:\RedisExcel)
            // can put RedisExcel.json next to it; that folder is searched first.
            string addIn = @"C:\RedisExcel";
            var paths = AppConfig.CandidatePaths(addIn, @"C:\ExcelFolder").Where(p => p != null).ToList();
            Assert.Equal(Path.Combine(addIn, "RedisExcel.json"), paths[0]);
            Assert.Contains(Path.Combine(@"C:\ExcelFolder", "RedisExcel.json"), paths);
        }

        [Fact]
        public void CandidatePaths_NoAddInFolder_SkipsItWithoutBreakingTheRest()
        {
            var paths = AppConfig.CandidatePaths(null, null).Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            Assert.All(paths, p => Assert.EndsWith("RedisExcel.json", p));
            Assert.DoesNotContain(@"C:\RedisExcel", string.Join(";", paths));
        }

        [Fact]
        public void Sanitize_FillsDefaults()
        {
            var config = AppConfig.Sanitize(new ConfigRoot());

            Assert.NotNull(config.RTD);
            Assert.Equal(AppConfig.DefaultHost, config.Default);
            Assert.NotNull(config.Timeout);
            Assert.NotNull(config.Servers);
            Assert.Equal(1000, config.RTD.RedisUpdateRateMs);
            Assert.Equal(100, config.RTD.ExcelUpdateRateMs);
            Assert.Equal(ENUMExcelUpdateStyle.Automatic, config.RTD.ExcelUpdateStyle);
            Assert.Equal(10000, config.RTD.MessageCounterThreshold);
            Assert.True(config.RTD.UseGetMultiple);
            Assert.Equal(10000, config.PublishDedupCacheSize);
        }

        [Fact]
        public void Sanitize_KeepsConfiguredValues()
        {
            var config = new ConfigRoot
            {
                Default = "prod",
                Servers = new Dictionary<string, ServerDef>
                {
                    ["prod"] = new ServerDef { host = "host1", connectTimeout = 8000, syncTimeout = 9000 },
                    ["cache"] = new ServerDef { host = "host2" }
                },
                RTD = new RTDConfig
                {
                    RedisUpdateRateMs = 7,
                    ExcelUpdateRateMs = 3,
                    ExcelUpdateStyle = ENUMExcelUpdateStyle.Timer,
                    MessageCounterThreshold = 5,
                    UseGetMultiple = false
                }
            };

            var sanitized = AppConfig.Sanitize(config);

            Assert.Equal("prod", sanitized.Default);
            Assert.Equal("host1", sanitized.Servers["prod"].host);
            Assert.Equal(8000, sanitized.Servers["prod"].connectTimeout);
            Assert.Equal(9000, sanitized.Servers["prod"].syncTimeout);
            Assert.Equal("host2", sanitized.Servers["cache"].host);
            Assert.Equal(7, sanitized.RTD.RedisUpdateRateMs);
            Assert.Equal(3, sanitized.RTD.ExcelUpdateRateMs);
            Assert.Equal(ENUMExcelUpdateStyle.Timer, sanitized.RTD.ExcelUpdateStyle);
            Assert.Equal(5, sanitized.RTD.MessageCounterThreshold);
            Assert.False(sanitized.RTD.UseGetMultiple);
        }

        [Fact]
        public void Sanitize_DropsServerEntriesWithNoHost()
        {
            var config = new ConfigRoot
            {
                Servers = new Dictionary<string, ServerDef>
                {
                    ["ok"] = new ServerDef { host = "h:1" },
                    ["empty"] = new ServerDef { host = "" },
                    ["null"] = null
                }
            };

            var sanitized = AppConfig.Sanitize(config);

            Assert.True(sanitized.Servers.ContainsKey("ok"));
            Assert.False(sanitized.Servers.ContainsKey("empty"));
            Assert.False(sanitized.Servers.ContainsKey("null"));
        }

        [Fact]
        public void Sanitize_FixesNonPositiveRates()
        {
            var config = new ConfigRoot
            {
                RTD = new RTDConfig { RedisUpdateRateMs = -5, ExcelUpdateRateMs = 0 }
            };

            var sanitized = AppConfig.Sanitize(config);

            Assert.Equal(1000, sanitized.RTD.RedisUpdateRateMs);
            Assert.Equal(100, sanitized.RTD.ExcelUpdateRateMs);
        }

        // ---- resolution -------------------------------------------------------

        private static ConfigRoot Config(string defaultRef, params (string name, string host)[] servers)
        {
            var map = new Dictionary<string, ServerDef>(StringComparer.Ordinal);
            foreach (var s in servers)
                map[s.name] = new ServerDef { host = s.host };
            return new ConfigRoot { Default = defaultRef, Servers = map };
        }

        [Theory]
        [InlineData(null, "localhost:6379")]      // blank -> Default (raw string)
        [InlineData("", "localhost:6379")]
        [InlineData("   ", "localhost:6379")]
        [InlineData("localhost:6379", "localhost:6379")] // raw string as-is
        [InlineData("prod", "h-prod")]            // name -> Servers[str]
        [InlineData("PROD", "h-prod")]            // case-insensitive name
        public void ResolveServer_ResolvesDefaultAndNames(string reference, string expected)
        {
            var cfg = Config("localhost:6379", ("prod", "h-prod"));
            Assert.Equal(expected, AppConfig.ResolveServer(cfg, reference)?.host);
        }

        [Fact]
        public void ResolveServer_ExactNameWinsOverCaseVariant()
        {
            var cfg = Config(null,
                ("prod", "exact"), ("Prod", "variant"));
            Assert.Equal("exact", AppConfig.ResolveServer(cfg, "prod").host);
            Assert.Equal("variant", AppConfig.ResolveServer(cfg, "Prod").host);
        }

        [Fact]
        public void ResolveServer_UnknownBareName_FallsBackToRawString()
        {
            // A name that is not defined is used as a raw connection string (so a
            // bare "myhost:6379" keeps working); it is NOT a silent default.
            var cfg = Config("localhost:6379");
            Assert.Equal("myhost:6379", AppConfig.ResolveServer(cfg, "myhost:6379").host);
        }

        [Fact]
        public void ServerTimeouts_ServerObjectWinsThenGlobalThenDefault()
        {
            var cfg = new ConfigRoot
            {
                Default = "prod",
                Timeout = new TimeoutConfig { connectTimeout = 3000, syncTimeout = 4000 },
                Servers = new Dictionary<string, ServerDef>
                {
                    ["prod"] = new ServerDef { host = "h", connectTimeout = 8000 }
                }
            };
            AppConfig.CurrentOverrideForTests = cfg;
            try
            {
                // connect: server object (8000); sync: global (4000)
                var t = AppConfig.ServerTimeouts(null);
                Assert.Equal(8000, t.connectTimeout);
                Assert.Equal(4000, t.syncTimeout);
            }
            finally { AppConfig.CurrentOverrideForTests = null; }

            // With no config at all: the built-in defaults.
            AppConfig.CurrentOverrideForTests = new ConfigRoot();
            try
            {
                var t = AppConfig.ServerTimeouts("localhost:6379");
                Assert.Equal(AppConfig.DefaultConnectTimeoutMs, t.connectTimeout);
                Assert.Equal(AppConfig.DefaultSyncTimeoutMs, t.syncTimeout);
            }
            finally { AppConfig.CurrentOverrideForTests = null; }
        }

        [Fact]
        public void ServerTimeouts_SameHostTwoServers_DefaultWins()
        {
            // Two Servers entries with the SAME host: the connection belongs to
            // Default, so its timeout must win for the flattened host string
            // (audit D1), regardless of declaration order.
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot
                {
                    Default = "secondary",
                    Servers = new Dictionary<string, ServerDef>
                    {
                        ["primary"] = new ServerDef { host = "localhost:6379", syncTimeout = 1111 },
                        ["secondary"] = new ServerDef { host = "localhost:6379", syncTimeout = 2222 }
                    }
                });
                AppConfig.CurrentOverrideForTests = cfg;

                var t = AppConfig.ServerTimeouts("localhost:6379");
                Assert.Equal(2222, t.syncTimeout); // Default (secondary), not the first entry
            }
            finally { AppConfig.CurrentOverrideForTests = original; }
        }

        [Fact]
        public void ServerTimeouts_ResolvesByServerHostToo_NotJustTheName()
        {
            // The connection manager passes the RESOLVED host string (the name is
            // flattened), so a per-server timeout must still be found by matching
            // the server whose host equals that string - otherwise the per-server
            // timeout is dead config.
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot
                {
                    Default = "prod",
                    Timeout = new TimeoutConfig { connectTimeout = 3000, syncTimeout = 3000 },
                    Servers = new Dictionary<string, ServerDef>
                    {
                        ["prod"] = new ServerDef { host = "prod-host:6379", connectTimeout = 12000, syncTimeout = 13000 }
                    }
                });
                AppConfig.CurrentOverrideForTests = cfg;

                // By name (alias text):
                var byName = AppConfig.ServerTimeouts("prod");
                Assert.Equal(12000, byName.connectTimeout);
                Assert.Equal(13000, byName.syncTimeout);
                // By the RESOLVED host string (what Connect actually passes):
                var byHost = AppConfig.ServerTimeouts("prod-host:6379");
                Assert.Equal(12000, byHost.connectTimeout);
                Assert.Equal(13000, byHost.syncTimeout);
                // The default (blank reference) also resolves to the prod server:
                var byDefault = AppConfig.ServerTimeouts(null);
                Assert.Equal(12000, byDefault.connectTimeout);
            }
            finally { AppConfig.CurrentOverrideForTests = original; }
        }

        [Fact]
        public void LoadFromPaths_LegacyKeys_RejectTheFileNotSilentlyDefault()
        {
            // A pre-1.6 file must FAIL CLOSED (a removed key is an error), never
            // silently reload on localhost while the validator says OK.
            using (var temp = new TempFiles())
            {
                string legacy = temp.Write("legacy.json",
                    "{\"RTD\":{\"host\":\"prod.redis.internal:6379\",\"timeout\":4000,\"RedisUpdateRateMs\":500}," +
                    "\"UDF\":{\"host\":\"prod.redis.internal:6379\",\"timeout\":4000}," +
                    "\"CoalesceRealtimeUpdates\":false,\"SyncWrite\":\"sync\"}");

                var config = AppConfig.LoadFromPaths(new[] { legacy });

                Assert.True(config.Rejected);
                Assert.Contains("RTD/host", config.RejectedReason);
            }
        }

        [Fact]
        public void LoadFromPaths_BareDefaultName_IsAWarningNotARejection()
        {
            // A bare word is a valid port-less host ("localhost", a service name),
            // so a Default that is not a Servers key must NOT reject the file - it
            // is used as the host, with a warning (audit F1). A silent localhost
            // fallback is still impossible: the value itself becomes the host.
            using (var temp = new TempFiles())
            {
                string bare = temp.Write("bare.json",
                    "{\"Default\":\"prod\",\"Servers\":{\"production\":\"h:6379\"}}");

                var config = AppConfig.LoadFromPaths(new[] { bare });

                Assert.False(config.Rejected);
                Assert.Equal("prod", config.Default);
                var issues = new System.Collections.Generic.List<ConfigIssue>();
                AppConfig.ParseTolerant(System.IO.File.ReadAllText(bare), null, issues);
                Assert.Contains(issues, i => i.Path == "/Default" && i.Severity == ConfigSeverity.Warning);
            }
        }

        [Fact]
        public void LoadFromPaths_DefaultRawConnectionString_IsAccepted()
        {
            using (var temp = new TempFiles())
            {
                string raw = temp.Write("raw.json", "{\"Default\":\"localhost:6379\"}");
                Assert.False(AppConfig.LoadFromPaths(new[] { raw }).Rejected);
                string dotted = temp.Write("dotted.json", "{\"Default\":\"redis.example.com\"}");
                Assert.False(AppConfig.LoadFromPaths(new[] { dotted }).Rejected);
            }
        }

        [Fact]
        public void Sanitize_UndefinedExcelUpdateStyle_ResetsToAutomatic()
        {
            var config = new ConfigRoot
            {
                RTD = new RTDConfig { ExcelUpdateStyle = (ENUMExcelUpdateStyle)99 }
            };

            Assert.Equal(
                ENUMExcelUpdateStyle.Automatic,
                AppConfig.Sanitize(config).RTD.ExcelUpdateStyle);
        }

        [Fact]
        public void Sanitize_UpdateCheckDefaultsToEnabled()
        {
            Assert.True(AppConfig.Sanitize(new ConfigRoot()).UpdateCheck);
            Assert.False(AppConfig.Sanitize(new ConfigRoot { UpdateCheck = false }).UpdateCheck);
        }

        [Fact]
        public void ConfigDefaults_PerformanceOptionsEnabled()
        {
            Assert.True(new ConfigRoot().SkipRepeatedMessages);
            Assert.False(new ConfigRoot { SkipRepeatedMessages = false }.SkipRepeatedMessages);
            Assert.True(new RTDConfig().CoalesceRealtimeUpdates);
            Assert.False(new RTDConfig { CoalesceRealtimeUpdates = false }.CoalesceRealtimeUpdates);
        }

        [Fact]
        public void Sanitize_PublishDedupCacheSize_FallsBackToDefault()
        {
            Assert.Equal(10000, new ConfigRoot().PublishDedupCacheSize);
            Assert.Equal(10000, AppConfig.Sanitize(new ConfigRoot { PublishDedupCacheSize = 0 }).PublishDedupCacheSize);
            Assert.Equal(10000, AppConfig.Sanitize(new ConfigRoot { PublishDedupCacheSize = -3 }).PublishDedupCacheSize);
            Assert.Equal(500, AppConfig.Sanitize(new ConfigRoot { PublishDedupCacheSize = 500 }).PublishDedupCacheSize);
        }

        [Fact]
        public void Sanitize_SyncWrite_DefaultsToFireForgetWhenAbsent()
        {
            Assert.Equal("fireforget", new ConfigRoot().SyncWrite);
            Assert.Equal("fireforget", AppConfig.Sanitize(new ConfigRoot()).SyncWrite);
            Assert.Equal("fireforget", AppConfig.Sanitize(new ConfigRoot { SyncWrite = null }).SyncWrite);
            Assert.Equal("fireforget", AppConfig.Sanitize(new ConfigRoot { SyncWrite = "" }).SyncWrite);
            Assert.Equal("fireforget", AppConfig.Sanitize(new ConfigRoot { SyncWrite = "   " }).SyncWrite);
        }

        [Theory]
        [InlineData("sync", "sync")]
        [InlineData("SYNC", "sync")]
        [InlineData("fireforget", "fireforget")]
        [InlineData("FireForget", "fireforget")]
        [InlineData("FIREFORGET-ALL", "fireforget-all")]
        [InlineData("FireForget-All", "fireforget-all")]
        public void Sanitize_SyncWrite_NormalizesKnownValuesToLowercase(string configured, string expected)
        {
            Assert.Equal(expected, AppConfig.Sanitize(new ConfigRoot { SyncWrite = configured }).SyncWrite);
        }

        [Fact]
        public void Sanitize_SyncWrite_UnknownValueFallsBackToFireForget()
        {
            Assert.Equal("fireforget", AppConfig.Sanitize(new ConfigRoot { SyncWrite = "bogus" }).SyncWrite);
            Assert.Equal("fireforget", AppConfig.NormalizeSyncWrite("bogus"));
            Assert.Equal("sync", AppConfig.NormalizeSyncWrite(" sync "));
        }

        [Fact]
        public void Sanitize_AsyncWrites_DefaultsToFalseAndRoundTrips()
        {
            Assert.False(new ConfigRoot().AsyncWrites);
            Assert.False(AppConfig.Sanitize(new ConfigRoot()).AsyncWrites);
            Assert.True(AppConfig.Sanitize(new ConfigRoot { AsyncWrites = true }).AsyncWrites);
            Assert.False(AppConfig.Sanitize(new ConfigRoot { AsyncWrites = false }).AsyncWrites);
        }

        // ---- loader (new shape) ----------------------------------------------

        [Fact]
        public void LoadFromPaths_MalformedFirstExisting_FailsClosedAndStops()
        {
            using (var temp = new TempFiles())
            {
                string malformed = temp.Write("malformed.json", "{ this is not json");
                string wellFormed = temp.Write("well-formed.json", "{\"Default\":\"second-host:6379\"}");

                var config = AppConfig.LoadFromPaths(new[] { malformed, wellFormed });

                // A first existing file that cannot be parsed fails CLOSED (the
                // lower-priority file is never silently used).
                Assert.True(config.Rejected);
                Assert.Contains("malformed.json", config.RejectedReason);
            }
        }

        [Fact]
        public void LoadFromPaths_NonObjectRoot_FailsClosed()
        {
            // A file whose ROOT is not a JSON object (array/number/string/bool)
            // must fail closed too, not silently fall back to localhost (audit C3
            // showed this branch was unpinned).
            using (var temp = new TempFiles())
            {
                foreach (var root in new[] { "[1,2]", "123", "\"x\"", "true" })
                {
                    string file = temp.Write("bad-root.json", root);
                    var config = AppConfig.LoadFromPaths(new[] { file });
                    Assert.True(config.Rejected, "root " + root + " was not rejected");
                    Assert.Contains("bad-root.json", config.RejectedReason);
                }
            }
        }

        [Fact]
        public void LoadFromPaths_LeadingComment_IsAccepted()
        {
            // Comments are documented as allowed anywhere; a file whose FIRST token
            // is a comment must still load (a regression fixed after audit B1).
            using (var temp = new TempFiles())
            {
                string lineComment = temp.Write("lead-line.json", "// leading note\n{ \"Default\":\"myhost:6379\" }");
                Assert.False(AppConfig.LoadFromPaths(new[] { lineComment }).Rejected);
                Assert.Equal("myhost:6379", AppConfig.LoadFromPaths(new[] { lineComment }).Default);

                string blockComment = temp.Write("lead-block.json", "/* leading */ { \"Default\":\"myhost:6379\" }");
                Assert.Equal("myhost:6379", AppConfig.LoadFromPaths(new[] { blockComment }).Default);
            }
        }

        [Fact]
        public void LoadFromPaths_NoExistingFiles_UsesLegacyDefaults()
        {
            using (var temp = new TempFiles())
            {
                var config = AppConfig.LoadFromPaths(new[]
                {
                    temp.PathOf("missing-1.json"),
                    temp.PathOf("missing-2.json")
                });

                Assert.Equal(AppConfig.DefaultHost, config.Default);
                Assert.Equal(1000, config.RTD.ExcelUpdateRateMs);
                Assert.Equal(10000, config.PublishDedupCacheSize);
            }
        }

        [Fact]
        public void LoadFromPaths_ValidFile_ParsesConfiguredValues()
        {
            using (var temp = new TempFiles())
            {
                string valid = temp.Write("valid.json",
                    "{\"Default\":\"prod\"," +
                    "\"Servers\":{\"prod\":{\"host\":\"h1\",\"connectTimeout\":8000},\"cache\":\"h2\"}," +
                    "\"RTD\":{\"RedisUpdateRateMs\":42,\"ExcelUpdateRateMs\":7}," +
                    "\"PublishDedupCacheSize\":7,\"SkipRepeatedMessages\":false}");

                var config = AppConfig.LoadFromPaths(new[] { valid });

                Assert.Equal("prod", config.Default);
                Assert.Equal("h1", config.Servers["prod"].host);
                Assert.Equal(8000, config.Servers["prod"].connectTimeout);
                Assert.Equal("h2", config.Servers["cache"].host);   // string shorthand
                Assert.Equal(42, config.RTD.RedisUpdateRateMs);
                Assert.Equal(7, config.RTD.ExcelUpdateRateMs);
                Assert.Equal(7, config.PublishDedupCacheSize);
                Assert.False(config.SkipRepeatedMessages);
            }
        }

        [Fact]
        public void LoadFromPaths_WrongValueOnKnownKey_RejectsTheWholeFile()
        {
            // 1.6 policy: a wrong VALUE on a known key fails closed (the file is
            // rejected), unlike the old per-value tolerance. Extra keys are not
            // values and stay warnings.
            using (var temp = new TempFiles())
            {
                string bad = temp.Write("bad.json",
                    "{\"Default\":\"myhost:6379\",\"RTD\":{\"RedisUpdateRateMs\":\"nope\"},\"SyncWrite\":\"sync\"}");

                var config = AppConfig.LoadFromPaths(new[] { bad });

                Assert.True(config.Rejected);
                Assert.NotNull(config.RejectedReason);
                Assert.Contains("RedisUpdateRateMs", config.RejectedReason);
            }
        }

        [Fact]
        public void LoadFromPaths_UnknownExtraKey_IsToleratedAsWarning()
        {
            // An extra key we do not use must not reject the file.
            using (var temp = new TempFiles())
            {
                string extra = temp.Write("extra.json",
                    "{\"Default\":\"myhost:6379\",\"Nope\":1,\"SyncWrite\":\"sync\"}");

                var config = AppConfig.LoadFromPaths(new[] { extra });

                Assert.False(config.Rejected);
                Assert.Equal("myhost:6379", config.Default);
                Assert.Equal("sync", config.SyncWrite);
            }
        }

        [Fact]
        public void LoadFromPaths_ParsesSyncWriteAndAsyncWrites()
        {
            using (var temp = new TempFiles())
            {
                string valid = temp.Write("writes.json",
                    "{\"SyncWrite\":\"FIREFORGET-ALL\",\"AsyncWrites\":true}");

                var config = AppConfig.LoadFromPaths(new[] { valid });

                Assert.Equal("fireforget-all", config.SyncWrite);
                Assert.True(config.AsyncWrites);
            }
        }

        [Fact]
        public void LoadFromPaths_ServerStringOrObject_BothBind()
        {
            using (var temp = new TempFiles())
            {
                string valid = temp.Write("servers.json",
                    "{\"Servers\":{\"a\":\"h:1\",\"b\":{\"host\":\"h:2\",\"syncTimeout\":9000}}}");

                var config = AppConfig.LoadFromPaths(new[] { valid });

                Assert.Equal("h:1", config.Servers["a"].host);
                Assert.Equal("h:2", config.Servers["b"].host);
                Assert.Equal(9000, config.Servers["b"].syncTimeout);
            }
        }

        [Fact]
        public void LoadFromPaths_AcceptsCommentsInTheConfigFile()
        {
            using (var temp = new TempFiles())
            {
                string withComments = temp.Write("comments.json",
                    "{\n// the production host\n\"Default\":\"myhost:6379\",\n" +
                    "/* write mode */ \"SyncWrite\":\"sync\"\n}");

                var config = AppConfig.LoadFromPaths(new[] { withComments });

                Assert.Equal("myhost:6379", config.Default);
                Assert.Equal("sync", config.SyncWrite);
            }
        }

        [Fact]
        public void LoadFromPaths_AliasWithADot_KeepsTheRestOfTheFile()
        {
            using (var temp = new TempFiles())
            {
                string partlyBad = temp.Write("dot-alias.json",
                    "{\"Servers\":{\"prod.eu\":\"prodhost:6379\"},\"SyncWrite\":\"sync\"," +
                    "\"Default\":\"myhost:6379\"}");

                var config = AppConfig.LoadFromPaths(new[] { partlyBad });

                Assert.Equal("sync", config.SyncWrite);
                Assert.Equal("myhost:6379", config.Default);
                Assert.Equal("prodhost:6379", config.Servers["prod.eu"].host);
            }
        }

        [Fact]
        public void LoadFromPaths_ServersBadValue_DropsJustThatServer()
        {
            using (var temp = new TempFiles())
            {
                string partlyBad = temp.Write("server-val.json",
                    "{\"Servers\":{\"ok\":\"h:6379\",\"bad\":123},\"SyncWrite\":\"sync\"}");

                var config = AppConfig.LoadFromPaths(new[] { partlyBad });

                // A Servers value that is neither a string nor an object is a value
                // error -> reject the file (fail-closed).
                Assert.True(config.Rejected);
            }
        }

        [Fact]
        public void ValidateConfiguration_ReportsErrorsWithoutThrowing()
        {
            using (var temp = new TempFiles())
            {
                string bad = temp.Write("bad.json", "{\"RTD\":{\"RedisUpdateRateMs\":true},\"AsyncWrites\":\"yes\"}");
                var issues = new List<ConfigIssue>();
                var parsed = AppConfig.ParseTolerant(File.ReadAllText(bad), null, issues);

                Assert.NotNull(parsed);
                Assert.True(parsed.Rejected);
                Assert.Contains(issues, i => i.Path.Contains("RedisUpdateRateMs"));
                Assert.Contains(issues, i => i.Path.Contains("AsyncWrites"));
            }
        }

        [Fact]
        public void ValidateConfiguration_ByPath_ReportsOkOrProblems()
        {
            using (var temp = new TempFiles())
            {
                string good = temp.Write("good.json", "{\"SyncWrite\":\"sync\"}");
                string bad = temp.Write("bad.json", "{\"ConflationMs\":\"soon\"}");

                Assert.StartsWith("OK", AppConfig.ValidateConfiguration(good));
                string report = AppConfig.ValidateConfiguration(bad);
                Assert.StartsWith("Error:", report);
                Assert.Contains("ConflationMs", report);
                // An out-of-range ConflationMs is VALID (Sanitize clamps it).
                string clamped = temp.Write("clamped.json", "{\"ConflationMs\":9999999}");
                Assert.StartsWith("OK", AppConfig.ValidateConfiguration(clamped));
            }
        }

        [Fact]
        public void ValidateConfiguration_MissingFile_ReturnsError()
        {
            Assert.StartsWith("Error: file not found", AppConfig.ValidateConfiguration(@"Z:\nope\does-not-exist.json"));
        }

        [Fact]
        public void RedisCheckConfigFile_BlankValidatesInUseAndPathValidatesThatFile()
        {
            using (var temp = new TempFiles())
            {
                string good = temp.Write("good.json", "{\"SyncWrite\":\"sync\"}");
                Assert.StartsWith("OK", (string)RedisUDF.RedisCheckConfigFile(good));
                Assert.StartsWith("Error: file not found", (string)RedisUDF.RedisCheckConfigFile(@"Z:\nope.json"));
                string inUse = (string)RedisUDF.RedisCheckConfigFile(ExcelDna.Integration.ExcelMissing.Value);
                Assert.True(inUse.StartsWith("OK") || inUse.StartsWith("Error:"), "unexpected report: " + inUse);
            }
        }

        /// <summary>Creates a unique temp directory for config files and removes
        /// it when the test finishes.</summary>
        private sealed class TempFiles : IDisposable
        {
            private readonly string _directory;

            public TempFiles()
            {
                _directory = Path.Combine(Path.GetTempPath(), "redisexcel-config-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_directory);
            }

            public string PathOf(string name) => Path.Combine(_directory, name);

            public string Write(string name, string content)
            {
                string path = PathOf(name);
                File.WriteAllText(path, content);
                return path;
            }

            public void Dispose()
            {
                try
                {
                    Directory.Delete(_directory, recursive: true);
                }
                catch
                {
                    // Best-effort cleanup; a locked temp file must not fail the test.
                }
            }
        }
    }
}

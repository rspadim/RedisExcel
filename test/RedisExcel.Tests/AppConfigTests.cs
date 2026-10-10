using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace RedisExcel.Tests
{
    public class AppConfigTests
    {
        [Fact]
        public void Sanitize_FillsMissingSectionsWithDefaults()
        {
            var config = AppConfig.Sanitize(new ConfigRoot());

            Assert.NotNull(config.RTD);
            Assert.NotNull(config.UDF);
            Assert.Equal(AppConfig.DefaultHost, config.RTD.host);
            Assert.Equal(AppConfig.DefaultHost, config.UDF.host);
            Assert.Equal(1000, config.RTD.timeout);
            Assert.Equal(1000, config.UDF.timeout);
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
                RTD = new RTDConfig
                {
                    host = "host1",
                    timeout = 42,
                    RedisUpdateRateMs = 7,
                    ExcelUpdateRateMs = 3,
                    ExcelUpdateStyle = ENUMExcelUpdateStyle.Timer,
                    MessageCounterThreshold = 5,
                    UseGetMultiple = false
                },
                UDF = new ConfigSection { host = "host2", timeout = 9 }
            };

            var sanitized = AppConfig.Sanitize(config);

            Assert.Equal("host1", sanitized.RTD.host);
            Assert.Equal(42, sanitized.RTD.timeout);
            Assert.Equal(7, sanitized.RTD.RedisUpdateRateMs);
            Assert.Equal(3, sanitized.RTD.ExcelUpdateRateMs);
            Assert.Equal(ENUMExcelUpdateStyle.Timer, sanitized.RTD.ExcelUpdateStyle);
            Assert.Equal(5, sanitized.RTD.MessageCounterThreshold);
            Assert.False(sanitized.RTD.UseGetMultiple);
            Assert.Equal("host2", sanitized.UDF.host);
            Assert.Equal(9, sanitized.UDF.timeout);
        }

        [Fact]
        public void Sanitize_FixesNonPositiveTimeoutsAndRates()
        {
            var config = new ConfigRoot
            {
                RTD = new RTDConfig { timeout = 0, RedisUpdateRateMs = -5, ExcelUpdateRateMs = 0 },
                UDF = new ConfigSection { timeout = -1 }
            };

            var sanitized = AppConfig.Sanitize(config);

            Assert.Equal(1000, sanitized.RTD.timeout);
            Assert.Equal(1000, sanitized.RTD.RedisUpdateRateMs);
            Assert.Equal(100, sanitized.RTD.ExcelUpdateRateMs);
            Assert.Equal(1000, sanitized.UDF.timeout);
        }

        [Theory]
        [InlineData(null, "default", null, "default")]
        [InlineData("", "default", null, "default")]
        [InlineData("   ", "default", null, "default")]
        [InlineData("alias", "default", "alias=resolved", "resolved")]
        [InlineData(" prod ", "default", "prod=resolved", "resolved")] // trimmed before lookup
        [InlineData("PROD", "default", "prod=resolved", "resolved")] // case-insensitive fallback
        [InlineData("unknown", "default", null, "unknown")]
        [InlineData("alias", "alias", "alias=resolved", "resolved")]
        public void ResolveHostCore_ResolvesAliasesAndDefaults(string host, string defaultHost, string mapping, string expected)
        {
            var servers = new Dictionary<string, string>();
            if (mapping != null)
            {
                var parts = mapping.Split('=');
                servers[parts[0]] = parts[1];
            }

            Assert.Equal(expected, AppConfig.ResolveHostCore(host, defaultHost, servers));
        }

        [Fact]
        public void ResolveHostCore_ExactAliasWinsOverCaseVariant()
        {
            var servers = new Dictionary<string, string>
            {
                { "prod", "exact" },
                { "Prod", "variant" }
            };

            // Each spelling hits its own exact key (ordinal) instead of the
            // case-insensitive scan picking an arbitrary dictionary order.
            Assert.Equal("exact", AppConfig.ResolveHostCore("prod", null, servers));
            Assert.Equal("variant", AppConfig.ResolveHostCore("Prod", null, servers));
        }

        [Fact]
        public void ResolveHostCore_NullHostAndDefault_ReturnsNull()
        {
            var servers = new Dictionary<string, string> { { "prod", "x" } };

            // With no default host there is nothing to resolve: null in, null
            // out (an alias mapping must not invent a host).
            Assert.Null(AppConfig.ResolveHostCore(null, null, servers));
            Assert.Null(AppConfig.ResolveHostCore("   ", null, servers));
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
            Assert.True(new ConfigRoot().CoalesceRealtimeUpdates);
            Assert.False(new ConfigRoot { CoalesceRealtimeUpdates = false }.CoalesceRealtimeUpdates);
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
            // Absent (fresh instance), null, empty and whitespace all fall back.
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

        [Fact]
        public void LoadFromPaths_MalformedFirstExisting_StopsAndUsesDefaults()
        {
            using (var temp = new TempFiles())
            {
                string malformed = temp.Write("malformed.json", "{ this is not json");
                string wellFormed = temp.Write("well-formed.json", "{\"RTD\":{\"host\":\"second-file-host\"}}");

                var config = AppConfig.LoadFromPaths(new[] { malformed, wellFormed });

                // The first existing file wins even when malformed: no fall-through.
                Assert.Equal(AppConfig.DefaultHost, config.RTD.host);
                // No usable file: the legacy no-file flush timer applies.
                Assert.Equal(1000, config.RTD.ExcelUpdateRateMs);
                Assert.Equal(10000, config.PublishDedupCacheSize);
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

                Assert.Equal(AppConfig.DefaultHost, config.RTD.host);
                Assert.Equal(1000, config.RTD.ExcelUpdateRateMs);
                Assert.Equal(10000, config.PublishDedupCacheSize);
            }
        }

        [Fact]
        public void LoadFromPaths_MissingFirstThenValid_LoadsSecond()
        {
            using (var temp = new TempFiles())
            {
                string missing = temp.PathOf("missing.json");
                string valid = temp.Write("valid.json", "{\"RTD\":{\"host\":\"loaded-host\"}}");

                var config = AppConfig.LoadFromPaths(new[] { missing, null, valid });

                Assert.Equal("loaded-host", config.RTD.host);
                // A usable file keeps the standard default, not the legacy fallback.
                Assert.Equal(100, config.RTD.ExcelUpdateRateMs);
            }
        }

        [Fact]
        public void LoadFromPaths_ValidFile_ParsesConfiguredValues()
        {
            using (var temp = new TempFiles())
            {
                string valid = temp.Write("valid.json",
                    "{\"RTD\":{\"host\":\"h1\",\"RedisUpdateRateMs\":42,\"ExcelUpdateRateMs\":7}," +
                    "\"UDF\":{\"host\":\"h2\",\"timeout\":9}," +
                    "\"PublishDedupCacheSize\":7," +
                    "\"SkipRepeatedMessages\":false}");

                var config = AppConfig.LoadFromPaths(new[] { valid });

                Assert.Equal("h1", config.RTD.host);
                Assert.Equal(42, config.RTD.RedisUpdateRateMs);
                Assert.Equal(7, config.RTD.ExcelUpdateRateMs);
                Assert.Equal("h2", config.UDF.host);
                Assert.Equal(9, config.UDF.timeout);
                Assert.Equal(7, config.PublishDedupCacheSize);
                Assert.False(config.SkipRepeatedMessages);
            }
        }

        [Fact]
        public void LoadFromPaths_OneBadValue_KeepsEveryOtherValue()
        {
            using (var temp = new TempFiles())
            {
                // A wrong-typed/unknown-enum value must NOT discard the rest of the
                // file: that silently repointed the host to localhost and changed
                // the write mode.
                string partlyBad = temp.Write("partly-bad.json",
                    "{\"RTD\":{\"host\":\"myhost\",\"ExcelUpdateStyle\":\"Bogus\",\"RedisUpdateRateMs\":42}," +
                    "\"SyncWrite\":\"sync\",\"PublishDedupCacheSize\":9,\"UpdateCheck\":\"no\"}");

                var config = AppConfig.LoadFromPaths(new[] { partlyBad });

                Assert.Equal("myhost", config.RTD.host);            // kept
                Assert.Equal(42, config.RTD.RedisUpdateRateMs);      // kept
                Assert.Equal("sync", config.SyncWrite);              // kept
                Assert.Equal(9, config.PublishDedupCacheSize);       // kept
                Assert.Equal(ENUMExcelUpdateStyle.Automatic, config.RTD.ExcelUpdateStyle); // bad value -> default
            }
        }

        [Fact]
        public void LoadFromPaths_ParsesSyncWriteAndAsyncWrites()
        {
            using (var temp = new TempFiles())
            {
                // The config may use any case; the sanitized value is lowercase.
                string valid = temp.Write("writes.json",
                    "{\"SyncWrite\":\"FIREFORGET-ALL\",\"AsyncWrites\":true}");

                var config = AppConfig.LoadFromPaths(new[] { valid });

                Assert.Equal("fireforget-all", config.SyncWrite);
                Assert.True(config.AsyncWrites);
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

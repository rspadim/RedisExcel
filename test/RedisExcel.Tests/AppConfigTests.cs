using System.Collections.Generic;
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
        public void Sanitize_UpdateCheckDefaultsToEnabled()
        {
            Assert.True(AppConfig.Sanitize(new ConfigRoot()).UpdateCheck);
            Assert.False(AppConfig.Sanitize(new ConfigRoot { UpdateCheck = false }).UpdateCheck);
        }

        [Fact]
        public void ConfigDefaults_SkipRepeatedMessagesEnabled()
        {
            Assert.True(new ConfigRoot().SkipRepeatedMessages);
            Assert.False(new ConfigRoot { SkipRepeatedMessages = false }.SkipRepeatedMessages);
        }
    }
}

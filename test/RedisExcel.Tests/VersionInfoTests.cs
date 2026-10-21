using System;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The version/diagnostic worksheet function. It answers the first question a
    /// support case asks ("which RedisExcel is running, and from where?"), and it
    /// is the cheap way to spot TWO installed copies: each one reports a
    /// different xll path. It must never throw into Excel, and it must work
    /// without a Redis server (a broken install is exactly when it is used).
    /// </summary>
    public class VersionInfoTests
    {
        [Fact]
        public void BuildInfo_VersionIsTheTagWithoutTheVPrefix()
        {
            // The tag is embedded by the CI (InformationalVersion); the readable
            // version is that tag minus the "v" and any build metadata. A local
            // build reports "dev", so the assertion is on the SHAPE, not on the
            // number - and it must never be blank.
            Assert.False(string.IsNullOrWhiteSpace(BuildInfo.Version));
            Assert.StartsWith("RedisExcel " + BuildInfo.Version, RedisUDF.RedisUDFVersion());
            Assert.DoesNotContain("+", BuildInfo.Version);
            Assert.False(BuildInfo.Version.StartsWith("v", StringComparison.OrdinalIgnoreCase)
                && BuildInfo.Version.Length > 1 && char.IsDigit(BuildInfo.Version[1]),
                "the leading 'v' must be stripped: " + BuildInfo.Version);
        }

        [Fact]
        public void RedisUDFVersion_ReportsTheXllPathAndConfigInUse()
        {
            string text = RedisUDF.RedisUDFVersion();

            // xll= is the point of the function: without it two copies are
            // indistinguishable. Under the test runner there is no Excel host,
            // so ExcelDnaUtil.XllPath is unavailable and it reports "(unknown)";
            // inside Excel it carries the loaded file.
            Assert.Contains("| xll=", text, StringComparison.Ordinal);
            Assert.Contains("| config=", text, StringComparison.Ordinal);
            Assert.Contains("| connections=", text, StringComparison.Ordinal);

            // Single line on purpose: a multi-line volatile cell makes the sheet
            // jump every recalculation.
            Assert.DoesNotContain("\n", text);
        }

        [Fact]
        public void RedisUDFVersion_DoesNotNeedARedisServer()
        {
            // No connection is opened by the function itself; the counter is a
            // neutral read (0 when nothing connected). A live host that is down
            // must not turn this cell into an error.
            Assert.Contains("connections=0", RedisUDF.RedisUDFVersion());
        }
    }
}

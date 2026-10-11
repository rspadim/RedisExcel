using System;
using Xunit;
using Xunit.Abstractions;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Reports the environment's antivirus state so a unit run is attributable
    /// like the smoke/liveness/load/E2E suites (which print the same banner), and
    /// verifies the probe itself is well-formed.
    /// </summary>
    public class EnvironmentInfoTests
    {
        private readonly ITestOutputHelper _output;

        public EnvironmentInfoTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Antivirus_ProbeReturnsAWellFormedStatus()
        {
            string status = AntivirusInfo.Describe();
            _output.WriteLine("Antivirus: " + status);

            // The probe must always return a non-empty, single-line report: either
            // "none reported ..." or "ACTIVE: <products>". This is falsifiable - a
            // probe that threw, returned null/empty or produced a multi-line blob
            // (which would break the suite banner) fails.
            Assert.False(string.IsNullOrWhiteSpace(status));
            Assert.DoesNotContain("\n", status);
            Assert.True(
                status.StartsWith("none reported", StringComparison.Ordinal)
                || status.StartsWith("ACTIVE:", StringComparison.Ordinal),
                "unexpected AV status format: " + status);
        }
    }
}

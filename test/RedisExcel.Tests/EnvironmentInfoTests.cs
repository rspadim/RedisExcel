using Xunit;
using Xunit.Abstractions;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Reports the environment's antivirus state so a unit run is attributable
    /// like the smoke/liveness/load/E2E suites (which print the same banner).
    /// The unit suite is offline and AV-irrelevant, but recording the state
    /// keeps the environment explanation consistent across every suite. The
    /// line is visible with <c>dotnet test --logger "console;verbosity=detailed"</c>
    /// (or in the trx artifacts).
    /// </summary>
    public class EnvironmentInfoTests
    {
        private readonly ITestOutputHelper _output;

        public EnvironmentInfoTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Antivirus_EnvironmentIsReported()
        {
            string status = AntivirusInfo.Describe();
            _output.WriteLine("Antivirus: " + status);
            // Informational by design: the value depends on the machine, so the
            // test only asserts the probe does not throw and returns text.
            Assert.False(string.IsNullOrWhiteSpace(status));
        }
    }
}

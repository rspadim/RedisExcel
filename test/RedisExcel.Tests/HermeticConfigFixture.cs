using System;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Makes the unit suite hermetic with respect to the machine's configuration.
    /// <c>AppConfig.Current</c> normally reads %USERPROFILE%\RedisExcel.json, so a
    /// developer (or a leftover E2E run) with <c>AsyncWrites:true</c> there would
    /// change the UDF dispatch and fail tests for the wrong reason. This fixture
    /// pins a known default for the whole write-mode collection.
    ///
    /// Tests that exercise parsing call <c>LoadFromPaths</c>/<c>ParseTolerant</c>
    /// directly and are unaffected; the write-mode seams
    /// (<c>SyncWriteOverrideForTests</c>/<c>AsyncWritesOverrideForTests</c>) still
    /// take precedence over the configured values.
    /// </summary>
    [CollectionDefinition("write-mode-seam")]
    public class WriteModeSeamCollection : ICollectionFixture<HermeticConfigFixture>
    {
    }

    public sealed class HermeticConfigFixture : IDisposable
    {
        private readonly ConfigRoot _previous;

        public HermeticConfigFixture()
        {
            _previous = AppConfig.CurrentOverrideForTests;
            AppConfig.CurrentOverrideForTests = AppConfig.Sanitize(new ConfigRoot());
        }

        public void Dispose()
        {
            AppConfig.CurrentOverrideForTests = _previous;
        }
    }
}

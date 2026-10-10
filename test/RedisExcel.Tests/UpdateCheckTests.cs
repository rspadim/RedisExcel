using System;
using System.Globalization;
using System.Reflection;
using Xunit;

namespace RedisExcel.Tests
{
    public class UpdateCheckTests
    {
        [Theory]
        [InlineData("v1.1.1", "v1.1.0", true)]
        [InlineData("v1.1.0", "v1.1.0", false)]
        [InlineData("v1.0.10", "v1.0.9", true)]
        [InlineData("v1.2", "v1.10", false)] // numeric comparison, not lexicographic
        [InlineData("v1.0.7", "v1.0.6.8", true)]
        [InlineData("v1.1.0-beta", "v1.1.0", false)] // suffix ignored
        [InlineData("v1.4.1 -rc", "v1.4.0", true)] // space before the suffix still resolves
        [InlineData("v2147483648.0.0", "v1.4.0", true)] // over int.MaxValue: parsed as long
        [InlineData("v1.1.0", "dev", false)] // unknown current version: never alert
        [InlineData("", "v1.0.0", false)]
        public void IsNewer_ComparesNumericVersions(string candidate, string current, bool expected)
        {
            Assert.Equal(expected, UpdateCheck.IsNewer(candidate, current));
        }

        [Theory]
        [InlineData("v1.1.0", "1.1.0")]
        [InlineData("V1.1.0", "1.1.0")]
        [InlineData(" 1.1.0 ", "1.1.0")]
        [InlineData("v1.1.0-beta+meta", "1.1.0")]
        [InlineData("v1.4.1 -rc", "1.4.1")] // the space left by the suffix strip is re-trimmed
        [InlineData(null, "")]
        public void NormalizeTag_StripsPrefixAndSuffix(string input, string expected)
        {
            Assert.Equal(expected, UpdateCheck.NormalizeTag(input));
        }

        // RedisUDFUpdateAvailable is a non-blocking status read: it only
        // schedules a background refresh (EnsureFresh, gated by
        // AppConfig.UpdateCheck and the in-flight/last-attempt windows) and then
        // reports the last KNOWN tag. It performs no network I/O on the calling
        // thread and never throws into Excel. Offline there is no known tag, so
        // the value is deterministically false; once a tag is known it is a pure
        // comparison against the running build (BuildInfo.Tag).
        private static readonly FieldInfo LatestTagField =
            typeof(UpdateCheck).GetField("_latestTag", BindingFlags.NonPublic | BindingFlags.Static);

        [Fact]
        public void RedisUDFUpdateAvailable_ReadsTheLastKnownTagWithoutNetwork()
        {
            Assert.NotNull(LatestTagField); // pins the private state this test seeds
            object originalTag = LatestTagField.GetValue(null);
            bool originalUpdateCheck = AppConfig.Current.UpdateCheck;
            try
            {
                // Make EnsureFresh a no-op so no background Refresh task can
                // overwrite the seeded tag while the assertions run.
                AppConfig.Current.UpdateCheck = false;

                // Offline (nothing known yet): boolean false, no exception.
                LatestTagField.SetValue(null, null);
                Assert.False(RedisUDF.RedisUDFUpdateAvailable());

                // The running build is never "newer than itself".
                LatestTagField.SetValue(null, UpdateCheck.CurrentTag);
                Assert.False(RedisUDF.RedisUDFUpdateAvailable());

                // A strictly newer tag flips it to true. "dev" (and any other
                // unparseable current tag) can never be beaten by design, so the
                // true case only runs for a parseable version.
                string newer = NewerThan(UpdateCheck.CurrentTag);
                if (newer != null)
                {
                    LatestTagField.SetValue(null, newer);
                    Assert.True(RedisUDF.RedisUDFUpdateAvailable());
                }
            }
            finally
            {
                LatestTagField.SetValue(null, originalTag);
                AppConfig.Current.UpdateCheck = originalUpdateCheck;
            }
        }

        /// <summary>Returns a "v{major+1}.0.0" tag guaranteed newer than the
        /// running build, or null when the build tag is not a plain numeric
        /// version (e.g. "dev", for which IsNewer always answers false).</summary>
        private static string NewerThan(string currentTag)
        {
            string normalized = UpdateCheck.NormalizeTag(currentTag);
            if (string.IsNullOrEmpty(normalized))
                return null;
            string majorText = normalized.Split('.')[0];
            if (!long.TryParse(majorText, NumberStyles.None, CultureInfo.InvariantCulture, out long major))
                return null;
            return "v" + (major + 1) + ".0.0";
        }
    }
}

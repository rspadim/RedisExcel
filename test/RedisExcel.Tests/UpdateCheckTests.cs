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
    }
}

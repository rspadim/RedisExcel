using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Offline checks for the SyncWrite mode helpers: the fire-and-forget
    /// decision per mode and the exact marker texts returned instead of a
    /// reply (the 24 call sites use these helpers, so pinning the helpers
    /// pins the mode semantics).
    /// </summary>
    public class SyncWriteModeTests
    {
        [Theory]
        [InlineData("sync", false, false)]
        [InlineData("sync", true, false)]
        [InlineData("fireforget", false, true)]
        [InlineData("fireforget", true, false)]
        [InlineData("fireforget-all", false, true)]
        [InlineData("fireforget-all", true, true)]
        public void ShouldFireAndForget_FollowsTheMode(string mode, bool replyDependent, bool expected)
        {
            string original = RedisUDF.SyncWriteOverrideForTests;
            try
            {
                RedisUDF.SyncWriteOverrideForTests = mode;
                Assert.Equal(expected, RedisUDF.ShouldFireAndForget(replyDependent));
            }
            finally
            {
                RedisUDF.SyncWriteOverrideForTests = original;
            }
        }

        [Theory]
        [InlineData(false, "OK FireForget")]
        [InlineData(true, "OK-FireForgetAll")]
        public void FireAndForgetMarker_IsExact(bool replyDependent, string expected)
        {
            Assert.Equal(expected, RedisUDF.FireAndForgetMarker(replyDependent));
        }
    }
}

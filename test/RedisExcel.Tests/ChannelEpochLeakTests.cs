using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The per-channel epoch map detects an unsubscribe racing a ChannelLatest
    /// install. It is dead weight for a channel with no live listener, so it is
    /// pruned opportunistically. The audit (round L) found the prune trigger was
    /// the KEY's own epoch counter, which never reaches 1024 when every channel is
    /// distinct - so the map grew without bound. These pin the global-bump prune.
    /// </summary>
    public class ChannelEpochLeakTests
    {
        private static readonly FieldInfo Epochs = typeof(RedisUDF)
            .GetField("_channelEpochs", BindingFlags.Static | BindingFlags.NonPublic);
        private static readonly MethodInfo Bump = typeof(RedisUDF)
            .GetMethod("BumpChannelEpoch", BindingFlags.Static | BindingFlags.NonPublic);

        private static int EpochCount()
        {
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            return dict.Count;
        }

        [Fact]
        public void DistinctChannelChurn_DoesNotGrowTheEpochMapUnboundedly()
        {
            Assert.NotNull(Epochs);
            Assert.NotNull(Bump);
            // No live listeners (a fresh process/test), so every entry is prunable.
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            dict.Clear();

            const int channels = 5000;
            for (int i = 0; i < channels; i++)
                Bump.Invoke(null, new object[] { "epoch:leak:" + i });

            // The global-bump prune sweeps every 1024 bumps, so after 5000 distinct
            // channels the map must be far below 5000 (it was exactly 5001 before
            // the fix). Allow one sweep's slack plus the current window.
            int count = EpochCount();
            Assert.True(count < 2048,
                "the epoch map must be pruned (bounded), got " + count + " entries for " + channels + " channels");
            dict.Clear();
        }

        [Fact]
        public void ActuallyBumpsTheEpochPerKey()
        {
            // Sanity: the value still increments (the racing-install detection relies
            // on it), independent of the prune change.
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            dict.Clear();
            string key = "epoch:bump:" + System.Guid.NewGuid().ToString("N");
            long first = (long)Bump.Invoke(null, new object[] { key });
            long second = (long)Bump.Invoke(null, new object[] { key });
            Assert.Equal(1, first);
            Assert.Equal(2, second);
            dict.Clear();
        }
    }
}

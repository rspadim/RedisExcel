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
        private static readonly MethodInfo BeginWatch = typeof(RedisUDF)
            .GetMethod("BeginSubscribeWatch", BindingFlags.Static | BindingFlags.NonPublic);
        private static readonly MethodInfo EndWatch = typeof(RedisUDF)
            .GetMethod("EndSubscribeWatch", BindingFlags.Static | BindingFlags.NonPublic);

        private static int EpochCount()
        {
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            return dict.Count;
        }

        private static bool HasEpoch(string key)
        {
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            return dict.Contains(key);
        }

        // One sweep fires per 1024 global bumps, and any 1024 consecutive bumps
        // contain exactly one multiple of 1024, so this guarantees a sweep.
        private static void ForceSweep(string prefix)
        {
            for (int i = 0; i < 1024; i++)
                Bump.Invoke(null, new object[] { prefix + ":" + i });
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

        [Fact]
        public void EpochSweep_DoesNotPruneInFlightSubscribe()
        {
            // Audit round M: the sweep used to delete any entry with no live
            // listener, INCLUDING one for a ChannelLatest subscribe that had read
            // epochBefore but not yet installed. Erasing it would make the install
            // re-check (ChannelEpoch == epochBefore) wrongly pass (0 == 0) and leak
            // a ghost listener. The in-flight watcher must veto the prune; without
            // it the key would be gone after the sweep.
            Assert.NotNull(BeginWatch);
            Assert.NotNull(EndWatch);
            var dict = (System.Collections.IDictionary)Epochs.GetValue(null);
            dict.Clear();

            string key = "epoch:watch:" + System.Guid.NewGuid().ToString("N");
            Bump.Invoke(null, new object[] { key });          // a racing unsubscribe/reset
            BeginWatch.Invoke(null, new object[] { key });     // subscribe in flight
            ForceSweep("epoch:watch:sweep");                   // must NOT sweep this key
            Assert.True(HasEpoch(key),
                "a key with an in-flight subscribe must survive the epoch sweep");

            EndWatch.Invoke(null, new object[] { key });       // subscribe resolved
            ForceSweep("epoch:watch:sweep2");                  // now prunable (no listener)
            Assert.False(HasEpoch(key),
                "once the subscribe resolves with no listener, the key is prunable");
            dict.Clear();
        }
    }
}

using System;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Real-time conflation window: the latest value of a topic reaches Excel at
    /// most once per window, 0 disables it (per-message push), and an absent
    /// <c>ConflationMs</c> inherits the legacy <c>CoalesceRealtimeUpdates</c>
    /// behaviour (true = the Excel tick interval).
    /// </summary>
    public class ConflationTests
    {
        private const long Ms = TimeSpan.TicksPerMillisecond;

        [Theory]
        [InlineData(0, 0)]          // 0 = off
        [InlineData(-5, 0)]         // negative sanitizes to off
        [InlineData(200, 200)]
        [InlineData(5000, 5000)]
        public void Sanitize_ClampsToValidRange(int input, int expected)
        {
            Assert.Equal(expected, Conflation.Sanitize(input));
        }

        [Fact]
        public void Sanitize_CapsAtMaxWindow()
        {
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Sanitize(Conflation.MaxWindowMs + 1));
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Sanitize(int.MaxValue));
        }

        [Fact]
        public void Resolve_ExplicitValueWins()
        {
            Assert.Equal(200, Conflation.Resolve(200, coalesceRealtimeUpdates: true, excelUpdateRateMs: 100));
            Assert.Equal(0, Conflation.Resolve(0, coalesceRealtimeUpdates: true, excelUpdateRateMs: 100));
        }

        [Fact]
        public void Resolve_AbsentFallsBackToTheLegacyBoolean()
        {
            // CoalesceRealtimeUpdates=true keeps the pre-existing window (the Excel tick).
            Assert.Equal(100, Conflation.Resolve(null, coalesceRealtimeUpdates: true, excelUpdateRateMs: 100));
            // ...=false keeps the pre-existing per-message push.
            Assert.Equal(0, Conflation.Resolve(null, coalesceRealtimeUpdates: false, excelUpdateRateMs: 100));
        }

        [Fact]
        public void Resolve_AbsentWithNonPositiveTickIsOff()
        {
            Assert.Equal(0, Conflation.Resolve(null, coalesceRealtimeUpdates: true, excelUpdateRateMs: 0));
        }

        [Fact]
        public void IsDue_NoWindow_AlwaysPushes()
        {
            Assert.True(Conflation.IsDue(lastPushTicks: 1_000_000, nowTicks: 1_000_000, windowMs: 0));
        }

        [Fact]
        public void IsDue_NeverPushed_AlwaysPushes()
        {
            Assert.True(Conflation.IsDue(lastPushTicks: 0, nowTicks: 12345, windowMs: 200));
        }

        [Fact]
        public void IsDue_InsideWindow_Waits()
        {
            long last = 1000 * Ms;
            Assert.False(Conflation.IsDue(last, last + 199 * Ms, 200));
        }

        [Fact]
        public void IsDue_AtOrPastWindow_Pushes()
        {
            long last = 1000 * Ms;
            Assert.True(Conflation.IsDue(last, last + 200 * Ms, 200));
            Assert.True(Conflation.IsDue(last, last + 5000 * Ms, 200));
        }

        [Fact]
        public void IsDue_ClockMovedBackwards_DoesNotStall()
        {
            // A backwards clock (NTP/manual change) must never freeze the cell.
            Assert.True(Conflation.IsDue(lastPushTicks: 5000 * Ms, nowTicks: 1000 * Ms, windowMs: 200));
        }
    }
}

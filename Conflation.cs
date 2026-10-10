using System;

namespace RedisExcel
{
    /// <summary>
    /// Real-time delivery conflation window. A topic that received a newer value
    /// keeps only the latest one and pushes it to Excel at most once per window
    /// (the latest value wins when the window elapses), instead of one push per
    /// incoming message. This cuts Excel repaints and the recalculation cascade
    /// of dependent formulas on fast feeds.
    ///
    /// Configuration: <c>ConflationMs</c> (root) enables the window explicitly
    /// (0 = off, per-message push). When it is absent, the legacy
    /// <c>CoalesceRealtimeUpdates</c> boolean decides:
    /// true = the Excel tick interval (the pre-existing coalescing behaviour),
    /// false = 0 (no conflation).
    /// </summary>
    internal static class Conflation
    {
        /// <summary>Largest accepted window (one hour); larger values are clamped.</summary>
        internal const int MaxWindowMs = 3600000;

        /// <summary>Sanitizes a configured window: negative becomes 0 (off).</summary>
        internal static int Sanitize(int windowMs)
        {
            if (windowMs <= 0)
                return 0;
            return windowMs > MaxWindowMs ? MaxWindowMs : windowMs;
        }

        /// <summary>
        /// Effective window in milliseconds. An explicit <paramref name="configuredMs"/>
        /// wins (already sanitized); otherwise the legacy boolean maps to the Excel
        /// tick interval (on) or 0 (off), preserving the previous behaviour.
        /// </summary>
        internal static int Resolve(int? configuredMs, bool coalesceRealtimeUpdates, int excelUpdateRateMs)
        {
            if (configuredMs.HasValue)
                return Sanitize(configuredMs.Value);
            if (!coalesceRealtimeUpdates)
                return 0;
            return excelUpdateRateMs > 0 ? excelUpdateRateMs : 0;
        }

        /// <summary>
        /// Whether a push is due for a topic whose last push happened at
        /// <paramref name="lastPushTicks"/>. A window &lt;= 0 always pushes
        /// (conflation off) and a topic that never pushed always pushes. A clock
        /// that moved backwards never stalls a topic (it pushes).
        /// </summary>
        internal static bool IsDue(long lastPushTicks, long nowTicks, int windowMs)
        {
            if (windowMs <= 0)
                return true;
            if (lastPushTicks == 0)
                return true;
            long elapsedTicks = nowTicks - lastPushTicks;
            if (elapsedTicks < 0)
                return true; // clock moved backwards: do not stall the cell
            return elapsedTicks >= windowMs * TimeSpan.TicksPerMillisecond;
        }
    }
}

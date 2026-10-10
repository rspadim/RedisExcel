using System;
using System.Threading;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Shared assertions for the offline async-queue tests.
    /// </summary>
    internal static class AsyncQueueAssert
    {
        /// <summary>
        /// Waits until the idle per-host queue was removed. The release
        /// continuation runs after the item completes, so this is a bounded
        /// spin rather than an immediate read.
        /// </summary>
        public static void Removed(string host)
        {
            Assert.True(
                SpinWait.SpinUntil(() => !RedisUdfAsync.HasQueueForTests(host), 5000),
                "The idle per-host queue was not removed.");
        }
    }
}

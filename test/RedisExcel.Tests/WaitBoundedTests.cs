using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Unit tests for <see cref="RedisUDF.WaitBounded"/>: the hard bound around
    /// a pipelined task wait. A completed task returns true; a faulted task
    /// also returns true so the caller's GetResult can surface the original
    /// failure; a task that never completes returns false at the bound instead
    /// of blocking the calling thread forever.
    /// </summary>
    public class WaitBoundedTests
    {
        [Fact]
        public void WaitBounded_CompletedTask_ReturnsTrue()
        {
            Assert.True(RedisUDF.WaitBounded(Task.CompletedTask, 200));
        }

        [Fact]
        public async Task WaitBounded_AlreadyFaultedTask_ReturnsTrueAndKeepsTheFailure()
        {
            var failure = new InvalidOperationException("pipeline failed");
            Task task = Task.FromException(failure);

            // A faulted task counts as done: true, while the failure stays
            // observed on the task (awaiting it throws the original exception,
            // the GetResult contract the callers rely on).
            Assert.True(RedisUDF.WaitBounded(task, 200));

            var thrown = await Assert.ThrowsAsync<InvalidOperationException>(() => task);
            Assert.Same(failure, thrown);
        }

        [Fact]
        public void WaitBounded_NeverCompletingTask_ReturnsFalseAtTheBound()
        {
            var tcs = new TaskCompletionSource<object>(
                TaskCreationOptions.RunContinuationsAsynchronously);

            var watch = Stopwatch.StartNew();
            bool completed = RedisUDF.WaitBounded(tcs.Task, 200);
            watch.Stop();

            Assert.False(completed, "A task that never completes must report the timeout.");
            // The wait itself is bounded by the 200 ms timeout; the elapsed can
            // still overshoot under a loaded machine (thread scheduling), so the
            // assertion only pins "it really waited, it did not return early".
            // A hanging implementation is caught by the test's overall timeout.
            Assert.InRange(watch.ElapsedMilliseconds, 100, 20000);

            tcs.SetResult(null);
        }
    }
}

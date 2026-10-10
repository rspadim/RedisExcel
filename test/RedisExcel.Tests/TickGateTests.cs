using System;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace RedisExcel.Tests
{
    public class TickGateTests
    {
        [Fact]
        public void TryEnter_IsExclusive()
        {
            var gate = new TickGate();

            Assert.True(gate.TryEnter());
            Assert.False(gate.TryEnter());

            gate.Exit();
            Assert.True(gate.TryEnter());
        }

        [Fact]
        public void Exit_WhenIdle_KeepsGateFree()
        {
            var gate = new TickGate();

            // Documents the idle (never-entered) behavior only: Exit on a free gate
            // leaves it free. This is not a general guarantee - Exit while another
            // owner holds the gate would reopen the section early.
            gate.Exit();
            gate.Exit();

            Assert.True(gate.TryEnter());
        }

        [Fact]
        public void TryEnter_RejectedWhileHeld()
        {
            var gate = new TickGate();

            // The gate is held for the whole parallel burst: no attempt may win.
            Assert.True(gate.TryEnter());

            int winners = 0;
            Parallel.For(0, 64, _ =>
            {
                if (gate.TryEnter())
                {
                    Interlocked.Increment(ref winners);
                }
            });

            Assert.Equal(0, winners);

            // Once released, exactly the next attempt can win.
            gate.Exit();
            Assert.True(gate.TryEnter());
        }

        [Fact]
        public void TryEnter_SingleWinnerWhenGateFree()
        {
            var gate = new TickGate();

            int winners = 0;
            Parallel.For(0, 64, _ => { if (gate.TryEnter()) Interlocked.Increment(ref winners); });
            Assert.Equal(1, winners);
        }

        [Fact]
        public void TryEnter_ReleasedWhenGuardedRegionThrows()
        {
            var gate = new TickGate();

            // Timer callbacks guard their body with try/finally (see
            // RedisRtd.CreateTimer): a throwing tick must release the gate, or
            // every later tick would be skipped forever.
            var failure = Assert.Throws<InvalidOperationException>(() => RunTick(gate));
            Assert.Equal("tick failed", failure.Message);

            Assert.True(gate.TryEnter());
            gate.Exit();
        }

        private static void RunTick(TickGate gate)
        {
            if (!gate.TryEnter())
                throw new InvalidOperationException("the gate was already held");

            try
            {
                Assert.False(gate.TryEnter()); // exclusive while the tick runs
                throw new InvalidOperationException("tick failed");
            }
            finally
            {
                gate.Exit();
            }
        }
    }
}

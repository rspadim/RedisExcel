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
        public void Exit_WithoutEnter_KeepsGateFree()
        {
            var gate = new TickGate();

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
    }
}

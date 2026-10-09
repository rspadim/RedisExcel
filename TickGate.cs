using System.Threading;

namespace RedisExcel
{
    /// <summary>
    /// Non-blocking reentrancy gate for timer callbacks: the first caller enters,
    /// concurrent callers are rejected until Exit() is called. Used so a slow
    /// tick is skipped instead of overlapping the next one.
    /// </summary>
    internal sealed class TickGate
    {
        private int _busy;

        /// <summary>TRUE when the gate was free and is now held by the caller.</summary>
        public bool TryEnter()
        {
            return Interlocked.CompareExchange(ref _busy, 1, 0) == 0;
        }

        /// <summary>
        /// Releases the gate. Only call it after a successful TryEnter by the
        /// current holder: calling Exit while another owner holds the gate would
        /// reopen the section early.
        /// </summary>
        public void Exit()
        {
            Interlocked.Exchange(ref _busy, 0);
        }
    }
}

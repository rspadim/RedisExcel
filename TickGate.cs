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

    /// <summary>
    /// Fixed set of lock stripes selected by key hash: callers serialize
    /// per-key work without a lock per key. The stripe count is fixed at
    /// construction, so the same key always maps to the same stripe.
    /// </summary>
    internal sealed class StripedLocks
    {
        private readonly object[] _stripes;

        public StripedLocks(int count)
        {
            _stripes = new object[count];
            for (int i = 0; i < _stripes.Length; i++)
                _stripes[i] = new object();
        }

        /// <summary>Stripe for the key: its hash masked non-negative, modulo the count.</summary>
        public object For(string key)
        {
            return _stripes[(key.GetHashCode() & 0x7FFFFFFF) % _stripes.Length];
        }
    }
}

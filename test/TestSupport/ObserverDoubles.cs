using System;
using System.Collections.Generic;
using System.Threading;
using ExcelDna.Integration;

namespace RedisExcel.TestSupport
{
    /// <summary>
    /// IExcelObserver test double recording OnNext values, OnError exceptions
    /// and the completion count, with a wait handle. Shared by the unit and
    /// liveness suites (linked source, see the Compile items in their csproj).
    /// </summary>
    public sealed class RecordingObserver : IExcelObserver
    {
        private readonly object _gate = new object();
        private readonly List<object> _values = new List<object>();
        private readonly List<Exception> _errors = new List<Exception>();
        private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
        private int _completedCount;

        public void OnNext(object value)
        {
            lock (_gate)
                _values.Add(value);
        }

        public void OnError(Exception exception)
        {
            lock (_gate)
                _errors.Add(exception);
        }

        public void OnCompleted()
        {
            lock (_gate)
                _completedCount++;
            _completed.Set();
        }

        public object[] Values
        {
            get { lock (_gate) return _values.ToArray(); }
        }

        public Exception[] Errors
        {
            get { lock (_gate) return _errors.ToArray(); }
        }

        public int CompletedCount
        {
            get { lock (_gate) return _completedCount; }
        }

        public bool WaitCompleted(int timeoutMs = 5000) => _completed.Wait(timeoutMs);
    }

    /// <summary>
    /// IExcelObserver test double whose OnNext always throws: pins that the
    /// delivery path swallows an observer failure and still completes.
    /// </summary>
    public sealed class OnNextThrowingObserver : IExcelObserver
    {
        private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);
        private int _nextCalls;
        private int _completedCount;

        public void OnNext(object value)
        {
            Interlocked.Increment(ref _nextCalls);
            throw new InvalidOperationException("observer rejected the value");
        }

        public void OnError(Exception exception)
        {
        }

        public void OnCompleted()
        {
            Interlocked.Increment(ref _completedCount);
            _completed.Set();
        }

        public int OnNextCalls => Volatile.Read(ref _nextCalls);

        public int CompletedCount => Volatile.Read(ref _completedCount);

        public bool WaitCompleted(int timeoutMs = 5000) => _completed.Wait(timeoutMs);
    }
}

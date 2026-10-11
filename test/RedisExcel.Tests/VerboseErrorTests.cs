using System;
using System.Reflection;
using StackExchange.Redis;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The cell text of a RUNTIME failure (unreachable host, timeout, wrong
    /// Redis type) is verbose: the operation, the identifiers it acted on, the
    /// cause and a short hint. Argument-validation failures stay terse and are
    /// pinned by UdfArgumentValidationTests instead. The private runtime helper
    /// is driven directly with synthesized exceptions, so the test is
    /// deterministic and independent of the process-wide connection cache.
    /// </summary>
    public class VerboseErrorTests
    {
        private static readonly MethodInfo VerboseError =
            typeof(RedisUDF).GetMethod("VerboseError", BindingFlags.Static | BindingFlags.NonPublic);

        private static string Format(Exception ex, string context)
        {
            Assert.NotNull(VerboseError);
            return (string)VerboseError.Invoke(null, new object[] { ex, context });
        }

        [Fact]
        public void ConnectionFailure_ReturnsContextCauseAndUnreachableHint()
        {
            var ex = new RedisConnectionException(ConnectionFailureType.UnableToConnect, "boom");
            string text = Format(ex, "key=k, host=dead:1");

            Assert.Equal("Error: key=k, host=dead:1: boom | host unreachable or wrong port; check the host / RedisExcel.json", text);
        }

        [Fact]
        public void Timeout_ReturnsTimeoutHint()
        {
            var ex = new RedisTimeoutException("Timeout performing INCR", CommandStatus.Unknown);
            string text = Format(ex, "key=k, host=h");

            Assert.Contains("Redis did not answer in time", text);
            Assert.Contains("raise UDF.timeout", text);
        }

        [Fact]
        public void WrongType_ReturnsTypeHint()
        {
            // A server-side WRONGTYPE reaches the cell as a plain message, not a
            // typed exception, so the hint is driven by the text.
            var ex = new Exception("WRONGTYPE Operation against a key holding the wrong kind of value");
            string text = Format(ex, "key=k, host=h"); 

            Assert.Contains("already holds a different Redis type", text);
        }

        [Fact]
        public void SelfExplanatoryCause_AppendsNoHint()
        {
            string text = Format(new Exception("some plain failure"), "key=k, host=h");

            Assert.Equal("Error: key=k, host=h: some plain failure", text);
        }

        [Fact]
        public void LongCauseAndContext_AreTruncated()
        {
            string text = Format(new Exception(new string('x', 500)), new string('c', 500));

            Assert.Contains("...", text);
            Assert.True(text.Length < 500, "the cell text must stay bounded");
        }
    }
}

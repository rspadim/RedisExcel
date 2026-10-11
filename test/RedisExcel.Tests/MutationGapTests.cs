using System;
using System.Reflection;
using StackExchange.Redis;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Pins production lines the mutation audit showed were unverified: helpers
    /// whose exact value/branch no test asserted, so a mutation of the line did
    /// not fail the suite. Each test here kills at least one such mutant.
    /// </summary>
    [Collection("write-mode-seam")]
    public class MutationGapTests
    {
        // ---- ResponseTimeoutMs: pool selection, slack, clamps, fallback -------

        [Fact]
        public void ResponseTimeoutMs_UsesThePoolTimeoutPlus500()
        {
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot());
                cfg.RTD.timeout = 3000;
                cfg.UDF.timeout = 7000;
                AppConfig.CurrentOverrideForTests = cfg;

                // The pool dimension matters: UDF uses UDF.timeout, RTD uses RTD.timeout.
                Assert.Equal(7500, RedisUDF.ResponseTimeoutMs(RedisPool.UdfData));
                Assert.Equal(3500, RedisUDF.ResponseTimeoutMs(RedisPool.RtdData));
                Assert.Equal(3500, RedisUDF.ResponseTimeoutMs(RedisPool.RtdSub));
                // The no-arg overload defaults to the RTD pool.
                Assert.Equal(3500, RedisUDF.ResponseTimeoutMs());
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = original;
            }
        }

        [Fact]
        public void ResponseTimeoutMs_ClampsTheOverflowBound()
        {
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot());
                cfg.UDF.timeout = int.MaxValue; // +500 would overflow a 32-bit int
                AppConfig.CurrentOverrideForTests = cfg;

                int bound = RedisUDF.ResponseTimeoutMs(RedisPool.UdfData);
                Assert.True(bound > 0, "the bound must never go negative on overflow, got " + bound);
                Assert.Equal(int.MaxValue, bound); // clamped to int.MaxValue, not wrapped
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = original;
            }
        }

        [Fact]
        public void ResponseTimeoutMs_HasAMinimumOf500()
        {
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot());
                cfg.UDF.timeout = 1; // 1 + 500 = 501, above the 500 floor
                AppConfig.CurrentOverrideForTests = cfg;
                Assert.Equal(501, RedisUDF.ResponseTimeoutMs(RedisPool.UdfData));

                cfg.UDF.timeout = -1000; // would be below the floor -> 500
                Assert.Equal(500, RedisUDF.ResponseTimeoutMs(RedisPool.UdfData));
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = original;
            }
        }

        // ---- ReaderCountText: singular/plural ---------------------------------

        [Theory]
        [InlineData(0, "0 readers")]
        [InlineData(1, "1 reader")]
        [InlineData(2, "2 readers")]
        [InlineData(100, "100 readers")]
        public void ReaderCountText_IsExact(long readers, string expected)
        {
            var method = typeof(RedisUDF).GetMethod("ReaderCountText", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);
            Assert.Equal(expected, (string)method.Invoke(null, new object[] { readers }));
        }

        // ---- VerboseError: credential masking + truncation of cause/context ---

        private static string VerboseError(Exception ex, string context)
        {
            var method = typeof(RedisUDF).GetMethod("VerboseError", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);
            return (string)method.Invoke(null, new object[] { ex, context });
        }

        [Fact]
        public void VerboseError_MasksCredentialsInsideTheCause()
        {
            // Defense in depth: even if a call site forgets to mask, the cause text
            // (an exception message embedding a connection string) is masked here.
            string text = VerboseError(
                new Exception("connect failed host:6379,password=topsecret,ssl=true"),
                "key=k");
            Assert.DoesNotContain("topsecret", text);
            Assert.Contains("password=****", text);
        }

        [Fact]
        public void VerboseError_MasksCredentialsInsideTheContext()
        {
            string text = VerboseError(new Exception("boom"), "host=h:6379,password=ctxsecret");
            Assert.DoesNotContain("ctxsecret", text);
            Assert.Contains("password=****", text);
        }

        [Fact]
        public void VerboseError_TruncatesLongCauseAndContext()
        {
            string cause = new string('c', 500);
            string context = new string('x', 500);
            string text = VerboseError(new Exception(cause), context);
            Assert.True(text.Length < 500, "the cell text must stay bounded, got " + text.Length);
            Assert.Contains("...", text);
        }

        [Fact]
        public void VerboseError_BlankContextUsesAPlaceholder()
        {
            string text = VerboseError(new Exception("boom"), "   ");
            Assert.Contains("(no context)", text);
        }

        // ---- Connect-failure memo: window + pool dimension --------------------

        [Fact]
        public void ConnectFailureMemo_HasTheDocumentedWindow()
        {
            var field = typeof(RedisConnectionManager)
                .GetField("ConnectFailureMemoMs", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(field);
            Assert.Equal(2000, (int)field.GetRawConstantValue());
        }

        [Fact]
        public void ConnectionPoolKey_IncludesThePoolDimension()
        {
            // The memo key must distinguish pools: a failure for RtdSub must not
            // fail-fast a UdfData connect to the same host.
            var method = typeof(RedisConnectionManager)
                .GetMethod("PoolKey", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);
            string a = (string)method.Invoke(null, new object[] { "h:6379", RedisPool.UdfData });
            string b = (string)method.Invoke(null, new object[] { "h:6379", RedisPool.RtdSub });
            string c = (string)method.Invoke(null, new object[] { "h:6379", RedisPool.RtdData });
            Assert.NotEqual(a, b);
            Assert.NotEqual(b, c);
            Assert.NotEqual(a, c);
            // Same host+pool is stable (used as a dictionary key).
            Assert.Equal(a, (string)method.Invoke(null, new object[] { "h:6379", RedisPool.UdfData }));
        }
    }
}

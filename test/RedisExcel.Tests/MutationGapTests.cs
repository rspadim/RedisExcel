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
        // ---- ResponseTimeoutMs: unified server syncTimeout + slack, clamps -----

        [Fact]
        public void ResponseTimeoutMs_UsesTheServerSyncTimeoutPlus500()
        {
            var original = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new ConfigRoot());
                cfg.Default = "h";
                cfg.Servers = new System.Collections.Generic.Dictionary<string, ServerDef>
                {
                    ["h"] = new ServerDef { host = "h", syncTimeout = 7000 }
                };
                AppConfig.CurrentOverrideForTests = cfg;

                // The unified server timeout (no pool dimension): 7000 + 500.
                Assert.Equal(7500, RedisUDF.ResponseTimeoutMs(RedisPool.UdfData));
                Assert.Equal(7500, RedisUDF.ResponseTimeoutMs(RedisPool.RtdData));
                Assert.Equal(7500, RedisUDF.ResponseTimeoutMs());
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
                cfg.Default = "h";
                cfg.Servers = new System.Collections.Generic.Dictionary<string, ServerDef>
                {
                    ["h"] = new ServerDef { host = "h", syncTimeout = int.MaxValue }
                };
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
                // A tiny syncTimeout (1) -> 1 + 500 = 501.
                var cfg = AppConfig.Sanitize(new ConfigRoot());
                cfg.Default = "h";
                cfg.Servers = new System.Collections.Generic.Dictionary<string, ServerDef>
                {
                    ["h"] = new ServerDef { host = "h", syncTimeout = 1 }
                };
                AppConfig.CurrentOverrideForTests = cfg;
                Assert.Equal(501, RedisUDF.ResponseTimeoutMs(RedisPool.UdfData));
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

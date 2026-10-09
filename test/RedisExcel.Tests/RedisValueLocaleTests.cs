using System;
using System.Globalization;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    public class RedisValueLocaleTests
    {
        [Fact]
        public void ToRedisString_UsesInvariantCultureForNumbers()
        {
            // Force a comma-decimal culture: the fact only proves invariance if it
            // would fail when the implementation falls back to CurrentCulture.
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");

                Assert.Equal("67000.5", RedisUDF.ToRedisString(67000.5));
                Assert.Equal("0.1", RedisUDF.ToRedisString(0.1));
                Assert.Equal("2", RedisUDF.ToRedisString(2.0));
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void ToRedisString_PassesThroughStringsAndNulls()
        {
            Assert.Equal("abc", RedisUDF.ToRedisString("abc"));
            Assert.Null(RedisUDF.ToRedisString(null));
        }

        [Fact]
        public void ToRedisString_ReturnsNullForExcelSentinels()
        {
            Assert.Null(RedisUDF.ToRedisString(ExcelMissing.Value));
            Assert.Null(RedisUDF.ToRedisString(ExcelEmpty.Value));
            Assert.Null(RedisUDF.ToRedisString(ExcelError.ExcelErrorValue));
        }

        [Fact]
        public void ToRedisString_FormatsBooleansInvariantly()
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");
                Assert.Equal("True", RedisUDF.ToRedisString(true));
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void ToInt64Invariant_ParsesInvariantNumbers()
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");

                Assert.Equal(5L, RedisUDF.ToInt64Invariant(5.0));
                Assert.Equal(7L, RedisUDF.ToInt64Invariant("7"));

                // Under de-DE a culture-sensitive parse would accept "1,5";
                // the invariant implementation must reject it.
                Assert.Throws<FormatException>(() => RedisUDF.ToInt64Invariant("1,5"));
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }
    }
}

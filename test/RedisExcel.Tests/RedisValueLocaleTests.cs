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
        }

        [Fact]
        public void ToRedisString_FormatsBooleansInvariantly()
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");
                Assert.Equal("true", RedisUDF.ToRedisString(true));
                Assert.Equal("false", RedisUDF.ToRedisString(false));
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void ToRedisString_MultiCellRange_Throws()
        {
            var range = new object[,] { { 1, 2 } };
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToRedisString(range));
            Assert.Equal("A multi-cell range is not a valid scalar argument", ex.Message);
        }

        [Fact]
        public void ToRedisString_ExcelError_Throws()
        {
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToRedisString(ExcelError.ExcelErrorNA));
            Assert.Equal("Excel error cells are not valid arguments", ex.Message);
        }

        [Fact]
        public void ToRedisString_DoublesRoundTrip()
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");

                // G15 is not enough for these; the implementation must fall back to
                // G17 so the invariant parse returns the exact same double.
                foreach (var value in new[] { 0.84551240822557006, double.MaxValue })
                {
                    var result = RedisUDF.ToRedisString(value);
                    Assert.Equal(value, double.Parse(result, CultureInfo.InvariantCulture));
                }
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
                // the invariant implementation must reject it with a stable message.
                var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant("1,5"));
                Assert.Equal("numeric argument is not valid", ex.Message);
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void ToInt64Invariant_ExcelError_Throws()
        {
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(ExcelError.ExcelErrorNA));
            Assert.Equal("Excel error cells are not valid numeric arguments", ex.Message);
        }

        [Fact]
        public void ToInt64Invariant_OutOfRange_Throws()
        {
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(1e19));
            Assert.Equal("numeric argument is out of range", ex.Message);
        }

        [Fact]
        public void ToInt64Invariant_MultiCellRange_Throws()
        {
            var range = new object[,] { { 1, 2 } };
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(range));
            Assert.Equal("A multi-cell range is not a valid numeric argument", ex.Message);
        }
    }
}

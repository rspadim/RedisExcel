using System;
using System.Globalization;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    public class RedisValueLocaleTests
    {
        /// <summary>
        /// Runs <paramref name="test"/> under the given culture and restores the
        /// original one afterwards.
        /// </summary>
        private static void WithCulture(string culture, Action test)
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo(culture);
                test();
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void ToRedisString_UsesInvariantCultureForNumbers()
        {
            // Force a comma-decimal culture: the fact only proves invariance if it
            // would fail when the implementation falls back to CurrentCulture.
            WithCulture("de-DE", () =>
            {
                Assert.Equal("67000.5", RedisUDF.ToRedisString(67000.5));
                Assert.Equal("0.1", RedisUDF.ToRedisString(0.1));
                Assert.Equal("2", RedisUDF.ToRedisString(2.0));
            });
        }

        [Fact]
        public void ToRedisString_PassesThroughStringsAndNulls()
        {
            Assert.Equal("abc", RedisUDF.ToRedisString("abc"));
            Assert.Null(RedisUDF.ToRedisString(null));
        }

        [Fact]
        public void ToRedisString_PassesThroughEmptyAndWhitespaceStrings()
        {
            // An explicit empty or whitespace-only cell is a valid Redis name,
            // not a missing argument.
            Assert.Equal("", RedisUDF.ToRedisString(""));
            Assert.Equal("   ", RedisUDF.ToRedisString("   "));
        }

        [Fact]
        public void ToRedisString_DateTime_UsesIso8601RoundTrip()
        {
            // "o" keeps DateTime values parseable for the JSON wrappers.
            var value = new DateTime(2026, 10, 9, 14, 30, 15, 250, DateTimeKind.Utc);
            Assert.Equal("2026-10-09T14:30:15.2500000Z", RedisUDF.ToRedisString(value));
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
            WithCulture("de-DE", () =>
            {
                Assert.Equal("true", RedisUDF.ToRedisString(true));
                Assert.Equal("false", RedisUDF.ToRedisString(false));
            });
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
            WithCulture("de-DE", () =>
            {
                // G15 is not enough for these; the implementation must fall back to
                // G17 so the invariant parse returns the exact same double.
                foreach (var value in new[] { 0.84551240822557006, double.MaxValue })
                {
                    var result = RedisUDF.ToRedisString(value);
                    Assert.Equal(value, double.Parse(result, CultureInfo.InvariantCulture));
                }
            });
        }

        [Fact]
        public void ToInt64Invariant_ParsesInvariantNumbers()
        {
            WithCulture("de-DE", () =>
            {
                Assert.Equal(5L, RedisUDF.ToInt64Invariant(5.0));
                Assert.Equal(7L, RedisUDF.ToInt64Invariant("7"));

                // Convert.ToInt64(string) parses with NumberStyles.Integer, which
                // never accepts a group separator, so "1,5" fails under any
                // culture; the stable error message is what this asserts.
                var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant("1,5"));
                Assert.Equal("numeric argument is not valid", ex.Message);
            });
        }

        [Fact]
        public void ToInt64Invariant_ExcelError_Throws()
        {
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(ExcelError.ExcelErrorNA));
            Assert.Equal("Excel error cells are not valid numeric arguments", ex.Message);
        }

        [Fact]
        public void ToInt64Invariant_RejectsBooleansAndFractions()
        {
            var boolEx = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(true));
            Assert.Equal("numeric argument is not valid", boolEx.Message);

            var doubleEx = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(2.5));
            Assert.Equal("numeric argument is not an integer", doubleEx.Message);

            var decimalEx = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(1.5m));
            Assert.Equal("numeric argument is not an integer", decimalEx.Message);

            Assert.Equal(2L, RedisUDF.ToInt64Invariant(2.0));
        }

        [Fact]
        public void ToInt64Invariant_OutOfRange_Throws()
        {
            var ex = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(1e19));
            Assert.Equal("numeric argument is out of range", ex.Message);
        }

        [Fact]
        public void ToInt64Invariant_NullAndExcelSentinels_Throw()
        {
            // null is not reachable from an Excel cell (a blank cell arrives as
            // ExcelEmpty), but the helper contract mirrors ToRedisString.
            var nullEx = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(null));
            Assert.Equal("numeric argument is not valid", nullEx.Message);

            var missingEx = Assert.Throws<ArgumentException>(() => RedisUDF.ToInt64Invariant(ExcelMissing.Value));
            Assert.Equal("numeric argument is not valid", missingEx.Message);
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

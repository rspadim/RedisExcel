using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    public class RedisValueLocaleTests
    {
        [Fact]
        public void ToRedisString_UsesInvariantCultureForNumbers()
        {
            Assert.Equal("67000.5", RedisUDF.ToRedisString(67000.5));
            Assert.Equal("0.1", RedisUDF.ToRedisString(0.1));
            Assert.Equal("2", RedisUDF.ToRedisString(2.0));
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
        public void ToRedisString_UsesInvariantCultureForBooleans()
        {
            Assert.Equal("True", RedisUDF.ToRedisString(true));
        }

        [Fact]
        public void ToInt64Invariant_ParsesInvariantNumbers()
        {
            Assert.Equal(5L, RedisUDF.ToInt64Invariant(5.0));
            Assert.Equal(7L, RedisUDF.ToInt64Invariant("7"));
        }
    }
}

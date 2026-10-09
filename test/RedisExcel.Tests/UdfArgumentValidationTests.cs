using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// UDF argument guards that must reject invalid input before any Redis I/O;
    /// these tests run offline, with no Redis server involved.
    /// </summary>
    public class UdfArgumentValidationTests
    {
        [Fact]
        public void GetMultiple_NoNonBlankKeys_ReturnsNoValidKeyError()
        {
            var result = RedisUDF.RedisUDFGetMultiple(new object[,] { { "" } }, false, ExcelMissing.Value);
            Assert.Equal("Error: No valid key", (string)result[0, 0]);
        }

        [Fact]
        public void GetMultiple_ExcelErrorFlag_ReturnsExcelErrorArgumentMessage()
        {
            var result = RedisUDF.RedisUDFGetMultiple(new object[,] { { "k" } }, ExcelError.ExcelErrorNA, ExcelMissing.Value);
            Assert.Equal("Error: Excel error cells are not valid arguments", (string)result[0, 0]);
        }

        [Fact]
        public void GetMultiple_MultiCellFlag_ReturnsMultiCellRangeMessage()
        {
            var result = RedisUDF.RedisUDFGetMultiple(new object[,] { { "k" } }, new object[,] { { true } }, ExcelMissing.Value);
            Assert.Equal("Error: A multi-cell range is not a valid scalar argument", (string)result[0, 0]);
        }

        [Fact]
        public void GetMultiple_NonBooleanTextFlag_ReturnsFlagError()
        {
            var result = RedisUDF.RedisUDFGetMultiple(new object[,] { { "k" } }, "maybe", ExcelMissing.Value);
            Assert.Equal("Error: multipleColumns must be TRUE or FALSE", (string)result[0, 0]);
        }

        [Fact]
        public void SetKV_DifferentCellCounts_ReturnsCellCountMismatch()
        {
            var result = RedisUDF.RedisUDFSetKV(
                new object[,] { { "k1", "k2" } },
                new object[,] { { "v1" } },
                ExcelMissing.Value);
            Assert.Equal("Error: keys and values must have the same number of cells", result);
        }

        [Fact]
        public void SetKVPair_RangeWithoutPairShape_ReturnsRangeShapeError()
        {
            var result = RedisUDF.RedisUDFSetKVPair(
                new object[,] { { "k1", "v1", "extra" } },
                ExcelMissing.Value);
            Assert.Equal("Error: expected a range with 2 columns or 2 rows", result);
        }

        [Fact]
        public void HashSetMultiple_RangeWithoutPairShape_ReturnsRangeShapeError()
        {
            var result = RedisUDF.RedisUDFHashSetMultiple(
                "hash",
                new object[,] { { "f1", "v1", "extra" } },
                ExcelMissing.Value);
            Assert.Equal("Error: expected a range with 2 columns or 2 rows", result);
        }

        [Fact]
        public void ChannelUnsubscribe_BlankChannel_ReturnsChannelRequired()
        {
            var result = RedisUDF.RedisUDFChannelUnsubscribe("");
            Assert.Equal("Error: a channel is required", result);
        }

        [Fact]
        public void Keys_EmptyPattern_ReturnsKeyPatternRequiredMessage()
        {
            var result = RedisUDF.RedisUDFKeys("", ExcelMissing.Value);
            Assert.Equal("Error: a key pattern is required; use \"*\" to match all keys", (string)result[0, 0]);
        }

        [Fact]
        public void Get_NonTextHost_ReturnsHostMustBeTextMessage()
        {
            var result = RedisUDF.RedisUDFGet("k", 42);
            Assert.Equal("Error: host must be a text value", result);
        }
    }
}

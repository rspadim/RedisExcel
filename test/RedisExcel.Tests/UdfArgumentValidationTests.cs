using System;
using System.Collections.Generic;
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
        public void GetMultiple_NullKeysRange_ReturnsRangeRequired()
        {
            var result = RedisUDF.RedisUDFGetMultiple(null, false, ExcelMissing.Value);
            Assert.Equal("Error: a range is required", (string)result[0, 0]);
        }

        [Fact]
        public void ExistsMultiples_NullKeysRange_ReturnsRangeRequired()
        {
            var result = RedisUDF.RedisUDFExistsMultiples(null, ExcelMissing.Value);
            Assert.Equal("Error: a range is required", (string)result[0, 0]);
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
        public void SetKV_AllBlankKeys_ReturnsNoEntriesToWrite()
        {
            // ExcelEmpty is a truly blank cell; blank cells are skipped.
            var result = RedisUDF.RedisUDFSetKV(
                new object[,] { { ExcelEmpty.Value, ExcelEmpty.Value } },
                new object[,] { { "v1", "v2" } },
                ExcelMissing.Value);
            Assert.Equal("Error: no entries to write", result);
        }

        [Fact]
        public void SetKV_EmptyStringKey_IsKeptAsAnEntry()
        {
            // "" and whitespace-only cells are valid Redis names; only truly
            // missing cells (ExcelEmpty/ExcelMissing/null) are skipped.
            var entries = RedisUDF.CollectStringSetEntries(new[]
            {
                new KeyValuePair<object, object>("", "v1"),
                new KeyValuePair<object, object>("   ", "v2"),
                new KeyValuePair<object, object>(ExcelEmpty.Value, "v3"),
                new KeyValuePair<object, object>(ExcelMissing.Value, "v4")
            });

            Assert.Equal(2, entries.Count);
            Assert.Equal("", entries[0].Key.ToString());
            Assert.Equal("v1", entries[0].Value.ToString());
            Assert.Equal("   ", entries[1].Key.ToString());
            Assert.Equal("v2", entries[1].Value.ToString());
        }

        [Fact]
        public void HashSetMultiple_WhitespaceField_IsKeptAsAnEntry()
        {
            var entries = RedisUDF.CollectHashEntries(new[]
            {
                new KeyValuePair<object, object>(" ", "v1"),
                new KeyValuePair<object, object>(ExcelEmpty.Value, "v2")
            });

            Assert.Single(entries);
            Assert.Equal(" ", entries[0].Name.ToString());
            Assert.Equal("v1", entries[0].Value.ToString());
        }

        [Fact]
        public void SetKV_NullKeysRange_ReturnsRangeRequired()
        {
            var result = RedisUDF.RedisUDFSetKV(
                null,
                new object[,] { { "v1" } },
                ExcelMissing.Value);
            Assert.Equal("Error: a range is required", result);
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

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void ChannelUnsubscribe_BlankChannel_ReturnsChannelRequired(string channel)
        {
            // Omitted host (Excel passes ExcelMissing) = default host.
            var result = RedisUDF.RedisUDFChannelUnsubscribe(channel, ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Fact]
        public void ChannelPublish_EmptyChannel_ReturnsChannelRequired()
        {
            var result = RedisUDF.RedisUDFChannelPublish("", "x", ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Theory]
        [InlineData(" ")]
        [InlineData("\t")]
        public void ChannelPublish_WhitespaceChannel_ReturnsChannelRequired(string channel)
        {
            var result = RedisUDF.RedisUDFChannelPublish(channel, "x", ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Theory]
        [InlineData(" ")]
        [InlineData("\t")]
        public void ChannelPublishIfChanged_WhitespaceChannel_ReturnsChannelRequired(string channel)
        {
            var result = RedisUDF.RedisUDFChannelPublishIfChanged(channel, "x", ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Theory]
        [InlineData(" ")]
        [InlineData("\t")]
        public void ChannelLatest_WhitespaceChannel_ReturnsChannelRequired(string channel)
        {
            var result = RedisUDF.RedisUDFChannelLatest(channel, ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Fact]
        public void ChannelLatest_EmptyChannel_ReturnsChannelRequired()
        {
            var result = RedisUDF.RedisUDFChannelLatest("", ExcelMissing.Value);
            Assert.Equal("Error: a channel is required", result);
        }

        [Fact]
        public void Get_MissingKey_ReturnsKeyRequiredMessage()
        {
            // A truly blank cell must be rejected before any Redis I/O, with the
            // friendly message instead of StackExchange.Redis' "null key" error.
            var result = RedisUDF.RedisUDFGet(ExcelEmpty.Value, ExcelMissing.Value);
            Assert.Equal("Error: a key is required", result);
        }

        [Fact]
        public void HashGet_MissingField_ReturnsFieldRequiredMessage()
        {
            var result = RedisUDF.RedisUDFHashGet("hash", ExcelEmpty.Value, ExcelMissing.Value);
            Assert.Equal("Error: a field is required", result);
        }

        [Fact]
        public void HashSet_MissingHashKey_ReturnsHashKeyRequiredMessage()
        {
            var result = RedisUDF.RedisUDFHashSet(ExcelEmpty.Value, "field", "value", ExcelMissing.Value);
            Assert.Equal("Error: a hash key is required", result);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void SetEx_NonPositiveTtl_ReturnsTtlMustBePositive(int ttl)
        {
            // Must fail before any Redis I/O: this test runs with no server.
            var result = RedisUDF.RedisUDFSetEx("k", "v", ttl, ExcelMissing.Value);
            Assert.Equal("Error: ttl must be a positive number of seconds", result);
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

    /// <summary>Offline tests for the bounded LRU cache backing the
    /// PublishIfChanged dedup marker.</summary>
    public class PublishDedupCacheTests
    {
        [Fact]
        public void Set_OverCapacity_EvictsLeastRecentlyUsed()
        {
            var cache = new PublishDedupCache(2);
            cache.Set("a", "1");
            cache.Set("b", "2");
            Assert.True(cache.TryGet("a", out var valueA)); // refresh "a"
            Assert.Equal("1", valueA);

            cache.Set("c", "3");

            Assert.Equal(2, cache.Count);
            Assert.True(cache.TryGet("a", out var a));
            Assert.Equal("1", a);
            Assert.True(cache.TryGet("c", out var c));
            Assert.Equal("3", c);
            Assert.False(cache.TryGet("b", out _)); // "b" was the least recently used
        }

        [Fact]
        public void Set_ExistingKey_UpdatesValueWithoutGrowing()
        {
            var cache = new PublishDedupCache(1);
            cache.Set("a", "1");
            cache.Set("a", "2");

            Assert.Equal(1, cache.Count);
            Assert.True(cache.TryGet("a", out var value));
            Assert.Equal("2", value);
        }

        [Fact]
        public void Remove_DropsEntry()
        {
            var cache = new PublishDedupCache(4);
            cache.Set("a", "1");

            Assert.True(cache.Remove("a"));
            Assert.False(cache.TryGet("a", out _));
            Assert.False(cache.Remove("a"));
        }

        [Fact]
        public void Constructor_NonPositiveCapacity_Throws()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new PublishDedupCache(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => new PublishDedupCache(-1));
        }
    }
}

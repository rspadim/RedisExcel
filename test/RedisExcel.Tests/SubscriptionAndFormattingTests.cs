using Newtonsoft.Json.Linq;
using StackExchange.Redis;
using Xunit;

namespace RedisExcel.Tests
{
    public class SubscriptionKeyTests
    {
        [Fact]
        public void MakeKey_IsStableForSameInputs()
        {
            Assert.Equal(
                RedisSubscriptionManager.MakeKey("host:6379", "canal", false),
                RedisSubscriptionManager.MakeKey("host:6379", "canal", false));
        }

        [Fact]
        public void MakeKey_DistinguishesPatternMode()
        {
            Assert.NotEqual(
                RedisSubscriptionManager.MakeKey("host:6379", "canal", false),
                RedisSubscriptionManager.MakeKey("host:6379", "canal", true));
        }

        [Fact]
        public void MakeKey_DistinguishesHosts()
        {
            Assert.NotEqual(
                RedisSubscriptionManager.MakeKey("h1", "c", false),
                RedisSubscriptionManager.MakeKey("h2", "c", false));
        }

        [Fact]
        public void MakeKey_HostAndChannelBoundariesDoNotCollide()
        {
            Assert.NotEqual(
                RedisSubscriptionManager.MakeKey("a", "b|c", false),
                RedisSubscriptionManager.MakeKey("a|b", "c", false));
        }

        [Fact]
        public void MakeKey_EscapesControlCharacterBoundaries()
        {
            Assert.NotEqual(
                RedisSubscriptionManager.MakeKey("a\u0001b", "c", false),
                RedisSubscriptionManager.MakeKey("a", "b\u0001c", false));
        }
    }

    public class RedisResultFormatterTests
    {
        [Fact]
        public void FormatHash_NullOrEmptyReturnsEmptyJsonObject()
        {
            // A Redis hash is never empty: a missing hash is expressed as "{}"
            // (valid JSON) instead of the "(no value)" sentinel, which stays for
            // GET/HGET only.
            Assert.Equal("{}", RedisResultFormatter.FormatHash(null));
            Assert.Equal("{}", RedisResultFormatter.FormatHash(new HashEntry[0]));
        }

        [Fact]
        public void FormatHash_ProducesValidJson()
        {
            var entries = new[]
            {
                new HashEntry("campo1", "valor1"),
                new HashEntry("quote", "a\"b\n")
            };

            var json = RedisResultFormatter.FormatHash(entries);
            var parsed = JObject.Parse(json);

            Assert.Equal("valor1", parsed["campo1"].ToString());
            Assert.Equal("a\"b\n", parsed["quote"].ToString());
        }

        [Fact]
        public void HashEquals_ComparesContent()
        {
            var a = new[] { new HashEntry("f1", "v1"), new HashEntry("f2", "v2") };
            var b = new[] { new HashEntry("f1", "v1"), new HashEntry("f2", "v2") };
            var changed = new[] { new HashEntry("f1", "v1"), new HashEntry("f2", "other") };
            var reordered = new[] { new HashEntry("f2", "v2"), new HashEntry("f1", "v1") };

            Assert.True(RedisResultFormatter.HashEquals(a, b));
            Assert.True(RedisResultFormatter.HashEquals(a, a));
            Assert.True(RedisResultFormatter.HashEquals(a, reordered));
            Assert.False(RedisResultFormatter.HashEquals(a, changed));
            Assert.False(RedisResultFormatter.HashEquals(a, new[] { new HashEntry("f1", "v1") }));
            Assert.False(RedisResultFormatter.HashEquals(null, a));
        }

        [Fact]
        public void HashEquals_ComparesFieldNamesAsTextNotNumericValue()
        {
            // "1" and "01" are distinct Redis hash fields; comparing the names
            // with RedisValue equality would normalize them to the same number
            // and wrongly suppress an update when a field is renamed.
            var a = new[] { new HashEntry("1", "v") };
            var b = new[] { new HashEntry("01", "v") };
            Assert.False(RedisResultFormatter.HashEquals(a, b));
        }

        [Fact]
        public void HashEquals_ValuesKeepRedisValueEquality()
        {
            // Values keep RedisValue semantics: numeric formatting differences
            // ("1" vs "1.00") are the same value and must not trigger an update.
            var a = new[] { new HashEntry("f", "1") };
            var b = new[] { new HashEntry("f", "1.00") };
            Assert.True(RedisResultFormatter.HashEquals(a, b));
        }
    }

    /// <summary>
    /// Offline tests for the publish-dedup markers invalidated when a
    /// subscription listener joins (RTD SUB/PSUB or UDF ChannelLatest), so an
    /// unchanged payload is still fanned out to the new listener. No Redis
    /// server is involved; only the marker cache and the join handler run.
    /// </summary>
    public class ListenerJoinDedupTests
    {
        public ListenerJoinDedupTests()
        {
            // Deterministic capacity for the shared marker cache these tests
            // seed directly: a tiny PublishDedupCacheSize from the machine's
            // RedisExcel.json could otherwise evict a seed mid-test. This is
            // the only test class that touches the shared cache (the
            // whitespace-channel guards in the other classes fail before the
            // cache is reached), so no cross-class collection is needed.
            RedisUDF.ResetDedupCacheForTests(256);
        }

        private static void Seed(string host, string channel, string payload)
        {
            RedisUDF.LastPublishedMessagesForTests.Set(RedisUDF.ChannelKey(host, channel), payload);
        }

        private static bool HasMarker(string host, string channel)
        {
            return RedisUDF.LastPublishedMessagesForTests.TryGet(RedisUDF.ChannelKey(host, channel), out _);
        }

        [Fact]
        public void ChannelKey_IsInjectiveAcrossHostChannelBoundaries()
        {
            // Without the length prefix both pairs would produce "a:b:c".
            Assert.NotEqual(RedisUDF.ChannelKey("a", "b:c"), RedisUDF.ChannelKey("a:b", "c"));
        }

        [Fact]
        public void ChannelKey_RoundTripsThroughTryParse()
        {
            Assert.True(RedisUDF.TryParseChannelKey(RedisUDF.ChannelKey("h:1", "c:2"), out var host, out var channel));
            Assert.Equal("h:1", host);
            Assert.Equal("c:2", channel);
        }

        [Fact]
        public void ClearForListener_Literal_RemovesOnlyThatChannelsMarker()
        {
            string host = "join-literal-host";
            Seed(host, "orders", "1");
            Seed(host, "other", "2");
            Seed(host + ":other", "orders", "3");

            // The HandleListenerJoined entry point is what the manager event
            // calls; use it here to cover the wrapper too.
            RedisUDF.HandleListenerJoined(host, "orders", pattern: false);

            Assert.False(HasMarker(host, "orders"));
            Assert.True(HasMarker(host, "other"));
            Assert.True(HasMarker(host + ":other", "orders"));
        }

        [Fact]
        public void ClearForListener_StarPattern_ClearsEveryMarkerOfThatHostOnly()
        {
            string host = "join-star-host";
            Seed(host, "orders:1", "1");
            Seed(host, "orders:2", "2");
            Seed(host, "other", "3");
            Seed(host + ":elsewhere", "orders:1", "4");

            // Pattern joins clear ALL markers of the host on purpose: the local
            // glob matcher could not reproduce Redis stringmatchlen semantics
            // exactly, and under-clearing starves the returning listener.
            RedisUDF.ClearForListener(host, "orders:*", pattern: true);

            Assert.False(HasMarker(host, "orders:1"));
            Assert.False(HasMarker(host, "orders:2"));
            Assert.False(HasMarker(host, "other"));
            Assert.True(HasMarker(host + ":elsewhere", "orders:1"));
        }

        [Fact]
        public void ClearForListener_QuestionMarkPattern_ClearsEveryMarkerOfThatHost()
        {
            string host = "join-qmark-host";
            Seed(host, "k:1", "1");
            Seed(host, "k:a", "2");
            Seed(host, "k:12", "3");
            Seed(host, "k:", "4");

            RedisUDF.ClearForListener(host, "k:?", pattern: true);

            Assert.False(HasMarker(host, "k:1"));
            Assert.False(HasMarker(host, "k:a"));
            Assert.False(HasMarker(host, "k:12"));
            Assert.False(HasMarker(host, "k:"));
        }

        [Fact]
        public void ClearForListener_CharacterClassPattern_ClearsEveryMarkerOfThatHost()
        {
            string host = "join-class-host";
            Seed(host, "c:a", "1");
            Seed(host, "c:b", "2");
            Seed(host, "c:c", "3");
            Seed(host, "c:d", "4");
            Seed(host, "c:a1", "5");

            RedisUDF.ClearForListener(host, "c:[abc]", pattern: true);

            Assert.False(HasMarker(host, "c:a"));
            Assert.False(HasMarker(host, "c:b"));
            Assert.False(HasMarker(host, "c:c"));
            Assert.False(HasMarker(host, "c:d"));
            Assert.False(HasMarker(host, "c:a1"));
        }

        [Theory]
        [InlineData("*", "", true)]
        [InlineData("*", "anything", true)]
        [InlineData("orders:*", "orders:", true)]
        [InlineData("orders:*", "orders:7", true)]
        [InlineData("orders:*", "other:7", false)]
        [InlineData("k:?", "k:1", true)]
        [InlineData("k:?", "k:12", false)]
        [InlineData("k:?", "k:", false)]
        [InlineData("[abc]", "b", true)]
        [InlineData("[abc]", "d", false)]
        [InlineData("[a-c]x", "bx", true)]
        [InlineData("[a-c]x", "dx", false)]
        [InlineData("[^abc]", "d", true)]
        [InlineData("[^abc]", "a", false)]
        [InlineData("a*b*c", "aXXbYYc", true)]
        [InlineData("a*b*c", "aXXcYYb", false)]
        public void GlobMatches_SupportsRedisWildcards(string pattern, string value, bool expected)
        {
            Assert.Equal(expected, RedisUDF.GlobMatches(pattern, value));
        }
    }
}

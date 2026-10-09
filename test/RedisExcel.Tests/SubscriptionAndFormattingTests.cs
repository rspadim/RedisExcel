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
        public void MakeKey_HostAndChannelBoundariesDoNotCollide()
        {
            Assert.NotEqual(
                RedisSubscriptionManager.MakeKey("a", "b|c", false),
                RedisSubscriptionManager.MakeKey("a|b", "c", false));
        }
    }

    public class RedisResultFormatterTests
    {
        [Fact]
        public void FormatHash_NullOrEmptyReturnsSentinel()
        {
            Assert.Equal("(no value)", RedisResultFormatter.FormatHash(null));
            Assert.Equal("(no value)", RedisResultFormatter.FormatHash(new HashEntry[0]));
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

            Assert.True(RedisResultFormatter.HashEquals(a, b));
            Assert.True(RedisResultFormatter.HashEquals(a, a));
            Assert.False(RedisResultFormatter.HashEquals(a, changed));
            Assert.False(RedisResultFormatter.HashEquals(a, new[] { new HashEntry("f1", "v1") }));
            Assert.False(RedisResultFormatter.HashEquals(null, a));
        }
    }
}

using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The host argument of a UDF/RTD call can be a full connection string, so
    /// it may carry "password=...". Every log line and every verbose error cell
    /// must mask that credential. MaskHost is the single redaction primitive;
    /// these tests pin it and the end-to-end cell text.
    /// </summary>
    public class HostMaskingTests
    {
        [Fact]
        public void MaskHost_MasksPasswordKeepingTheRestReadable()
        {
            Assert.Equal(
                "localhost:6379,password=****,abortConnect=False",
                AppConfig.MaskHost("localhost:6379,password=secret,abortConnect=False"));
        }

        [Theory]
        [InlineData("Password=secret")]
        [InlineData("PASSWORD=secret")]
        [InlineData("password=secret")]
        public void MaskHost_IsCaseInsensitiveOnTheKeyword(string pair)
        {
            string masked = AppConfig.MaskHost("host:6379," + pair + ",ssl=true");
            Assert.DoesNotContain("secret", masked);
            Assert.Contains("****", masked);
            Assert.Contains("ssl=true", masked); // the rest survives
        }

        [Fact]
        public void MaskHost_AlsoMasksTheShortPassAlias()
        {
            Assert.Equal("host:6379,pass=****", AppConfig.MaskHost("host:6379,pass=abc"));
        }

        [Theory]
        [InlineData("user=admin")]
        [InlineData("username=admin")]
        public void MaskHost_MasksTheUserName(string pair)
        {
            string masked = AppConfig.MaskHost("host:6379," + pair + ",password=secret");
            Assert.DoesNotContain("admin", masked);
            Assert.DoesNotContain("secret", masked);
            Assert.Contains("user", masked); // the key stays visible
        }

        [Theory]
        [InlineData("localhost:6379")]
        [InlineData("localhost:6379,abortConnect=False,ssl=true")]
        [InlineData("")]
        public void MaskHost_LeavesCredentialFreeHostsUnchanged(string host)
        {
            Assert.Equal(host, AppConfig.MaskHost(host));
        }

        [Fact]
        public void MaskHost_NullIsNull()
        {
            Assert.Null(AppConfig.MaskHost(null));
        }

        [Theory]
        // StackExchange.Redis tolerates whitespace around '=' and quoted values;
        // the regex must cover those or the secret leaks (found by audit).
        [InlineData("host:6379,password = secret")]
        [InlineData("host:6379,password =secret")]
        [InlineData("host:6379,password= my secret")]
        [InlineData("host:6379,password=ab;cd")]
        [InlineData("host:6379, password = secret")]
        public void MaskHost_ToleratesWhitespaceAndDelimiters(string input)
        {
            string masked = AppConfig.MaskHost(input);
            Assert.Contains("password", masked, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("****", masked);
            Assert.DoesNotContain("secret", masked);
            Assert.DoesNotContain("ab;cd", masked);
        }

        [Fact]
        public void MaskHost_MasksQuotedValuesIncludingTheSpace()
        {
            string doubleQuoted = AppConfig.MaskHost("host:6379,password=\"my secret\",ssl=true");
            Assert.DoesNotContain("my secret", doubleQuoted);
            Assert.Contains("ssl=true", doubleQuoted); // the value after the quote survives

            string singleQuoted = AppConfig.MaskHost("host:6379,password='q s',x=1");
            Assert.DoesNotContain("q s", singleQuoted);
            Assert.Contains("x=1", singleQuoted);
        }

        [Theory]
        // A user key/channel/pattern that happens to contain a credential-looking
        // token must NOT be rewritten: only a value after a connection-string
        // delimiter is a real credential (audit: the old regex corrupted these).
        [InlineData("key=user=abc, host=dead:1")]
        [InlineData("key=a password=b, host=dead:1")]
        [InlineData("channel=pass=x, host=h")]
        public void MaskHost_DoesNotCorruptUserIdentifiers(string context)
        {
            Assert.Equal(context, AppConfig.MaskHost(context));
        }

        [Fact]
        public void SafeMessage_MasksCredentialsInExceptionText_AsyncPath()
        {
            // The async dispatch path returns the exception text to the cell.
            string masked = RedisUdfAsync.SafeMessage(
                new System.Exception("connect failed: host:6379,password=topsecret,ssl=true"));
            Assert.DoesNotContain("topsecret", masked);
            Assert.Contains("password=****", masked);
        }

        [Fact]
        public void InvalidHostErrorCell_MasksCredentials()
        {
            // An invalid endpoint is rejected by RedisConnectionManager.ParseOptions
            // with "invalid Redis host '<host>'"; the host there is a connection
            // string and reaches the cell. Assert no credential leaks.
            const string secret = "cache-secret";
            string host = "localhost:99999,password=" + secret;
            object result = RedisUDF.RedisUDFGet("k", host);
            string text = (string)result;

            Assert.StartsWith("Error:", text);
            Assert.DoesNotContain(secret, text);
            Assert.Contains("password=****", text);
        }

        [Fact]
        public void VerboseErrorCell_NeverContainsThePassword()
        {
            // End-to-end: a runtime failure on a password-protected host must not
            // leak the password into the cell. The exception type varies with the
            // macha state, so only the "no secret" contract is asserted.
            const string secret = "s3cr3t-should-never-appear";
            string host = "127.0.0.1:1,password=" + secret + ",connectTimeout=300,connectRetry=0,abortConnect=False";
            string previous = RedisUDF.SyncWriteOverrideForTests;
            try
            {
                RedisUDF.SyncWriteOverrideForTests = "sync";
                object result = RedisUDF.RedisUDFIncr("mask:test", host);
                string text = Assert.IsType<string>(result);

                Assert.StartsWith("Error: ", text);
                Assert.DoesNotContain(secret, text);
                Assert.Contains("password=****", text);
            }
            finally
            {
                RedisUDF.SyncWriteOverrideForTests = previous;
            }
        }
    }
}

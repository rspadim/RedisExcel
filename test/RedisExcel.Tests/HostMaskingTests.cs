using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The host argument of a UDF/RTD call can be a full connection string, so
    /// it may carry "password=...". Every log line and every verbose error cell
    /// must mask that credential. MaskHost is the single redaction primitive;
    /// these tests pin it and the end-to-end cell text.
    /// </summary>
    [Collection("write-mode-seam")]
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
        // A credential AFTER WHITESPACE (no comma) must be masked too: it can reach
        // a cell or log line embedded after any whitespace (audit round P).
        [InlineData("localhost:6379 password=SECRET")]
        [InlineData("localhost:6379\tpassword=SECRET")]
        public void MaskHost_MasksCredentialsAfterWhitespace(string input)
        {
            string masked = AppConfig.MaskHost(input);
            Assert.DoesNotContain("SECRET", masked);
            Assert.Contains("****", masked);
        }

        [Theory]
        // A URI authority carries userinfo the key=value regex cannot see.
        [InlineData("redis://user:SECRET@127.0.0.1:6379")]
        [InlineData("rediss://:SECRET@127.0.0.1:6379")]
        [InlineData("redis://user:SECRET@127.0.0.1:6379,abortConnect=False")]
        public void MaskHost_MasksUriUserInfo(string input)
        {
            string masked = AppConfig.MaskHost(input);
            Assert.DoesNotContain("SECRET", masked);
            Assert.DoesNotContain("user:", masked);
            Assert.Contains("6379", masked); // the host survives
        }

        [Theory]
        // A user key/channel/pattern that happens to contain a credential-looking
        // token must NOT be rewritten: only a value after a connection-string
        // delimiter is a real credential (audit: the old regex corrupted these).
        // (A credential after whitespace-but-no-delimiter is a known, documented
        // limitation: masking it would corrupt such identifiers.)
        [InlineData("key=user=abc, host=dead:1")]
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

        [Theory]
        // A connection-string password may contain '@'; the key=value rule must
        // mask the WHOLE value even when a bare-userinfo-looking token follows
        // (audit G2: the bare rule ran first and leaked the tail).
        [InlineData("host:6379,password=admin@S3cret!", "host:6379,password=****")]
        [InlineData("host:6379,password=@topsecret", "host:6379,password=****")]
        [InlineData("local:6379,password=a@SECRET,ssl=true", "local:6379,password=****,ssl=true")]
        [InlineData("host:6379,password=p@ssword", "host:6379,password=****")]
        [InlineData("host:6379,password='p@ss'", "host:6379,password=****")]
        [InlineData("host:6379 user=admin@S3cret!", "host:6379 user=****")]
        public void MaskHost_MasksAKeyValueValueThatContainsAt(string input, string expected)
        {
            Assert.Equal(expected, AppConfig.MaskHost(input));
        }

        [Theory]
        // A bare user:pass@host (no scheme) is still masked, but a value with no
        // colon before '@' (an email, "key=a@b") must NOT be corrupted (audit G2).
        [InlineData("user:secret@127.0.0.1:6379", "****@127.0.0.1:6379")]
        [InlineData("key=a@b", "key=a@b")]
        [InlineData("someone@example.com", "someone@example.com")]
        public void MaskHost_BareUserInfoNeedsAColon(string input, string expected)
        {
            Assert.Equal(expected, AppConfig.MaskHost(input));
        }

        [Fact]
        public void MaskHost_IsLinearOnALongRunOfSpaces()
        {
            // Regression for the ReDoS (audit H1): the old delimiter regex let a
            // whitespace run be consumed by two alternatives, which is O(n^2).
            // A pathological value must return quickly, not hang for seconds.
            string input = new string(' ', 20000);
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string masked = AppConfig.MaskHost(input);
            sw.Stop();
            Assert.Equal(input, masked);
            Assert.True(sw.ElapsedMilliseconds < 1000,
                "MaskHost took " + sw.ElapsedMilliseconds + "ms on a whitespace run (likely ReDoS)");
        }

        [Fact]
        public void ConfigIssue_StripsUnicodeLineSeparators()
        {
            // A key name can contain U+0085/U+2028/U+2029 (not char.IsControl),
            // which would forge a second line/cell (audit H1/C3).
            var nel = new ConfigIssue("/Servers/evil\u0085FAKE", "x");
            Assert.DoesNotContain("\u0085", nel.ToString());
            var ls = new ConfigIssue("/a\u2028b", "x");
            Assert.DoesNotContain("\u2028", ls.ToString());
            var ps = new ConfigIssue("/a\u2029b", "x");
            Assert.DoesNotContain("\u2029", ps.ToString());
            var c1 = new ConfigIssue("/a\u009fb", "x");
            Assert.DoesNotContain("\u009f", c1.ToString());
        }

        [Fact]
        public void ServerDefConverter_ReadsKeysCaseInsensitively()
        {
            // The converter must not depend on Clean having canonicalized the keys
            // (audit H1/C4): a raw object with Host/ConnectTimeout still binds.
            var def = Newtonsoft.Json.JsonConvert.DeserializeObject<ServerDef>(
                "{\"Host\":\"h:1\",\"ConnectTimeout\":9,\"SyncTimeout\":8}");
            Assert.Equal("h:1", def.host);
            Assert.Equal(9, def.connectTimeout);
            Assert.Equal(8, def.syncTimeout);
        }

        [Theory]
        // A URI authority carries userinfo the key=value regex cannot see. The
        // full userinfo must be redacted even when it contains its own '@'
        // (audit D2: the old regex stopped at the first '@' and leaked the rest).
        [InlineData("redis://user:p@leakpart@127.0.0.1:6379")]
        [InlineData("redis://user:secret@127.0.0.1:6379,abortConnect=False")]
        [InlineData("user:secret@127.0.0.1:6379")]   // bare, no scheme (audit F1)
        [InlineData(":secret@127.0.0.1:6379")]        // password only
        public void MaskHost_MasksTheWholeUriUserInfo(string input)
        {
            string masked = AppConfig.MaskHost(input);
            Assert.DoesNotContain("leakpart", masked);
            Assert.DoesNotContain("secret", masked);
            Assert.DoesNotContain("MY PASSWORD", masked);
            Assert.DoesNotContain("user:", masked);
            Assert.Contains("127.0.0.1:6379", masked); // the host survives
        }

        [Fact]
        public void ConfigIssue_MasksASecretInThePathAndStripsControlCharacters()
        {
            // The issue PATH embeds raw JSON key names; a Servers key can be a
            // connection string with password=..., and a key can contain \n.
            var secret = new ConfigIssue("/Servers/h:6379,password=TOPSECRET", "expected a connection string");
            string s = secret.ToString();
            Assert.DoesNotContain("TOPSECRET", s);
            Assert.Contains("password=****", s);
            Assert.Contains("expected a connection string", s);

            var newline = new ConfigIssue("/Servers/srv\nFORGED|injected", "x");
            Assert.DoesNotContain("\n", newline.ToString());
            Assert.DoesNotContain("\r", newline.ToString());
        }

        [Fact]
        public void RejectedReason_MasksASecretInAServersKeyAndNeverSpansLines()
        {
            string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "redisexcel-mask-" + System.Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(dir);
            try
            {
                string path = System.IO.Path.Combine(dir, "RedisExcel.json");
                System.IO.File.WriteAllText(path,
                    "{\"Default\":\"localhost:6379\",\"Servers\":{\"h:6379,password=TOPSECRET\":12345}}");

                var config = AppConfig.LoadFromPaths(new[] { path });

                Assert.True(config.Rejected);
                Assert.DoesNotContain("TOPSECRET", config.RejectedReason);
                Assert.DoesNotContain("\n", config.RejectedReason);
                Assert.DoesNotContain("\r", config.RejectedReason);
            }
            finally
            {
                try { System.IO.Directory.Delete(dir, true); } catch { }
            }
        }
    }
}

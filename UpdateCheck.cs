using NLog;
using Newtonsoft.Json.Linq;
using System;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;

namespace RedisExcel
{
    /// <summary>
    /// Non-blocking GitHub release check: runs in a background task at add-in
    /// load and refreshes at most every 6 hours when the worksheet function
    /// recalculates. The result is stored so functions never block Excel.
    /// Failures (offline, rate limit) are silent and only logged.
    /// </summary>
    internal static class UpdateCheck
    {
        private const string LatestReleaseUrl = "https://api.github.com/repos/rspadim/RedisExcel/releases/latest";

        /// <summary>Retry delay after a failed check (offline, rate limit, ...).</summary>
        private static readonly TimeSpan FailureRetryDelay = TimeSpan.FromMinutes(5);

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();

        private static bool _inFlight;
        private static DateTime _lastAttemptUtc;
        private static DateTime _lastSuccessUtc;
        private static string _latestTag;

        internal static string CurrentTag => CurrentTagOverrideForTests ?? BuildInfo.Tag;

        /// <summary>Test-only override of the running build tag (null = the real
        /// InformationalVersion). Lets the offline tests exercise the positive
        /// "an update is available" path deterministically, which the local
        /// "dev" build tag otherwise skips.</summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static string CurrentTagOverrideForTests;
#pragma warning restore 0649

        /// <summary>Starts the check at add-in load time.</summary>
        internal static void Start() => EnsureFresh(TimeSpan.Zero);

        /// <summary>
        /// Schedules a background refresh when the last SUCCESSFUL check is older
        /// than <paramref name="maxAge"/>. After a failure the refresh is retried
        /// after <see cref="FailureRetryDelay"/> instead of consuming the full
        /// window silently. Safe to call from worksheet functions.
        /// </summary>
        internal static void EnsureFresh(TimeSpan maxAge)
        {
            if (!AppConfig.Current.UpdateCheck)
                return;
            lock (Sync)
            {
                if (_inFlight)
                    return;
                if (maxAge > TimeSpan.Zero && _lastAttemptUtc != default(DateTime))
                {
                    // Gate on the last success (full interval) once a check
                    // worked; while the last attempt failed, use the short
                    // backoff so recovery does not wait maxAge. maxAge=Zero
                    // (Start at add-in load) keeps its original "check now"
                    // semantics, unaffected by a previous failure's backoff.
                    bool lastAttemptSucceeded = _lastSuccessUtc >= _lastAttemptUtc;
                    DateTime gateFrom = lastAttemptSucceeded ? _lastSuccessUtc : _lastAttemptUtc;
                    TimeSpan window = lastAttemptSucceeded ? maxAge : FailureRetryDelay;
                    if (DateTime.UtcNow - gateFrom < window)
                        return;
                }
                _inFlight = true;
                _lastAttemptUtc = DateTime.UtcNow;
            }
            Task.Run(() =>
            {
                try
                {
                    Refresh();
                }
                finally
                {
                    lock (Sync) { _inFlight = false; }
                }
            });
        }

        /// <summary>TRUE when a newer release is known to exist (no I/O, never blocks).</summary>
        internal static bool IsUpdateAvailable()
        {
            lock (Sync)
                return _latestTag != null && IsNewer(_latestTag, CurrentTag);
        }

        private static void Refresh()
        {
            try
            {
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) })
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("RedisExcel-UpdateCheck");
                    client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
                    string json = client.GetStringAsync(LatestReleaseUrl).GetAwaiter().GetResult();
                    string tag = (string)JObject.Parse(json)["tag_name"];
                    string latest;
                    lock (Sync)
                    {
                        latest = _latestTag;
                        // A response without a tag_name is not a usable release:
                        // treat it as a failed check, i.e. keep the last known
                        // tag and skip the success marker, so the short retry
                        // backoff still applies.
                        if (tag != null)
                        {
                            _latestTag = tag;
                            latest = tag;
                            _lastSuccessUtc = DateTime.UtcNow;
                        }
                    }
                    if (logger.IsInfoEnabled)
                        logger.Info(tag == null
                            ? $"UpdateCheck: response without tag_name, keeping latest={latest}"
                            : $"UpdateCheck: current={CurrentTag}, latest={latest}");
                }
            }
            catch (Exception ex)
            {
                if (logger.IsInfoEnabled)
                    logger.Info($"UpdateCheck: {ex.Message}");
            }
        }

        internal static bool IsNewer(string candidate, string current)
        {
            var a = ParseVersion(NormalizeTag(candidate));
            var b = ParseVersion(NormalizeTag(current));
            if (a == null || b == null)
                return false; // unknown formats (e.g. "dev"): never alert
            return Compare(a, b) > 0;
        }

        internal static string NormalizeTag(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag))
                return "";
            tag = tag.Trim();
            if (tag.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                tag = tag.Substring(1);
            int suffix = tag.IndexOfAny(new[] { '-', '+' });
            if (suffix >= 0)
                tag = tag.Substring(0, suffix);
            // Re-trim: "v1.4.1 -rc" would otherwise keep the space before the
            // suffix and ParseVersion would reject the whole tag (missed
            // update alerts).
            return tag.Trim();
        }

        private static long[] ParseVersion(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;
            var parts = value.Split('.');
            var numbers = new long[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                // long: an int.TryParse failure (overflow) would silently
                // disable update alerts for a large-but-real version number.
                if (!long.TryParse(parts[i], NumberStyles.None, CultureInfo.InvariantCulture, out numbers[i]))
                    return null;
            }
            return numbers;
        }

        private static int Compare(long[] a, long[] b)
        {
            int length = Math.Max(a.Length, b.Length);
            for (int i = 0; i < length; i++)
            {
                long x = i < a.Length ? a[i] : 0;
                long y = i < b.Length ? b[i] : 0;
                if (x != y)
                    return x.CompareTo(y);
            }
            return 0;
        }
    }
}

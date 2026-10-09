using NLog;
using Newtonsoft.Json.Linq;
using System;
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

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();

        private static bool _inFlight;
        private static DateTime _lastAttemptUtc;
        private static string _latestTag;

        internal static string CurrentTag => BuildInfo.Tag;

        /// <summary>Starts the check at add-in load time.</summary>
        internal static void Start() => EnsureFresh(TimeSpan.Zero);

        /// <summary>
        /// Schedules a background refresh when the last attempt is older than
        /// <paramref name="maxAge"/>. Safe to call from worksheet functions.
        /// </summary>
        internal static void EnsureFresh(TimeSpan maxAge)
        {
            if (!AppConfig.Current.UpdateCheck)
                return;
            lock (Sync)
            {
                if (_inFlight)
                    return;
                if (_lastAttemptUtc != default(DateTime) && DateTime.UtcNow - _lastAttemptUtc < maxAge)
                    return;
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
                    lock (Sync) { _latestTag = tag; }
                    if (logger.IsInfoEnabled)
                        logger.Info($"UpdateCheck: current={CurrentTag}, latest={tag}");
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
            return tag;
        }

        private static int[] ParseVersion(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;
            var parts = value.Split('.');
            var numbers = new int[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                if (!int.TryParse(parts[i], out numbers[i]))
                    return null;
            }
            return numbers;
        }

        private static int Compare(int[] a, int[] b)
        {
            int length = Math.Max(a.Length, b.Length);
            for (int i = 0; i < length; i++)
            {
                int x = i < a.Length ? a[i] : 0;
                int y = i < b.Length ? b[i] : 0;
                if (x != y)
                    return x.CompareTo(y);
            }
            return 0;
        }
    }
}

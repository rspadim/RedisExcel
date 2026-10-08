using NLog;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace RedisExcel
{
    /// <summary>
    /// Non-blocking GitHub release check. It runs in a background task with a
    /// short timeout, never touches the Excel main thread, caches the result and
    /// exposes it through the RedisUDFUpdate* worksheet functions. Failures
    /// (offline, rate limit) are silent and only logged.
    /// </summary>
    internal static class UpdateCheck
    {
        private const string LatestReleaseUrl = "https://api.github.com/repos/rspadim/RedisExcel/releases/latest";

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();
        private static readonly TimeSpan MinInterval = TimeSpan.FromHours(6);

        private static bool _inFlight;
        private static DateTime _lastAttemptUtc;
        private static string _latestTag;
        private static string _releaseUrl;
        private static string _status = "checking...";

        internal static string CurrentTag => BuildInfo.Tag;

        /// <summary>Kicks off a check at add-in load time.</summary>
        internal static void Start() => RefreshIfStale(TimeSpan.Zero);

        /// <summary>
        /// Schedules a background refresh when the last attempt is older than
        /// <paramref name="maxAge"/>. Safe to call from worksheet functions.
        /// </summary>
        internal static void RefreshIfStale(TimeSpan maxAge)
        {
            if (!AppConfig.Current.UpdateCheck.enabled)
            {
                lock (Sync) { _status = "update check disabled"; }
                return;
            }

            lock (Sync)
            {
                if (_inFlight)
                    return;
                if (_lastAttemptUtc != default(DateTime) && DateTime.UtcNow - _lastAttemptUtc < maxAge)
                    return;
                _inFlight = true;
                _lastAttemptUtc = DateTime.UtcNow;
                if (_latestTag == null)
                    _status = "checking...";
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

        /// <summary>Single-line summary for the worksheet.</summary>
        internal static string Summary()
        {
            lock (Sync)
            {
                if (_latestTag == null)
                    return _status;
                return IsNewer(_latestTag, CurrentTag)
                    ? "update available: " + _latestTag
                    : "up to date (" + _latestTag + ")";
            }
        }

        /// <summary>Detailed 5x2 matrix for the worksheet.</summary>
        internal static object[,] Info()
        {
            lock (Sync)
            {
                string availability = _latestTag == null
                    ? _status
                    : (IsNewer(_latestTag, CurrentTag) ? "YES" : "no");
                return new object[,]
                {
                    { "Current version", CurrentTag },
                    { "Latest version", (object)_latestTag ?? "(unknown)" },
                    { "Update available", availability },
                    { "Release page", (object)_releaseUrl ?? "(unknown)" },
                    { "Status", _status }
                };
            }
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
                    var release = JObject.Parse(json);
                    string tag = (string)release["tag_name"];
                    string url = (string)release["html_url"];

                    lock (Sync)
                    {
                        _latestTag = tag;
                        _releaseUrl = url;
                        _status = IsNewer(tag, CurrentTag) ? "update available: " + tag : "up to date (" + tag + ")";
                    }
                    if (logger.IsInfoEnabled)
                        logger.Info($"UpdateCheck: current={CurrentTag}, latest={tag}, url={url}");
                }
            }
            catch (Exception ex)
            {
                lock (Sync) { _status = "release check failed"; }
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

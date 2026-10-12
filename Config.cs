using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace RedisExcel
{
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ENUMExcelUpdateStyle
    {
        Timer = 0,
        Realtime = 1,
        Automatic = 2
    }

    public class ConfigSection
    {
        public string host { get; set; }
        public int timeout { get; set; }
    }

    public class RTDConfig : ConfigSection
    {
        public int RedisUpdateRateMs { get; set; } = 1000;
        public int ExcelUpdateRateMs { get; set; } = 100;
        public ENUMExcelUpdateStyle ExcelUpdateStyle { get; set; } = ENUMExcelUpdateStyle.Automatic;
        public long MessageCounterThreshold { get; set; } = 10000;
        public bool UseGetMultiple { get; set; } = true;
    }

    public class ConfigRoot
    {
        public RTDConfig RTD { get; set; }
        public ConfigSection UDF { get; set; }
        public Dictionary<string, string> Servers { get; set; }
        public bool UpdateCheck { get; set; } = true;
        /// <summary>Skip identical consecutive payloads before decoding/delivering
        /// (feeds republish unchanged values constantly). Default on.</summary>
        public bool SkipRepeatedMessages { get; set; } = true;
        /// <summary>In real-time mode, send at most one update per topic per
        /// ExcelUpdateRateMs window instead of one per incoming message.
        /// Default on (the more performant option). Superseded by an explicit
        /// <see cref="ConflationMs"/>.</summary>
        public bool CoalesceRealtimeUpdates { get; set; } = true;
        /// <summary>Explicit real-time conflation window in milliseconds: the
        /// latest value of a topic is pushed to Excel at most once per window
        /// (0 = off, one push per message). When absent (null), the legacy
        /// <see cref="CoalesceRealtimeUpdates"/> boolean decides. Sanitized to
        /// [0, Conflation.MaxWindowMs].</summary>
        public int? ConflationMs { get; set; }
        /// <summary>Maximum number of host/channel entries remembered by the
        /// PublishIfChanged deduplication cache (LRU). Default 10000.</summary>
        public int PublishDedupCacheSize { get; set; } = 10000;
        /// <summary>Write mode: "sync", "fireforget" or "fireforget-all".
        /// Sanitized to the lowercase form; unknown or blank values fall back
        /// to "fireforget".</summary>
        public string SyncWrite { get; set; } = "fireforget";
        /// <summary>Whether writes are dispatched asynchronously. Default off.</summary>
        public bool AsyncWrites { get; set; } = false;
    }

    /// <summary>Write mode names accepted by <see cref="ConfigRoot.SyncWrite"/>.</summary>
    internal static class SyncWriteModes
    {
        public const string Sync = "sync";
        public const string FireForget = "fireforget";
        public const string FireForgetAll = "fireforget-all";
    }

    /// <summary>
    /// Loads RedisExcel.json once per process and resolves host aliases.
    /// Before: every call could re-read the file from disk, and a missing section
    /// (RTD/UDF) caused a NullReferenceException inside the functions.
    /// </summary>
    public static class AppConfig
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private const string ConfigFileName = "RedisExcel.json";

        public const string DefaultHost = "localhost:6379,password=,defaultDatabase=0,ssl=False,abortConnect=False";

        private static readonly Lazy<ConfigRoot> _current = new Lazy<ConfigRoot>(Load);

        /// <summary>Test-only override: when set, <see cref="Current"/> returns this
        /// instance instead of the machine's file, so the unit suite is hermetic and
        /// never depends on a developer's %USERPROFILE%\RedisExcel.json.</summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static ConfigRoot CurrentOverrideForTests;
#pragma warning restore 0649

        public static ConfigRoot Current => CurrentOverrideForTests ?? _current.Value;

        private static ConfigRoot Load() => LoadFromPaths(CandidatePaths());

        /// <summary>
        /// Deserializes a configuration file tolerating per-value mistakes: the
        /// JSON is first validated against the embedded schema (comments
        /// accepted) and every invalid value is REMOVED before deserializing, so
        /// the reader never sees a wrong-typed value. This closes the hole where a
        /// bool/object/array in a numeric position failed at the JSON reader level
        /// (the Error callback cannot repair that) and silently discarded the
        /// whole file - repointing the host to localhost and changing the write
        /// mode. Each dropped value is reported through <paramref name="onError"/>
        /// and collected in <paramref name="issues"/>. A structural failure (not
        /// JSON) still throws to the caller's catch and falls back to defaults.
        /// </summary>
        internal static ConfigRoot ParseTolerant(string json, Action<Exception> onError, IList<ConfigIssue> issues = null)
        {
            JObject root = ConfigSchema.ParseWithComments(json);
            if (root == null)
                return null;
            var collected = issues ?? new List<ConfigIssue>();
            JToken cleaned = ConfigSchema.Clean(root, collected);
            if (onError != null)
            {
                foreach (var issue in collected)
                    onError(new FormatException(issue.ToString()));
            }
            if (cleaned == null)
                return null;
            // Comments / unknown keys are gone and every remaining value matches the
            // schema, so the strict deserialize cannot fail on a typed property.
            return cleaned.ToObject<ConfigRoot>();
        }

        /// <summary>
        /// Loads the first existing candidate path. A file that exists but cannot
        /// be parsed stops the search: the process falls back to safe defaults
        /// instead of silently using a lower-priority file. When no candidate
        /// exists, the legacy no-file defaults apply (1000ms Excel flush timer).
        /// Extracted for unit tests.
        /// </summary>
        internal static ConfigRoot LoadFromPaths(IEnumerable<string> paths)
        {
            try
            {
                foreach (var path in paths)
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                            continue;

                        // First existing candidate wins: if it exists but cannot be
                        // parsed, stop here with safe defaults instead of silently
                        // falling through to a lower-priority file.
                        string json = File.ReadAllText(path);
                        var issues = new List<ConfigIssue>();
                        var config = ParseTolerant(json, ex =>
                            logger.Error(ex, $"AppConfig: value ignored in {path} (schema), keeping the rest of the file"), issues);
                        if (config != null)
                        {
                            logger.Info(issues.Count == 0
                                ? $"AppConfig: loaded configuration from {path}"
                                : $"AppConfig: loaded configuration from {path} with {issues.Count} value(s) ignored");
                            return Sanitize(config);
                        }
                        logger.Error($"AppConfig: config file {path} is not a JSON object, using defaults");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"AppConfig: error reading config file {path}, using defaults");
                    }
                    // The first existing candidate stops the search, even on failure.
                    break;
                }
            }
            catch (Exception ex)
            {
                // The enumeration itself can throw (path probing): never let the
                // Lazy cache that exception - every function would then fail
                // until Excel restarts. Fall back to the safe defaults instead.
                logger.Error(ex, "AppConfig: error while probing configuration paths, using defaults");
            }
            logger.Info("AppConfig: no configuration file loaded, using defaults");
            var fallback = Sanitize(new ConfigRoot());
            // The original no-file fallback used 1000ms for the Excel flush timer;
            // keep that legacy behavior when no configuration file could be loaded.
            fallback.RTD.ExcelUpdateRateMs = 1000;
            return fallback;
        }

        private static IEnumerable<string> CandidatePaths()
        {
            return CandidatePaths(AddInDirectory(), ExcelDirectory());
        }

        /// <summary>Pure ordering, testable without Excel: the add-in's own folder
        /// (where the XLL lives, e.g. C:\RedisExcel) is searched first, then the
        /// user profile, the Excel folder and C:\Windows (first existing wins).</summary>
        internal static IEnumerable<string> CandidatePaths(string addInDirectory, string excelDirectory)
        {
            yield return addInDirectory == null ? null : Path.Combine(addInDirectory, ConfigFileName);
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ConfigFileName);
            yield return excelDirectory == null ? null : Path.Combine(excelDirectory, ConfigFileName);
            var windowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            yield return string.IsNullOrEmpty(windowsDirectory) ? null : Path.Combine(windowsDirectory, ConfigFileName);
        }

        private static string AddInDirectory()
        {
            try
            {
                var xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                return string.IsNullOrWhiteSpace(xll) ? null : Path.GetDirectoryName(xll);
            }
            catch
            {
                return null;
            }
        }

        private static string ExcelDirectory()
        {
            try
            {
                return Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            }
            catch
            {
                return null;
            }
        }

        internal static ConfigRoot Sanitize(ConfigRoot config)
        {
            config = config ?? new ConfigRoot();
            config.RTD = config.RTD ?? new RTDConfig();
            config.UDF = config.UDF ?? new ConfigSection();
            if (string.IsNullOrWhiteSpace(config.RTD.host)) config.RTD.host = DefaultHost;
            if (string.IsNullOrWhiteSpace(config.UDF.host)) config.UDF.host = DefaultHost;
            if (config.RTD.timeout <= 0) config.RTD.timeout = 1000;
            if (config.UDF.timeout <= 0) config.UDF.timeout = 1000;
            if (config.RTD.RedisUpdateRateMs <= 0) config.RTD.RedisUpdateRateMs = 1000;
            if (config.RTD.ExcelUpdateRateMs <= 0) config.RTD.ExcelUpdateRateMs = 100;
            if (config.PublishDedupCacheSize <= 0) config.PublishDedupCacheSize = 10000;
            if (!Enum.IsDefined(typeof(ENUMExcelUpdateStyle), config.RTD.ExcelUpdateStyle))
                config.RTD.ExcelUpdateStyle = ENUMExcelUpdateStyle.Automatic;
            if (config.ConflationMs.HasValue)
                config.ConflationMs = Conflation.Sanitize(config.ConflationMs.Value);
            config.SyncWrite = NormalizeSyncWrite(config.SyncWrite);
            return config;
        }

        /// <summary>Normalizes a SyncWrite mode to its lowercase form; unknown
        /// or blank values fall back to "fireforget". Kept internal for tests.</summary>
        internal static string NormalizeSyncWrite(string mode)
        {
            switch (mode?.Trim().ToLowerInvariant())
            {
                case SyncWriteModes.Sync: return SyncWriteModes.Sync;
                case SyncWriteModes.FireForget: return SyncWriteModes.FireForget;
                case SyncWriteModes.FireForgetAll: return SyncWriteModes.FireForgetAll;
                default: return SyncWriteModes.FireForget;
            }
        }

        /// <summary>Pure alias/default resolution (kept internal for unit tests).</summary>
        internal static string ResolveHostCore(string host, string defaultHost, IDictionary<string, string> servers)
        {
            string candidate = string.IsNullOrWhiteSpace(host) ? defaultHost : host.Trim();
            if (candidate == null)
                return null;
            if (servers != null)
            {
                if (servers.TryGetValue(candidate, out var mapped) && !string.IsNullOrWhiteSpace(mapped))
                    return mapped;
                // "PROD"/" prod " should still hit the "prod" alias: fall back
                // to a case-insensitive scan (the exact match stays first, so a
                // config with two spellings keeps its own priority).
                foreach (var kv in servers)
                {
                    if (string.Equals(kv.Key, candidate, StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(kv.Value))
                        return kv.Value;
                }
            }
            return candidate;
        }

        /// <summary>Resolves aliases ("prod", "dev") defined in Servers; a blank host uses the provided default.</summary>
        public static string ResolveHost(string host, string defaultHost)
        {
            return ResolveHostCore(host, defaultHost, Current.Servers);
        }

        /// <summary>
        /// Redacts the credentials of a Redis host/connection string (or any
        /// text that embeds one, e.g. an exception message) before it reaches
        /// the log or a cell. The host argument of a UDF/RTD call may be a full
        /// connection string ("host:6379,password=secret,user=admin,..."), so a
        /// bare "host={host}" would leak a credential.
        ///
        /// The value is matched the way StackExchange.Redis accepts it: the
        /// keyword must follow a connection-string delimiter (start, ',' or ';',
        /// optionally with whitespace after it) so a user key/value like
        /// "key=user=abc" or "key=a password=b" is NOT rewritten, optional
        /// whitespace is allowed around '=', and quoted values ('...' or "...")
        /// are handled. The unquoted value runs to the next ',' (Redis does not
        /// split on ';' inside a value), so "password=ab;cd" is fully masked too.
        /// </summary>
        public static string MaskHost(string host)
        {
            if (string.IsNullOrEmpty(host))
                return host;
            // A URI authority (redis://user:pass@host) carries credentials the
            // key=value regex cannot see; redact its userinfo first.
            string masked = UriUserInfo.Replace(host, "****@");
            return PasswordPair.Replace(masked, "$1$2=****");
        }

        // redis://user:SECRET@host / rediss://:SECRET@host / redis://user@host:
        // whatever sits between '//' and '@' is userinfo and must be redacted.
        private static readonly Regex UriUserInfo = new Regex(
            @"(?<=//)[^/@\s]+@", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // A credential keyword is preceded by the string start, a connection-string
        // delimiter (comma/semicolon) OR whitespace: a credential can reach a cell
        // or log line embedded after any whitespace, not only after a comma, and a
        // leaked secret is worse than over-masking a user identifier that happens to
        // look like "x=password=y" (audit round P).
        private static readonly Regex PasswordPair = new Regex(
            @"(^|[,;\s])\s*(password|pass|user|username)\s*=\s*(?:""[^""]*""|'[^']*'|[^,]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static string ResolveRtdHost(string host) => ResolveHost(host, Current.RTD.host);

        public static string ResolveUdfHost(string host) => ResolveHost(host, Current.UDF.host);

        /// <summary>
        /// Validates a RedisExcel.json and returns a short report for a cell.
        /// <paramref name="path"/> = null/blank validates the configuration this
        /// process would load (same discovery order); otherwise the given file is
        /// validated. Comments (// and /* */) are accepted. Returns "OK (...)" when
        /// valid, otherwise one line per problem ("&lt;path&gt;: &lt;message&gt;").
        /// Never throws.
        /// </summary>
        public static string ValidateConfiguration(string path = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    path = FirstExistingPath();
                else
                    path = path.Trim();

                if (path == null)
                    return "OK (no RedisExcel.json found; using defaults)";
                if (!File.Exists(path))
                    return "Error: file not found: " + path;

                string json = File.ReadAllText(path);
                JObject root;
                try
                {
                    root = ConfigSchema.ParseWithComments(json);
                }
                catch (Exception ex)
                {
                    return "Error: " + AppConfig.MaskHost(ex.Message) + " (" + path + ")";
                }
                if (root == null)
                    return "Error: the config root is not a JSON object (" + path + ")";

                var issues = new List<ConfigIssue>();
                ConfigSchema.Clean(root, issues);
                if (issues.Count == 0)
                    return "OK (" + path + ")";

                var lines = new List<string> { "Error: " + issues.Count + " problem(s) in " + path };
                foreach (var issue in issues)
                    lines.Add(issue.ToString());
                return string.Join("\n", lines);
            }
            catch (Exception ex)
            {
                return "Error: " + AppConfig.MaskHost(ex.Message);
            }
        }

        /// <summary>The first existing candidate path, or null when none exists
        /// (same discovery order as the loader).</summary>
        internal static string FirstExistingPath()
        {
            foreach (var path in CandidatePaths())
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    return path;
            }
            return null;
        }

        /// <summary>Configured write mode ("sync", "fireforget" or "fireforget-all");
        /// convenience accessor for the process-wide configuration.</summary>
        public static string SyncWrite => Current.SyncWrite;

        /// <summary>Whether writes are dispatched asynchronously; convenience
        /// accessor for the process-wide configuration.</summary>
        public static bool AsyncWrites => Current.AsyncWrites;
    }
}

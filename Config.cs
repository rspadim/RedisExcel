using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;

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
        /// Default on (the more performant option).</summary>
        public bool CoalesceRealtimeUpdates { get; set; } = true;
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

        public static ConfigRoot Current => _current.Value;

        private static ConfigRoot Load() => LoadFromPaths(CandidatePaths());

        /// <summary>
        /// Loads the first existing candidate path. A file that exists but cannot
        /// be parsed stops the search: the process falls back to safe defaults
        /// instead of silently using a lower-priority file. When no candidate
        /// exists, the legacy no-file defaults apply (1000ms Excel flush timer).
        /// Extracted for unit tests.
        /// </summary>
        internal static ConfigRoot LoadFromPaths(IEnumerable<string> paths)
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
                    var config = JsonConvert.DeserializeObject<ConfigRoot>(File.ReadAllText(path));
                    if (config != null)
                    {
                        logger.Info($"AppConfig: loaded configuration from {path}");
                        return Sanitize(config);
                    }
                    logger.Error($"AppConfig: config file {path} deserialized to null, using defaults");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"AppConfig: error reading config file {path}, using defaults");
                }
                // The first existing candidate stops the search, even on failure.
                break;
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
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ConfigFileName);
            var excelDirectory = ExcelDirectory();
            yield return excelDirectory == null ? null : Path.Combine(excelDirectory, ConfigFileName);
            yield return Path.Combine("C:\\Windows", ConfigFileName);
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
            config.RTD = config.RTD ?? new RTDConfig();
            config.UDF = config.UDF ?? new ConfigSection();
            if (string.IsNullOrWhiteSpace(config.RTD.host)) config.RTD.host = DefaultHost;
            if (string.IsNullOrWhiteSpace(config.UDF.host)) config.UDF.host = DefaultHost;
            if (config.RTD.timeout <= 0) config.RTD.timeout = 1000;
            if (config.UDF.timeout <= 0) config.UDF.timeout = 1000;
            if (config.RTD.RedisUpdateRateMs <= 0) config.RTD.RedisUpdateRateMs = 1000;
            if (config.RTD.ExcelUpdateRateMs <= 0) config.RTD.ExcelUpdateRateMs = 100;
            if (config.PublishDedupCacheSize <= 0) config.PublishDedupCacheSize = 10000;
            config.SyncWrite = NormalizeSyncWrite(config.SyncWrite);
            return config;
        }

        /// <summary>Normalizes a SyncWrite mode to its lowercase form; unknown
        /// or blank values fall back to "fireforget". Kept internal for tests.</summary>
        internal static string NormalizeSyncWrite(string mode)
        {
            switch (mode?.Trim().ToLowerInvariant())
            {
                case "sync": return "sync";
                case "fireforget": return "fireforget";
                case "fireforget-all": return "fireforget-all";
                default: return "fireforget";
            }
        }

        /// <summary>Pure alias/default resolution (kept internal for unit tests).</summary>
        internal static string ResolveHostCore(string host, string defaultHost, IDictionary<string, string> servers)
        {
            string candidate = string.IsNullOrWhiteSpace(host) ? defaultHost : host;
            if (servers != null && servers.TryGetValue(candidate, out var mapped) && !string.IsNullOrWhiteSpace(mapped))
                return mapped;
            return candidate;
        }

        /// <summary>Resolves aliases ("prod", "dev") defined in Servers; a blank host uses the provided default.</summary>
        public static string ResolveHost(string host, string defaultHost)
        {
            return ResolveHostCore(host, defaultHost, Current.Servers);
        }

        public static string ResolveRtdHost(string host) => ResolveHost(host, Current.RTD.host);

        public static string ResolveUdfHost(string host) => ResolveHost(host, Current.UDF.host);

        /// <summary>Configured write mode ("sync", "fireforget" or "fireforget-all");
        /// convenience accessor for the process-wide configuration.</summary>
        public static string SyncWrite => Current.SyncWrite;

        /// <summary>Whether writes are dispatched asynchronously; convenience
        /// accessor for the process-wide configuration.</summary>
        public static bool AsyncWrites => Current.AsyncWrites;
    }
}

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

        private static ConfigRoot Load()
        {
            foreach (var path in CandidatePaths())
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                        continue;

                    var config = JsonConvert.DeserializeObject<ConfigRoot>(File.ReadAllText(path));
                    if (config != null)
                    {
                        logger.Info($"AppConfig: loaded configuration from {path}");
                        return Sanitize(config);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"AppConfig: error reading config file {path}");
                }
            }
            logger.Info("AppConfig: no configuration file found, using defaults");
            var fallback = Sanitize(new ConfigRoot());
            // The original no-file fallback used 1000ms for the Excel flush timer;
            // keep that legacy behavior when no configuration file exists.
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
            return config;
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
    }
}

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

    /// <summary>Timeout defaults used when a server does not declare its own.</summary>
    public class TimeoutConfig
    {
        /// <summary>Handshake/AUTH timeout (ms). A slow or VPN link needs more
        /// than the old 1s; default 5000.</summary>
        public int? connectTimeout { get; set; }
        /// <summary>Wait for a command reply (ms). Default 5000.</summary>
        public int? syncTimeout { get; set; }
    }

    /// <summary>A named server: a connection string plus optional per-server
    /// timeouts. In JSON, a value in "Servers" may be a plain string (shorthand,
    /// equivalent to this object with only <see cref="host"/>) or an object.</summary>
    [JsonConverter(typeof(ServerDefConverter))]
    public class ServerDef
    {
        public string host { get; set; }
        public int? connectTimeout { get; set; }
        public int? syncTimeout { get; set; }
    }

    /// <summary>Reads a "Servers" value that may be a string or an object.</summary>
    public class ServerDefConverter : JsonConverter
    {
        public override bool CanWrite => false;

        public override bool CanConvert(Type objectType) => objectType == typeof(ServerDef);

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.String)
                return new ServerDef { host = (string)reader.Value };
            if (reader.TokenType == JsonToken.StartObject)
            {
                // JObject.Load (not serializer.Deserialize) avoids recursing back
                // into this converter.
                var obj = JObject.Load(reader);
                return new ServerDef
                {
                    host = (string)obj["host"],
                    connectTimeout = (int?)obj["connectTimeout"],
                    syncTimeout = (int?)obj["syncTimeout"],
                };
            }
            throw new JsonSerializationException("a Servers value must be a string or an object");
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            => throw new NotSupportedException();
    }

    public class RTDConfig
    {
        public int RedisUpdateRateMs { get; set; } = 1000;
        public int ExcelUpdateRateMs { get; set; } = 100;
        public ENUMExcelUpdateStyle ExcelUpdateStyle { get; set; } = ENUMExcelUpdateStyle.Automatic;
        public long MessageCounterThreshold { get; set; } = 10000;
        public bool UseGetMultiple { get; set; } = true;
        /// <summary>In real-time mode, send at most one update per topic per
        /// ExcelUpdateRateMs window instead of one per incoming message
        /// (RTD-only, so it lives here). Default on. Superseded by an explicit
        /// <see cref="ConfigRoot.ConflationMs"/>.</summary>
        public bool CoalesceRealtimeUpdates { get; set; } = true;
    }

    public class ConfigRoot
    {
        /// <summary>The default server: a name defined in <see cref="Servers"/>,
        /// or a raw connection string used as-is. RTD and UDF both use it (there
        /// is no per-section host).</summary>
        public string Default { get; set; }
        /// <summary>Global timeout defaults for servers declared as a plain
        /// string (or that omit a field).</summary>
        public TimeoutConfig Timeout { get; set; }
        /// <summary>Named servers: each a string (shorthand) or an object.</summary>
        public Dictionary<string, ServerDef> Servers { get; set; }

        public RTDConfig RTD { get; set; }

        public bool UpdateCheck { get; set; } = true;
        /// <summary>Skip identical consecutive payloads before decoding/delivering
        /// (feeds republish unchanged values constantly). Default on.</summary>
        public bool SkipRepeatedMessages { get; set; } = true;
        /// <summary>Explicit real-time conflation window in milliseconds: the
        /// latest value of a topic is pushed to Excel at most once per window
        /// (0 = off, one push per message). When absent (null), the legacy
        /// <see cref="RTDConfig.CoalesceRealtimeUpdates"/> boolean decides.
        /// Sanitized to [0, Conflation.MaxWindowMs].</summary>
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

        /// <summary>Set by the loader when the file failed schema validation
        /// (a wrong VALUE on a known key). When true no Redis access happens and
        /// every UDF/RTD call returns this <see cref="RejectedReason"/> as a
        /// "#CONFIG" error. Not part of the JSON.</summary>
        [JsonIgnore]
        public bool Rejected { get; set; }
        /// <summary>Human-readable reason shown in the "#CONFIG" cell when
        /// <see cref="Rejected"/> is true.</summary>
        [JsonIgnore]
        public string RejectedReason { get; set; }
    }

    /// <summary>Raised for a configuration-level failure (rejected file); the
    /// cell shows a "#CONFIG: ..." line instead of a generic error.</summary>
    public sealed class ConfigException : Exception
    {
        public ConfigException(string message) : base(message) { }
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
            // Fail-closed on a wrong VALUE (error severity): the whole file is
            // rejected and every UDF/RTD call returns the "#CONFIG" reason. Unknown
            // /extra keys are warnings and do not reject (they are pruned anyway).
            var firstError = collected.FirstOrDefault(i => i.Severity == ConfigSeverity.Error);
            if (firstError != null)
            {
                var rejected = Sanitize(new ConfigRoot());
                rejected.Rejected = true;
                rejected.RejectedReason = firstError.ToString();
                return rejected;
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
                            if (config.Rejected)
                            {
                                // Name the FILE in the reason: with several candidate
                                // locations, a stale %USERPROFILE% copy shadowing a good
                                // C:\RedisExcel one must be identifiable from the cell.
                                config.RejectedReason = Path.GetFileName(path) + " (" + path + "): " + config.RejectedReason;
                                logger.Error($"AppConfig: configuration REJECTED from {path}: {config.RejectedReason} - every UDF/RTD call returns #CONFIG until it is fixed");
                            }
                            else
                                logger.Info(issues.Count == 0
                                    ? $"AppConfig: loaded configuration from {path}"
                                    : $"AppConfig: loaded configuration from {path} with {issues.Count} value(s) ignored");
                            return Sanitize(config);
                        }
                        // A file that exists but is not a JSON OBJECT (e.g. a broken
                        // root) must ALSO fail closed, not silently fall back to
                        // localhost - the same principle as a wrong value (#CONFIG).
                        logger.Error($"AppConfig: config file {path} is not a JSON object; every UDF/RTD call returns #CONFIG until it is fixed");
                        var badRoot = Sanitize(new ConfigRoot());
                        badRoot.Rejected = true;
                        badRoot.RejectedReason = Path.GetFileName(path) + " (" + path + "): the config root is not a JSON object";
                        return badRoot;
                    }
                    catch (Exception ex)
                    {
                        // Malformed JSON (the reader threw): fail closed too, so a
                        // typo cannot silently drop the user onto localhost.
                        logger.Error(ex, $"AppConfig: error parsing config file {path}; every UDF/RTD call returns #CONFIG until it is fixed");
                        var malformed = Sanitize(new ConfigRoot());
                        malformed.Rejected = true;
                        malformed.RejectedReason = Path.GetFileName(path) + " (" + path + "): invalid JSON (" + ex.Message + ")";
                        return malformed;
                    }
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

        public const int DefaultConnectTimeoutMs = 5000;
        public const int DefaultSyncTimeoutMs = 5000;

        internal static ConfigRoot Sanitize(ConfigRoot config)
        {
            config = config ?? new ConfigRoot();
            config.RTD = config.RTD ?? new RTDConfig();
            if (string.IsNullOrWhiteSpace(config.Default)) config.Default = DefaultHost;
            config.Timeout = config.Timeout ?? new TimeoutConfig();
            config.Servers = config.Servers ?? new Dictionary<string, ServerDef>(StringComparer.Ordinal);
            // Drop server entries that carry no host (a bad object): they can never
            // be resolved, so a name pointing at one is a dangling reference.
            foreach (var key in config.Servers.Where(kv => kv.Value == null || string.IsNullOrWhiteSpace(kv.Value.host)).Select(kv => kv.Key).ToList())
                config.Servers.Remove(key);
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

        /// <summary>Resolves a server reference (a "Servers" name or a raw
        /// connection string) to its connection string. Null/blank falls back to
        /// the configured default. Never throws.</summary>
        public static string ResolveServerHost(string reference)
        {
            return ResolveServer(Current, reference)?.host ?? DefaultHost;
        }

        /// <summary>The server definition for a reference (a name in "Servers" or
        /// a raw connection string), or the Default server when the reference is
        /// blank. Returns null when a NAME is used but not defined (dangling).</summary>
        internal static ServerDef ResolveServer(ConfigRoot config, string reference)
        {
            string name = string.IsNullOrWhiteSpace(reference) ? config?.Default : reference.Trim();
            if (string.IsNullOrWhiteSpace(name))
                return new ServerDef { host = DefaultHost };

            // A name defined in Servers?
            var def = FindServer(config?.Servers, name);
            if (def != null)
                return def;

            // Otherwise treat it as a raw connection string.
            return new ServerDef { host = name };
        }

        private static ServerDef FindServer(IDictionary<string, ServerDef> servers, string name)
        {
            if (servers == null)
                return null;
            if (servers.TryGetValue(name, out var exact) && exact != null)
                return exact;
            foreach (var kv in servers)
            {
                if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                    return kv.Value;
            }
            return null;
        }

        /// <summary>Both timeouts (ms) for a resolved server: the server's own
        /// value, else the global <c>Timeout</c> block, else the built-in default
        /// (5000 ms). The reference may be a <c>Servers</c> name OR the resolved
        /// connection string (the connection manager passes the flattened host, so
        /// the per-server timeouts are recovered by matching the server whose host
        /// equals it - otherwise the per-server timeout would be dead config).</summary>
        public static (int connectTimeout, int syncTimeout) ServerTimeouts(string reference)
        {
            var config = Current;
            ServerDef def;
            if (string.IsNullOrWhiteSpace(reference))
                def = ResolveServer(config, null);                       // the Default server (a name or a raw string)
            else
                def = FindServer(config?.Servers, reference)             // a Servers KEY (the alias text)
                      ?? FindServerByHost(config?.Servers, reference)    // a Servers VALUE's host (the resolved string)
                      ?? new ServerDef { host = reference };             // a raw connection string
            int connect = def?.connectTimeout
                ?? config?.Timeout?.connectTimeout
                ?? DefaultConnectTimeoutMs;
            int sync = def?.syncTimeout
                ?? config?.Timeout?.syncTimeout
                ?? DefaultSyncTimeoutMs;
            if (connect <= 0) connect = DefaultConnectTimeoutMs;
            if (sync <= 0) sync = DefaultSyncTimeoutMs;
            return (connect, sync);
        }

        /// <summary>Finds a server whose <c>host</c> equals the given (resolved)
        /// connection string, so a per-server timeout survives the name -> host
        /// flattening. When two servers share a host string the <c>Default</c>
        /// server wins (it is the one the connection actually belongs to),
        /// otherwise the first declared entry.</summary>
        private static ServerDef FindServerByHost(IDictionary<string, ServerDef> servers, string host)
        {
            if (servers == null || string.IsNullOrWhiteSpace(host))
                return null;
            // Prefer the Default server's entry: the caller has flattened Default
            // to this host string, so when two entries share the host the Default
            // one is the connection's true owner (audit D1/D2).
            var defaultDef = FindServer(servers, Current?.Default);
            if (defaultDef != null && string.Equals(defaultDef.host, host, StringComparison.Ordinal))
                return defaultDef;
            foreach (var kv in servers)
            {
                if (kv.Value != null && string.Equals(kv.Value.host, host, StringComparison.Ordinal))
                    return kv.Value;
            }
            return null;
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
        // whatever sits between '//' and the LAST '@' before the authority is
        // userinfo and must be redacted. Matching to the FIRST '@' leaked the
        // rest when the userinfo itself contained an '@' (e.g.
        // redis://user:p@leakpart@host); the greedy run up to the last '@' that
        // precedes a host character redacts the whole userinfo.
        private static readonly Regex UriUserInfo = new Regex(
            @"(?<=//)[^/\s]*@(?=[^@/\s]*$|/)", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // A credential keyword is preceded by the string start, a connection-string
        // delimiter (comma/semicolon) OR whitespace: a credential can reach a cell
        // or log line embedded after any whitespace, not only after a comma, and a
        // leaked secret is worse than over-masking a user identifier that happens to
        // look like "x=password=y" (audit round P).
        private static readonly Regex PasswordPair = new Regex(
            @"(^|[,;\s])\s*(password|pass|user|username)\s*=\s*(?:""[^""]*""|'[^']*'|[^,]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
                var errors = issues.Where(i => i.Severity == ConfigSeverity.Error).ToList();
                if (errors.Count == 0)
                {
                    if (issues.Count == 0)
                        return "OK (" + path + ")";
                    // Only warnings (e.g. unknown keys): the config still loads.
                    var warnLines = new List<string> { "OK (with " + issues.Count + " warning(s)) - " + path };
                    foreach (var issue in issues)
                        warnLines.Add("  warning: " + issue.ToString());
                    return string.Join("\n", warnLines);
                }

                var lines = new List<string> { "Error: " + errors.Count + " problem(s) in " + path };
                foreach (var issue in errors)
                    lines.Add(issue.ToString());
                foreach (var issue in issues.Where(i => i.Severity == ConfigSeverity.Warning))
                    lines.Add("  warning: " + issue.ToString());
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

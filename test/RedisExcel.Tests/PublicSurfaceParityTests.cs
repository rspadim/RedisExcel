using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Freezes the add-in's public Excel surface (AGENTS golden rule 2: "never
    /// change the public Excel surface"). It reflects over every production type
    /// that carries <c>[ExcelFunction]</c> and compares the result to the
    /// committed baseline <c>PublicSurface.baseline.txt</c>, so ANY rename,
    /// argument reorder, added/removed argument, optional/default change,
    /// description edit or volatility change FAILS the test until the baseline is
    /// intentionally refreshed.
    ///
    /// <para>
    /// The <b>shipped add-in assembly</b> (<c>bin\Release\net48\RedisExcel.dll</c>)
    /// is the baseline's source of truth: it holds the COMPLETE surface, including
    /// the <c>RedisRTD*</c> status helpers that live in <c>RedisRtd.cs</c> (not
    /// linked into the test assemblies). Like <see cref="DnaPackagingParityTests"/>,
    /// a Release build is a hard prerequisite (a missing assembly fails rather than
    /// skips). The baseline is therefore generated from the real code, never
    /// hand-written.
    /// </para>
    ///
    /// <para>
    /// <b>Refreshing the baseline.</b> A normal run never writes anything. To
    /// regenerate it from the real code (after an intentional, additive surface
    /// change), set <c>REDISEXCEL_REFRESH_PUBLIC_SURFACE=1</c> and run the unit
    /// suite; the test rewrites the file and passes. Review the diff and commit it
    /// on purpose.
    /// </para>
    /// </summary>
    public class PublicSurfaceParityTests
    {
        private const string BaselineFileName = "PublicSurface.baseline.txt";
        private const string RefreshEnvironmentVariable = "REDISEXCEL_REFRESH_PUBLIC_SURFACE";
        private const string TwinSuffix = "NonVolatile";
        private const string NoDescription = "(no description)";

        // The real, code-derived counts. Pinned so a rename that keeps the count
        // cannot slip through, and so the UDF/RTD/diagnostic split stays explicit.
        private const int ExpectedTotal = 79;
        private const int ExpectedUdf = 68;
        private const int ExpectedRtd = 9;
        private const int ExpectedDiagnostic = 2;
        private const int ExpectedTwins = 24;

        /// <summary>The nine frozen <c>RedisRTD*</c> status helpers, pinned by name
        /// so a rename cannot slip through the count alone.</summary>
        private static readonly string[] ExpectedRtdNames =
        {
            "RedisRTDChannelCount",
            "RedisRTDConnectionCount",
            "RedisRTDDefaultHost",
            "RedisRTDExcelUpdateInterval",
            "RedisRTDMessagesCounter",
            "RedisRTDRealTimeUpdates",
            "RedisRTDRedisUpdateInterval",
            "RedisRTDSubscriptionCount",
            "RedisRTDTopicCount",
        };

        /// <summary>The two non-UDF, non-RTD diagnostic functions, pinned by name.</summary>
        private static readonly string[] ExpectedDiagnosticNames =
        {
            "RedisCheckConfigFile",
            "RedisJsonSchemaContents",
        };

        [Fact]
        public void PublicSurface_MatchesTheCommittedBaseline()
        {
            List<FunctionEntry> shipped = CollectSurface(ShippedAssembly());

            List<string> actualLines = shipped
                .Select(entry => entry.Line)
                .OrderBy(line => line, StringComparer.Ordinal)
                .ToList();

            string baselinePath = BaselinePath();
            if (IsRefreshRequested())
            {
                WriteBaseline(baselinePath, actualLines);
                return;
            }

            Assert.True(File.Exists(baselinePath),
                "the public-surface baseline is missing: " + baselinePath
                + " (regenerate it with " + RefreshEnvironmentVariable + "=1)");

            List<string> baselineLines = ReadBaseline(baselinePath);
            Assert.True(baselineLines.SequenceEqual(actualLines, StringComparer.Ordinal),
                BuildDiff(baselineLines, actualLines));

            // The linked production sources must agree with the shipped assembly
            // for every RedisUDF* function; a disagreement means the build output
            // is stale relative to the sources.
            Dictionary<string, string> linked = CollectSurface(LinkedAssembly())
                .ToDictionary(entry => entry.Name, entry => entry.Line, StringComparer.Ordinal);

            foreach (FunctionEntry entry in shipped.Where(
                item => item.Name.StartsWith("RedisUDF", StringComparison.Ordinal)))
            {
                Assert.True(linked.TryGetValue(entry.Name, out string linkedLine),
                    "the shipped add-in declares '" + entry.Name
                    + "' but the linked production sources do not. Rebuild the add-in "
                    + "(`dotnet build RedisExcel.sln -c Release`) and re-run.");

                Assert.True(string.Equals(entry.Line, linkedLine, StringComparison.Ordinal),
                    "the shipped add-in and the linked production sources disagree for '"
                    + entry.Name + "'. Rebuild the add-in "
                    + "(`dotnet build RedisExcel.sln -c Release`) and re-run.\n"
                    + "  shipped: " + entry.Line + "\n"
                    + "  linked : " + linkedLine);
            }
        }

        [Fact]
        public void PublicSurface_CountsAndSplitArePinned()
        {
            List<FunctionEntry> shipped = CollectSurface(ShippedAssembly());
            Dictionary<string, FunctionEntry> byName = shipped
                .ToDictionary(entry => entry.Name, entry => entry, StringComparer.Ordinal);

            int total = shipped.Count;
            int udf = shipped.Count(entry => entry.Name.StartsWith("RedisUDF", StringComparison.Ordinal));
            int rtd = shipped.Count(entry => entry.Name.StartsWith("RedisRTD", StringComparison.Ordinal));
            int diagnostic = shipped.Count(entry => !entry.Name.StartsWith("RedisUDF", StringComparison.Ordinal)
                && !entry.Name.StartsWith("RedisRTD", StringComparison.Ordinal));

            List<FunctionEntry> twins = shipped
                .Where(entry => entry.Name.EndsWith(TwinSuffix, StringComparison.Ordinal))
                .ToList();

            Assert.Equal(ExpectedTotal, total);
            Assert.Equal(ExpectedUdf, udf);
            Assert.Equal(ExpectedRtd, rtd);
            Assert.Equal(ExpectedDiagnostic, diagnostic);
            Assert.Equal(ExpectedTwins, twins.Count);

            // The split must be exhaustive and disjoint: every function is either
            // a RedisUDF*, a RedisRTD* or one of the two pinned diagnostics.
            Assert.Equal(total, udf + rtd + diagnostic);

            // Pin the nine RTD helper names, not just the count.
            string[] rtdNames = shipped
                .Where(entry => entry.Name.StartsWith("RedisRTD", StringComparison.Ordinal))
                .Select(entry => entry.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(
                ExpectedRtdNames.OrderBy(name => name, StringComparer.Ordinal).ToArray(),
                rtdNames);

            // Pin the diagnostic names too.
            string[] diagnosticNames = shipped
                .Where(entry => !entry.Name.StartsWith("RedisUDF", StringComparison.Ordinal)
                    && !entry.Name.StartsWith("RedisRTD", StringComparison.Ordinal))
                .Select(entry => entry.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(
                ExpectedDiagnosticNames.OrderBy(name => name, StringComparer.Ordinal).ToArray(),
                diagnosticNames);

            // Every twin has a base with an IDENTICAL argument signature, the base
            // is volatile and the twin is not (a twin whose signature differs from
            // its base is a real surface defect, not test noise).
            foreach (FunctionEntry twin in twins)
            {
                string baseName = twin.Name.Substring(0, twin.Name.Length - TwinSuffix.Length);
                Assert.True(byName.ContainsKey(baseName),
                    "the NonVolatile twin '" + twin.Name + "' has no matching base '" + baseName + "'");

                FunctionEntry @base = byName[baseName];
                Assert.True(string.Equals(@base.ArgsSignature, twin.ArgsSignature, StringComparison.Ordinal),
                    "the twin '" + twin.Name + "' has a different argument signature from its base '"
                    + baseName + "'.\n  base : " + @base.ArgsSignature + "\n  twin : " + twin.ArgsSignature);

                Assert.True(@base.IsVolatile, "the base '" + baseName + "' of a NonVolatile twin must be volatile");
                Assert.False(twin.IsVolatile, "the twin '" + twin.Name + "' must not be volatile");
            }
        }

        /// <summary>One <c>[ExcelFunction]</c> as it appears in the baseline.</summary>
        private sealed class FunctionEntry
        {
            public string Name;
            public string Line;
            public string ArgsSignature;
            public bool IsVolatile;

            public static FunctionEntry Create(MethodInfo method, ExcelFunctionAttribute attribute)
            {
                ParameterInfo[] parameters = method.GetParameters();
                string args = parameters.Length == 0
                    ? "(no arguments)"
                    : string.Join("; ", parameters.Select(FormatParameter));

                string line = method.Name
                    + " | volatile=" + (attribute.IsVolatile ? "true" : "false")
                    + " | " + args
                    + " | \"" + Escape(attribute.Description ?? NoDescription) + "\"";

                return new FunctionEntry
                {
                    Name = method.Name,
                    Line = line,
                    ArgsSignature = args,
                    IsVolatile = attribute.IsVolatile,
                };
            }

            private static string FormatParameter(ParameterInfo parameter)
            {
                var argumentAttribute = parameter
                    .GetCustomAttributes(typeof(ExcelArgumentAttribute), false)
                    .OfType<ExcelArgumentAttribute>()
                    .FirstOrDefault();

                string optional = string.Empty;
                if (parameter.IsOptional)
                {
                    optional = parameter.HasDefaultValue
                        ? "=" + FormatDefault(parameter.DefaultValue)
                        : "=?";
                }

                return parameter.Name
                    + ":" + TypeName(parameter.ParameterType)
                    + optional
                    + " \"" + Escape(argumentAttribute?.Description ?? NoDescription) + "\"";
            }

            private static string TypeName(Type type)
            {
                if (type == typeof(object))
                    return "object";
                if (type == typeof(string))
                    return "string";
                if (type == typeof(double))
                    return "double";
                if (type == typeof(int))
                    return "int";
                if (type == typeof(bool))
                    return "bool";
                if (type.IsArray)
                    return TypeName(type.GetElementType()) + "[" + new string(',', type.GetArrayRank() - 1) + "]";
                return type.Name;
            }

            private static string FormatDefault(object value)
            {
                if (value == null)
                    return "null";
                if (value is string text)
                    return "\"" + Escape(text) + "\"";
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "(null)";
            }

            private static string Escape(string value)
            {
                return value
                    .Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\r", "\\r")
                    .Replace("\n", "\\n");
            }
        }

        private static List<FunctionEntry> CollectSurface(Assembly assembly)
        {
            var entries = new List<FunctionEntry>();
            Type attributeType = typeof(ExcelFunctionAttribute);

            foreach (Type type in SafeGetTypes(assembly))
            {
                foreach (MethodInfo method in type.GetMethods(
                    BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    var attributes = method.GetCustomAttributes(attributeType, false);
                    if (attributes.Length == 0)
                        continue;

                    entries.Add(FunctionEntry.Create(method, (ExcelFunctionAttribute)attributes[0]));
                }
            }

            return entries;
        }

        private static IEnumerable<Type> SafeGetTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(type => type != null);
            }
        }

        private static Assembly LinkedAssembly()
        {
            return typeof(PublicSurfaceParityTests).Assembly;
        }

        private static Assembly ShippedAssembly()
        {
            string path = Path.Combine(
                FindRepoRoot(), "bin", "Release", "net48", "RedisExcel.dll");
            Assert.True(File.Exists(path),
                "a Release build of the add-in is required to freeze the public surface "
                + "(including the RedisRtd.cs helpers): run "
                + "`dotnet build RedisExcel.sln -c Release` first. Missing: " + path);
            return Assembly.LoadFrom(path);
        }

        private static bool IsRefreshRequested()
        {
            string value = Environment.GetEnvironmentVariable(RefreshEnvironmentVariable);
            return !string.IsNullOrEmpty(value)
                && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        private static string BaselinePath()
        {
            return Path.Combine(
                FindRepoRoot(), "test", "RedisExcel.Tests", BaselineFileName);
        }

        private static IEnumerable<string> BaselineHeader()
        {
            yield return "# RedisExcel public Excel surface baseline.";
            yield return "# Frozen at the v1.6.0 release; additive changes only (AGENTS golden rule 2).";
            yield return "# Generated from the built add-in assembly + the linked production sources;";
            yield return "# never hand-write this file. Checked by PublicSurfaceParityTests.";
            yield return "# Refresh deliberately: set " + RefreshEnvironmentVariable + "=1 and run";
            yield return "#   dotnet test test\\RedisExcel.Tests\\RedisExcel.Tests.csproj -c Release";
            yield return "#";
            yield return "# Every non-comment line is:";
            yield return "#   <name> | volatile=<true|false> | <param:type[=default] \"<arg description>\"; ...> | \"<function description>\"";
            yield return "# Sorted ordinal by function name. No machine-specific data.";
        }

        private static void WriteBaseline(string path, List<string> lines)
        {
            string content = string.Join("\n", BaselineHeader().Concat(lines)) + "\n";
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, new UTF8Encoding(false));
        }

        private static List<string> ReadBaseline(string path)
        {
            return File.ReadAllLines(path)
                .Select(line => line.TrimEnd('\r'))
                .Where(line => line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal))
                .ToList();
        }

        private static Dictionary<string, string> ToNameMap(IEnumerable<string> lines)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string line in lines)
            {
                int separator = line.IndexOf(" | ", StringComparison.Ordinal);
                string name = separator < 0 ? line : line.Substring(0, separator);
                map[name] = line;
            }
            return map;
        }

        private static string BuildDiff(List<string> baseline, List<string> actual)
        {
            Dictionary<string, string> before = ToNameMap(baseline);
            Dictionary<string, string> after = ToNameMap(actual);

            var builder = new StringBuilder();
            builder.AppendLine("The public Excel surface drifted from " + BaselineFileName + ".");
            builder.AppendLine("If the change is intentional (additive only!), refresh the baseline: set "
                + RefreshEnvironmentVariable + "=1 and re-run this test.");

            foreach (string name in after.Keys.Except(before.Keys).OrderBy(key => key, StringComparer.Ordinal))
                builder.AppendLine("  ADDED   " + after[name]);

            foreach (string name in before.Keys.Except(after.Keys).OrderBy(key => key, StringComparer.Ordinal))
                builder.AppendLine("  REMOVED " + before[name]);

            foreach (string name in after.Keys.Intersect(before.Keys).OrderBy(key => key, StringComparer.Ordinal))
            {
                if (string.Equals(after[name], before[name], StringComparison.Ordinal))
                    continue;
                builder.AppendLine("  CHANGED " + name);
                builder.AppendLine("    baseline: " + before[name]);
                builder.AppendLine("    actual  : " + after[name]);
            }

            return builder.ToString();
        }

        /// <summary>Walks up from the test binary until the repository root,
        /// anchored by <c>RedisExcel.sln</c>. Mirrors
        /// <see cref="ProjectFileParityTests"/>.</summary>
        private static string FindRepoRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "RedisExcel.sln")))
                    return directory.FullName;
                directory = directory.Parent;
            }
            throw new InvalidOperationException(
                "RedisExcel.sln was not found above " + AppContext.BaseDirectory);
        }
    }
}

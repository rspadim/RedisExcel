using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Guards the SINGLE production-source list every test project imports
    /// (test\TestSupport\ProductionSources.props): the suites compile the
    /// production sources directly (no ProjectReference), so a new root
    /// production file silently missing there would never be exercised by any
    /// suite. RedisRtd.cs is the only intentional exclusion: its RTD/COM surface
    /// is covered by the E2E script and the shipped-assembly public-surface test.
    /// </summary>
    public class ProjectFileParityTests
    {
        [Fact]
        public void EveryTestProject_ImportsTheSharedProductionSourceList()
        {
            string repoRoot = FindRepoRoot();
            foreach (string project in new[]
            {
                Path.Combine("test", "RedisExcel.Tests", "RedisExcel.Tests.csproj"),
                Path.Combine("test", "SmokeTests", "RedisExcel.SmokeTests.csproj"),
                Path.Combine("test", "LivenessTests", "LivenessTests.csproj"),
                Path.Combine("test", "LoadTests", "RedisExcel.LoadTests.csproj"),
            })
            {
                string path = Path.Combine(repoRoot, project);
                Assert.True(File.Exists(path), "test project not found at " + path);
                var document = XDocument.Load(path);
                XNamespace ns = document.Root.Name.Namespace;
                bool imports = document.Descendants(ns + "Import")
                    .Any(item => ((string)item.Attribute("Project") ?? string.Empty)
                        .IndexOf("ProductionSources.props", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.True(imports,
                    project + " must import ..\\TestSupport\\ProductionSources.props "
                    + "(the single production-source list; a per-project copy drifts)");

                // A hand list in the project itself is what we removed: it must not
                // come back (the props is the only place the sources are declared).
                bool hasOwnProductionCompile = document.Descendants(ns + "Compile")
                    .Any(item => ((string)item.Attribute("Include") ?? string.Empty)
                        .Replace('\\', '/').StartsWith("../../", StringComparison.Ordinal));
                Assert.False(hasOwnProductionCompile,
                    project + " must not declare production sources itself; add them to "
                    + "ProductionSources.props instead");
            }
        }

        [Fact]
        public void SharedList_CompilesEveryRootProductionSourceExceptRedisRtd()
        {
            string repoRoot = FindRepoRoot();
            string propsPath = Path.Combine(repoRoot, "test", "TestSupport", "ProductionSources.props");
            Assert.True(File.Exists(propsPath), "shared props not found at " + propsPath);

            string propsDirectory = Path.GetDirectoryName(propsPath);
            var compiled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var document = XDocument.Load(propsPath);
            XNamespace ns = document.Root.Name.Namespace;
            foreach (var compile in document.Descendants(ns + "Compile"))
            {
                var include = (string)compile.Attribute("Include");
                if (string.IsNullOrWhiteSpace(include))
                    continue;
                // The props uses $(MSBuildThisFileDirectory) for the repo root.
                string normalized = include
                    .Replace("$(MSBuildThisFileDirectory)", propsDirectory + Path.DirectorySeparatorChar)
                    .Replace('\\', Path.DirectorySeparatorChar)
                    .Replace('/', Path.DirectorySeparatorChar);
                compiled.Add(Path.GetFullPath(normalized));
            }

            // Every top-level production source (the root is the only place
            // production .cs files live) must be listed, except the deliberately
            // excluded RTD/COM file.
            List<string> required = Directory.GetFiles(repoRoot, "*.cs", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFullPath)
                .Where(path => !string.Equals(
                    Path.GetFileName(path), "RedisRtd.cs", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<string> missing = required.Where(path => !compiled.Contains(path)).ToList();
            Assert.True(missing.Count == 0,
                "production sources missing from test\\TestSupport\\ProductionSources.props: "
                + string.Join(", ", missing));

            // The exclusion is intentional: adding RedisRtd.cs here would pull
            // the RTD/COM surface into every suite and mask that gap.
            Assert.DoesNotContain(
                Path.GetFullPath(Path.Combine(repoRoot, "RedisRtd.cs")), compiled);
        }

        /// <summary>
        /// RedisRtd.cs is not compiled by any test project, so a regression in a
        /// source-level invariant there would ship silently (audit E3). Pin the
        /// critical one: the RTD polled-command drain must bound its batch wait by
        /// the TOPIC host's syncTimeout, not the Default server's - otherwise a
        /// reply arriving within the connection's own timeout is abandoned.
        /// </summary>
        [Fact]
        public void RedisRtd_PollDrain_UsesTheTopicHostBound()
        {
            string repoRoot = FindRepoRoot();
            string source = File.ReadAllText(Path.Combine(repoRoot, "RedisRtd.cs"));

            Assert.Contains("int responseTimeoutMs = RedisUDF.ResponseTimeoutMs(host);", source);
            Assert.DoesNotContain("int responseTimeoutMs = RedisUDF.ResponseTimeoutMs();", source);
        }

        /// <summary>
        /// Resolves one Compile Include value to absolute paths, expanding the
        /// simple globs the csproj format allows (for example "..\..\*.cs").
        /// </summary>
        private static IEnumerable<string> ExpandCompileItem(string projectDirectory, string include)
        {
            string normalized = include
                .Replace('\\', Path.DirectorySeparatorChar)
                .Replace('/', Path.DirectorySeparatorChar);
            if (normalized.IndexOf('*') < 0 && normalized.IndexOf('?') < 0)
            {
                yield return Path.GetFullPath(Path.Combine(projectDirectory, normalized));
                yield break;
            }

            string directoryPart = Path.GetDirectoryName(normalized);
            bool recursive = directoryPart != null
                && directoryPart.IndexOf("**", StringComparison.Ordinal) >= 0;
            if (recursive)
                directoryPart = directoryPart.Replace("**", ".");
            string searchDirectory = Path.GetFullPath(Path.Combine(
                projectDirectory,
                string.IsNullOrEmpty(directoryPart) ? "." : directoryPart));
            string filePattern = Path.GetFileName(normalized);

            foreach (string file in Directory.GetFiles(
                searchDirectory,
                filePattern,
                recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
            {
                yield return Path.GetFullPath(file);
            }
        }

        /// <summary>Walks up from the test binary until the solution root.</summary>
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

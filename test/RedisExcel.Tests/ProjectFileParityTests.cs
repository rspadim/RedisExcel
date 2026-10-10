using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Guards the hand-maintained Compile list in RedisExcel.Tests.csproj: the
    /// unit tests compile the production sources directly (no ProjectReference),
    /// so a new root production file silently excluded there would never be
    /// exercised by the unit suite. RedisRtd.cs is the only intentional
    /// exclusion: its RTD/COM surface is covered by the E2E script instead.
    /// </summary>
    public class ProjectFileParityTests
    {
        [Fact]
        public void TestsProject_CompilesEveryRootProductionSourceExceptRedisRtd()
        {
            string repoRoot = FindRepoRoot();
            string projectPath = Path.Combine(
                repoRoot, "test", "RedisExcel.Tests", "RedisExcel.Tests.csproj");
            Assert.True(File.Exists(projectPath), "test project not found at " + projectPath);

            string projectDirectory = Path.GetDirectoryName(projectPath);
            var compiled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var project = XDocument.Load(projectPath);
            XNamespace ns = project.Root.Name.Namespace;
            foreach (var compile in project.Descendants(ns + "Compile"))
            {
                var include = (string)compile.Attribute("Include");
                if (string.IsNullOrWhiteSpace(include))
                    continue; // Remove/Update items carry no Include
                foreach (string file in ExpandCompileItem(projectDirectory, include))
                    compiled.Add(file);
            }

            // Every top-level production source (the root is the only place
            // production .cs files live) must be compiled by the test project,
            // except the deliberately excluded RTD/COM file.
            List<string> required = Directory.GetFiles(repoRoot, "*.cs", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFullPath)
                .Where(path => !string.Equals(
                    Path.GetFileName(path), "RedisRtd.cs", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<string> missing = required.Where(path => !compiled.Contains(path)).ToList();
            Assert.True(missing.Count == 0,
                "production sources missing from RedisExcel.Tests.csproj's Compile list: "
                + string.Join(", ", missing));

            // The exclusion is intentional: adding RedisRtd.cs here would pull
            // the RTD/COM surface into the unit suite and mask that gap.
            Assert.DoesNotContain(
                Path.GetFullPath(Path.Combine(repoRoot, "RedisRtd.cs")), compiled);
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

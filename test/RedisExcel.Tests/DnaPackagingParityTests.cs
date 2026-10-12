using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Guards the packed XLL against a silent missing dependency: <c>RedisExcel.dna</c>
    /// lists the assemblies the ExcelDna packer embeds with <c>Pack="true"</c>. A
    /// dependency that is referenced by the add-in but absent from that list is NOT
    /// packed, so at runtime the first use throws FileNotFoundException - and the
    /// config loader catches everything and silently falls back to defaults (the
    /// exact "lost my RedisExcel.json" failure). This test fails the build when a
    /// shipped assembly is not in the .dna list.
    /// </summary>
    public class DnaPackagingParityTests
    {
        /// <summary>Assemblies in the build output that must NOT be packed (the
        /// add-in itself is the ExternalLibrary; test-only/private ones are not
        /// shipped).</summary>
        private static readonly HashSet<string> Excluded = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "RedisExcel.dll",           // the ExternalLibrary, not a Reference
            "ExcelDna.Integration.dll", // provided by the ExcelDna host
        };

        [Fact]
        public void EveryShippedDependencyIsPackedInTheDna()
        {
            string repoRoot = FindRepoRoot();
            string outputDir = Path.Combine(repoRoot, "bin", "Release", "net48");
            string dnaPath = Path.Combine(repoRoot, "RedisExcel.dna");
            Assert.True(File.Exists(dnaPath), "RedisExcel.dna not found at " + dnaPath);

            if (!Directory.Exists(outputDir) || Directory.GetFiles(outputDir, "*.dll").Length == 0)
            {
                // The check compares the packed list against the actual build output;
                // without a built add-in there is nothing to compare, so it is
                // inconclusive here. CI builds the add-in before the unit step, so it
                // always runs there. Build locally with
                // `dotnet build RedisExcel.sln -c Release` to exercise it.
                Console.WriteLine("SKIP DnaPackagingParityTests: build the add-in first " +
                    "(dotnet build RedisExcel.sln -c Release); no output in " + outputDir);
                return;
            }

            // DLLs that the build copied next to the add-in (its real dependencies).
            var inOutput = new HashSet<string>(
                Directory.GetFiles(outputDir, "*.dll").Select(Path.GetFileName),
                StringComparer.OrdinalIgnoreCase);
            Assert.NotEmpty(inOutput);

            var packed = new HashSet<string>(
                XDocument.Load(dnaPath).Descendants("Reference")
                    .Where(e => IsPacked(e))
                    .Select(e => (string)e.Attribute("Path"))
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Select(Path.GetFileName),
                StringComparer.OrdinalIgnoreCase);

            var missing = inOutput
                .Where(dll => !Excluded.Contains(dll) && !packed.Contains(dll))
                .OrderBy(dll => dll, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Assert.True(missing.Count == 0,
                "assemblies in bin\\Release\\net48 but NOT packed in RedisExcel.dna " +
                "(they would be missing from the XLL and cause a silent runtime failure): "
                + string.Join(", ", missing));
        }

        /// <summary>A &lt;Reference&gt; is embedded in the XLL only when Pack is true
        /// (ExcelDna's default is false). A dependency marked <c>Pack="false"</c> is
        /// NOT shipped, so it must not count as packed (a mutation that flips it
        /// would otherwise leave the parity test green - audit round P).</summary>
        private static bool IsPacked(XElement reference)
        {
            var pack = (string)reference.Attribute("Pack");
            return string.Equals(pack, "true", StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DnaListsNoAssemblyThatIsNotShipped()
        {
            // The inverse: a stale .dna entry for a DLL no longer produced would make
            // the packer fail (or ship a wrong version).
            string repoRoot = FindRepoRoot();
            string outputDir = Path.Combine(repoRoot, "bin", "Release", "net48");
            Assert.True(Directory.Exists(outputDir),
                "the build output is missing; build the add-in first " +
                "(dotnet build RedisExcel.sln -c Release); no output in " + outputDir);
            var inOutput = new HashSet<string>(
                Directory.GetFiles(outputDir, "*.dll").Select(Path.GetFileName),
                StringComparer.OrdinalIgnoreCase);

            var packed = XDocument.Load(Path.Combine(repoRoot, "RedisExcel.dna")).Descendants("Reference")
                .Select(e => Path.GetFileName((string)e.Attribute("Path")))
                .Where(p => !string.IsNullOrEmpty(p))
                .ToList();

            var stale = packed
                .Where(dll => !inOutput.Contains(dll) && !Excluded.Contains(dll))
                .OrderBy(dll => dll, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Assert.True(stale.Count == 0,
                "RedisExcel.dna references assemblies that are not in the build output: "
                + string.Join(", ", stale));
        }

        private static string FindRepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                if (File.Exists(Path.Combine(dir.FullName, "RedisExcel.sln")))
                    return dir.FullName;
                dir = dir.Parent;
            }
            throw new InvalidOperationException("RedisExcel.sln not found above " + AppContext.BaseDirectory);
        }
    }
}

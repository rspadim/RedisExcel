using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// The unbounded-blocking tripwire: the Excel thread must never wait on a
    /// task without a bound. Every synchronous wait in the root production
    /// sources is a defect UNLESS it sits inside one of the few deliberately
    /// bounded/allowed sites below - a bounded wait helper, a continuation that
    /// only runs after its antecedent completed, a background thread with a
    /// client timeout, or the one-off local schema parse.
    ///
    /// <para>
    /// This is a pure source scan (no build, no Redis, no Excel) so it runs
    /// everywhere and cannot be defeated by a runtime path that never executes.
    /// Comments and string literals are stripped first, so prose never
    /// false-positives. Only the root <c>*.cs</c> files are scanned:
    /// <c>test\</c> and <c>tools\</c> legitimately block (their entry point is
    /// not the Excel thread).
    /// </para>
    ///
    /// <para>
    /// The block patterns are <c>.Result</c>, <c>.Wait(</c> and
    /// <c>.GetAwaiter().GetResult()</c>. Adding a new blocking site means adding
    /// an entry here WITH the reason - the allow-list is the audit trail.
    /// </para>
    /// </summary>
    public class NoUnboundedBlockingTests
    {
        /// <summary>Every blocking site we deliberately accept, with the reason.
        /// A blocking call is allowed only between the named method's declaration
        /// and its matching close brace in that file.</summary>
        private static readonly (string File, string Method, string Reason)[] Allowed =
        {
            ("RedisUDF.cs", "WaitBounded",
                "the ONE bounded wait primitive: task.Wait(timeoutMs)"),
            ("RedisUDF.cs", "WaitBoundedResult",
                "reads a task that WaitBounded just confirmed completed"),
            ("RedisRtd.cs", "DrainBatch",
                "reads each polled task right after WaitBounded confirmed it"),
            ("RedisUdfAsync.cs", "Subscribe",
                "inner.Result inside ContinueWith, which runs only after the antecedent completed"),
            ("RedisUdfAsync.cs", "ResultForExcel",
                "queued.Result guarded by Status == RanToCompletion"),
            ("UpdateCheck.cs", "Refresh",
                "runs on Task.Run (not the Excel thread) with HttpClient.Timeout = 10s"),
            ("ConfigSchema.cs", "LoadSchema",
                "one-off parse of the embedded schema text (no I/O); named for this allow-list"),
        };

        private static readonly Regex BlockingPattern = new Regex(
            @"\.Result\b|\.Wait\(|GetAwaiter\(\)\.GetResult\(\)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        [Fact]
        public void RootProductionSources_ContainNoUnboundedBlockingOutsideTheAllowedSites()
        {
            string repoRoot = FindRepoRoot();
            string[] sources = Directory.GetFiles(repoRoot, "*.cs", SearchOption.TopDirectoryOnly);

            var offenders = new List<string>();
            var allowedSeen = new List<string>();

            foreach (string source in sources)
            {
                string fileName = Path.GetFileName(source);
                string[] lines = StripCommentsAndStrings(File.ReadAllLines(source));

                // The line ranges (1-based, inclusive) that this file may block in.
                var ranges = new List<(int Start, int End)>();
                foreach (var entry in Allowed.Where(item =>
                    string.Equals(item.File, fileName, StringComparison.OrdinalIgnoreCase)))
                {
                    var range = MethodRange(lines, entry.Method);
                    if (range == null)
                    {
                        offenders.Add(fileName + ": the allow-list names method '"
                            + entry.Method + "' but it was not found (renamed? update the allow-list)");
                        continue;
                    }
                    ranges.Add(range.Value);
                    allowedSeen.Add(fileName + "." + entry.Method);
                }

                for (int index = 0; index < lines.Length; index++)
                {
                    if (!BlockingPattern.IsMatch(lines[index]))
                        continue;

                    int lineNumber = index + 1;
                    if (ranges.Any(range => lineNumber >= range.Start && lineNumber <= range.End))
                        continue;

                    offenders.Add(fileName + ":" + lineNumber + ": " + lines[index].Trim());
                }
            }

            // Every allow-list entry must actually exist (a stale entry would
            // silently widen the allow-list after a rename).
            string[] expectedSeen = Allowed
                .Select(item => item.File + "." + item.Method)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(expectedSeen, allowedSeen.OrderBy(name => name, StringComparer.Ordinal).ToArray());

            Assert.True(offenders.Count == 0,
                "unbounded blocking in the root production sources (the Excel thread must never "
                + "wait without a bound). If the site is deliberate, add it to the allow-list in "
                + "NoUnboundedBlockingTests with its reason:\n  "
                + string.Join("\n  ", offenders));
        }

        /// <summary>Finds a method declaration line (by name) and its matching
        /// close brace, ignoring braces inside comments/strings (already
        /// stripped). Returns null when the method is not found.</summary>
        private static (int Start, int End)? MethodRange(string[] strippedLines, string methodName)
        {
            var declaration = new Regex(
                @"(?:^|\s)" + Regex.Escape(methodName) + @"\s*(?:<[^>]*>)?\s*\(",
                RegexOptions.CultureInvariant);

            for (int index = 0; index < strippedLines.Length; index++)
            {
                if (!declaration.IsMatch(strippedLines[index]))
                    continue;
                // Skip a CALL site (e.g. `Refresh();`): a declaration does not end
                // with a semicolon.
                if (strippedLines[index].TrimEnd().EndsWith(";", StringComparison.Ordinal))
                    continue;
                // A declaration may be an interface member (no body) or the call
                // itself; find the first '{' at or after it.
                int depth = 0;
                bool opened = false;
                for (int scan = index; scan < strippedLines.Length; scan++)
                {
                    foreach (char c in strippedLines[scan])
                    {
                        if (c == '{')
                        {
                            depth++;
                            opened = true;
                        }
                        else if (c == '}')
                        {
                            depth--;
                            if (opened && depth == 0)
                                return (index + 1, scan + 1);
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>Removes // line comments, /* */ block comments and the CONTENT
        /// of string/char literals so prose and messages never match a pattern
        /// (replaced with spaces to keep the line/column shape irrelevant).</summary>
        private static string[] StripCommentsAndStrings(string[] lines)
        {
            var result = new string[lines.Length];
            bool inBlockComment = false;
            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index];
                var builder = new StringBuilder(line.Length);
                bool inString = false;
                bool inChar = false;
                bool inLineComment = false;
                for (int i = 0; i < line.Length; i++)
                {
                    char c = line[i];
                    char next = i + 1 < line.Length ? line[i + 1] : '\0';

                    if (inLineComment)
                    {
                        builder.Append(' ');
                        continue;
                    }
                    if (inBlockComment)
                    {
                        builder.Append(' ');
                        if (c == '*' && next == '/')
                        {
                            inBlockComment = false;
                            builder.Append(' ');
                            i++;
                        }
                        continue;
                    }
                    if (inString)
                    {
                        builder.Append(' ');
                        if (c == '\\' && next != '\0')
                        {
                            builder.Append(' ');
                            i++;
                        }
                        else if (c == '"')
                        {
                            inString = false;
                        }
                        continue;
                    }
                    if (inChar)
                    {
                        builder.Append(' ');
                        if (c == '\\' && next != '\0')
                        {
                            builder.Append(' ');
                            i++;
                        }
                        else if (c == '\'')
                        {
                            inChar = false;
                        }
                        continue;
                    }

                    if (c == '/' && next == '/')
                    {
                        inLineComment = true;
                        builder.Append(' ');
                        i++;
                        continue;
                    }
                    if (c == '/' && next == '*')
                    {
                        inBlockComment = true;
                        builder.Append(' ');
                        i++;
                        continue;
                    }
                    if (c == '"')
                    {
                        inString = true;
                        builder.Append(' ');
                        continue;
                    }
                    if (c == '\'')
                    {
                        inChar = true;
                        builder.Append(' ');
                        continue;
                    }
                    builder.Append(c);
                }
                result[index] = builder.ToString();
            }
            return result;
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

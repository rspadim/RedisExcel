using System;
using System.Collections.Generic;
using System.Management;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Reads the registered antivirus state from the Windows Security Center.
    /// Security software is the top cause of flaky Excel automation (see
    /// AGENTS.md) and can also slow file/network I/O, so every suite prints it
    /// up front: a run's result is then attributable to the environment. The
    /// lookup is best-effort (some domain policies hide SecurityCenter2) and
    /// never throws into a test.
    /// </summary>
    internal static class AntivirusInfo
    {
        /// <summary>Names of the registered products that are ACTIVE (the
        /// productState 0x?000 nibble is non-zero). Empty when none is
        /// reported or the query is unavailable.</summary>
        public static IList<string> ActiveProductNames()
        {
            var active = new List<string>();
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    @"root\SecurityCenter2", "SELECT displayName, productState FROM AntiVirusProduct"))
                using (var results = searcher.Get())
                {
                    foreach (ManagementBaseObject product in results)
                    {
                        using (product)
                        {
                            var state = product["productState"];
                            if (state == null)
                                continue;
                            int value = Convert.ToInt32(state);
                            int level = (value >> 12) & 0xF;
                            if (level >= 1)
                                active.Add(Convert.ToString(product["displayName"]));
                        }
                    }
                }
            }
            catch
            {
                // Unavailable (policy, WMI failure): report nothing rather than fail.
            }
            return active;
        }

        /// <summary>One-line status for the suite banner.</summary>
        public static string Describe()
        {
            IList<string> active = ActiveProductNames();
            return active.Count == 0
                ? "none reported by Windows Security Center"
                : "ACTIVE: " + string.Join(", ", active) + " (may slow I/O and flake COM)";
        }
    }
}

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
        /// <summary>Names of the registered products that are ACTIVE. The
        /// productState bit 12 (0x1000) is the real-time/on-access flag
        /// (authoritative decode: 0x41000 = enabled, 0x42000 = paused), so test
        /// that bit rather than the whole nibble. Empty when none is reported
        /// or the query is unavailable.</summary>
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
                            if ((value & 0x1000) != 0)
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

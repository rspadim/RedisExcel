using System;
using System.Reflection;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Pins the Excel-DNA integration build the async write dispatch was
    /// audited against: the lock-order and threading findings behind
    /// <see cref="RedisWriteObservable"/> / ExcelAsyncUtil.Observe were verified
    /// for this exact revision. A framework bump must be a deliberate change
    /// that re-runs that audit, so any version drift fails loudly here.
    /// </summary>
    public class ExcelDnaPinTests
    {
        [Fact]
        public void ExcelDnaIntegration_InformationalVersion_IsPinned()
        {
            var attribute = (AssemblyInformationalVersionAttribute)Attribute.GetCustomAttribute(
                typeof(ExcelAsyncUtil).Assembly,
                typeof(AssemblyInformationalVersionAttribute));

            Assert.NotNull(attribute);
            Assert.StartsWith("1.9.0.6+c631da4e", attribute.InformationalVersion);
        }
    }
}

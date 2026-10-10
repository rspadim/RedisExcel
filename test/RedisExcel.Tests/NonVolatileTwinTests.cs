using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;
using Xunit;

namespace RedisExcel.Tests
{
    /// <summary>
    /// Offline reflection checks for the 24 non-volatile write twins added in
    /// v1.3.0: every twin must exist, mirror the volatile base signature
    /// (parameter types, count, optionality and return type) and must not be
    /// marked volatile, so Excel evaluates it only on entry and when an
    /// argument changes.
    /// </summary>
    public class NonVolatileTwinTests
    {
        /// <summary>The (volatile base, non-volatile twin) UDF name pairs.</summary>
        public static IEnumerable<object[]> TwinPairs()
        {
            yield return new object[] { "RedisUDFSet", "RedisUDFSetNonVolatile" };
            yield return new object[] { "RedisUDFSetEx", "RedisUDFSetExNonVolatile" };
            yield return new object[] { "RedisUDFDel", "RedisUDFDelNonVolatile" };
            yield return new object[] { "RedisUDFExpire", "RedisUDFExpireNonVolatile" };
            yield return new object[] { "RedisUDFIncr", "RedisUDFIncrNonVolatile" };
            yield return new object[] { "RedisUDFIncrBy", "RedisUDFIncrByNonVolatile" };
            yield return new object[] { "RedisUDFRename", "RedisUDFRenameNonVolatile" };
            yield return new object[] { "RedisUDFSetJSON", "RedisUDFSetJSONNonVolatile" };
            yield return new object[] { "RedisUDFSetKV", "RedisUDFSetKVNonVolatile" };
            yield return new object[] { "RedisUDFSetKVPair", "RedisUDFSetKVPairNonVolatile" };
            yield return new object[] { "RedisUDFHashSet", "RedisUDFHashSetNonVolatile" };
            yield return new object[] { "RedisUDFHashSetMultiple", "RedisUDFHashSetMultipleNonVolatile" };
            yield return new object[] { "RedisUDFHashDel", "RedisUDFHashDelNonVolatile" };
            yield return new object[] { "RedisUDFListPushRight", "RedisUDFListPushRightNonVolatile" };
            yield return new object[] { "RedisUDFListPushLeft", "RedisUDFListPushLeftNonVolatile" };
            yield return new object[] { "RedisUDFListPopRight", "RedisUDFListPopRightNonVolatile" };
            yield return new object[] { "RedisUDFListPopLeft", "RedisUDFListPopLeftNonVolatile" };
            yield return new object[] { "RedisUDFSetAdd", "RedisUDFSetAddNonVolatile" };
            yield return new object[] { "RedisUDFSetRemove", "RedisUDFSetRemoveNonVolatile" };
            yield return new object[] { "RedisUDFChannelPublish", "RedisUDFChannelPublishNonVolatile" };
            yield return new object[] { "RedisUDFChannelPublishJSON", "RedisUDFChannelPublishJSONNonVolatile" };
            yield return new object[] { "RedisUDFChannelPublishIfChanged", "RedisUDFChannelPublishIfChangedNonVolatile" };
            yield return new object[] { "RedisUDFChannelPublishIfChangedJSON", "RedisUDFChannelPublishIfChangedJSONNonVolatile" };
            yield return new object[] { "RedisUDFChannelUnsubscribe", "RedisUDFChannelUnsubscribeNonVolatile" };
        }

        [Fact]
        public void TwinPairs_CoverAll24WriteFunctions()
        {
            List<object[]> pairs = TwinPairs().ToList();
            Assert.Equal(24, pairs.Count);
            Assert.Equal(24, pairs.Select(p => (string)p[0]).Distinct().Count());
            Assert.Equal(24, pairs.Select(p => (string)p[1]).Distinct().Count());
        }

        [Theory]
        [MemberData(nameof(TwinPairs))]
        public void Twin_MatchesBaseSignatureAndIsNotVolatile(string baseName, string twinName)
        {
            MethodInfo baseMethod = FindUdf(baseName);
            MethodInfo twinMethod = FindUdf(twinName);
            Assert.True(baseMethod != null, $"base UDF '{baseName}' was not found");
            Assert.True(twinMethod != null, $"twin UDF '{twinName}' was not found");

            ExcelFunctionAttribute baseAttribute = baseMethod.GetCustomAttribute<ExcelFunctionAttribute>();
            ExcelFunctionAttribute twinAttribute = twinMethod.GetCustomAttribute<ExcelFunctionAttribute>();
            Assert.True(baseAttribute != null, $"base UDF '{baseName}' is not an Excel function");
            Assert.True(baseAttribute.IsVolatile, $"base UDF '{baseName}' must be volatile");
            Assert.True(twinAttribute != null, $"twin UDF '{twinName}' is not an Excel function");
            Assert.False(twinAttribute.IsVolatile, $"twin UDF '{twinName}' must not be volatile");

            // The twin help must be the base description plus the fixed suffix.
            Assert.Equal(
                baseAttribute.Description + "; runs once per entry/argument change (non-volatile)",
                twinAttribute.Description);

            // Delegation makes the twin return the same type as its base.
            Assert.Equal(baseMethod.ReturnType, twinMethod.ReturnType);

            ParameterInfo[] baseParameters = baseMethod.GetParameters();
            ParameterInfo[] twinParameters = twinMethod.GetParameters();
            Assert.Equal(baseParameters.Length, twinParameters.Length);
            for (int i = 0; i < baseParameters.Length; i++)
            {
                Assert.Equal(baseParameters[i].ParameterType, twinParameters[i].ParameterType);
                Assert.Equal(baseParameters[i].HasDefaultValue, twinParameters[i].HasDefaultValue);
                Assert.Equal(baseParameters[i].DefaultValue, twinParameters[i].DefaultValue);
                // Names and Excel argument metadata must match so the twin
                // shows the same help as the base.
                Assert.Equal(baseParameters[i].Name, twinParameters[i].Name);
                ExcelArgumentAttribute baseArgument = baseParameters[i].GetCustomAttribute<ExcelArgumentAttribute>();
                ExcelArgumentAttribute twinArgument = twinParameters[i].GetCustomAttribute<ExcelArgumentAttribute>();
                Assert.Equal(baseArgument != null, twinArgument != null);
                if (baseArgument != null)
                {
                    Assert.Equal(baseArgument.Name, twinArgument.Name);
                    Assert.Equal(baseArgument.Description, twinArgument.Description);
                    Assert.Equal(baseArgument.AllowReference, twinArgument.AllowReference);
                }
            }
        }

        /// <summary>
        /// Every volatile [ExcelFunction] that owns a ...NonVolatile twin must
        /// be one of the 24 write functions in the explicit list, and vice
        /// versa. The suite has 42 volatile functions in total: 24 writes with
        /// twins plus 18 read/status helpers that are volatile by design and
        /// intentionally have no twin. The total is pinned so a new volatile
        /// function cannot enter unclassified: it either needs a twin and a
        /// TwinPairs entry, or an update to these counts.
        /// </summary>
        [Fact]
        public void VolatileFunctions_WithTwins_AreExactlyThe24Writes()
        {
            List<MethodInfo> volatileFunctions = typeof(RedisUDF)
                .GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.GetCustomAttribute<ExcelFunctionAttribute>() != null)
                .Where(m => m.GetCustomAttribute<ExcelFunctionAttribute>().IsVolatile)
                .ToList();

            List<string> volatileWithTwin = volatileFunctions
                .Where(m => FindUdf(m.Name + "NonVolatile") != null)
                .Select(m => m.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();
            List<string> explicitBases = TwinPairs()
                .Select(pair => (string)pair[0])
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();

            Assert.Equal(42, volatileFunctions.Count);
            Assert.Equal(24, volatileWithTwin.Count);
            Assert.Equal(18, volatileFunctions.Count - volatileWithTwin.Count);
            Assert.Equal(explicitBases, volatileWithTwin);

            foreach (string name in volatileWithTwin)
            {
                MethodInfo baseMethod = FindUdf(name);
                MethodInfo twinMethod = FindUdf(name + "NonVolatile");
                Assert.True(
                    baseMethod.GetCustomAttribute<ExcelFunctionAttribute>().IsVolatile,
                    $"base UDF '{name}' must be volatile");
                Assert.False(
                    twinMethod.GetCustomAttribute<ExcelFunctionAttribute>().IsVolatile,
                    $"twin UDF '{name}NonVolatile' must not be volatile");
                Assert.Equal(baseMethod.ReturnType, twinMethod.ReturnType);
            }
        }

        /// <summary>
        /// Behavioral delegation check: calling the twin must return exactly
        /// what its base returns. A non-text host argument is rejected by
        /// ResolveHost before the connection manager or the async dispatch is
        /// touched, so the comparison is fully offline and deterministic (it
        /// cannot race the AsyncWrites/SyncWrite seams or the RedisRuntime
        /// singleton used by other test collections).
        /// </summary>
        [Fact]
        public void Twins_DelegateToBase_WithIdenticalOfflineErrorText()
        {
            object invalidHost = 42; // host must be a text value

            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFSet("key", "value", invalidHost),
                () => RedisUDF.RedisUDFSetNonVolatile("key", "value", invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFSetEx("key", "value", 60, invalidHost),
                () => RedisUDF.RedisUDFSetExNonVolatile("key", "value", 60, invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFIncr("key", invalidHost),
                () => RedisUDF.RedisUDFIncrNonVolatile("key", invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFListPushRight("key", "value", invalidHost),
                () => RedisUDF.RedisUDFListPushRightNonVolatile("key", "value", invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFHashSet("hash", "field", "value", invalidHost),
                () => RedisUDF.RedisUDFHashSetNonVolatile("hash", "field", "value", invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFChannelPublish("channel", "message", invalidHost),
                () => RedisUDF.RedisUDFChannelPublishNonVolatile("channel", "message", invalidHost));
            AssertTwinMatchesBase(
                () => RedisUDF.RedisUDFDel("key", invalidHost),
                () => RedisUDF.RedisUDFDelNonVolatile("key", invalidHost));

            // Pin the exact diagnostic once: the invalid host must be
            // reported, not swallowed or replaced by a dispatch error.
            Assert.Equal("Error: host must be a text value", RedisUDF.RedisUDFSet("key", "value", invalidHost));
        }

        private static void AssertTwinMatchesBase(Func<object> baseCall, Func<object> twinCall)
        {
            string baseText = Assert.IsType<string>(baseCall());
            string twinText = Assert.IsType<string>(twinCall());

            Assert.StartsWith("Error: ", baseText);
            Assert.Equal(baseText, twinText);
        }

        private static MethodInfo FindUdf(string name) =>
            typeof(RedisUDF).GetMethod(name, BindingFlags.Public | BindingFlags.Static);
    }
}

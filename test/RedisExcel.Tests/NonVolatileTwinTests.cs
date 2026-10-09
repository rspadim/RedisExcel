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
            }
        }

        private static MethodInfo FindUdf(string name) =>
            typeof(RedisUDF).GetMethod(name, BindingFlags.Public | BindingFlags.Static);
    }
}

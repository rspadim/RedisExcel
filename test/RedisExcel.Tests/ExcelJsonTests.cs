using Newtonsoft.Json.Linq;
using System;
using Xunit;

namespace RedisExcel.Tests
{
    public class ExcelJsonTests
    {
        [Fact]
        public void MatrixToJSON_IntegralDoublesBecomeLongs()
        {
            var matrix = new object[,] { { 1.0, "a" }, { 2.5, null } };
            Assert.Equal("[[1,\"a\"],[2.5,null]]", ExcelJson.RedisUDFMatrixToJSON(matrix));
        }

        [Fact]
        public void JSONToMatrix_ArrayOfArrays()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[[1,\"a\"],[2,\"b\"]]", "");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal(1L, Convert.ToInt64(result[0, 0]));
            Assert.Equal("a", result[0, 1]);
            Assert.Equal(2L, Convert.ToInt64(result[1, 0]));
            Assert.Equal("b", result[1, 1]);
        }

        [Fact]
        public void JSONToMatrix_RaggedRowsArePaddedWithFill()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[[1],[2,3]]", "vazio");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal("vazio", result[0, 1]);
            Assert.Equal(3L, Convert.ToInt64(result[1, 1]));
        }

        [Fact]
        public void JSONToMatrix_ObjectWithArraysAndScalars()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":[1,2],\"b\":3}", "");
            Assert.Equal(3, result.GetLength(0)); // header + max(2, 1) rows
            Assert.Equal(2, result.GetLength(1)); // two keys
            Assert.Equal("a", result[0, 0]);
            Assert.Equal("b", result[0, 1]);
            Assert.Equal(1, ((JToken)result[1, 0]).Value<int>());
            Assert.Equal(3, ((JToken)result[1, 1]).Value<int>());
            Assert.Equal(2, ((JToken)result[2, 0]).Value<int>());
            Assert.Equal("", result[2, 1]); // "b" has a single value
        }

        [Fact]
        public void JSONToMatrix_FlatArrayAndNulls()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[1,null,3]", "X");
            Assert.Equal(1, result.GetLength(0));
            Assert.Equal(3, result.GetLength(1));
            Assert.Equal("1", ((JToken)result[0, 0]).ToString());
            Assert.Equal("X", result[0, 1]);
            Assert.Equal("3", ((JToken)result[0, 2]).ToString());
        }

        [Fact]
        public void JSONToMatrix_ScalarAndNull()
        {
            var scalar = ExcelJson.RedisUDFJSONToMatrix("123", "");
            Assert.Equal("123", ((JToken)scalar[0, 0]).ToString());

            var nullToken = ExcelJson.RedisUDFJSONToMatrix("null", "-");
            Assert.Equal("-", nullToken[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_InvalidJsonReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("{oops", "");
            Assert.StartsWith("Error:", result[0, 0].ToString());
        }
    }
}

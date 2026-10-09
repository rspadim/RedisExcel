using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
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
        public void MatrixToJSON_EmptyAndErrorCellsBecomeNull()
        {
            var matrix = new object[,] { { ExcelEmpty.Value, ExcelError.ExcelErrorValue } };
            Assert.Equal("[[null,null]]", ExcelJson.RedisUDFMatrixToJSON(matrix));
        }

        [Fact]
        public void MatrixToJSON_BooleanStaysLowercase()
        {
            // Pins the documented true/false casing against a ToString refactor.
            Assert.Equal("[[true]]", ExcelJson.RedisUDFMatrixToJSON(new object[,] { { true } }));
        }

        [Fact]
        public void MatrixToJSON_HugeDoubleIsNotCastToLong()
        {
            var json = ExcelJson.RedisUDFMatrixToJSON(new object[,] { { 1e19 } });
            var parsed = JArray.Parse(json);
            Assert.Equal(1e19, parsed[0][0].Value<double>());
        }

        [Fact]
        public void MatrixToJSON_ExactlyTwoToThe63StaysDouble()
        {
            // 2^63 (9223372036854775808) does not fit in a long; the cast guard
            // must serialize it as a double instead of wrapping to long.MinValue.
            var json = ExcelJson.RedisUDFMatrixToJSON(new object[,] { { 9.2233720368547758E18 } });
            var parsed = JArray.Parse(json);
            Assert.Equal(JTokenType.Float, parsed[0][0].Type);
            Assert.Equal(9.2233720368547758E18, parsed[0][0].Value<double>());
        }

        [Fact]
        public void MatrixToJSON_LongMinValueStaysLong()
        {
            // Negative boundary: long.MinValue must stay an integer, not be
            // widened to a double.
            var json = ExcelJson.RedisUDFMatrixToJSON(new object[,] { { long.MinValue } });
            var parsed = JArray.Parse(json);
            Assert.Equal(JTokenType.Integer, parsed[0][0].Type);
            Assert.Equal(long.MinValue, parsed[0][0].Value<long>());
        }

        [Fact]
        public void JSONToMatrix_BigIntegerBecomesInvariantText()
        {
            // Integers beyond Int64 are parsed as BigInteger and cannot be
            // marshalled into a cell; they must become invariant decimal text.
            var result = ExcelJson.RedisUDFJSONToMatrix("[123456789012345678901234567890]", "");
            Assert.Equal("123456789012345678901234567890", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_IsoDateStringsStayText()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[\"2026-10-09T12:00:00Z\"]", "");
            Assert.Equal("2026-10-09T12:00:00Z", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_LeadingBomIsIgnored()
        {
            // Pasted JSON often carries a UTF-8 BOM (U+FEFF); it must not fail
            // the parse with "Unexpected character".
            var result = ExcelJson.RedisUDFJSONToMatrix("\uFEFF[1,2]", "");
            Assert.Equal(1, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal(1L, result[0, 0]);
            Assert.Equal(2L, result[0, 1]);
        }

        [Fact]
        public void JSONToMatrix_WhitespaceBeforeBomIsTolerated()
        {
            // Pasted JSON can carry stray whitespace before the BOM (e.g. a
            // leading space); both are insignificant and must be stripped.
            var result = ExcelJson.RedisUDFJSONToMatrix(" \uFEFF[1,2]", "");
            Assert.Equal(1, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal(1L, result[0, 0]);
            Assert.Equal(2L, result[0, 1]);
        }

        [Fact]
        public void JSONToMatrix_WhitespaceOrBomOnlyReturnsFill()
        {
            // Input that trims down to nothing is equivalent to a blank cell
            // and must return the fill value instead of a parser error.
            var whitespace = ExcelJson.RedisUDFJSONToMatrix(" \t\r\n\uFEFF ", "vazio");
            Assert.Equal(1, whitespace.GetLength(0));
            Assert.Equal(1, whitespace.GetLength(1));
            Assert.Equal("vazio", whitespace[0, 0]);

            var empty = ExcelJson.RedisUDFJSONToMatrix("", "vazio");
            Assert.Equal("vazio", empty[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_ModeratelyDeepJsonIsAccepted()
        {
            // The Newtonsoft default MaxDepth is 64; the raised ceiling (256)
            // must accept ~100 nested arrays instead of returning an error cell.
            const int depth = 100;
            string deep = new string('[', depth) + "1" + new string(']', depth);
            // Object values that are not arrays are kept as compact JSON text,
            // so this pins the exact round-trip of a deeply nested container.
            string nested = "{\"b\":" + deep + "}";
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":" + nested + "}", "");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(1, result.GetLength(1));
            Assert.Equal("a", result[0, 0]);
            Assert.Equal(nested, result[1, 0]);
        }

        [Fact]
        public void JSONToMatrix_ObjectWithTooManyColumnsReturnsError()
        {
            // 20000 keys exceed the 16384 Excel column limit; the result must be
            // rejected before the matrix is allocated, not built and discarded.
            var builder = new System.Text.StringBuilder("{");
            for (int i = 0; i < 20000; i++)
            {
                if (i > 0)
                    builder.Append(',');
                builder.Append("\"k").Append(i).Append("\":").Append(i);
            }
            builder.Append('}');

            var result = ExcelJson.RedisUDFJSONToMatrix(builder.ToString(), "");
            Assert.Equal("Error: JSON is too large for an Excel sheet", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_EmptyInnerArraysReturnFill()
        {
            var single = ExcelJson.RedisUDFJSONToMatrix("[[]]", "vazio");
            Assert.Equal(1, single.GetLength(0));
            Assert.Equal(1, single.GetLength(1));
            Assert.Equal("vazio", single[0, 0]);

            var repeated = ExcelJson.RedisUDFJSONToMatrix("[[],[]]", "vazio");
            Assert.Equal(1, repeated.GetLength(0));
            Assert.Equal(1, repeated.GetLength(1));
            Assert.Equal("vazio", repeated[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_ArrayOfArrays()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[[1,\"a\"],[2,\"b\"]]", "");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal(1L, result[0, 0]);
            Assert.Equal("a", result[0, 1]);
            Assert.Equal(2L, result[1, 0]);
            Assert.Equal("b", result[1, 1]);
        }

        [Fact]
        public void JSONToMatrix_RaggedRowsArePaddedWithFill()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[[1],[2,3]]", "vazio");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal("vazio", result[0, 1]);
            Assert.Equal(3L, result[1, 1]);
        }

        [Fact]
        public void JSONToMatrix_ObjectWithArraysAndScalars()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":[1,2],\"b\":3}", "");
            Assert.Equal(3, result.GetLength(0)); // header + max(2, 1) rows
            Assert.Equal(2, result.GetLength(1)); // two keys
            Assert.Equal("a", result[0, 0]);
            Assert.Equal("b", result[0, 1]);
            Assert.Equal(1L, result[1, 0]);
            Assert.Equal(3L, result[1, 1]);
            Assert.Equal(2L, result[2, 0]);
            Assert.Equal("", result[2, 1]); // "b" has a single value
        }

        [Fact]
        public void JSONToMatrix_ObjectWithScalarsOnly()
        {
            // Regression: the old row-fill loop used keys.Count and could fail.
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":1,\"b\":2}", "");
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal("a", result[0, 0]);
            Assert.Equal("b", result[0, 1]);
            Assert.Equal(1L, result[1, 0]);
            Assert.Equal(2L, result[1, 1]);
        }

        [Fact]
        public void JSONToMatrix_FlatArrayAndNulls()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[1,null,3]", "X");
            Assert.Equal(1, result.GetLength(0));
            Assert.Equal(3, result.GetLength(1));
            Assert.Equal(1L, result[0, 0]);
            Assert.Equal("X", result[0, 1]);
            Assert.Equal(3L, result[0, 2]);
        }

        [Fact]
        public void JSONToMatrix_NestedContainersBecomeJsonText()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":{\"x\":1}}", "");
            Assert.Equal("{\"x\":1}", result[1, 0]);
        }

        [Fact]
        public void JSONToMatrix_EmptyArrayAndEmptyObjectReturnSingleFillCell()
        {
            var emptyArray = ExcelJson.RedisUDFJSONToMatrix("[]", "vazio");
            Assert.Equal(1, emptyArray.GetLength(0));
            Assert.Equal(1, emptyArray.GetLength(1));
            Assert.Equal("vazio", emptyArray[0, 0]);

            var emptyObject = ExcelJson.RedisUDFJSONToMatrix("{}", "vazio");
            Assert.Equal("vazio", emptyObject[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_ScalarAndNull()
        {
            var scalar = ExcelJson.RedisUDFJSONToMatrix("123", "");
            Assert.Equal(123L, scalar[0, 0]);

            var nullToken = ExcelJson.RedisUDFJSONToMatrix("null", "-");
            Assert.Equal("-", nullToken[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_NullJsonReturnsFill()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix((object)null, "vazio");
            Assert.Equal(1, result.GetLength(0));
            Assert.Equal(1, result.GetLength(1));
            Assert.Equal("vazio", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_ErrorCellJsonReturnsError()
        {
            // An Excel error cell as the JSON argument is rejected by
            // ToRedisString; the caught exception becomes an "Error:" cell.
            var result = ExcelJson.RedisUDFJSONToMatrix((object)ExcelError.ExcelErrorNA, "");
            Assert.Equal("Error: Excel error cells are not valid arguments", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_MultiCellJsonReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix(new object[,] { { "[]" } }, "");
            Assert.Equal("Error: A multi-cell range is not a valid scalar argument", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_ErrorCellFillReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[1]", (object)ExcelError.ExcelErrorNA);
            Assert.Equal("Error: Excel error cells are not valid arguments", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_MultiCellFillReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[1]", new object[,] { { "X" } });
            Assert.Equal("Error: A multi-cell range is not a valid scalar argument", result[0, 0]);
        }

        [Fact]
        public void JSONToMatrix_InvalidJsonReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("{oops", "");
            Assert.StartsWith("Error:", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_NaNReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[NaN]", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_InfinityReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[Infinity]", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_NegativeInfinityReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[-Infinity]", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_NestedNaNReturnsError()
        {
            // Nested values also go through JTokenToValue, which must reject the
            // non-finite literal instead of handing a raw double to Excel.
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":NaN}", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_NestedInfinityInContainerReturnsError()
        {
            // A container that would be stringified as JSON text must still be
            // rejected when a non-finite number is hidden inside it.
            var result = ExcelJson.RedisUDFJSONToMatrix("{\"a\":{\"b\":-Infinity}}", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_DeeplyNestedNaNReturnsError()
        {
            var result = ExcelJson.RedisUDFJSONToMatrix("[[[NaN]]]", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_BareNaNScalarReturnsError()
        {
            // Bare NaN (not inside an array/object) takes the scalar path and
            // must hit the same non-finite guard.
            var result = ExcelJson.RedisUDFJSONToMatrix("NaN", "");
            Assert.Equal("Error: JSON contains a non-finite number", result[0, 0].ToString());
        }

        [Fact]
        public void JSONToMatrix_QuotedNonFiniteLiteralsStayText()
        {
            var nanText = ExcelJson.RedisUDFJSONToMatrix("[\"NaN\"]", "");
            Assert.Equal("NaN", nanText[0, 0]);

            var infinityText = ExcelJson.RedisUDFJSONToMatrix("{\"a\":\"Infinity\"}", "");
            Assert.Equal("Infinity", infinityText[1, 0]);
        }

        [Fact]
        public void JSONToMatrix_OverflowingNumberLiteralReturnsError()
        {
            // 1e999 is rejected by Newtonsoft itself (JsonReaderException), so
            // the exact message is not stable; only the "Error:" prefix is.
            var result = ExcelJson.RedisUDFJSONToMatrix("[1e999]", "");
            Assert.StartsWith("Error:", result[0, 0].ToString());
        }
    }
}

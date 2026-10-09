using ExcelDna.Integration;
using NLog;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RedisExcel
{
    /// <summary>Conversion between Excel matrices and JSON (worksheet functions).</summary>
    public static class ExcelJson
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        [ExcelFunction(Description = "Converts a 2D Excel matrix to a compact JSON array of arrays.")]
        public static string RedisUDFMatrixToJSON(
            [ExcelArgument(Description = "2D Excel range to convert")] object[,] range)
        {
            try
            {
                int rows = range.GetLength(0);
                int cols = range.GetLength(1);
                var array = new object[rows][];

                for (int i = 0; i < rows; i++)
                {
                    var row = new object[cols];
                    for (int j = 0; j < cols; j++)
                    {
                        var value = range[i, j];
                        if (value is ExcelEmpty || value is ExcelMissing || value is ExcelError || value == null)
                            row[j] = null;
                        else if (value is double d)
                        {
                            // Excel numbers are doubles; serialize whole numbers as longs,
                            // but only when the cast cannot overflow. long.MaxValue is not
                            // representable as a double (it rounds up to 2^63), so compare
                            // strictly: any double >= 2^63 would wrap to long.MinValue.
                            if (d % 1 == 0 && d >= long.MinValue && d < long.MaxValue)
                                row[j] = (long)d;
                            else
                                row[j] = d;
                        }
                        else
                            row[j] = value;
                    }
                    array[i] = row;
                }

                var settings = new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Include,
                    Formatting = Formatting.None
                };
                return JsonConvert.SerializeObject(array, settings);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUDFMatrixToJSON");
                return $"Error: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Converts a JSON string to a 2D Excel matrix. Accepts arrays, objects, or single values.")]
        public static object[,] RedisUDFJSONToMatrix(
            [ExcelArgument(Description = "JSON string to convert to Excel matrix")] object json,
            [ExcelArgument(Description = "Value to insert for nulls (default is empty string)")] object nullValue = null)
        {
            try
            {
                object fill;
                if (nullValue == null || nullValue is ExcelMissing || nullValue is ExcelEmpty)
                    fill = "";
                else if (nullValue is ExcelError)
                    throw new ArgumentException("Excel error cells are not valid arguments");
                else if (nullValue is Array)
                    throw new ArgumentException("A multi-cell range is not a valid scalar argument");
                else
                    fill = nullValue;
                string jsonText = RedisUDF.ToRedisString(json);
                // Empty/missing cells convert to null and return the fill value.
                // ExcelError cells and multi-cell ranges throw inside ToRedisString;
                // that exception is caught below and reported as a standard
                // "Error: ..." cell instead of reaching Excel raw.
                if (jsonText == null)
                    return new object[,] { { fill } };
                var token = JsonConvert.DeserializeObject<JToken>(jsonText, new JsonSerializerSettings
                {
                    // Keep date-like strings as text so cells receive the original
                    // JSON literal instead of a DateTime.
                    DateParseHandling = DateParseHandling.None
                });
                if (token is JArray array)
                    return JArrayToMatrix(array, fill);
                if (token is JObject obj)
                    return JObjectToMatrix(obj, fill);
                return new object[,] { { JTokenToValue(token, fill) } };
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUDFJSONToMatrix");
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }

        /// <summary>
        /// Converts a JToken to a CLR value ExcelDna can marshal into a cell
        /// (long/double/bool/string). Containers are returned as compact JSON text.
        /// </summary>
        private static object JTokenToValue(JToken token, object fill)
        {
            if (token == null || token.Type == JTokenType.Null)
                return fill;
            // Containers are returned as JSON text; reject a non-finite number
            // hidden anywhere inside before it gets stringified (e.g.
            // {"a":{"b":NaN}} would otherwise become the text {"b":"NaN"}).
            if (token is JContainer container &&
                container.Descendants().OfType<JValue>().Any(
                    v => v.Value is double nested && (double.IsNaN(nested) || double.IsInfinity(nested))))
                throw new ArgumentException("JSON contains a non-finite number");
            if (token is JValue value)
            {
                // JSON integers beyond Int64 arrive as BigInteger and cannot be
                // marshalled into a cell; send the invariant decimal text.
                if (value.Value is System.Numerics.BigInteger big)
                    return big.ToString(System.Globalization.CultureInfo.InvariantCulture);
                // JSON does not allow NaN/Infinity, but Newtonsoft accepts the
                // non-standard literals; Excel would render them as #NUM!, so
                // reject them as an error instead of returning the raw double.
                // Every raw-double path (top-level scalar, array element, object
                // value) funnels through here.
                if (value.Value is double d && (double.IsNaN(d) || double.IsInfinity(d)))
                    throw new ArgumentException("JSON contains a non-finite number");
                return value.Value ?? fill;
            }
            return token.ToString(Formatting.None);
        }

        private static object[,] JArrayToMatrix(JArray array, object fill)
        {
            // Matrix: [[1,2],[3,4]]
            if (array.Count > 0 && array.All(x => x is JArray))
            {
                var rows = array.Cast<JArray>().ToList();
                int cols = rows.Max(r => r.Count);
                // Empty inner arrays ([[]]) would build a 0-column matrix, which
                // Excel cannot size; return a single fill cell instead.
                if (cols == 0)
                    return new object[,] { { fill } };
                var result = new object[rows.Count, cols];
                for (int i = 0; i < rows.Count; i++)
                {
                    for (int j = 0; j < cols; j++)
                        result[i, j] = j < rows[i].Count ? JTokenToValue(rows[i][j], fill) : fill;
                }
                if (logger.IsDebugEnabled)
                    logger.Debug($"RedisUDFJSONToMatrix: array of arrays [{rows.Count}, {cols}]");
                return result;
            }

            // Flat vector: [1,2,3,4]. Empty [] returns a single fill cell.
            if (array.Count == 0)
                return new object[,] { { fill } };
            var flat = new object[1, array.Count];
            for (int i = 0; i < array.Count; i++)
                flat[0, i] = JTokenToValue(array[i], fill);
            if (logger.IsDebugEnabled)
                logger.Debug($"RedisUDFJSONToMatrix: flat array [{array.Count}]");
            return flat;
        }

        private static object[,] JObjectToMatrix(JObject obj, object fill)
        {
            var keys = obj.Properties().Select(p => p.Name).ToList();
            if (keys.Count == 0)
                return new object[,] { { fill } };

            int rows = obj.Properties().Max(p => p.Value is JArray ja ? ja.Count : 1);
            var result = new object[rows + 1, keys.Count];
            for (int j = 0; j < keys.Count; j++)
            {
                string key = keys[j];
                result[0, j] = key;
                for (int i = 0; i < rows; i++)
                    result[i + 1, j] = fill;

                var value = obj[key];
                if (value is JArray arr)
                {
                    for (int i = 0; i < arr.Count; i++)
                        result[i + 1, j] = JTokenToValue(arr[i], fill);
                }
                else
                {
                    result[1, j] = JTokenToValue(value, fill);
                }
            }
            if (logger.IsDebugEnabled)
                logger.Debug($"RedisUDFJSONToMatrix: object [{rows + 1}, {keys.Count}]");
            return result;
        }
    }
}

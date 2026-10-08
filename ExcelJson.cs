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
                            row[j] = d % 1 == 0 ? (object)(long)d : d;
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
                return $"Error: {ex}";
            }
        }

        [ExcelFunction(Description = "Converts a JSON string to a 2D Excel matrix. Accepts arrays, objects, or single values.")]
        public static object[,] RedisUDFJSONToMatrix(
            [ExcelArgument(Description = "JSON string to convert to Excel matrix")] string json,
            [ExcelArgument(Description = "Value to insert for nulls (default is empty string)")] object nullValue = null)
        {
            object fill = nullValue ?? "";
            try
            {
                var token = JsonConvert.DeserializeObject<JToken>(json);
                if (token is JArray array)
                    return JArrayToMatrix(array, fill);
                if (token is JObject obj)
                    return JObjectToMatrix(obj, fill);
                if (token == null || token.Type == JTokenType.Null)
                    return new object[,] { { fill } };
                return new object[,] { { token } };
            }
            catch (Exception ex)
            {
                logger.Error(ex, "RedisUDFJSONToMatrix");
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }

        private static object[,] JArrayToMatrix(JArray array, object fill)
        {
            // Matrix: [[1,2],[3,4]]
            if (array.Count > 0 && array.All(x => x is JArray))
            {
                var rows = array.Select(x => x.ToObject<List<object>>() ?? new List<object>()).ToList();
                int cols = rows.Count == 0 ? 0 : rows.Max(r => r.Count);
                var result = new object[rows.Count, cols];
                for (int i = 0; i < rows.Count; i++)
                {
                    for (int j = 0; j < cols; j++)
                        result[i, j] = j < rows[i].Count ? rows[i][j] ?? fill : fill;
                }
                if (logger.IsDebugEnabled)
                    logger.Debug($"RedisUDFJSONToMatrix: array of arrays [{rows.Count}, {cols}]");
                return result;
            }

            // Flat vector: [1,2,3,4]
            var flat = new object[1, array.Count];
            for (int i = 0; i < array.Count; i++)
                flat[0, i] = array[i].Type == JTokenType.Null ? fill : (object)array[i];
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
                        result[i + 1, j] = arr[i].Type == JTokenType.Null ? fill : (object)arr[i];
                }
                else
                {
                    result[1, j] = value == null || value.Type == JTokenType.Null ? fill : (object)value;
                }
            }
            if (logger.IsDebugEnabled)
                logger.Debug($"RedisUDFJSONToMatrix: object [{rows + 1}, {keys.Count}]");
            return result;
        }
    }
}

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using NLog;
using System;
using System.IO;

namespace RedisExcel
{
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ENUMExcelUpdateStyle
    {
        Timer = 0,
        Realtime = 1,
        Automatic = 2
    }
    public class ConfigSection
    {
        public string host { get; set; }
        public int timeout { get; set; }
    }

    public class RTDConfig : ConfigSection
    {
        public int RedisUpdateRateMs { get; set; } = 1000;
        public int ExcelUpdateRateMs { get; set; } = 100;
        public ENUMExcelUpdateStyle ExcelUpdateStyle { get; set; } = ENUMExcelUpdateStyle.Automatic;
        public long MessageCounterThreshold { get; set; } = 10000;
        public bool UseGetMultiple { get; set; } = true;
    }

    public class ConfigRoot
    {
        public RTDConfig RTD { get; set; }
        public ConfigSection UDF { get; set; }
    }

    public static class ConfigHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly string ConfigFileName = "RedisExcel.json";
        public static string ConfigDefaultHost = "localhost:6379,password=,defaultDatabase=0,ssl=False,abortConnect=False";
        private static string GetExcelDirectory()
        {
            try
            {
                var process = System.Diagnostics.Process.GetCurrentProcess();
                var path = process.MainModule.FileName;
                return Path.GetDirectoryName(path);
            }
            catch
            {
                return null;
            }
        }
        public static ConfigRoot GetConfig()
        {
            string[] possiblePaths = new[] {
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                GetExcelDirectory(),
                "C:\\Windows"
            };

            foreach (var dir in possiblePaths)
            {
                string path = Path.Combine(dir, ConfigFileName);
                try
                {
                    if (!File.Exists(path))
                        continue;
                    var config = JsonConvert.DeserializeObject<ConfigRoot>(
                        File.ReadAllText(path)
                    );
                    if (config != null)
                        return config;
                }
                catch (Exception ex)
                {
                    if (logger.IsErrorEnabled)
                        logger.Error(ex, $"GetConfig: Error reading config file {path}");
                }
            }
            return new ConfigRoot {
                RTD = new RTDConfig {
                    host = ConfigHelper.ConfigDefaultHost,
                    timeout = 1000,
                    RedisUpdateRateMs = 1000,
                    ExcelUpdateRateMs = 1000,
                    MessageCounterThreshold = 10000,
                    ExcelUpdateStyle = ENUMExcelUpdateStyle.Automatic,
                    UseGetMultiple = true
                },
                UDF = new ConfigSection
                {
                    host = ConfigHelper.ConfigDefaultHost,
                    timeout = 1000
                }
            };
        }


    }
}

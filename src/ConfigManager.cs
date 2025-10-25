using System;
using System.IO;
using System.Text.Json;

namespace BimAutomationTool
{
    public static class ConfigManager
    {
        public static void SaveConfig(ExportConfig config, string filePath)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save configuration: {ex.Message}", ex);
            }
        }

        public static ExportConfig LoadConfig(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("Configuration file not found", filePath);

                string json = File.ReadAllText(filePath);
                var config = JsonSerializer.Deserialize<ExportConfig>(json);

                if (config == null)
                    throw new Exception("Failed to deserialize configuration");

                return config;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load configuration: {ex.Message}", ex);
            }
        }

        public static string GetDefaultConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string configFolder = Path.Combine(appData, "BimAutomationTool");
            Directory.CreateDirectory(configFolder);
            return Path.Combine(configFolder, "default_config.json");
        }

        public static void SaveDefaultConfig(ExportConfig config)
        {
            SaveConfig(config, GetDefaultConfigPath());
        }

        public static ExportConfig LoadDefaultConfig()
        {
            string defaultPath = GetDefaultConfigPath();
            if (File.Exists(defaultPath))
            {
                return LoadConfig(defaultPath);
            }
            return new ExportConfig();
        }
    }
}

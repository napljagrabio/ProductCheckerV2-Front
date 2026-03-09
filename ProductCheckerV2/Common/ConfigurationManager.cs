using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Configuration;

namespace ProductCheckerV2.Common
{
    public static class ConfigurationManager
    {
        private const string StageEnvironment = "Stage";
        private const string LiveEnvironment = "Live";
        private const string PreferencesFileName = "user-preferences.json";
        private const string PreferencesEnvironmentKey = "Environment";

        private static string _overriddenEnvironment;
        private static readonly Lazy<string> StartupEnvironment = new(ResolveStartupEnvironment);

        private static readonly Lazy<IConfigurationRoot> Configuration = new(() =>
            new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
                .Build());

        private static string ConfiguredEnvironment =>
            NormalizeEnvironment(Configuration.Value["AppSettings:Environment"] ?? StageEnvironment);

        private static string EffectiveEnvironment =>
            _overriddenEnvironment ?? StartupEnvironment.Value;

        public static string ApplicationName =>
            Configuration.Value["AppSettings:ApplicationName"] ?? "Product Checker";

        public static int MaxFileSizeMB =>
            int.TryParse(Configuration.Value["AppSettings:MaxFileSizeMB"], out var value) ? value : 10;

        public static int BatchSize =>
            int.TryParse(Configuration.Value["AppSettings:BatchSize"], out var value) ? value : 100;
        public static int AutoRefreshTime =>
            int.TryParse(Configuration.Value["AppSettings:AutoRefreshDataTime"], out var value) ? value : 5;

        public static string GetEnvironmentSwitchPassword()
        {
            return Configuration.Value["AppSettings:EnvironmentSwitchPassword"] ?? "admin";
        }

        public static string GetEnvironment()
        {
            return EffectiveEnvironment;
        }

        public static string GetConnectionString(string connectionStringName)
        {
            return GetConnectionStringForEnvironment(EffectiveEnvironment, connectionStringName);
        }

        public static string GetArtemisApiBaseUrl()
        {
            return Configuration.Value[$"{EffectiveEnvironment}:ARTEMIS:BaseUrl"]
                ?? "http://localhost:8000/";
        }

        public static string GetArtemisLoginUsername()
        {
            return Configuration.Value[$"{EffectiveEnvironment}:ARTEMIS:Username"]
                ?? Configuration.Value["ARTEMIS:Username"]
                ?? string.Empty;
        }

        public static string GetArtemisLoginPassword()
        {
            return Configuration.Value[$"{EffectiveEnvironment}:ARTEMIS:Password"]
                ?? Configuration.Value["ARTEMIS:Password"]
                ?? string.Empty;
        }

        public static void SetEnvironment(string environment)
        {
            if (string.IsNullOrWhiteSpace(environment))
            {
                throw new ArgumentException("Environment is required", nameof(environment));
            }

            var normalizedEnvironment = NormalizeEnvironment(environment);
            if (!IsSupportedEnvironment(normalizedEnvironment))
            {
                throw new ArgumentException("Environment must be 'Stage' or 'Live'", nameof(environment));
            }

            _overriddenEnvironment = normalizedEnvironment;
            PersistEnvironmentSelection(normalizedEnvironment);
        }

        private static void PersistEnvironmentSelection(string environment)
        {
            try
            {
                PersistEnvironmentInUserPreferences(environment);

                var runtimeConfigPath = Path.Combine(AppContext.BaseDirectory, "appSettings.json");
                UpdateEnvironmentInJsonFile(runtimeConfigPath, environment);

                var projectConfigPath = FindProjectAppSettingsPath(runtimeConfigPath);
                if (!string.IsNullOrWhiteSpace(projectConfigPath) &&
                    !string.Equals(projectConfigPath, runtimeConfigPath, StringComparison.OrdinalIgnoreCase))
                {
                    UpdateEnvironmentInJsonFile(projectConfigPath, environment);
                }
            }
            catch
            {
                // Environment switching should still work even if persistence fails.
            }
        }

        private static string ResolveStartupEnvironment()
        {
            var cachedEnvironment = ReadEnvironmentFromUserPreferences();
            if (!string.IsNullOrWhiteSpace(cachedEnvironment))
            {
                return cachedEnvironment;
            }

            return ConfiguredEnvironment;
        }

        private static string GetConnectionStringForEnvironment(string environment, string connectionStringName)
        {
            var normalizedEnvironment = NormalizeEnvironment(environment);
            return Configuration.Value[$"{normalizedEnvironment}:ConnectionStrings:{connectionStringName}"]
                ?? Configuration.Value[$"{normalizedEnvironment}:ConnectionStrings:DefaultConnection"]
                ?? string.Empty;
        }

        private static string NormalizeEnvironment(string environment)
        {
            if (environment != null && environment.Equals(LiveEnvironment, StringComparison.OrdinalIgnoreCase))
            {
                return LiveEnvironment;
            }

            return StageEnvironment;
        }

        private static bool IsSupportedEnvironment(string environment)
        {
            return environment == StageEnvironment || environment == LiveEnvironment;
        }

        private static string FindProjectAppSettingsPath(string startFromPath)
        {
            var directory = Path.GetDirectoryName(startFromPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return string.Empty;
            }

            var current = new DirectoryInfo(directory);
            for (var i = 0; i < 6 && current != null; i++)
            {
                var candidate = Path.Combine(current.FullName, "appSettings.json");
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                var nestedCandidate = current.GetDirectories("ProductCheckerV2").FirstOrDefault();
                if (nestedCandidate != null)
                {
                    var nestedPath = Path.Combine(nestedCandidate.FullName, "appSettings.json");
                    if (File.Exists(nestedPath))
                    {
                        return nestedPath;
                    }
                }

                current = current.Parent;
            }

            return string.Empty;
        }

        private static void UpdateEnvironmentInJsonFile(string path, string environment)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return;
            }

            var jsonText = File.ReadAllText(path);
            var rootNode = JsonNode.Parse(jsonText) as JsonObject;
            if (rootNode == null)
            {
                return;
            }

            if (rootNode["AppSettings"] is JsonObject appSettings)
            {
                appSettings["Environment"] = environment;
            }

            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            File.WriteAllText(path, rootNode.ToJsonString(options));
        }

        private static string ReadEnvironmentFromUserPreferences()
        {
            try
            {
                var preferencesPath = GetUserPreferencesPath();
                if (!File.Exists(preferencesPath))
                {
                    return string.Empty;
                }

                var jsonText = File.ReadAllText(preferencesPath);
                var rootNode = JsonNode.Parse(jsonText) as JsonObject;
                if (rootNode == null)
                {
                    return string.Empty;
                }

                var environment = rootNode[PreferencesEnvironmentKey]?.GetValue<string>();
                return IsSupportedEnvironment(NormalizeEnvironment(environment))
                    ? NormalizeEnvironment(environment)
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void PersistEnvironmentInUserPreferences(string environment)
        {
            try
            {
                var preferencesPath = GetUserPreferencesPath();
                var preferencesDirectory = Path.GetDirectoryName(preferencesPath);
                if (!string.IsNullOrWhiteSpace(preferencesDirectory))
                {
                    Directory.CreateDirectory(preferencesDirectory);
                }

                JsonObject rootNode;
                if (File.Exists(preferencesPath))
                {
                    var existingText = File.ReadAllText(preferencesPath);
                    rootNode = JsonNode.Parse(existingText) as JsonObject ?? new JsonObject();
                }
                else
                {
                    rootNode = new JsonObject();
                }

                rootNode[PreferencesEnvironmentKey] = environment;

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                File.WriteAllText(preferencesPath, rootNode.ToJsonString(options));
            }
            catch
            {
                // Environment switching should still work even if preference persistence fails.
            }
        }

        private static string GetUserPreferencesPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ProductCheckerV2",
                "cache",
                PreferencesFileName);
        }
    }
}

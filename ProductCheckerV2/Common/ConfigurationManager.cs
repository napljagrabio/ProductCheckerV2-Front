using System;
using Microsoft.Extensions.Configuration;

namespace ProductCheckerV2.Common
{
    public static class ConfigurationManager
    {
        private static readonly Lazy<IConfigurationRoot> Configuration = new(() =>
            new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
                .Build());

        public static string ApplicationName =>
            Configuration.Value["AppSettings:ApplicationName"] ?? "Product Checker";

        public static int MaxFileSizeMB =>
            int.TryParse(Configuration.Value["AppSettings:MaxFileSizeMB"], out var value) ? value : 10;

        public static int BatchSize =>
            int.TryParse(Configuration.Value["AppSettings:BatchSize"], out var value) ? value : 100;

        public static string GetConnectionString(string connectionString = "DefaultConnection")
        {
            return Configuration.Value.GetConnectionString(connectionString)
                ?? "server=localhost;port=3306;database=product_checker;user=root;password=;charset=utf8mb4;AllowUserVariables=True";
        }

        public static string GetArtemisApiBaseUrl()
        {
            return Configuration.Value["ARTEMIS:BaseUrl"] ?? "http://localhost:8000/";
        }

        public static string GetArtemisLoginUsername()
        {
            return Configuration.Value["ARTEMIS:Username"] ?? string.Empty;
        }

        public static string GetArtemisLoginPassword()
        {
            return Configuration.Value["ARTEMIS:Password"] ?? string.Empty;
        }
    }
}

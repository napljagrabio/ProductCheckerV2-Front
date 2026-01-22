using System;

namespace ProductCheckerV2.Common
{
    public static class ConfigurationManager
    {
        public static string ApplicationName => "Product Checker V2";
        public static int MaxFileSizeMB => 10;
        public static int BatchSize => 100;

        public static string GetConnectionString()
        {
            return "server=localhost;port=3306;database=product_checker;user=root;password=;charset=utf8mb4;AllowUserVariables=True";
        }
    }
}
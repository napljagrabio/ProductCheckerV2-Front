using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ProductCheckerV2.Common
{
    internal static class HelperExtension
    {
        public static string GetColspan { get; set; }
        //public static string AbsPath(this string path)
        //{
        //    if (path == null) return null;
        //    return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path).Replace("\\", "/");
        //}
        public static string AbsPath(this string path)
        {
            if (path == null) return null;

            // If it's already an absolute path (starts with /), return as-is
            if (path.StartsWith("/"))
                return path;

            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path).Replace("\\", "/");
        }
    }
}

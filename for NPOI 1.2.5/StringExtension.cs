using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace NPOI.CSS
{
    internal static class StringExtension
    {
        public static bool IsNullOrWhiteSpace(this string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        public static IEnumerable<Type> DefinedTypes(this Assembly ass)
        {
            return ass.GetTypes();
        }

        public static bool NoncaseEqual(this string s, string compareString)
        {
            return s.Equals(compareString, StringComparison.OrdinalIgnoreCase);
        }

        public static string MD5(this string input)
        {
            if (input == null)
                input = string.Empty;
            byte[] data = Encoding.UTF8.GetBytes(input.Trim().ToLowerInvariant());
            using (var md5 = new MD5CryptoServiceProvider())
            {
                data = md5.ComputeHash(data);
            }

            var ret = new StringBuilder();
            foreach (byte b in data)
            {
                ret.Append(b.ToString("x2").ToLowerInvariant());
            }
            return ret.ToString();
        }
    }
}
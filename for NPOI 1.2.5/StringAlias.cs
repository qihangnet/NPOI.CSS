using System.Collections.Generic;

namespace NPOI.CSS
{
    internal static class StringAlias
    {
        private static readonly Dictionary<string, string> cssKeyDic;

        static StringAlias()
        {
            cssKeyDic = new Dictionary<string, string>();

            cssKeyDic.Add("color", "font-color");
            cssKeyDic.Add("fc", "font-color");

            cssKeyDic.Add("fw", "font-weight");

            cssKeyDic.Add("fn", "font-name");

            cssKeyDic.Add("fs", "font-size");

            cssKeyDic.Add("italic", "font-italic");
            cssKeyDic.Add("i", "font-italic");

            cssKeyDic.Add("underline", "font-underline");
            cssKeyDic.Add("u", "font-underline");

            cssKeyDic.Add("deleteline", "font-strikeout");
            cssKeyDic.Add("d-line", "font-strikeout");
            cssKeyDic.Add("strikeout", "font-strikeout");
            cssKeyDic.Add("d", "font-strikeout");

            cssKeyDic.Add("font-offset", "font-superscript");
            cssKeyDic.Add("superscript", "font-superscript");
            cssKeyDic.Add("fss", "font-superscript");
            cssKeyDic.Add("ss", "font-superscript");

            cssKeyDic.Add("bg-color", "background-color");
            cssKeyDic.Add("bg-c", "background-color");
            cssKeyDic.Add("bgc", "background-color");

            cssKeyDic.Add("align", "text-align");

            cssKeyDic.Add("v-align", "vertical-align");

            cssKeyDic.Add("format", "data-format");
        }

        public static string StandardCssKey(this string csskey)
        {
            if (cssKeyDic.ContainsKey(csskey))
            {
                var s_key = cssKeyDic[csskey];
                return s_key;
            }
            return csskey;
        }

        //public static string GetStandardCssKey(string keyAlias)
        //{
        //}
    }
}
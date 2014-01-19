using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace NPOI.CSS
{
    public static class CellCss
    {
        private const string defaultFontStyle = "font-color:black;font-name:Arial;font-size:10;font-weight:normal;font-underline:none;font-italic:false;font-strikeout:false;font-superscript:none;";

        public static ICell CSS(this ICell cell, string style)
        {
            var styleDic = new SortedDictionary<string, string>();
            InitStyleDic(styleDic);

            var sortedCss = GetCleanStyle(style, styleDic);

            var cssKey = string.Format("CellStyle_{0}", sortedCss.MD5());
            var wb = cell.Sheet.Workbook;
            var cellStyle = wb.GetCellStyle(cssKey);
            if (cellStyle == null)
            {
                cellStyle = GetCellStyle(cell, styleDic);
                wb.AttachedCellStyle(cssKey, cellStyle);
            }
            
            cell.CellStyle = cellStyle;

            return cell;//返回单元格方便流水式编程
        }

        public static ICellStyle DefaultCSS(this ICell cell)
        {
            var defaultDic = new SortedDictionary<string, string>();
            InitStyleDic(defaultDic);
            var sortedCss = GetCleanStyle(defaultFontStyle, defaultDic);

            var cssKey = string.Format("CellStyle_{0}", sortedCss.MD5());
            var wb = cell.Sheet.Workbook;
            var defaultStyle = wb.GetCellStyle(cssKey);
            if (defaultStyle == null)
            {
                var font = wb.GetFont(defaultDic);
                defaultStyle = wb.CreateCellStyle();
                defaultStyle.SetFont(font);
                wb.AttachedCellStyle(cssKey, defaultStyle);
            }

            return defaultStyle;
        }

        private static void InitStyleDic(SortedDictionary<string, string> dic)
        {
            var cssItems = GetCssItems(defaultFontStyle);

            foreach (var cssitem in cssItems)
            {
                var kvp = GetCssKeyValue(cssitem);
                dic.Add(kvp.Key, kvp.Value);
            }
        }

        private static ICellStyle GetCellStyle(ICell cell, SortedDictionary<string, string> dic)
        {
            var wb = cell.Sheet.Workbook;
            ICellStyle cellStyle = wb.CreateCellStyle();

            cellStyle.CloneStyleFrom(cell.CellStyle);
            var fontStyles = dic.Where(w => w.Key.StartsWith("font-")).ToArray();
            var fontDic = new SortedDictionary<string, string>();
            foreach (var kv in fontStyles)
            {
                fontDic.Add(kv.Key, kv.Value);
            }
            var font = wb.GetFont(fontDic);
            var xdic = dic.Except(fontDic);
            cellStyle.SetFont(font);//TODO 在基于style.xls基础的样式上增加css时，会造成原字体设置的丢失
            foreach (var kvp in xdic)
            {
                FireCssAccess(cellStyle, wb, kvp);
            }
            return cellStyle;
        }

        private static string GetCleanStyle(string style, SortedDictionary<string, string> dic)
        {
            style = Regex.Replace(style.Trim(), "\\s+", " ");
            style = Regex.Replace(style, "\\s;\\s", ";");
            style = Regex.Replace(style, "\\s:\\s", ":");

            var cssItems = GetCssItems(style.TrimEnd(';'));

            foreach (var cssitem in cssItems)
            {
                var kvp = GetCssKeyValue(cssitem);
                if (dic.ContainsKey(kvp.Key))
                {
                    dic[kvp.Key] = kvp.Value; //覆盖相同key的值
                }
                else
                {
                    dic.Add(kvp.Key, kvp.Value);
                }
            }

            return GetSortedCss(dic);
        }

        private static string GetSortedCss(SortedDictionary<string, string> dic)
        {
            var sortedCss = string.Join(";", dic.Select(s => string.Format("{0}:{1}", s.Key, s.Value)).ToArray());
            return sortedCss;
        }

        private static string[] GetCssItems(string style)
        {
            var cssItems = Regex.Split(style, ";");
            cssItems = cssItems.Where(w => !w.IsNullOrWhiteSpace()).ToArray();
            return cssItems;
        }

        private static KeyValuePair<string, string> GetCssKeyValue(string css)
        {
            var cssKeyValueArray = Regex.Split(css, ":").ToArray();
            var cssKey = cssKeyValueArray[0].StandardCssKey();
            var cssValue = cssKey == "font-name" ? cssKeyValueArray[1] : cssKeyValueArray[1].ToUpper(); //字体不应变大写
            var kv = new KeyValuePair<string, string>(cssKey, cssValue);
            return kv;
        }

        private static void FireCssAccess(ICellStyle style, IWorkbook wb, KeyValuePair<string, string> kvp)
        {
            switch (kvp.Key)
            {
                case "text-align":
                    style.TextAlign(wb, kvp.Value);
                    break;

                case "vertical-align":
                    style.VerticalAlign(wb, kvp.Value);
                    break;

                case "background-color":
                    style.BackgroundColor(wb, kvp.Value);
                    break;

                case "border-type":
                    style.BorderTypes(wb, kvp.Value);
                    break;

                case "data-format":
                    style.DataFormat(wb, kvp.Value);
                    break;

                default:
                    break;
            }
        }
    }
}
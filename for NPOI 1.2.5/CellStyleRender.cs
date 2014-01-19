using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace NPOI.CSS
{
    internal static partial class CellStyleRender
    {
        public static IFont GetFont(this IWorkbook wb, SortedDictionary<string, string> fontdic)
        {
            var weight = fontdic.FontWeight();
            var color = fontdic.FontColor();
            var size = fontdic.FontSize();
            var name = fontdic.FontName();
            var underline = fontdic.FontUnderline();
            var italic = fontdic.FontItalic();
            var strikeout = fontdic.FontStrikeout();
            var offset = fontdic.ConvertToSuperScript();

            var findHeight = (short)(size * 20);
            var font = wb.FindFont(weight, color, findHeight, name, italic, strikeout, offset, underline);

            if (font == null)
            {
                font = wb.CreateFont();
                font.Boldweight = weight;
                font.Color = color;
                font.FontHeightInPoints = size;
                font.FontName = name;
                font.Underline = underline;
                font.IsItalic = italic;
                font.IsStrikeout = strikeout;
                font.TypeOffset = offset;
            }
            return font;
        }

        #region font-weight

        public static short FontWeight(this SortedDictionary<string, string> fontdic)
        {
            switch (fontdic["font-weight"])
            {
                case "NORMAL":
                    return 400;

                case "BOLD":
                    return 700;

                default:
                    return 0;
            }
        }

        #endregion font-weight

        #region font-name

        public static string FontName(this SortedDictionary<string, string> fontdic)
        {
            return fontdic["font-name"];
        }

        #endregion font-name

        #region font-size

        public static short FontSize(this SortedDictionary<string, string> fontdic)
        {
            return short.Parse(fontdic["font-size"]);
        }

        #endregion font-size

        #region font-color

        public static short FontColor(this SortedDictionary<string, string> fontdic)
        {
            var color = fontdic["font-color"].ConvertToColor();
            return color;
        }

        #endregion font-color

        #region font-italic

        public static bool FontItalic(this SortedDictionary<string, string> fontdic)
        {
            return fontdic["font-italic"] == "TRUE";
        }

        #endregion font-italic

        #region font-strikeout

        public static bool FontStrikeout(this SortedDictionary<string, string> fontdic)
        {
            return fontdic["font-strikeout"] == "TRUE";
        }

        #endregion font-strikeout

        #region text-align

        public static void TextAlign(this ICellStyle style, IWorkbook wb, string v)
        {
            style.Alignment =v.ConvertToHorizontalAlignment();
        }

        #endregion text-align

        #region vertical-align

        public static void VerticalAlign(this ICellStyle style, IWorkbook wb, string v)
        {
            style.VerticalAlignment = v.ConvertToVerticalAlignment();
        }

        #endregion vertical-align

        #region boder-type

        public static void BorderTypes(this ICellStyle style, IWorkbook wb, string v)
        {
            string[] borderTypeNames = new[] { string.Empty, string.Empty, string.Empty, string.Empty };
            v = v.ToUpper();
            var vs = v.Split(' ');
            switch (vs.Length)
            {
                case 1:
                    borderTypeNames[0] = borderTypeNames[1] = borderTypeNames[2] = borderTypeNames[3] = vs[0];
                    break;

                case 2:
                    borderTypeNames[0] = borderTypeNames[2] = vs[0];
                    borderTypeNames[1] = borderTypeNames[3] = vs[1];
                    break;

                case 3:
                    borderTypeNames[0] = vs[0];
                    borderTypeNames[1] = borderTypeNames[3] = vs[1];
                    borderTypeNames[2] = vs[2];
                    break;

                case 4:
                    borderTypeNames[0] = vs[0];
                    borderTypeNames[1] = vs[1];
                    borderTypeNames[2] = vs[2];
                    borderTypeNames[3] = vs[3];
                    break;

                default:
                    break;
            }
            var borderTopTypeName = borderTypeNames[0];
            var borderRightTypeName = borderTypeNames[1];
            var borderBottomTypeName = borderTypeNames[2];
            var borderLeftTypeName = borderTypeNames[3];

            if (!borderTopTypeName.IsNullOrWhiteSpace())
            {
                style.BorderTop = borderTopTypeName.ConvertToBorderStyle();
            }
            if (!borderRightTypeName.IsNullOrWhiteSpace())
            {
                style.BorderRight = borderRightTypeName.ConvertToBorderStyle();
            }
            if (!borderBottomTypeName.IsNullOrWhiteSpace())
            {
                style.BorderBottom = borderBottomTypeName.ConvertToBorderStyle();
            }
            if (!borderLeftTypeName.IsNullOrWhiteSpace())
            {
                style.BorderLeft = borderLeftTypeName.ConvertToBorderStyle();
            }
        }

        #endregion boder-type

        #region data-format

        public static void DataFormat(this ICellStyle style, IWorkbook wb, string v)
        {
            var df = wb.CreateDataFormat();
            style.DataFormat = df.GetFormat(v); ;
        }

        #endregion data-format

    }
}
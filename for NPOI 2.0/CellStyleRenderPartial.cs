using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace NPOI.CSS
{
    internal static partial class CellStyleRender
    {
        public static FontSuperScript ConvertToSuperScript(this SortedDictionary<string, string> fontdic)
        {
            var v = fontdic["font-superscript"];
            switch (v)
            {
                case "SUPER":
                    return FontSuperScript.Super;

                case "SUB":
                    return FontSuperScript.Sub;

                default:
                    return FontSuperScript.None;
            }
        }

        public static FontUnderlineType FontUnderline(this SortedDictionary<string, string> fontdic)
        {
            var v = fontdic["font-underline"];
            switch (v)
            {
                case "SINGLE":
                    return FontUnderlineType.Single;

                case "DOUBLE":
                    return FontUnderlineType.Double;

                case "SINGLEACCOUNTING":
                case "SINGLE_ACCOUNTING":
                    return FontUnderlineType.SingleAccounting;

                case "DOUBLEACCOUNTING":
                case "DOUBLE_ACCOUNTING":
                    return FontUnderlineType.DoubleAccounting;

                default:
                    return FontUnderlineType.None;
            }
        }

        public static void BackgroundColor(this ICellStyle style, IWorkbook wb, string v)
        {
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = v.ConvertToColor();
        }

        public static short ConvertToColor(this string v)
        {
            switch (v)
            {
                case "AQUA":
                    return 49;

                case "AUTOMATIC":
                    return 64;

                case "BLACK":
                    return 8;

                case "BLUE":
                    return 39;

                case "BLUE_GREY":
                case "BLUEGREY":
                    return 54;

                case "BRIGHT_GREEN":
                case "BRIGHTGREEN":
                    return 35;

                case "BROWN":
                    return 60;

                case "CORAL":
                    return 29;

                case "CORNFLOWER_BLUE":
                case "CORNFLOWERBLUE":
                    return 24;

                case "DARK_BLUE":
                case "DARKBLUE":
                    return 32;

                case "DARK_GREEN":
                case "DARKGREEN":
                    return 58;
                case "DARK_RED":
                case "DARKRED":
                    return 37;

                case "DARK_TEAL":
                case "DARKTEAL":
                    return 56;

                case "DARK_YELLOW":
                case "DARKYELLOW":
                    return 19;

                case "GOLD":
                    return 51;

                case "GREEN":
                    return 17;
                case "GREY_25_PERCENT":
                case "GREY25PERCENT":
                    return 22;

                case "GREY_40_PERCENT":
                case "GREY40PERCENT":
                    return 55;

                case "GREY_50_PERCENT":
                case "GREY50PERCENT":
                    return 23;

                case "GREY_80_PERCENT":
                case "GREY80PERCENT":
                    return 63;

                case "INDIGO":
                    return 62;

                case "LAVENDER":
                    return 46;

                case "LEMON_CHIFFON":
                case "LEMONCHIFFON":
                    return 26;

                case "LIGHT_BLUE":
                case "LIGHTBLUE":
                    return 48;

                case "LIGHT_CORNFLOWERBLUE":
                case "LIGHTCORNFLOWERBLUE":
                    return 31;

                case "LIGHT_GREEN":
                case "LIGHTGREEN":
                    return 42;

                case "LIGHT_ORANGE":
                case "LIGHTORANGE":
                    return 52;

                case "LIGHT_TURQUOISE":
                case "LIGHTTURQUOISE":
                    return 27;

                case "LIGHT_YELLOW":
                case "LIGHTYELLOW":
                    return 43;

                case "LIME":
                    return 50;

                case "MAROON":
                    return 25;

                case "OLIVE_GREEN":
                case "OLIVEGREEN":
                    return 59;

                case "ORANGE":
                    return 53;

                case "ORCHID":
                    return 28;

                case "PALE_BLUE":
                case "PALEBLUE":
                    return 44;

                case "PINK":
                    return 33;

                case "PLUM":
                    return 25;

                case "RED":
                    return 10;

                case "ROSE":
                    return 45;

                case "ROYAL_BLUE":
                case "ROYALBLUE":
                    return 30;

                case "SEA_GREEN":
                case "SEAGREEN":
                    return 57;

                case "SKY_BLUE":
                case "SKYBLUE":
                    return 40;

                case "TAN":
                    return 47;

                case "TEAL":
                    return 38;

                case "TURQUOISE":
                    return 35;

                case "VIOLET":
                    return 36;

                case "WHITE":
                    return 9;

                case "YELLOW":
                    return 34;

                default:
                    return 32767;
            }
        }

        public static HorizontalAlignment ConvertToHorizontalAlignment(this string v)
        {
            switch (v)
            {
                case "LEFT":
                    return HorizontalAlignment.Left;

                case "CENTER":
                    return HorizontalAlignment.Center;

                case "CENTERSELECTION":
                case "CENTER_SELECTION":
                    return HorizontalAlignment.CenterSelection;

                case "RIGHT":
                    return HorizontalAlignment.Right;

                case "DISTRIBUTED":
                    return HorizontalAlignment.Distributed;

                case "FILL":
                    return HorizontalAlignment.Fill;

                case "JUSTIFY":
                    return HorizontalAlignment.Justify;

                default:
                    return HorizontalAlignment.General;
            }
        }

        public static VerticalAlignment ConvertToVerticalAlignment(this string v)
        {
            switch (v)
            {
                case "TOP":
                    return VerticalAlignment.Top;

                case "CENTER":
                    return VerticalAlignment.Center;

                case "BOTTOM":
                    return VerticalAlignment.Bottom;

                case "DISTRIBUTED":
                    return VerticalAlignment.Distributed;

                default:
                    return VerticalAlignment.Justify;
            }
        }

        public static BorderStyle ConvertToBorderStyle(this string v)
        {
            switch (v)
            {
                case "THIN":
                    return BorderStyle.Thin;

                case "MEDIUM":
                    return BorderStyle.Medium;

                case "DASHED":
                    return BorderStyle.Dashed;

                case "HAIR":
                    return BorderStyle.Hair;

                case "THICK":
                    return BorderStyle.Thick;

                case "DOUBLE":
                    return BorderStyle.Double;

                case "DOTTED":
                    return BorderStyle.Dotted;

                case "MEDIUMDASHED":
                case "MEDIUM_DASHED":
                    return BorderStyle.MediumDashed;

                case "DASHDOT":
                case "DASH_DOT":
                    return BorderStyle.DashDot;

                case "MEDIUMDASHDOT":
                case "MEDIUM_DASH_DOT":
                    return BorderStyle.MediumDashDot;

                case "DASHDOTDOT":
                case "DASH_DOT_DOT":
                    return BorderStyle.DashDotDot;

                case "MEDIUMDASHDOTDOT":
                case "MEDIUM_DASH_DOT_DOT":
                    return BorderStyle.MediumDashDotDot;

                case "SLANTEDDASHDOT":
                case "SLANTED_DASH_DOT":
                    return BorderStyle.SlantedDashDot;

                default:
                    return BorderStyle.None;
            }
        }
    }
}
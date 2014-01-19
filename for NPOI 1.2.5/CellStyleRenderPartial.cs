using FastReflectionLib;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace NPOI.CSS
{
    internal static partial class CellStyleRender
    {
        public static short ConvertToSuperScript(this SortedDictionary<string, string> fontdic)
        {
            var v = fontdic["font-superscript"];
            switch (v)
            {
                case "SUPER":
                    return 1;

                case "SUB":
                    return 2;

                default:
                    return 0;
            }
        }

        public static byte FontUnderline(this SortedDictionary<string, string> fontdic)
        {
            var v = fontdic["font-underline"];
            switch (v)
            {
                case "SINGLE":
                    return 1;

                case "DOUBLE":
                    return 2;

                case "SINGLEACCOUNTING":
                case "SINGLE_ACCOUNTING":
                    return 33;

                case "DOUBLEACCOUNTING":
                case "DOUBLE_ACCOUNTING":
                    return 34;

                default:
                    return 0;
            }
        }

        public static void BackgroundColor(this ICellStyle style, IWorkbook wb, string v)
        {
            style.FillPattern = FillPatternType.SOLID_FOREGROUND;
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
                    return 12;

                case "BLUE_GREY":
                case "BLUEGREY":
                    return 54;

                case "BRIGHT_GREEN":
                case "BRIGHTGREEN":
                    return 11;

                case "BROWN":
                    return 60;

                case "CORAL":
                    return 29;

                case "CORNFLOWER_BLUE":
                case "CORNFLOWERBLUE":
                    return 24;

                case "DARK_BLUE":
                case "DARKBLUE":
                    return 18;

                case "DARK_GREEN":
                case "DARKGREEN":
                    return 58;

                case "DARK_RED":
                case "DARKRED":
                    return 16;

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
                    return 41;

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
                    return 14;

                case "PLUM":
                    return 61;

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
                    return 21;

                case "TURQUOISE":
                    return 15;

                case "VIOLET":
                    return 20;

                case "WHITE":
                    return 9;

                case "YELLOW":
                    return 13;

                default:
                    return 32767;
            }
        }

        public static HorizontalAlignment ConvertToHorizontalAlignment(this string v)
        {
            switch (v)
            {
                case "LEFT":
                    return HorizontalAlignment.LEFT;

                case "CENTER":
                    return HorizontalAlignment.CENTER;

                case "CENTERSELECTION":
                case "CENTER_SELECTION":
                    return HorizontalAlignment.CENTER_SELECTION;

                case "RIGHT":
                    return HorizontalAlignment.RIGHT;

                case "DISTRIBUTED":
                    return HorizontalAlignment.DISTRIBUTED;

                case "FILL":
                    return HorizontalAlignment.FILL;

                case "JUSTIFY":
                    return HorizontalAlignment.JUSTIFY;

                default:
                    return HorizontalAlignment.GENERAL;
            }
        }

        public static VerticalAlignment ConvertToVerticalAlignment(this string v)
        {
            switch (v)
            {
                case "TOP":
                    return VerticalAlignment.TOP;

                case "CENTER":
                    return VerticalAlignment.CENTER;

                case "BOTTOM":
                    return VerticalAlignment.BOTTOM;

                case "DISTRIBUTED":
                    return VerticalAlignment.DISTRIBUTED;

                default:
                    return VerticalAlignment.JUSTIFY;
            }
        }

        public static BorderStyle ConvertToBorderStyle(this string v)
        {
            switch (v)
            {
                case "THIN":
                    return BorderStyle.THIN;

                case "MEDIUM":
                    return BorderStyle.MEDIUM;

                case "DASHED":
                    return BorderStyle.DASHED;

                case "HAIR":
                    return BorderStyle.HAIR;

                case "THICK":
                    return BorderStyle.THICK;

                case "DOUBLE":
                    return BorderStyle.DOUBLE;

                case "DOTTED":
                    return BorderStyle.DOTTED;

                case "MEDIUMDASHED":
                case "MEDIUM_DASHED":
                    return BorderStyle.MEDIUM_DASHED;

                case "DASHDOT":
                case "DASH_DOT":
                    return BorderStyle.DASH_DOT;

                case "MEDIUMDASHDOT":
                case "MEDIUM_DASH_DOT":
                    return BorderStyle.MEDIUM_DASH_DOT;

                case "DASHDOTDOT":
                case "DASH_DOT_DOT":
                    return BorderStyle.DASH_DOT_DOT;

                case "MEDIUMDASHDOTDOT":
                case "MEDIUM_DASH_DOT_DOT":
                    return BorderStyle.MEDIUM_DASH_DOT_DOT;

                case "SLANTEDDASHDOT":
                case "SLANTED_DASH_DOT":
                    return BorderStyle.SLANTED_DASH_DOT;

                default:
                    return BorderStyle.NONE;
            }
        }
    }
}
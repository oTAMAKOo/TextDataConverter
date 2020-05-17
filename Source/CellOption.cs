
using System;
using System.Drawing;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace GameTextConverter
{
    public static class CellOption
    {
        //----- params -----
        
        //----- field -----

        //----- property -----

        //----- method -----

        public static void Set(ExcelRange cell, string comment, string fontColor, string backgroundColor)
        {
            if (!string.IsNullOrEmpty(comment))
            {
                cell.AddComment(comment, "REF");
            }

            if (!string.IsNullOrEmpty(fontColor))
            {
                cell.Style.Font.Color.SetColor(ColorTranslator.FromHtml(fontColor));
            }

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(backgroundColor));
            }
        }

        public static Tuple<string, string, string> Get(ExcelRange cell)
        {
            var comment = cell.Comment != null ? cell.Comment.Text : null;

            var fontColor = GetColorCode(cell, cell.Style.Font.Color);
            var backgroundColor = GetColorCode(cell, cell.Style.Fill.BackgroundColor);

            var changed = false;
            
            changed |= !string.IsNullOrEmpty(comment);
            changed |= !string.IsNullOrEmpty(fontColor) && fontColor != "#FF000000";
            changed |= !string.IsNullOrEmpty(backgroundColor) && backgroundColor != "#FFFFFFFF";

            return changed ? Tuple.Create(comment, fontColor, backgroundColor) : null;
        }

        private static string GetColorCode(ExcelRange cell, ExcelColor color)
        {
            string colorCode = null;

            if (!string.IsNullOrEmpty(color.Rgb))
            {
                colorCode = "#" + color.Rgb;
            }

            if (!string.IsNullOrEmpty(color.Theme))
            {
                colorCode = null;

                ConsoleUtility.Warning("Theme color not support saved default color.\n[{0}] {1}", cell.Address, cell.Text);
            }

            return colorCode;
        }
    }
}

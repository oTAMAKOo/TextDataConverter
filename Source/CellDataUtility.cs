
using System;
using System.Drawing;
using System.Linq;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace GameTextConverter
{
    public static class CellDataUtility
    {
        public static void Set(ExcelWorksheet worksheet, CellData cellData)
        {
            if (cellData == null) { return; }

            var address = cellData.address.Split(',');

            var row = address.ElementAtOrDefault(0);
            var column = address.ElementAtOrDefault(1);

            if (string.IsNullOrEmpty(column) || string.IsNullOrEmpty(row)) { return; }

            var c = Convert.ToInt32(column);
            var r = Convert.ToInt32(row);

            var cell = worksheet.Cells[r, c];

            if (!string.IsNullOrEmpty(cellData.comment))
            {
                cell.AddComment(cellData.comment, "REF");
            }

            cell.Style.Fill.PatternType = cellData.patternType;

            SetColor(cell.Style.Font.Color, cellData.fontColor);
            SetColor(cell.Style.Fill.BackgroundColor, cellData.backgroundColor);
        }

        public static CellData Get(ExcelWorksheet worksheet, int column, int row)
        {
            var cell = worksheet.Cells[row, column];

            var cellData = new CellData()
            {
                address = string.Format("{0},{1}", row, column)
            };

            if (cell.Comment != null)
            {
                var comment = cell.Comment.Text;

                var author = cell.Comment.Author;

                var removeText = string.Format("{0}:", author);

                if (!string.IsNullOrEmpty(author) && comment.StartsWith(removeText))
                {
                    comment = comment.Substring(removeText.Length);
                }

                cellData.comment = comment.Trim('\n');
            }
            
            cellData.fontColor = GetColor(cell.Style.Font.Color);
            cellData.backgroundColor = GetColor(cell.Style.Fill.BackgroundColor);
            cellData.patternType = cell.Style.Fill.PatternType;

            if (IsEmptyCellData(cellData)){ return null; }

            return cellData;
        }

        private static CellData.Color GetColor(ExcelColor excelColor)
        {
            var color = new CellData.Color()
            {
                rgb = string.IsNullOrEmpty(excelColor.Rgb) ? null : excelColor.Rgb,
                theme = excelColor.Theme,
            };

            if (color.theme.HasValue)
            {
                color.tint = excelColor.Tint;
            }

            if (IsEmptyColor(color)) { return null; }

            return color;
        }

        private static void SetColor(ExcelColor excelColor, CellData.Color color)
        {
            if (IsEmptyColor(color)) { return; }

            if (color.theme.HasValue)
            {
                excelColor.SetColor(color.theme.Value);

                if (color.tint.HasValue)
                {
                    excelColor.Tint = color.tint.Value;
                }
            }
            else
            {
                excelColor.SetColor(ColorTranslator.FromHtml("#" + color.rgb));
            }
        }

        private static bool IsEmptyCellData(CellData cellData)
        {
            if (cellData == null){ return true; }

            var hasValue = false;

            hasValue |= !string.IsNullOrEmpty(cellData.comment);
            hasValue |= cellData.fontColor != null;
            hasValue |= cellData.backgroundColor != null;

            return !hasValue;
        }

        private static bool IsEmptyColor(CellData.Color color)
        {
            if (color == null) { return true; }

            var hasValue = false;

            hasValue |= !string.IsNullOrEmpty(color.rgb);
            hasValue |= color.theme.HasValue;
            hasValue |= color.tint.HasValue;

            return !hasValue;
        }
    }
}

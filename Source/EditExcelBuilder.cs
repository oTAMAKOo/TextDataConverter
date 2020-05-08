
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace GameTextConverter
{
    public sealed class EditExcelBuilder
    {
        //----- params -----

        //----- field -----


        //----- property -----

        //----- method -----

        public static void Build(string workspace, ExcelData excelData, Settings settings)
        {
            var originExcelPath = Path.GetFullPath(settings.ExcelPath);

            var editExcelPath = PathUtility.Combine(workspace, Constants.EditExcelFile);

            Console.WriteLine("------ Build edit excel file ------");

            //------ エディット用にエクセルファイルを複製 ------

            if (!File.Exists(originExcelPath))
            {
                throw new FileNotFoundException(string.Format("{0} is not exists.", originExcelPath));
            }

            var originXlsxFile = new FileInfo(originExcelPath);

            var editXlsxFile = originXlsxFile.CopyTo(editExcelPath, true);

            //------ レコード情報を書き込み ------

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var worksheets = excel.Workbook.Worksheets;

                // テンプレートシート.

                var templateSheet = worksheets.FirstOrDefault(x => x.Name.ToLower() == settings.TemplateSheetName);

                if (templateSheet == null)
                {
                    throw new Exception(string.Format("Template worksheet {0} not found.", settings.TemplateSheetName));
                }

                // シート作成.

                foreach (var sheet in excelData.sheets)
                {
                    if (string.IsNullOrEmpty(sheet.displayName)) { continue; }

                    if (worksheets.Any(x => x.Name == sheet.displayName))
                    {
                        throw new Exception(string.Format("Worksheet create failed. Worksheet {0} already exists", sheet.displayName));
                    }

                    // テンプレートシートを複製.                    
                    var newWorksheet = worksheets.Add(sheet.displayName, templateSheet);

                    // 保護解除.
                    newWorksheet.Protection.IsProtected = false;
                    // タブ選択状態解除.
                    newWorksheet.View.TabSelected = false;
                }

                Console.WriteLine("Create worksheet.");

                // 先頭のシートをアクティブ化.

                var firstWorksheet = worksheets.FirstOrDefault();

                if (firstWorksheet != null)
                {
                    firstWorksheet.View.TabSelected = true;
                }

                // シート順番入れ替え.

                var loop = true;

                while (loop)
                {
                    loop = false;
                    
                    foreach (var sheet in excelData.sheets)
                    {
                        if (worksheets.Count <= sheet.index) { continue; }

                        var worksheet = worksheets.FirstOrDefault(x => x.Name == sheet.displayName);

                        if (worksheet != null && worksheet.Index != sheet.index)
                        {
                            loop = true;
                            break;
                        }
                    }

                    foreach (var worksheet in worksheets)
                    {
                        var sheet = excelData.sheets.FirstOrDefault(x => x.index == worksheet.Index + 1);

                        if (sheet == null) { continue; }

                        if (sheet.displayName == worksheet.Name && sheet.index == worksheet.Index) { continue; }

                        var moveSheet = worksheets.FirstOrDefault(x => x.Name == sheet.displayName);

                        worksheets.MoveAfter(moveSheet.Index, worksheet.Index);
                    }
                }

                Console.WriteLine("Sort worksheet.");

                Console.WriteLine("Import worksheet.");

                // レコード情報設定.

                var graphics = Graphics.FromImage(new Bitmap(1, 1));

                foreach (var sheet in excelData.sheets)
                {
                    var worksheet = worksheets.FirstOrDefault(x => x.Name == sheet.displayName);

                    var dimension = worksheet.Dimension;

                    var records = excelData.records.GetValueOrDefault(sheet.sheetName);

                    worksheet.SetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X, sheet.sheetName);

                    SetGuid(worksheet, Constants.SheetGuidAddress.Y, Constants.SheetGuidAddress.X, sheet.guid);

                    // レコード投入用セルを用意.

                    for (var i = 0; i < records.Length; i++)
                    {
                        var recordRow = Constants.RecordStartRow + i;

                        // 行追加.
                        if (worksheet.Cells.End.Row < recordRow)
                        {
                            worksheet.InsertRow(recordRow, 1);
                        }

                        // セル情報コピー.
                        for (var column = 1; column < dimension.End.Column; column++)
                        {
                            CloneCellFormat(worksheet, Constants.RecordStartRow, recordRow, column);
                        }
                    }

                    // 値設定.

                    foreach (var record in records)
                    {
                        var row = record.line;

                        // Guid.
                        SetGuid(worksheet, row, Constants.GuidColumn, record.guid);

                        // Enum名.
                        worksheet.SetValue(row, Constants.EnumNameColumn, record.enumName);

                        // 説明.
                        worksheet.SetValue(row, Constants.DescriptionColumn, record.description);

                        // テキスト.
                        for (var i = 0; i < record.texts.Length; i++)
                        {
                            worksheet.SetValue(row, Constants.TextStartColumn + i, record.texts[i]);
                        }                        
                    }

                    // セルサイズを調整.

                    var maxRow = records.Max(x => x.line) + 1;
                    
                    worksheet.Cells.AutoFitColumns(20f, 100f);

                    for (var c = 1; c < dimension.End.Column; c++)
                    {
                        var columnWidth = worksheet.Column(c).Width;

                        for (var r = 1; r <= maxRow; r++)
                        {
                            var cell = worksheet.Cells[r, c];

                            if (string.IsNullOrEmpty(cell.Text)) { continue; }

                            cell.Style.WrapText = true;
                            cell.Style.ShrinkToFit = false;

                            var width = CalcTextWidth(graphics, cell);

                            if (columnWidth < width)
                            {
                                columnWidth = width;
                            }
                        }

                        worksheet.Column(c).Width = columnWidth;
                    }

                    // GUID行は幅固定.
                    worksheet.Column(Constants.GuidColumn).Width = 20d;

                    // 高さ.
                    for (var r = 1; r <= maxRow; r++)
                    {
                        for (var c = 1; c <= dimension.End.Column; c++)
                        {
                            var cell = worksheet.Cells[r, c];

                            if (string.IsNullOrEmpty(cell.Text)) { continue; }

                            var columnWidth = (int)worksheet.Column(c).Width;

                            var height = CalcTextHeight(graphics, cell, columnWidth);

                            if (worksheet.Row(r).Height < height)
                            {
                                worksheet.Row(r).Height = height;
                            }
                        }
                    }

                    Console.WriteLine("- {0}", sheet.displayName);
                }

                // 保存.
                excel.Save();
            }
        }

        private static void CloneCellFormat(ExcelWorksheet worksheet, int recordStartRow, int row, int column)
        {
            var srcCell = worksheet.Cells[recordStartRow, column];
            var destCell = worksheet.Cells[row, column];

            srcCell.Copy(destCell);
        }

        private static void SetGuid(ExcelWorksheet worksheet, int row, int column, string guid)
        {
            worksheet.SetValue(row, column, guid);

            worksheet.Cells[row, column].Style.Font.Size = 5;
        }

        private static double CalcTextWidth(Graphics graphics, ExcelRange cell)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            var drawingFont = new Font(font.Name, font.Size);

            var size = graphics.MeasureString(cell.Text, drawingFont);
            
            return Convert.ToDouble(size.Width) / 5.7;
        }

        private static double CalcTextHeight(Graphics graphics, ExcelRange cell, int width)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            var pixelWidth = Convert.ToInt32(width * 7.5);

            var drawingFont = new Font(font.Name, font.Size);

            var size = graphics.MeasureString(cell.Text, drawingFont, pixelWidth);
            
            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96 * 1.2, 409) + 2;
        }
    }
}

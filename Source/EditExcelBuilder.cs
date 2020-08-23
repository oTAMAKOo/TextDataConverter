
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

        public static void Build(string workspace, IndexData indexData, SheetData[] sheetData, Settings settings)
        {
            var originExcelPath = Path.GetFullPath(settings.ExcelPath);

            var editExcelPath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            ConsoleUtility.Progress("------ Build edit excel file ------");

            //------ エディット用にエクセルファイルを複製 ------

            if (!File.Exists(originExcelPath))
            {
                throw new FileNotFoundException(string.Format("{0} is not exists.", originExcelPath));
            }

            var originXlsxFile = new FileInfo(originExcelPath);

            var editXlsxFile = originXlsxFile.CopyTo(editExcelPath, true);

            //------ レコード情報を書き込み ------

            if (sheetData == null){ return; }

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var worksheets = excel.Workbook.Worksheets;

                // テンプレートシート.

                var templateSheet = worksheets.FirstOrDefault(x => x.Name.ToLower() == settings.TemplateSheetName);

                if (templateSheet == null)
                {
                    throw new Exception(string.Format("Template worksheet {0} not found.", settings.TemplateSheetName));
                }

                templateSheet.Cells.AutoFitColumns(20f, 100f);

                // シート作成.

                foreach (var data in sheetData)
                {
                    if (string.IsNullOrEmpty(data.displayName)) { continue; }

                    if (worksheets.Any(x => x.Name == data.displayName))
                    {
                        throw new Exception(string.Format("Worksheet create failed. Worksheet {0} already exists", data.displayName));
                    }

                    // テンプレートシートを複製.                    
                    var newWorksheet = worksheets.Add(data.displayName, templateSheet);

                    // 保護解除.
                    newWorksheet.Protection.IsProtected = false;
                    // タブ選択状態解除.
                    newWorksheet.View.TabSelected = false;
                }

                // シート順番入れ替え.

                if (worksheets.Any() && indexData != null)
                {
                    for (var i = indexData.sheetNames.Length - 1 ; 0 <= i ; i--)
                    {
                        var sheetName = indexData.sheetNames[i];

                        worksheets.MoveToStart(sheetName);
                    }
                }

                // 先頭のシートをアクティブ化.

                var firstWorksheet = worksheets.FirstOrDefault();

                if (firstWorksheet != null)
                {
                    firstWorksheet.View.TabSelected = true;
                }

                // レコード情報設定.

                var graphics = Graphics.FromImage(new Bitmap(1, 1));

                foreach (var data in sheetData)
                {
                    var worksheet = worksheets.FirstOrDefault(x => x.Name == data.displayName);

                    if (worksheet == null)
                    {
                        ConsoleUtility.Error("Worksheet:{0} not found.", data.displayName);
                        continue;
                    }

                    var dimension = worksheet.Dimension;

                    var records = data.records;

                    if (records == null){ continue; }

                    worksheet.SetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X, data.sheetName);

                    SetGuid(worksheet, Constants.SheetGuidAddress.Y, Constants.SheetGuidAddress.X, data.guid);

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

                    for (var i = 0; i < records.Length; i++)
                    {
                        var r = Constants.RecordStartRow + i;

                        var record = records[i];

                        // Guid.
                        SetGuid(worksheet, r, Constants.GuidColumn, record.guid);

                        // Enum名.
                        worksheet.SetValue(r, Constants.EnumNameColumn, record.enumName);

                        // 説明.
                        worksheet.SetValue(r, Constants.DescriptionColumn, record.description);

                        // テキスト・オプション情報.
                        for (var c = 0; c < record.contents.Length; c++)
                        {
                            var content = record.contents[c];

                            if (content == null) { continue; }
                            
                            worksheet.SetValue(r, Constants.TextStartColumn + c, content.text);

                            var cell = worksheet.Cells[r, Constants.TextStartColumn + c];
                            
                            CellOption.Set(cell, content.comment, content.fontColor, content.backgroundColor);
                        }
                    }

                    // セルサイズを調整.

                    var maxRow = Constants.RecordStartRow + records.Length + 1;
                    
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

                    ConsoleUtility.Task("- {0}", data.displayName);
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

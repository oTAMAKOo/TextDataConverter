
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
                    // セルサイズ調整.
                    newWorksheet.Cells.AutoFitColumns();

                    // エラー無視.
                    var excelIgnoredError = newWorksheet.IgnoredErrors.Add(newWorksheet.Dimension);

                    excelIgnoredError.NumberStoredAsText = true;
                }

                // シート順番入れ替え.

                if (worksheets.Any() && indexData != null)
                {
                    for (var i = indexData.sheetNames.Length - 1 ; 0 <= i ; i--)
                    {
                        var sheetName = indexData.sheetNames[i];

                        if(worksheets.All(x => x.Name != sheetName)){ continue; }

                        worksheets.MoveToStart(sheetName);
                    }
                }

                // 先頭のシートをアクティブ化.

                var firstWorksheet = worksheets.FirstOrDefault();

                if (firstWorksheet != null)
                {
                    firstWorksheet.View.TabSelected = true;
                }

                // コールバック作成.

                var ignoreWrapColumn = new int[]
                {
                    Constants.GuidColumn,
                    Constants.EnumNameColumn,
                };

                Func<int, int, string, bool> wrapTextCallback = (r, c, text) =>
                {
                    var result = true;

                    // 除外対象に含まれていない.
                    result &= !ignoreWrapColumn.Contains(c);
                    // 改行が含まれている.
                    result &= text.FixLineEnd().Contains("\n");

                    return result;
                };

                // レコード情報設定.

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

                    if (records == null) { continue; }

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

                        // テキスト.
                        for (var j = 0; j < record.texts.Length; j++)
                        {
                            var text = record.texts[j];

                            if (string.IsNullOrEmpty(text)) { continue; }

                            worksheet.SetValue(r, Constants.TextStartColumn + j, text);
                        }
                        
                        // セル情報.
                        if (record.cells != null)
                        {
                            foreach (var cellData in record.cells)
                            {
                                if (cellData == null) { continue; }

                                var address = cellData.address.Split(',');

                                var rowStr = address.ElementAtOrDefault(0);
                                var columnStr = address.ElementAtOrDefault(1);

                                if (string.IsNullOrEmpty(rowStr) || string.IsNullOrEmpty(columnStr)) { continue; }

                                var row = Convert.ToInt32(rowStr);
                                var column = Convert.ToInt32(columnStr);

                                ExcelCellUtility.Set<ExcelCell>(worksheet, row, column, cellData);
                            }
                        }
                    }

                    // セルサイズを調整.

                    var maxRow = Constants.RecordStartRow + records.Length + 1;

                    var celFitRange = worksheet.Cells[1, 1, maxRow, dimension.End.Column];

                    ExcelUtility.FitColumnSize(worksheet, celFitRange, null, 150, wrapTextCallback);

                    // GUID行は幅固定.
                    worksheet.Column(Constants.GuidColumn).Width = 20d;

                    ExcelUtility.FitRowSize(worksheet, celFitRange);

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
    }
}

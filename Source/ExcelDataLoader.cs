
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;

namespace GameTextConverter
{
    public static class ExcelDataLoader
    {
        /// <summary> シート名一覧読み込み(.xlsx) </summary>
        public static string[] LoadSheetNames(string workspace, Settings settings)
        {
            var excelFilePath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            if (!File.Exists(excelFilePath)) { return null; }
            
            var sheetNames = new List<string>();

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    sheetNames.Add(worksheet.Name);
                }
            }

            return sheetNames.ToArray();
        }

        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static SheetData[] LoadSheetData(string workspace, Settings settings)
        {
            var excelFilePath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            if (!File.Exists(excelFilePath)) { return null; }

            ConsoleUtility.Progress("------ LoadExcelData ------");

            var sheets = new List<SheetData>();

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    if (worksheet.Name == settings.TemplateSheetName) { continue; }

                    if (settings.IgnoreSheetNames.Contains(worksheet.Name)) { continue; }

                    var sheetEnumNameValue = worksheet.GetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X);

                    var sheetEnumName = ExcelUtility.ConvertValue<string>(sheetEnumNameValue);

                    if (string.IsNullOrEmpty(sheetEnumName)) { continue; }

                    // 既にシートデータがある場合は読み込み.
                    var existSheetData = LoadExistSheetData(workspace, sheetEnumName, settings);
                    
                    // データが出力されていない場合は新規Guidを割り当て.
                    var sheetGuid = existSheetData == null ? Guid.NewGuid().ToString("N") : existSheetData.guid;

                    var sheetData = new SheetData()
                    {
                        guid = sheetGuid,
                        displayName = worksheet.Name,
                        sheetName = sheetEnumName,
                    };

                    var records = new List<RecordData>();

                    for (var r = Constants.RecordStartRow; r <= worksheet.Dimension.End.Row; r++)
                    {
                        var rowValues = ExcelUtility.GetRowValues(worksheet, r).ToArray();

                        var enumName = ExcelUtility.ConvertValue<string>(rowValues, Constants.EnumNameColumn - 1);

                        if (string.IsNullOrEmpty(enumName)) { continue; }

                        var recordGuid = string.Empty;

                        // データが出力されていない場合は新規Guidを割り当て.

                        if (existSheetData != null && existSheetData.records != null)
                        {
                            var existRecordData = existSheetData.records.FirstOrDefault(x => x.enumName == enumName);

                            if (existRecordData != null)
                            {
                                recordGuid = existRecordData.guid;
                            }
                        }

                        if (string.IsNullOrEmpty(recordGuid))
                        {
                            recordGuid = Guid.NewGuid().ToString("N");
                        }

                        var description = ExcelUtility.ConvertValue<string>(rowValues, Constants.DescriptionColumn - 1);

                        var record = new RecordData()
                        {
                            guid = recordGuid,
                            enumName = enumName,
                            description = description,
                        };

                        // 言語タイプ数取得.

                        var textEndColumn = Constants.TextStartColumn;

                        for (var c = Constants.TextStartColumn; c < rowValues.Length; c++)
                        {
                            var textTypeValue = worksheet.GetValue(Constants.TextTypeStartRow, c);

                            var textTypeName = ExcelUtility.ConvertValue<string>(textTypeValue);

                            if (string.IsNullOrEmpty(textTypeName)) { break; }

                            textEndColumn++;
                        }

                        // 実テキスト取得.

                        var contents = new List<ContentData>();

                        for (var c = Constants.TextStartColumn; c < textEndColumn; c++)
                        {
                            var text = ExcelUtility.ConvertValue<string>(rowValues, c - 1);

                            var option = CellOption.Get(worksheet.Cells[r, c]);

                            var data = new ContentData()
                            {
                                text = text,
                                comment = option != null ? option.Item1 : null,
                                fontColor = option != null ? option.Item2 : null,
                                backgroundColor = option != null ? option.Item3 : null,
                            };

                            contents.Add(data);
                        }

                        record.contents = contents.ToArray();

                        records.Add(record);
                    }

                    sheetData.records = records.ToArray();

                    sheets.Add(sheetData);

                    ConsoleUtility.Task("- {0}", sheetData.displayName);
                }
            }

            return sheets.ToArray();
        }

        private static SheetData LoadExistSheetData(string workspace, string sheetName, Settings settings)
        {
            var extension = settings.GetFileExtension();

            var sheetFilePath = PathUtility.Combine(new string[] { workspace, Constants.ContentsFolderName, sheetName }) + extension;

            if (!File.Exists(sheetFilePath)) { return null; }

            return DataLoader.LoadSheetData(sheetFilePath, settings);
        }
    }
}

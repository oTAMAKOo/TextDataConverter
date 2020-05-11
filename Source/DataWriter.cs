using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Extensions;
using Newtonsoft.Json;
using OfficeOpenXml;
using YamlDotNet.Serialization;

namespace GameTextConverter
{
    public sealed class DataWriter
    {
        //----- params -----
        
        //----- field -----

        //----- property -----

        //----- method -----

        public static void Write(string workspace, ExcelData excelData, Settings settings)
        {
            CreateCleanDirectory(workspace);

            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            ConsoleUtility.Progress("------ WriteData ------");

            foreach (var worksheet in excelData.sheets)
            {
                if (string.IsNullOrEmpty(worksheet.sheetName)) { continue; }

                var records = excelData.records.GetValueOrDefault(worksheet.sheetName, new RecordData[0]);

                if (records.IsEmpty()) { continue; }

                var directory = PathUtility.Combine(rootDirectory, worksheet.sheetName);

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // シート情報書き出し.

                if (!string.IsNullOrEmpty(worksheet.sheetName))
                {
                    var fileName = worksheet.sheetName + Constants.SheetFileExtension;

                    var filePath = PathUtility.Combine(rootDirectory, fileName);

                    FileSystem.WriteFile(filePath, worksheet, settings.FileFormat);
                }

                // レコード情報書き出し.

                foreach (var record in records)
                {
                    if (string.IsNullOrEmpty(record.enumName)) { continue; }

                    var fileName = record.enumName + Constants.RecordFileExtension;

                    var filePath = PathUtility.Combine(directory, fileName);

                    FileSystem.WriteFile(filePath, record, settings.FileFormat);
                }

                ConsoleUtility.Task("- {0}", worksheet.sheetName);
            }
        }
    
        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static ExcelData LoadExcelData(string workspace, Settings settings)
        {
            var excelFilePath = PathUtility.Combine(workspace, Constants.EditExcelFile);

            if (!File.Exists(excelFilePath)) { return null; }

            ConsoleUtility.Progress("------ LoadExcelData ------");

            var sheets = new List<SheetData>();
            var records = new Dictionary<string, RecordData[]>();

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    if (worksheet.Name == settings.TemplateSheetName) { continue; }

                    if (settings.IgnoreSheetNames.Contains(worksheet.Name)) { continue; }
                    
                    var sheetEnumNameValue = worksheet.GetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X);

                    var sheetEnumName = ExcelUtility.ConvertValue<string>(sheetEnumNameValue);

                    if (string.IsNullOrEmpty(sheetEnumName)) { continue; }
                    
                    var sheetGuidValue = worksheet.GetValue(Constants.SheetGuidAddress.Y, Constants.SheetGuidAddress.X);

                    var sheetGuid = ExcelUtility.ConvertValue<string>(sheetGuidValue);
                    
                    if (string.IsNullOrEmpty(sheetGuid))
                    {
                        sheetGuid = Guid.NewGuid().ToString("N");
                    }

                    var sheetData = new SheetData()
                    {
                        guid = sheetGuid,
                        index = worksheet.Index,
                        displayName = worksheet.Name,
                        sheetName = sheetEnumName,
                    };

                    var recordList = new List<RecordData>();

                    for (var r = Constants.RecordStartRow; r <= worksheet.Dimension.End.Row; r++)
                    {
                        var rowValues = ExcelUtility.GetRowValues(worksheet, r).ToArray();

                        var enumName = ConvertRowValue<string>(rowValues, Constants.EnumNameColumn);

                        if (string.IsNullOrEmpty(enumName)) { continue; }

                        var recordGuid = string.Empty;

                        recordGuid = ConvertRowValue<string>(rowValues, Constants.GuidColumn);

                        if (string.IsNullOrEmpty(recordGuid))
                        {
                            recordGuid = Guid.NewGuid().ToString("N");
                        }

                        var description = ConvertRowValue<string>(rowValues, Constants.DescriptionColumn);

                        var record = new RecordData()
                        {
                            guid = recordGuid,
                            sheet = sheetData.guid,
                            line = r,
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

                        var texts = new List<string>();

                        for (var c = Constants.TextStartColumn; c < textEndColumn; c++)
                        {
                            var text = ConvertRowValue<string>(rowValues, c);

                            texts.Add(text);
                        }

                        record.texts = texts.ToArray();

                        recordList.Add(record);
                    }

                    records.Add(sheetEnumName, recordList.ToArray());

                    sheets.Add(sheetData);

                    ConsoleUtility.Task("- {0}", sheetData.displayName);
                }
            }

            var excelData = new ExcelData()
            {
                sheets = sheets.ToArray(),
                records = records,
            };

            return excelData;
        }

        private static void CreateCleanDirectory(string exportPath)
        {
            if (string.IsNullOrEmpty(exportPath)) { throw new ArgumentException("exportPath is null"); }

            var directory = PathUtility.Combine(exportPath, Constants.RecordFolderName);

            if (Directory.Exists(directory))
            {
                DirectoryUtility.Delete(directory);

                // ディレクトリの削除は非同期で実行される為、削除完了するまで待機する.
                while (Directory.Exists(directory))
                {
                    Thread.Sleep(10);
                }
            }

            Directory.CreateDirectory(directory);
        }

        private static T ConvertRowValue<T>(object[] values, int index)
        {
            // OfficeOpenXmlはアドレスが(1,1)開始なので合わせる.
            index--;

            if (index < 0 || values.Length <= index)
            {
                throw new ArgumentOutOfRangeException();
            }

            var value = values[index];

            return ExcelUtility.ConvertValue<T>(value);
        }
    }
}

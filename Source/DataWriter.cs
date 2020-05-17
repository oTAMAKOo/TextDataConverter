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

        public static void Write(string workspace, SheetData[] sheetData, Settings settings)
        {
            CreateCleanDirectory(workspace);

            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            var extension = string.Empty;

            switch (settings.FileFormat)
            {
                case FileSystem.Format.Json:
                    extension = Constants.JsonFileExtension;
                    break;
                case FileSystem.Format.Yaml:
                    extension = Constants.YamlFileExtension;
                    break;
            }

            ConsoleUtility.Progress("------ WriteData ------");

            foreach (var data in sheetData)
            {
                if (string.IsNullOrEmpty(data.sheetName)) { continue; }

                var records = data.records;

                if (records == null || records.IsEmpty()) { continue; }

                // シート情報書き出し.

                if (!string.IsNullOrEmpty(data.sheetName))
                {
                    var fileName = data.sheetName + extension;

                    var filePath = PathUtility.Combine(rootDirectory, fileName);

                    FileSystem.WriteFile(filePath, data, settings.FileFormat);
                }

                ConsoleUtility.Task("- {0}", data.sheetName);
            }
        }
    
        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static SheetData[] LoadExcelData(string workspace, Settings settings)
        {
            var excelFilePath = PathUtility.Combine(workspace, Constants.EditExcelFile);

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

                    var records = new List<RecordData>();

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
                            var text = ConvertRowValue<string>(rowValues, c);

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


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;

namespace TextDataConverter
{
    public static class ExcelDataLoader
    {
        private sealed class RecordEntry
        {
            public SheetData sheet = null;

            public RecordData record = null;
        }

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
                    if (string.Equals(worksheet.Name, settings.TemplateSheetName, StringComparison.OrdinalIgnoreCase)) { continue; }

                    if (settings.IgnoreSheetNames.Contains(worksheet.Name)) { continue; }

                    var sheetEnumNameValue = worksheet.GetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X);

                    var sheetEnumName = ExcelUtility.ConvertValue<string>(sheetEnumNameValue);

                    if (string.IsNullOrEmpty(sheetEnumName)) { continue; }
                    
                    // 既にシートデータがある場合は読み込み.
                    var existSheetData = LoadExistSheetData(workspace, sheetEnumName, settings);
                    
                    // データが出力されていない場合は新規Guidを割り当て.
                    var sheetGuid = existSheetData == null || string.IsNullOrEmpty(existSheetData.guid) ? 
                            Guid.NewGuid().ToString("N") : 
                            existSheetData.guid;

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

                        var guid = ExcelUtility.ConvertValue<string>(rowValues, Constants.GuidColumn - 1);

                        var enumName = ExcelUtility.ConvertValue<string>(rowValues, Constants.EnumNameColumn - 1);

                        if (string.IsNullOrEmpty(enumName)) { continue; }

                        // 既存のデータがある場合.

                        if (existSheetData != null && existSheetData.records != null)
                        {
                            RecordData existRecordData = null;

                            // 既に出力済みだがインポートしていない状態でシートにGUIDが登録されていない.

                            existRecordData = existSheetData.records.FirstOrDefault(x => x.enumName == enumName);

                            if (existRecordData != null)
                            {
                                guid = existRecordData.guid;
                            }

                            // 既に出力済みでGUIDが割り振り済み.

                            existRecordData = existSheetData.records.FirstOrDefault(x => x.guid == guid);

                            if (existRecordData != null)
                            {
                                guid = existRecordData.guid;
                            }
                        }

                        if (string.IsNullOrEmpty(guid))
                        {
                            guid = Guid.NewGuid().ToString("N");
                        }

                        var description = ExcelUtility.ConvertValue<string>(rowValues, Constants.DescriptionColumn - 1);

                        var record = new RecordData()
                        {
                            guid = guid,
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

                        // テキスト取得.

                        var texts = new List<string>();

                        for (var c = Constants.TextStartColumn; c < textEndColumn; c++)
                        {
                            var text = ExcelUtility.ConvertValue<string>(rowValues, c - 1);
                            
                            texts.Add(text);
                        }

                        record.texts = texts.ToArray();

                        // セル情報取得.

                        var cells = new List<ExcelCell>();

                        for (var c = Constants.TextStartColumn; c < textEndColumn; c++)
                        {
                            var cellData = ExcelCellUtility.Get<ExcelCell>(worksheet, r, c);

                            if (cellData == null){ continue; }

                            cellData.address = string.Format("{0},{1}", r, c);

                            cells.Add(cellData);
                        }

                        record.cells = cells.Any() ? cells.ToArray() : null;

                        records.Add(record);
                    }

                    sheetData.records = records.ToArray();

                    sheets.Add(sheetData);

                    ConsoleUtility.Task("- {0}", sheetData.displayName);
                }
            }

            return sheets.ToArray();
        }

        /// <summary> テキストGuidの重複を解消 </summary>
        public static void ResolveDuplicateGuid(string workspace, SheetData[] sheets, Settings settings)
        {
            // 出力済みデータでGuidを保有しているレコード. (Guidが競合した場合に元の保有者へGuidを残す為)

            var guidOwners = new Dictionary<string, string>();

            foreach (var sheet in sheets)
            {
                var existSheetData = LoadExistSheetData(workspace, sheet.sheetName, settings);

                if (existSheetData == null) { continue; }

                if (existSheetData.records == null) { continue; }

                foreach (var record in existSheetData.records)
                {
                    if (string.IsNullOrEmpty(record.guid)) { continue; }

                    if (guidOwners.ContainsKey(record.guid)) { continue; }

                    guidOwners.Add(record.guid, BuildRecordKey(existSheetData.sheetName, record.enumName));
                }
            }

            // Guidが重複するレコードには新規Guidを発行.

            var assignedRecords = new Dictionary<string, RecordEntry>();

            var reassignedLogs = new List<string>();

            foreach (var sheet in sheets)
            {
                if (sheet.records == null) { continue; }

                foreach (var record in sheet.records)
                {
                    var entry = new RecordEntry() { sheet = sheet, record = record };

                    var assignedRecord = assignedRecords.GetValueOrDefault(record.guid);

                    if (assignedRecord == null)
                    {
                        assignedRecords.Add(record.guid, entry);

                        continue;
                    }

                    // 出力済みデータでGuidを保有していたレコードにGuidを残す. (既存のテキスト参照を壊さない為)

                    var recordKey = BuildRecordKey(sheet.sheetName, record.enumName);

                    var target = entry;

                    if (guidOwners.GetValueOrDefault(record.guid) == recordKey)
                    {
                        target = assignedRecord;

                        assignedRecords[record.guid] = entry;
                    }

                    var newGuid = GenerateGuid(assignedRecords);

                    reassignedLogs.Add(string.Format("- {0} : {1} -> {2}", BuildRecordKey(target.sheet.sheetName, target.record.enumName), target.record.guid, newGuid));

                    target.record.guid = newGuid;

                    assignedRecords.Add(newGuid, target);
                }
            }

            if (reassignedLogs.IsEmpty()) { return; }

            ConsoleUtility.Warning("------ Duplicate text guid ------");

            foreach (var log in reassignedLogs)
            {
                ConsoleUtility.Warning(log);
            }

            ConsoleUtility.Warning("");
            ConsoleUtility.Warning("New guid was assigned. Run import to apply to excel.");
        }

        /// <summary> 未使用のGuidを発行 </summary>
        private static string GenerateGuid(Dictionary<string, RecordEntry> assignedRecords)
        {
            var guid = Guid.NewGuid().ToString("N");

            while (assignedRecords.ContainsKey(guid))
            {
                guid = Guid.NewGuid().ToString("N");
            }

            return guid;
        }

        /// <summary> レコード識別名を取得 </summary>
        private static string BuildRecordKey(string sheetName, string enumName)
        {
            return string.Format("{0}-{1}", sheetName, enumName);
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

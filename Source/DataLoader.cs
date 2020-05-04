
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Extensions;

namespace GameTextConverter
{
    public sealed class DataLoader
    {
        //----- params -----
        
        //----- field -----

        //----- property -----

        //----- method -----
        
        /// <summary> エクセル情報読み込み </summary>
        public static ExcelData Load(string workspace, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            if (!Directory.Exists(rootDirectory)) { throw new DirectoryNotFoundException(); }

            // シート情報読み込み.

            var worksheets = LoadSheetData(rootDirectory, settings);

            // レコード情報読み込み.

            Console.WriteLine("------ LoadRecordData ------");

            var worksheetRecords = new Dictionary<string, RecordData[]>();

            foreach (var worksheet in worksheets)
            {
                var records = LoadRecordData(rootDirectory, worksheet, settings);

                if (records != null)
                {
                    worksheetRecords.Add(worksheet.sheetName, records);
                }

                Console.WriteLine("- {0}", worksheet.displayName);
            }

            var excelData = new ExcelData()
            {
                sheets = worksheets,
                records = worksheetRecords,
            };

            return excelData;
        }

        private static SheetData[] LoadSheetData(string rootDirectory, Settings settings)
        {
            var sheetFiles = Directory.EnumerateFiles(rootDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(x => Path.GetExtension(x) == Constants.SheetFileExtension)
                .ToArray();

            var sheets = new List<SheetData>();

            Console.WriteLine("------ LoadSheetData ------");

            foreach (var sheetFile in sheetFiles)
            {
                var sheet = FileSystem.LoadFile<SheetData>(sheetFile, settings.FileFormat);

                if (sheet != null)
                {
                    Console.WriteLine("- {0}", sheet.displayName);

                    sheets.Add(sheet);
                }
            }

            return sheets.ToArray();
        }

        private static RecordData[] LoadRecordData(string rootDirectory, SheetData worksheet, Settings settings)
        {
            var worksheetDirectory = PathUtility.Combine(rootDirectory, worksheet.sheetName);

            if (!Directory.Exists(worksheetDirectory)) { return new RecordData[0]; }

            var recordFiles = Directory.EnumerateFiles(worksheetDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(x => Path.GetExtension(x) == Constants.RecordFileExtension)
                .ToArray();

            var records = new List<RecordData>();

            foreach (var recordFile in recordFiles)
            {
                var record = FileSystem.LoadFile<RecordData>(recordFile, settings.FileFormat);

                if (record != null)
                {
                    records.Add(record);
                }
            }

            return records.ToArray();
        }
    }
}

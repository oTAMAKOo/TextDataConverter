
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
        public static SheetData[] Load(string workspace, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            if (!Directory.Exists(rootDirectory)) { throw new DirectoryNotFoundException(); }

            // シート情報読み込み.

            var sheetData = LoadSheetData(rootDirectory, settings);
            
            return sheetData;
        }

        private static SheetData[] LoadSheetData(string rootDirectory, Settings settings)
        {
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
            
            var sheetFiles = Directory.EnumerateFiles(rootDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(x => Path.GetExtension(x) == extension)
                .ToArray();

            var sheets = new List<SheetData>();

            ConsoleUtility.Progress("------ LoadSheetData ------");

            foreach (var sheetFile in sheetFiles)
            {
                var sheet = FileSystem.LoadFile<SheetData>(sheetFile, settings.FileFormat);

                if (sheet != null)
                {
                    ConsoleUtility.Task("- {0}", sheet.displayName);

                    sheets.Add(sheet);
                }
            }

            return sheets.ToArray();
        }
    }
}

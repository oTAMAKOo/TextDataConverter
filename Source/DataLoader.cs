
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Extensions;

namespace TextDataConverter
{
    public sealed class DataLoader
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----
        
        public static IndexData LoadSheetIndex(string workspace, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.ContentsFolderName);

            if (!Directory.Exists(rootDirectory)) { return null; }

            var fileName = Path.ChangeExtension(settings.EditExcelFileName, Constants.IndexFileExtension);

            var indexFilePath = PathUtility.Combine(rootDirectory, fileName);

            return FileSystem.LoadFile<IndexData>(indexFilePath, settings.FileFormat);
        }
        
        public static SheetData[] LoadAllSheetData(string workspace, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.ContentsFolderName);

            if (!Directory.Exists(rootDirectory)) { return null; }

            // シート情報読み込み.

            var extension = settings.GetFileExtension();

            var sheetFiles = Directory.EnumerateFiles(rootDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(x => Path.GetExtension(x) == extension)
                .ToArray();

            var sheets = new List<SheetData>();

            if (sheetFiles.IsEmpty()){ return new SheetData[0]; }
            
            ConsoleUtility.Progress("------ Load SheetData ------");

            foreach (var sheetFile in sheetFiles)
            {
                var sheet = LoadSheetData(sheetFile, settings);

                if (sheet != null)
                {
                    ConsoleUtility.Task("- {0}", sheet.displayName);

                    sheets.Add(sheet);
                }
            }

            return sheets.ToArray();
        }

        public static SheetData LoadSheetData(string filePath, Settings settings)
        {
            return FileSystem.LoadFile<SheetData>(filePath, settings.FileFormat);
        }
    }
}

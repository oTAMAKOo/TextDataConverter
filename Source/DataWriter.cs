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

        public static void WriteSheetIndex(string workspace, string[] sheetNames, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.ContentsFolderName);

            if (!Directory.Exists(rootDirectory)) { throw new DirectoryNotFoundException(); }

            var filePath = PathUtility.Combine(rootDirectory, Constants.SheetIndexFileName);
            
            var indexData = new IndexData()
            {
                sheetNames = sheetNames
            };

            FileSystem.WriteFile(filePath, indexData, settings.FileFormat);
        }

        public static void WriteAllSheetData(string workspace, SheetData[] sheetData, Settings settings)
        {
            CreateCleanDirectory(workspace);
            
            var rootDirectory = PathUtility.Combine(workspace, Constants.ContentsFolderName);

            var extension = settings.GetFileExtension();

            if (sheetData.IsEmpty()){ return; }

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

        private static void CreateCleanDirectory(string exportPath)
        {
            if (string.IsNullOrEmpty(exportPath)) { throw new ArgumentException("exportPath is null"); }

            var directory = PathUtility.Combine(exportPath, Constants.ContentsFolderName);

            DirectoryUtility.Clean(directory);
        }
    }
}

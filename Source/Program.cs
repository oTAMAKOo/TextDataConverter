
using System;
using System.IO;
using System.Linq;
using System.Text;
using CommandLine;
using OfficeOpenXml;
using Extensions;

namespace TextDataConverter
{
    class Program
    {
        //----- params -----

        private class CommandLineOptions
        {
            [Option("workspace", Required = false)]
            public string Workspace { get; set; } = "";
            [Option("mode", Required = false, HelpText = "Convert mode. (import or export).")]
            public string Mode { get; set; } = "import";
        }

        //----- field -----

        //----- property -----

        //----- method -----

        static void Main(string[] args)
        {
            // コマンドライン引数.

            var options = Parser.Default.ParseArguments<CommandLineOptions>(args) as Parsed<CommandLineOptions>;

            if (options == null)
            {
                Exit(1, "Arguments parse failed.");

                return;
            }

            var workspace = options.Value.Workspace;

            var mode = options.Value.Mode;

            // 設定ファイル.

            var settings = new Settings();

            if (!settings.Load(workspace))
            {
                Exit(1, "Settings load failed.");
            }

            // EPPlus License setup.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            /*=== 開発用 ========================================

            #if DEBUG

            var currentDirectory = @"";

            workspace = @"";

            mode = "import";
            
            Directory.SetCurrentDirectory(currentDirectory);

            #endif

            //==================================================*/

            Console.WriteLine();

            ConsoleUtility.Info("Workspace : {0}", workspace);

            try
            {
                switch (mode)
                {
                    case "import":
                        Import(workspace, settings);
                        break;

                    case "export":
                        Export(workspace, settings);
                        break;

                    default:
                        throw new NotSupportedException("Unknown mode selection.");
                }

                ConsoleUtility.Info("Complete!");
            }
            catch (Exception e)
            {
                Exit(1, e.ToString());
            }
            
            Exit(0);
        }

        private static void Import(string workspace, Settings settings)
        {
            var editExcelPath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            if (ExcelUtility.IsLocked(editExcelPath))
            {
                throw new Exception($"Excel file locked.\nPath : {editExcelPath}");
            }
            
            var indexData = DataLoader.LoadSheetIndex(workspace, settings);

            var sheetData = DataLoader.LoadAllSheetData(workspace, settings);

            EditExcelBuilder.Build(workspace, indexData, sheetData, settings);
        }

        private static void Export(string workspace, Settings settings)
        {
            var sheetData = ExcelDataLoader.LoadSheetData(workspace, settings);

            var duplicates = sheetData.GroupBy(x => x.sheetName)
                .Where(x => 1 < x.Count())
                .Select(g => g.Key)
                .ToArray();

            if (duplicates.Any())
            {
                var builder = new StringBuilder();

                builder.AppendLine();

                foreach (var item in duplicates)
                {
                    builder.AppendFormat("Duplicate sheet name exists. SheetName = {0}", item).AppendLine();
                }

                Exit(1, builder.ToString());

                return;
            }

            DataWriter.WriteAllSheetData(workspace, sheetData, settings);

            var sheetNames = ExcelDataLoader.LoadSheetNames(workspace, settings);

            DataWriter.WriteSheetIndex(workspace, sheetNames, settings);
        }

        private static void Exit(int exitCode, string message = "")
        {
            if (!string.IsNullOrEmpty(message))
            {
                ConsoleUtility.Error(message);
            }

            // 正常終了以外ならコンソールを閉じない.
            if (exitCode != 0)
            {
                Console.ReadLine();
            }

            Environment.Exit(exitCode);
        }
    }
}

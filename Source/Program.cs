
using System;
using System.IO;
using CommandLine;
using Extensions;

namespace GameTextConverter
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
            }

            // 設定ファイル.

            var settings = new Settings();

            if (!settings.Load())
            {
                Exit(1, "Settings load failed.");
            }

            // メイン処理.

            var workspace = options.Value.Workspace;

            var mode = options.Value.Mode;

            /*=== 開発用 ========================================

            #if DEBUG

            workspace = @"";

            Directory.SetCurrentDirectory(workspace);

            mode = "import";

            #endif

            //==================================================*/

            Console.WriteLine();

            ConsoleUtility.Info("Workspace : {0}", workspace);

            try
            {
                switch (mode)
                {
                    case "import":
                        {
                            if (!IsEditExcelFileLocked(workspace, settings))
                            {
                                var indexData = DataLoader.LoadSheetIndex(workspace, settings);

                                var sheetData = DataLoader.LoadAllSheetData(workspace, settings);

                                EditExcelBuilder.Build(workspace, indexData, sheetData, settings);
                            }
                        }
                        break;

                    case "export":
                        {
                            var sheetData = ExcelDataLoader.LoadSheetData(workspace, settings);

                            DataWriter.WriteAllSheetData(workspace, sheetData, settings);

                            var sheetNames = ExcelDataLoader.LoadSheetNames(workspace, settings);

                            DataWriter.WriteSheetIndex(workspace, sheetNames, settings);
                        }
                        break;

                    default:
                        throw new NotSupportedException("Unknown mode selection.");
                }
            }
            catch (Exception e)
            {
                Exit(1, e.ToString());
            }

            ConsoleUtility.Info("Complete!");

            // 終了.

            Exit(0);
        }

        private static bool IsEditExcelFileLocked(string workspace, Settings settings)
        {
            var editExcelPath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            // ファイルが存在＋ロック時はエラー.
            if (File.Exists(editExcelPath))
            {
                if (FileUtility.IsFileLocked(editExcelPath))
                {
                    Exit(1, string.Format("File locked. {0}", editExcelPath));
                    return true;
                }
            }

            return false;
        }
        
        // レコードファイルのディレクトリ取得.
        private static string GetRecordFileDirectory(string directory)
        {
            return PathUtility.Combine(directory, Constants.ContentsFolderName);
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

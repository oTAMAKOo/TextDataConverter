
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace GameTextConverter
{
    public sealed class RecordOption
    {
        //----- params -----

        public class CellOption
        {
            public string recordName;
            public CellInfo[] cellInfos;
        }

        public class CellInfo
        {
            public string comment;
            public string fontColor;
            public string backgroundColor;
        }

        //----- field -----

        //----- property -----

        //----- method -----
        
        public static void Write(string workspace, ExcelData excelData, Settings settings)
        {
            ConsoleUtility.Progress("------ WriteCellOption ------");

            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            var excelFilePath = PathUtility.Combine(workspace, Constants.EditExcelFile);

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excelData.sheets)
                {
                    if (string.IsNullOrEmpty(worksheet.sheetName)) { continue; }

                    var records = excelData.records.GetValueOrDefault(worksheet.sheetName, new RecordData[0]);

                    if (records.IsEmpty()) { continue; }

                    var directory = PathUtility.Combine(rootDirectory, worksheet.sheetName);

                    var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == worksheet.displayName);

                    if (sheet != null)
                    {
                        foreach (var record in records)
                        {
                            if (string.IsNullOrEmpty(record.enumName)) { continue; }

                            var cellOption = new CellOption()
                            {
                                recordName = record.enumName,
                                cellInfos = new CellInfo[record.texts.Length],
                            };

                            var r = record.line;

                            for (var i = 0; i < record.texts.Length; i++)
                            {
                                var c = Constants.TextStartColumn + i;

                                cellOption.cellInfos[i] = GetCellInfo(sheet.Cells[r, c]);
                            }

                            if (cellOption.cellInfos.Any(x => x != null))
                            {
                                var fileName = record.enumName + Constants.CellOptionFileExtension;

                                var filePath = PathUtility.Combine(directory, fileName);

                                FileSystem.WriteFile(filePath, cellOption, settings.FileFormat);                                
                            }
                        }

                        ConsoleUtility.Task("- {0}", sheet.Name);
                    }
                }
            }
        }
        
        /// <summary> セルオプション情報読み込み </summary>
        public static void Load(string workspace, ExcelData excelData, Settings settings)
        {
            var rootDirectory = PathUtility.Combine(workspace, Constants.RecordFolderName);

            var excelFilePath = PathUtility.Combine(workspace, Constants.EditExcelFile);

            ConsoleUtility.Progress("------ LoadCellOption ------");

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excelData.sheets)
                {
                    if (string.IsNullOrEmpty(worksheet.sheetName)) { continue; }

                    var records = excelData.records.GetValueOrDefault(worksheet.sheetName, new RecordData[0]);

                    if (records.IsEmpty()) { continue; }

                    var directory = PathUtility.Combine(rootDirectory, worksheet.sheetName);

                    var cellOptionFiles = Directory.EnumerateFiles(directory, "*.*", SearchOption.TopDirectoryOnly)
                        .Where(x => Path.GetExtension(x) == Constants.CellOptionFileExtension)
                        .ToArray();

                    var cellOptions = cellOptionFiles
                        .Select(x => FileSystem.LoadFile<CellOption>(x, settings.FileFormat))
                        .Where(x => x != null)
                        .ToArray();

                    if (cellOptions.Any())
                    {
                        var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == worksheet.displayName);

                        if (sheet != null)
                        {
                            foreach (var record in records)
                            {
                                if (string.IsNullOrEmpty(record.enumName)) { continue; }

                                var cellOption = cellOptions.FirstOrDefault(x => x.recordName == record.enumName);

                                if (cellOption == null) { continue; }

                                var r = record.line;

                                for (var i = 0; i < record.texts.Length; i++)
                                {
                                    var c = Constants.TextStartColumn + i;

                                    var cellInfo = cellOption.cellInfos.ElementAtOrDefault(i);

                                    if (cellInfo != null)
                                    {
                                        SetCellInfos(sheet.Cells[r, c], cellInfo);
                                    }
                                }
                            }

                            ConsoleUtility.Task("- {0}", sheet.Name);
                        }
                    }
                }

                excel.Save();
            }
        }

        private static void SetCellInfos(ExcelRange cell, CellInfo cellInfo)
        {
            if (cellInfo == null) { return; }

            if (!string.IsNullOrEmpty(cellInfo.comment))
            {
                cell.AddComment(cellInfo.comment, "REF");
            }

            if (!string.IsNullOrEmpty(cellInfo.fontColor))
            {
                cell.Style.Font.Color.SetColor(ColorTranslator.FromHtml(cellInfo.fontColor));
            }

            if (!string.IsNullOrEmpty(cellInfo.backgroundColor))
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(cellInfo.backgroundColor));
            }
        }

        private static CellInfo GetCellInfo(ExcelRange cell)
        {
            CellInfo cellInfo = null;
            
            var comment = cell.Comment != null ? cell.Comment.Text : null;

            var fontColor = GetColorCode(cell, cell.Style.Font.Color);
            var backgroundColor = GetColorCode(cell, cell.Style.Fill.BackgroundColor);

            var changed = false;
            
            changed |= !string.IsNullOrEmpty(comment);
            changed |= !string.IsNullOrEmpty(fontColor) && fontColor != "#FF000000";
            changed |= !string.IsNullOrEmpty(backgroundColor) && backgroundColor != "#FFFFFFFF";

            if (changed)
            {
                cellInfo = new CellInfo()
                {
                    comment = comment,
                    fontColor = fontColor,
                    backgroundColor = backgroundColor,
                };
            }

            return cellInfo;
        }

        private static string GetColorCode(ExcelRange cell, ExcelColor color)
        {
            string colorCode = null;

            if (!string.IsNullOrEmpty(color.Rgb))
            {
                colorCode = "#" + color.Rgb;
            }

            if (!string.IsNullOrEmpty(color.Theme))
            {
                colorCode = null;

                ConsoleUtility.Warning("Theme color not support saved default color.\n[{0}] {1}", cell.Address, cell.Text);
            }

            return colorCode;
        }
    }
}

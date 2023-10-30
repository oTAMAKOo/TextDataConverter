
using Extensions;
using OfficeOpenXml;

namespace TextDataConverter
{
    public sealed class EditExcelBuilder
    {
        //----- params -----

        //----- field -----


        //----- property -----

        //----- method -----

        public static void Build(string workspace, IndexData indexData, SheetData[] sheetData, Settings settings)
        {
            var originExcelPath = Path.GetRelativePath(workspace, settings.ExcelPath);

            var editExcelPath = PathUtility.Combine(workspace, settings.EditExcelFileName);

            ConsoleUtility.Progress("------ Build edit excel file ------");

            //------ エディット用にエクセルファイルを複製 ------

            if (!File.Exists(originExcelPath))
            {
                throw new FileNotFoundException($"{originExcelPath} is not exists.");
            }

            var originXlsxFile = new FileInfo(originExcelPath);

            var editXlsxFile = originXlsxFile.CopyTo(editExcelPath, true);

            //------ レコード情報を書き込み ------

            if (sheetData == null){ return; }

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var worksheets = excel.Workbook.Worksheets;

                // テンプレートシート.

                var templateSheet = worksheets.FirstOrDefault(x => x.Name.ToLower() == settings.TemplateSheetName);

                if (templateSheet == null)
                {
                    throw new Exception($"Template worksheet {settings.TemplateSheetName} not found.");
                }

                // シート作成.

                foreach (var data in sheetData)
                {
                    if (string.IsNullOrEmpty(data.displayName)) { continue; }

                    if (worksheets.Any(x => x.Name == data.displayName))
                    {
                        throw new Exception($"Worksheet create failed. Worksheet {data.displayName} already exists");
                    }

                    // テンプレートシートを複製.                    
                    var newWorksheet = worksheets.Add(data.displayName, templateSheet);

                    // 保護解除.
                    newWorksheet.Protection.IsProtected = false;
                    // タブ選択状態解除.
                    newWorksheet.View.TabSelected = false;

                    // エラー無視.
                    var excelIgnoredError = newWorksheet.IgnoredErrors.Add(newWorksheet.Dimension);

                    excelIgnoredError.NumberStoredAsText = true;
                }

                // シート順番入れ替え.

                if (worksheets.Any() && indexData != null)
                {
                    for (var i = indexData.sheetNames.Length - 1 ; 0 <= i ; i--)
                    {
                        var sheetName = indexData.sheetNames[i];

                        if(worksheets.All(x => x.Name != sheetName)){ continue; }

                        worksheets.MoveToStart(sheetName);
                    }
                }

                // 先頭のシートをアクティブ化.

                var firstWorksheet = worksheets.FirstOrDefault();

                if (firstWorksheet != null)
                {
                    firstWorksheet.View.TabSelected = true;
                }

                // レコード情報設定.

                foreach (var data in sheetData)
                {
                    var worksheet = worksheets.FirstOrDefault(x => x.Name == data.displayName);

                    if (worksheet == null)
                    {
                        ConsoleUtility.Error("Worksheet:{0} not found.", data.displayName);
                        continue;
                    }

                    var dimension = worksheet.Dimension;

                    var records = data.records;

                    if (records == null) { continue; }

                    worksheet.SetValue(Constants.SheetNameAddress.Y, Constants.SheetNameAddress.X, data.sheetName);

                    // カラム初期幅.

                    var columnsWidth = new Dictionary<int, double>();

                    for (var c = dimension.Start.Column; c <= dimension.End.Column; c++)
                    {
                        var width = worksheet.Columns[c].Width;

                        columnsWidth[c] = width;
                    }

                    // レコード投入用セルを用意.

                    for (var i = 0; i < records.Length; i++)
                    {
                        var recordRow = Constants.RecordStartRow + i;

                        // 行追加.
                        if (worksheet.Cells.End.Row < recordRow)
                        {
                            worksheet.InsertRow(recordRow, 1);
                        }

                        // セル情報コピー.
                        for (var column = 1; column < dimension.End.Column; column++)
                        {
                            CloneCellFormat(worksheet, Constants.RecordStartRow, recordRow, column);
                        }
                    }

                    // 値設定.

                    for (var i = 0; i < records.Length; i++)
                    {
                        var r = Constants.RecordStartRow + i;

                        var record = records[i];

                        // Enum名.
                        worksheet.SetValue(r, Constants.EnumNameColumn, record.enumName);

                        // 説明.
                        worksheet.SetValue(r, Constants.DescriptionColumn, record.description);

                        // テキスト.
                        for (var j = 0; j < record.texts.Length; j++)
                        {
                            var text = record.texts[j];

                            if (string.IsNullOrEmpty(text)) { continue; }

                            // 改行を含む場合は折り畳む.
                            worksheet.Cells[r, Constants.TextStartColumn + j].Style.WrapText = text.FixLineEnd().Contains("\n");

                            // 設定.
                            worksheet.SetValue(r, Constants.TextStartColumn + j, text);
                        }
                        
                        // セル情報.
                        if (record.cells != null)
                        {
                            foreach (var cellData in record.cells)
                            {
                                if (cellData == null) { continue; }

                                var address = cellData.address.Split(',');

                                var rowStr = address.ElementAtOrDefault(0);
                                var columnStr = address.ElementAtOrDefault(1);

                                if (string.IsNullOrEmpty(rowStr) || string.IsNullOrEmpty(columnStr)) { continue; }

                                var row = Convert.ToInt32(rowStr);
                                var column = Convert.ToInt32(columnStr);

                                ExcelCellUtility.Set<ExcelCell>(worksheet, row, column, cellData);
                            }
                        }
                    }

                    // セルサイズを調整.
                
                    var celFitRange = worksheet.Cells[1, 1, dimension.End.Row, dimension.End.Column];

                    celFitRange.AutoFitColumns();

                    for (var c = celFitRange.Start.Column; c <= celFitRange.End.Column; c++)
                    {
                        var baseWidth = columnsWidth.GetValueOrDefault(c);
                        var currentWidth = worksheet.Column(c).Width;

                        if (currentWidth < baseWidth)
                        {
                            worksheet.Column(c).Width = baseWidth;
                        }
                    }

                    ConsoleUtility.Task("- {0}", data.displayName);                    
                }

                // 保存.
                excel.Save();
            }
        }

        private static void CloneCellFormat(ExcelWorksheet worksheet, int recordStartRow, int row, int column)
        {
            var srcCell = worksheet.Cells[recordStartRow, column];
            var destCell = worksheet.Cells[row, column];

            srcCell.Copy(destCell);
        }
    }
}


using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace GameTextConverter
{
    public sealed class IndexData
    {
        public string[] sheetNames = null;
    }

    public sealed class SheetData
    {
        public string sheetName = null;

        public string displayName = null;

        public string guid = null;
        
        public RecordData[] records = null;
    }

    public sealed class RecordData
    {
        public string enumName = null;

        public string description = null;
        
        public string guid = null;

        public string[] texts = null;

        public CellData[] cells = null;
    }

    public sealed class CellData
    {
        public class Color
        {
            public string rgb = null;
            public eThemeSchemeColor? theme = null;
            public decimal? tint = null;
        }

        public string address = null;

        public string comment = null;

        public Color fontColor = null;

        public Color backgroundColor = null;

        public ExcelFillStyle patternType = ExcelFillStyle.Solid;
    }
}

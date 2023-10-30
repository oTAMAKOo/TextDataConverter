
namespace TextDataConverter
{
    public sealed class IndexData
    {
        public string[] sheetNames = null;
    }

    public sealed class SheetData
    {
        public string sheetName = null;

        public string displayName = null;
        
        public RecordData[] records = null;
    }

    public sealed class RecordData
    {
        public string identifier = null;

        public string enumName = null;

        public string description = null;

        public string[] texts = null;

        public ExcelCell[] cells = null;
    }

    public sealed class ExcelCell : Extensions.ExcelCell
    {
        public string address = null;
    }
}

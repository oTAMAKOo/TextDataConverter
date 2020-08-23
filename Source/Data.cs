
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

        public ContentData[] contents = null;
    }

    public sealed class ContentData
    {
        public string text = null;

        public string comment = null;

        public string fontColor = null;

        public string backgroundColor = null;
    }
}

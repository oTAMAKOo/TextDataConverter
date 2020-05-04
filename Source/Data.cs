
using System;
using System.Collections.Generic;

namespace GameTextConverter
{
    public sealed class ExcelData
    {
        public SheetData[] sheets = null;

        public Dictionary<string, RecordData[]> records = null;
    }

    public sealed class SheetData
    {
        public string guid = null;

        public int index = 0;

        public string sheetName = null;

        public string displayName = null;
    }

    public sealed class RecordData
    {
        public string guid = null;

        public string sheet = null;

        public int line = 0;

        public string enumName = null;

        public string description = null;

        public string[] texts = null;
    }
}

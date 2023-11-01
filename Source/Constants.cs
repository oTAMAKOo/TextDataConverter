
using System.Drawing;

namespace TextDataConverter
{
    public static class Constants
    {
        /// <summary>インデックスファイル拡張子 </summary>
        public const string IndexFileExtension = ".index";

        /// <summary> データフォルダ名 </summary>
        public const string ContentsFolderName = "Contents";

        /// <summary> Json拡張子 </summary>
        public const string JsonFileExtension = ".json";

        /// <summary> Yaml拡張子 </summary>
        public const string YamlFileExtension = ".yaml";

        /// <summary> Excel拡張子 </summary>
        public const string ExcelExtension = ".xlsx";

        /// <summary> シートGuid定義アドレス </summary>
        public static readonly Point SheetGuidAddress = new Point(1, 1);

        /// <summary> シートEnum名定義アドレス </summary>
        public static readonly Point SheetNameAddress = new Point(2, 1);
        
        /// <summary> Guid列 </summary>
        public const int GuidColumn = 1;

        /// <summary> 区分列 </summary>
        public const int DescriptionColumn = 2;

        /// <summary> Enum名列 </summary>0
        public const int EnumNameColumn = 3;

        /// <summary> テキスト開始列 </summary>
        public const int TextStartColumn = 4;

        /// <summary> テキストタイプ開始行 </summary>
        public const int TextTypeStartRow = 2;

        /// <summary> データ開始行 </summary>
        public const int RecordStartRow = 3;
    }
}

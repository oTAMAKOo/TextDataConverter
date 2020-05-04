
using System;
using System.Drawing;

namespace GameTextConverter
{
    public static class Constants
    {
        /// <summary> レコードフォルダ名 </summary>
        public const string RecordFolderName = "Records";

        /// <summary> シートファイル拡張子 </summary>
        public const string SheetFileExtension = ".sheet";

        /// <summary> レコードファイル拡張子 </summary>
        public const string RecordFileExtension = ".record";

        /// <summary> セルオプションファイル拡張子 </summary>
        public const string CellOptionFileExtension = ".option";

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

        /// <summary> 編集エクセルファイル </summary>
        public const string EditExcelFile = "GameText.xlsx";
    }
}

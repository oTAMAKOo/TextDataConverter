﻿
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;

namespace Extensions
{
    public static class ExcelUtility
    {
        public static bool IsLocked(string excelPath)
        {
            // ファイルが存在＋ロック時はエラー.
            if (File.Exists(excelPath))
            {
                if (FileUtility.IsFileLocked(excelPath)) { return true; }
            }

            return false;
        }

        public static T ConvertValue<T>(object value)
        {
            if (value is T) { return (T)value; }

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch (InvalidCastException)
            {
                return default(T);
            }
        }


        public static T ConvertValue<T>(object[] values, int index)
        {
            if (index < 0 || values.Length <= index)
            {
                throw new ArgumentOutOfRangeException();
            }

            var value = values[index];

            return ConvertValue<T>(value);
        }

        /// <summary> 1行取得. </summary>
        public static IEnumerable<object> GetRowValues(ExcelWorksheet sheet, int row)
        {
            var address = sheet.Dimension;

            var values = new List<object>();

            for (var i = address.Start.Column; i <= address.End.Column; i++)
            {
                var value = sheet.GetValue(row, i);

                values.Add(value);
            }

            return values;
        }

        /// <summary> 1行取得(文字列). </summary>
        public static IEnumerable<string> GetRowValueTexts(ExcelWorksheet sheet, int row)
        {
            var address = sheet.Dimension;

            var values = new List<string>();

            for (var i = address.Start.Column; i <= address.End.Column; i++)
            {
                values.Add(sheet.Cells[row, i].Text);
            }

            return values;
        }
    }
}

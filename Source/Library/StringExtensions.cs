﻿
using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace Extensions
{
    public static partial class StringExtensions
    {
        /// <summary> SringBuilderを使って複数の文字列を連結. </summary>
        public static string Combine(this string value, string[] targets)
        {
            var builder = new StringBuilder(value);

            foreach (var target in targets)
            {
                if (string.IsNullOrEmpty(target)) { continue; }

                builder.Append(target);
            }

            return builder.ToString();
        }

        /// <summary> 改行コードを統一 </summary>
        public static string FixLineEnd(this string value, string newLineStr = "\n")
        {
            if (string.IsNullOrEmpty(value)) { return null; }

            var regex = new Regex(@"\r|\r\n");

            if (regex.IsMatch(value))
            {
                value = value.Replace("\r\n", "\r").Replace("\r", newLineStr);
            }

            return value;
        }

        /// <summary> エスケープシーケンス(\t、\nなど)を制御コード(\\t、\\n)に変換 </summary>
        public static string Escape(this string value)
        {
            return Regex.Escape(value);
        }

        /// <summary> 制御コード(\\t、\\n)をエスケープシーケンス(\t、\nなど)に変換 </summary>
        public static string Unescape(this string value)
        {
            return Regex.Unescape(value);
        }

        /// <summary> 指定された文字列が指定範囲と一致するか判定 </summary>
        public static bool SubstringEquals(this string value, int startIndex, int length, string target)
        {
            if (value.Length < startIndex + length) { return false; }

            return value.SafeSubstring(startIndex, length) == target;
        }

        /// <summary>
        /// このインスタンスから部分文字列を取得します.
        /// 部分文字列は、文字列中の指定した文字の位置で開始し、文字列の末尾まで続きます.
        /// </summary>
        public static string SafeSubstring(this string value, int startIndex, int? length = null)
        {
            var len = length.HasValue ? length.Value : int.MaxValue;

            return new string((value ?? string.Empty).Skip(startIndex).Take(len).ToArray());
        }

        /// <summary> 指定された文字列をSHA256でハッシュ化 </summary>
        public static string GetHash(this string value)
        {
            return CalcSHA256(value, Encoding.UTF8);
        }

        /// <summary> 指定された文字列をSHA256でハッシュ化 </summary>
        public static string GetHash(this string value, Encoding enc)
        {
            return CalcSHA256(value, enc);
        }

        // SHA256ハッシュ生成.
        private static string CalcSHA256(string value, Encoding enc)
        {
            #if NET6_0_OR_GREATER

            var hashAlgorithm = SHA256.Create();

            #else

            var hashAlgorithm = new SHA256CryptoServiceProvider();

            #endif

            return string.Join("", hashAlgorithm.ComputeHash(enc.GetBytes(value)).Select(x => $"{x:x2}"));
        }

        /// <summary> 指定された文字列をCRC32でハッシュ化 </summary>
        public static string GetCRC(this string value)
        {
            return CalcCRC32(value, Encoding.UTF8);
        }

        /// <summary> 指定された文字列をCRC32でハッシュ化 </summary>
        public static string GetCRC(this string value, Encoding enc)
        {
            return CalcCRC32(value, enc);
        }

        // CRC32ハッシュ生成.
        private static string CalcCRC32(string value, Encoding enc)
        {
            var byteValues = enc.GetBytes(value);

            var crc32 = new CRC32();

            var hashValue = crc32.ComputeHash(byteValues);

            var hashedText = new StringBuilder();

            for (var i = 0; i < hashValue.Length; i++)
            {
                hashedText.AppendFormat("{0:x2}", hashValue[i]);
            }

            return hashedText.ToString();
        }

        /// <summary> 文字列に指定されたキーワード群が含まれるか判定 </summary>
        public static bool IsMatch(this string text, string[] keywords)
        {
            keywords = keywords.Select(x => x.ToLower()).ToArray();

            if (!string.IsNullOrEmpty(text))
            {
                var tl = text.ToLower();
                var matches = 0;

                for (var b = 0; b < keywords.Length; ++b)
                {
                    if (tl.Contains(keywords[b])) ++matches;
                }

                if (matches == keywords.Length)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary> タグ文字列を除外した文字列を返す </summary>
        public static string RemoveTag(this string text)
        {
            return Regex.Replace(text, "<.*?>", string.Empty);
        }        
    }
}

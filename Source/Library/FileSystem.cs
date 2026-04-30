
using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using YamlDotNet.Core;
using YamlDotNet.Serialization;

namespace Extensions
{
    public static class FileSystem
    {
        public enum Format
        {
            Yaml,
            Json,
        }

        public static void WriteFile<T>(string filePath, T value, Format format)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                using (var writer = new StreamWriter(file, new UTF8Encoding(false)))
                {
                    switch (format)
                    {
                        case Format.Json:
                            {
                                using (var jsonTextWriter = new JsonTextWriter(writer))
                                {
                                    jsonTextWriter.Formatting = Formatting.Indented;

                                    var jsonSerializer = new JsonSerializer()
                                    {
                                        Formatting = Formatting.Indented,
                                        NullValueHandling = NullValueHandling.Ignore,
                                    };

                                    jsonSerializer.Serialize(jsonTextWriter, value);
                                }
                            }
                            break;

                        case Format.Yaml:
                            {
                                var builder = new SerializerBuilder();

                                builder.ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull);

                                var yamlSerializer = builder.Build();

                                yamlSerializer.Serialize(writer, value);
                            }
                            break;
                    }
                }
            }
        }

        public static T LoadFile<T>(string filePath, Format format) where T : class
        {
            if (!File.Exists(filePath)) { return null; }

            T result = null;

            string contents = null;

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = new StreamReader(file, new UTF8Encoding(false)))
                {
                    contents = reader.ReadToEnd();
                }
            }

            switch (format)
            {
                case Format.Json:
                    {
                        try
                        {
                            using (var stringReader = new StringReader(contents))
                            using (var jsonTextReader = new JsonTextReader(stringReader))
                            {
                                var jsonSerializer = new JsonSerializer()
                                {
                                    Formatting = Formatting.Indented,
                                    NullValueHandling = NullValueHandling.Ignore,
                                };

                                result = jsonSerializer.Deserialize<T>(jsonTextReader);
                            }
                        }
                        catch (JsonException e)
                        {
                            var line = 0;
                            var column = 0;

                            if (e is JsonReaderException readerEx)
                            {
                                line = readerEx.LineNumber;
                                column = readerEx.LinePosition;
                            }
                            else if (e is JsonSerializationException serializeEx)
                            {
                                line = serializeEx.LineNumber;
                                column = serializeEx.LinePosition;
                            }

                            throw new Exception(BuildLoadErrorMessage(filePath, contents, line, column, e.Message), e);
                        }
                    }
                    break;

                case Format.Yaml:
                    {
                        try
                        {
                            var builder = new DeserializerBuilder();

                            builder.IgnoreUnmatchedProperties();

                            var yamlDeserializer = builder.Build();

                            result = yamlDeserializer.Deserialize<T>(contents);
                        }
                        catch (YamlException e)
                        {
                            var line = e.Start.Line;
                            var column = e.Start.Column;
                            var innerMessage = e.InnerException != null ? e.InnerException.Message : e.Message;

                            throw new Exception(BuildLoadErrorMessage(filePath, contents, line, column, innerMessage), e);
                        }
                    }
                    break;
            }

            return result;
        }

        // デシリアライズエラー用メッセージ生成(エラー位置の周辺行を含む).
        private static string BuildLoadErrorMessage(string filePath, string contents, int line, int column, string innerMessage)
        {
            var builder = new StringBuilder();

            builder.AppendLine();
            builder.AppendFormat("Failed to deserialize file.").AppendLine();
            builder.AppendFormat("  File    : {0}", filePath).AppendLine();
            builder.AppendFormat("  Position: Line {0}, Column {1}", line, column).AppendLine();
            builder.AppendFormat("  Reason  : {0}", innerMessage).AppendLine();

            // エラー位置周辺の内容を表示.
            if (!string.IsNullOrEmpty(contents) && 0 < line)
            {
                var lines = contents.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

                var startLine = Math.Max(1, line - 3);
                var endLine = Math.Min(lines.Length, line + 3);

                builder.AppendLine();
                builder.AppendLine("--- Context ---");

                for (var i = startLine; i <= endLine; i++)
                {
                    var marker = (i == line) ? ">>" : "  ";

                    builder.AppendFormat("{0} {1,5}: {2}", marker, i, lines[i - 1]).AppendLine();

                    // エラー列を指す矢印を出力.
                    if (i == line && 0 < column)
                    {
                        builder.Append("        ").Append(' ', column - 1).AppendLine("^");
                    }
                }
            }

            return builder.ToString();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DinkCel
{
    internal sealed class CsvDocument
    {
        public IList<string[]> Rows;
        public Encoding Encoding;
        public string NewLine;
        public bool EndsWithNewLine;
        public int DataRows;
        public int DataColumns;
    }

    internal static class CsvFile
    {
        public static IList<string[]> Read(string path, int maxRows, int maxColumns)
        {
            return ReadDocument(path, maxRows, maxColumns).Rows;
        }

        public static CsvDocument ReadDocument(string path, int maxRows, int maxColumns)
        {
            string content;
            Encoding encoding;
            try
            {
                using (var reader = new StreamReader(path,
                    new UTF8Encoding(false, true), true))
                {
                    content = reader.ReadToEnd();
                    encoding = reader.CurrentEncoding;
                }
            }
            catch (DecoderFallbackException)
            {
                // Older Windows CSV exports may use the machine's ANSI code page.
                using (var reader = new StreamReader(path, Encoding.Default, true))
                {
                    content = reader.ReadToEnd();
                    encoding = reader.CurrentEncoding;
                }
            }
            IList<string[]> rows = Parse(new StringReader(content), maxRows, maxColumns);
            int columns = 0;
            foreach (string[] row in rows)
                columns = Math.Max(columns, row.Length);
            return new CsvDocument
            {
                Rows = rows,
                Encoding = encoding,
                NewLine = content.Contains("\r\n") ? "\r\n" :
                    content.Contains("\n") ? "\n" :
                    content.Contains("\r") ? "\r" : Environment.NewLine,
                EndsWithNewLine = content.EndsWith("\n", StringComparison.Ordinal) ||
                    content.EndsWith("\r", StringComparison.Ordinal),
                DataRows = rows.Count,
                DataColumns = columns
            };
        }

        public static void Write(string path, IList<string[]> rows, CsvDocument format)
        {
            string directory = Path.GetDirectoryName(Path.GetFullPath(path));
            string temporary = Path.Combine(directory, "." + Path.GetFileName(path) +
                "." + Guid.NewGuid().ToString("N") + ".tmp");
            Encoding encoding = format == null ? new UTF8Encoding(true) : format.Encoding;
            string newLine = format == null ? Environment.NewLine : format.NewLine;
            bool trailingNewLine = format == null || format.EndsWithNewLine;
            try
            {
                using (var writer = new StreamWriter(temporary, false, encoding))
                {
                    for (int row = 0; row < rows.Count; row++)
                    {
                        for (int column = 0; column < rows[row].Length; column++)
                        {
                            if (column > 0)
                                writer.Write(',');
                            writer.Write(Escape(rows[row][column] ?? ""));
                        }
                        if (row < rows.Count - 1 || trailingNewLine)
                            writer.Write(newLine);
                    }
                }
                if (File.Exists(path))
                    File.Replace(temporary, path, null);
                else
                    File.Move(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        private static string Escape(string value)
        {
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0 &&
                !value.StartsWith(" ", StringComparison.Ordinal) &&
                !value.EndsWith(" ", StringComparison.Ordinal))
                return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        public static IList<string[]> Parse(TextReader reader, int maxRows, int maxColumns)
        {
            var rows = new List<string[]>();
            var fields = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            bool afterQuote = false;
            bool atFieldStart = true;
            bool anyInput = false;

            while (true)
            {
                int code = reader.Read();
                if (code < 0)
                    break;
                anyInput = true;
                char ch = (char)code;

                if (quoted)
                {
                    if (ch == '"')
                    {
                        if (reader.Peek() == '"')
                        {
                            reader.Read();
                            field.Append('"');
                        }
                        else
                        {
                            quoted = false;
                            afterQuote = true;
                        }
                    }
                    else
                        field.Append(ch);
                    continue;
                }

                if (ch == ',' || ch == '\r' || ch == '\n')
                {
                    AddField(fields, field, maxColumns);
                    atFieldStart = true;
                    afterQuote = false;
                    if (ch != ',')
                    {
                        if (ch == '\r' && reader.Peek() == '\n')
                            reader.Read();
                        AddRow(rows, fields, maxRows);
                    }
                    continue;
                }

                if (ch == '"' && atFieldStart)
                {
                    quoted = true;
                    atFieldStart = false;
                    continue;
                }
                if (ch == '"' || afterQuote)
                    throw new InvalidDataException("CSV có dấu ngoặc kép không hợp lệ.");
                field.Append(ch);
                atFieldStart = false;
            }

            if (quoted)
                throw new InvalidDataException("CSV có ô chưa đóng dấu ngoặc kép.");
            if (anyInput && (fields.Count > 0 || field.Length > 0 || !atFieldStart || afterQuote))
            {
                AddField(fields, field, maxColumns);
                AddRow(rows, fields, maxRows);
            }
            return rows;
        }

        private static void AddField(List<string> fields, StringBuilder field, int maxColumns)
        {
            if (fields.Count >= maxColumns)
                throw new InvalidDataException("CSV có quá nhiều cột (tối đa " + maxColumns + ").");
            fields.Add(field.ToString());
            field.Length = 0;
        }

        private static void AddRow(List<string[]> rows, List<string> fields, int maxRows)
        {
            if (rows.Count >= maxRows)
                throw new InvalidDataException("CSV có quá nhiều hàng (tối đa " + maxRows + ").");
            rows.Add(fields.ToArray());
            fields.Clear();
        }
    }
}

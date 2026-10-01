using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DinkCel
{
    internal static class CsvFileTests
    {
        private static int checks;

        private static void Main()
        {
            IList<string[]> rows = CsvFile.Parse(new StringReader(
                "name,note,empty\r\n\"A,B\",\"line 1\r\nline \"\"2\"\"\",\r\n"), 200, 26);
            Equal(2, rows.Count);
            Equal("name", rows[0][0]);
            Equal("A,B", rows[1][0]);
            Equal("line 1\r\nline \"2\"", rows[1][1]);
            Equal("", rows[1][2]);
            Equal(0, CsvFile.Parse(new StringReader(""), 200, 26).Count);
            Equal(1, CsvFile.Parse(new StringReader("\"\""), 200, 26).Count);
            Equal(2, CsvFile.Parse(new StringReader("a,"), 200, 26)[0].Length);
            Fails("a,b,c", 200, 2);
            Fails("a\nb", 1, 26);
            Fails("\"unfinished", 200, 26);

            string temp = Path.Combine(Path.GetTempPath(), "DinkCelCsvTest_" +
                Guid.NewGuid().ToString("N") + ".csv");
            try
            {
                File.WriteAllText(temp, "tên,giá\nBút,12", new UTF8Encoding(true));
                rows = CsvFile.Read(temp, 200, 26);
                Equal("tên", rows[0][0]);
                Equal("Bút", rows[1][0]);
            }
            finally
            {
                File.Delete(temp);
            }
            temp = Path.Combine(Path.GetTempPath(), "DinkCelCsvTest_" +
                Guid.NewGuid().ToString("N") + ".csv");
            try
            {
                File.WriteAllText(temp, "a,b\n1,\"x,y\"", new UTF8Encoding(false));
                CsvDocument document = CsvFile.ReadDocument(temp, 200, 26);
                document.Rows[1][1] = "line 1\n\"line 2\"";
                CsvFile.Write(temp, document.Rows, document);
                byte[] bytes = File.ReadAllBytes(temp);
                Equal((byte)'a', bytes[0]);
                string raw = File.ReadAllText(temp, new UTF8Encoding(false, true));
                Equal("a,b\n1,\"line 1\n\"\"line 2\"\"\"", raw);
                Equal("line 1\n\"line 2\"", CsvFile.Read(temp, 200, 26)[1][1]);
            }
            finally
            {
                File.Delete(temp);
            }
            temp = Path.Combine(Path.GetTempPath(), "DinkCelCsvTest_" +
                Guid.NewGuid().ToString("N") + ".csv");
            try
            {
                File.WriteAllText(temp, "tên,giá\r\nBút,12\r\n", Encoding.Unicode);
                CsvDocument document = CsvFile.ReadDocument(temp, 200, 26);
                document.Rows[1][1] = "15";
                CsvFile.Write(temp, document.Rows, document);
                byte[] bytes = File.ReadAllBytes(temp);
                Equal((byte)0xFF, bytes[0]);
                Equal((byte)0xFE, bytes[1]);
                Equal("15", CsvFile.Read(temp, 200, 26)[1][1]);
            }
            finally
            {
                File.Delete(temp);
            }
            Console.WriteLine("CSV: " + checks + " checks passed.");
        }

        private static void Equal(object expected, object actual)
        {
            checks++;
            if (!object.Equals(expected, actual))
                throw new Exception("Expected " + expected + ", got " + actual);
        }

        private static void Fails(string csv, int maxRows, int maxColumns)
        {
            checks++;
            try
            {
                CsvFile.Parse(new StringReader(csv), maxRows, maxColumns);
            }
            catch (InvalidDataException)
            {
                return;
            }
            throw new Exception("Expected invalid CSV or size error.");
        }
    }
}

using System;
using System.Drawing;
using System.IO;
using PdfSharp.Pdf.IO;

namespace DinkCel
{
    internal static class InterchangeTests
    {
        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual);
        }

        private static void CheckPdf(string path)
        {
            using (var reader = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                Equal(2, reader.PageCount);
        }

        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            EmbeddedDependencies.Install();
            var book = new WorkbookSnapshot();
            book.Sheets[0].Name = "Data";
            book.Sheets[0].Cells[0] = new CellSnapshot { Text = "Nhãn" };
            book.Sheets[0].Cells[26] = new CellSnapshot { Text = "12.5", NumberFormat = "0.00" };
            book.Sheets[0].Cells[27] = new CellSnapshot { Text = "=A2*2" };
            book.Sheets[0].Merges.Add(new Rectangle(0, 0, 2, 1));
            var second = new SheetSnapshot { Name = "More" };
            second.Cells[0] = new CellSnapshot { Text = "=Data!A2" };
            book.Sheets.Add(second);
            book.NamedRanges.Add(new NamedRange { Name = "Revenue", Sheet = "Data",
                Range = new Rectangle(0, 1, 1, 1) });
            string basePath = Path.Combine(Path.GetTempPath(), "DinkCel_v03_" + Guid.NewGuid().ToString("N"));
            try
            {
                string xls = basePath + ".xls", ods = basePath + ".ods", pdf = basePath + ".pdf";
                XlsFile.Write(xls, book, 200, 26);
                var legacy = XlsFile.Read(xls, 200, 26);
                Equal(2, legacy.Sheets.Count);
                Equal("Nhãn", legacy.Sheets[0].Cells[0].Text);
                Equal("=A2*2", legacy.Sheets[0].Cells[27].Text);
                Equal("0.00", legacy.Sheets[0].Cells[26].NumberFormat);
                Equal(1, legacy.Sheets[0].Merges.Count);
                Equal("=Data!A2", legacy.Sheets[1].Cells[0].Text);
                Equal("Revenue", legacy.NamedRanges[0].Name);
                var outsideXls = XlsFile.Read(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "fixtures", "thirdparty.xls"), 200, 26);
                Equal(2, outsideXls.Sheets.Count);
                Equal("Alpha", outsideXls.Sheets[0].Cells[26].Text);
                Equal("=B2*2", outsideXls.Sheets[0].Cells[28].Text);
                OdsFile.Write(ods, book, 200, 26);
                var open = OdsFile.Read(ods, 200, 26);
                Equal(2, open.Sheets.Count);
                Equal("Nhãn", open.Sheets[0].Cells[0].Text);
                Equal("=A2*2", open.Sheets[0].Cells[27].Text);
                Equal(1, open.Sheets[0].Merges.Count);
                Equal("=Data!A2", open.Sheets[1].Cells[0].Text);
                Equal("Revenue", open.NamedRanges[0].Name);
                var outsideOds = OdsFile.Read(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "fixtures", "thirdparty.ods"), 200, 26);
                Equal("OpenData", outsideOds.Sheets[0].Name);
                Equal("Alpha", outsideOds.Sheets[0].Cells[26].Text);
                Equal("=B2*2", outsideOds.Sheets[0].Cells[28].Text);
                var repeatedOds = OdsFile.Read(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "fixtures", "repeated.ods"), 200, 26);
                Equal("=B1", repeatedOds.Sheets[0].Cells[26].Text);
                Equal("=A2", repeatedOds.Sheets[0].Cells[27].Text);
                Equal("=B2", repeatedOds.Sheets[0].Cells[28].Text);
                Equal("=B2", repeatedOds.Sheets[0].Cells[52].Text);
                Equal("=A3", repeatedOds.Sheets[0].Cells[53].Text);
                Equal("=B3", repeatedOds.Sheets[0].Cells[54].Text);
                PdfFile.Write(pdf, book, 200, 26);
                CheckPdf(pdf);
                Console.WriteLine("XLS and ODS: multi-sheet round trips passed.");
                Console.WriteLine("PDF: workbook export passed.");
                if (Environment.GetEnvironmentVariable("DINKCEL_KEEP_INTERCHANGE") == "1")
                    Console.WriteLine(basePath);
            }
            finally
            {
                if (Environment.GetEnvironmentVariable("DINKCEL_KEEP_INTERCHANGE") != "1")
                {
                    if (File.Exists(basePath + ".xls")) File.Delete(basePath + ".xls");
                    if (File.Exists(basePath + ".ods")) File.Delete(basePath + ".ods");
                    if (File.Exists(basePath + ".pdf")) File.Delete(basePath + ".pdf");
                }
            }
        }
    }
}

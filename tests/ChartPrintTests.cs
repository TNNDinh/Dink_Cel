using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using PdfSharp.Pdf.IO;

namespace DinkCel
{
    internal static class ChartPrintTests
    {
        private static void Check(bool condition, string message)
        { if (!condition) throw new Exception(message); }

        private static CellSnapshot Cell(string value)
        { return new CellSnapshot { Text = value }; }

        [STAThread]
        private static void Main()
        {
            try { EmbeddedDependencies.Install(); Run(); Console.WriteLine("v0.9 charts and printing passed."); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            var book = new WorkbookSnapshot();
            SheetSnapshot sheet = book.Sheets[0];
            sheet.Name = "Sales";
            string[,] data = { { "Month", "Revenue", "Cost" },
                { "1", "10", "4" }, { "2", "20", "8" }, { "3", "15", "6" } };
            for (int r = 0; r < 4; r++)
                for (int c = 0; c < 3; c++) sheet.Cells[r * 26 + c] = Cell(data[r, c]);
            string[] kinds = { "Column", "Line", "Pie", "Bar", "Area", "Scatter",
                "Stacked", "100% Stacked", "Combo" };
            foreach (string kind in kinds)
            {
                var chart = new ChartDefinition { Kind = kind, Title = kind + " chart",
                    Range = new Rectangle(0, 0, 3, 4), Placement = new Rectangle(4, 2, 8, 14),
                    AxisTitleX = "Month", AxisTitleY = "Amount", Legend = false,
                    DataLabels = true, Gridlines = false };
                chart.SeriesColumns.Add(1); chart.SeriesColumns.Add(2);
                chart.SeriesNames[1] = "Sales";
                chart.SeriesColors[1] = Color.Crimson;
                using (Chart rendered = ChartRendering.Build(chart, (r, c) => data[r, c],
                    Color.White, Color.Black, Color.Crimson))
                using (var image = new MemoryStream())
                {
                    rendered.Size = new Size(640, 400);
                    rendered.SaveImage(image, ChartImageFormat.Png);
                    Check(image.Length > 1000, "render " + kind);
                    Check(rendered.Series.Count == (kind == "Pie" ? 1 : 2), "series " + kind);
                }
                sheet.Charts.Add(chart);
            }
            sheet.Print.Paper = "Letter";
            sheet.Print.Landscape = false;
            sheet.Print.MarginLeft = 75;
            sheet.Print.PrintArea = new Rectangle(0, 0, 3, 4);
            sheet.Print.TitleRows = 1;
            sheet.Print.FitToOnePage = true;
            sheet.Print.Header = "&F — &D";
            sheet.Print.Footer = "Page &P of &N";
            sheet.Print.Gridlines = false;
            sheet.Print.PageBreakRows.Add(2);
            var plan = PrintLayout.Plan(sheet, 200, 26, Rectangle.Empty, false);
            Check(plan.Count == 1 && plan[0].TitleRows == 1 &&
                PrintLayout.HeaderFooter(sheet.Print.Footer, plan[0]) == "Page 1 of 1",
                "fit page and title rows");
            sheet.Print.FitToOnePage = false;
            plan = PrintLayout.Plan(sheet, 200, 26, Rectangle.Empty, false);
            Check(plan.Count == 2, "manual page break");
            string xlsx = Path.Combine(Path.GetTempPath(), "DinkCel_chart_" +
                Guid.NewGuid().ToString("N") + ".xlsx");
            string pdf = Path.ChangeExtension(xlsx, ".pdf");
            string native = Path.ChangeExtension(xlsx, ".dinkcel");
            try
            {
                XlsxFile.Write(xlsx, book, 200, 26);
                SheetSnapshot loaded = XlsxFile.Read(xlsx, 200, 26).Sheets[0];
                Check(loaded.Charts.Count == kinds.Length, "XLSX chart count");
                for (int i = 0; i < kinds.Length; i++)
                {
                    ChartDefinition chart = loaded.Charts[i];
                    Check(chart.Kind == kinds[i] && (kinds[i] == "Pie" ||
                        chart.AxisTitleX == "Month" && chart.AxisTitleY == "Amount") &&
                        !chart.Legend && chart.DataLabels &&
                        !chart.Gridlines && chart.SeriesColors[1].ToArgb() == Color.Crimson.ToArgb() &&
                        chart.Placement.Width == 8, "XLSX " + kinds[i] + " " + chart.Kind + " " +
                        chart.SeriesColors.ContainsKey(1) + " " + chart.Placement.Width);
                }
                Check(loaded.Print.Paper == "Letter" && !loaded.Print.Landscape &&
                    loaded.Print.PrintArea == sheet.Print.PrintArea && loaded.Print.TitleRows == 1 &&
                    loaded.Print.PageBreakRows.Contains(2), "XLSX page setup");
                sheet.Charts.RemoveRange(1, sheet.Charts.Count - 1);
                sheet.Print.PageBreakRows.Clear(); sheet.Print.FitToOnePage = true;
                PdfFile.Write(pdf, book, 200, 26);
                using (var reader = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    Check(reader.PageCount == 2 && reader.Pages[0].Width < reader.Pages[0].Height,
                        "PDF chart page and portrait");
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    ((System.Collections.Generic.List<ChartDefinition>)typeof(SpreadsheetForm)
                        .GetField("charts", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form))
                        .Add(sheet.Charts[0].Copy());
                    typeof(SpreadsheetForm).GetField("printSettings", BindingFlags.Instance |
                        BindingFlags.NonPublic).SetValue(form, sheet.Print.Copy());
                    Check((bool)typeof(SpreadsheetForm).GetMethod("WriteWorkbook",
                        BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { native }),
                        "native write");
                    form.Close();
                }
                using (var reopened = new SpreadsheetForm(native))
                {
                    reopened.Show(); Application.DoEvents();
                    var print = (PrintSettings)typeof(SpreadsheetForm).GetField("printSettings",
                        BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    var charts = (System.Collections.Generic.List<ChartDefinition>)typeof(SpreadsheetForm)
                        .GetField("charts", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    var overlays = (System.Collections.IDictionary)typeof(SpreadsheetForm)
                        .GetField("chartOverlays", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    Check(print.Paper == "Letter" && print.PrintArea == sheet.Print.PrintArea &&
                        charts.Count == 1 && charts[0].AxisTitleY == "Amount" && overlays.Count == 1,
                        "native page and chart overlay");
                    reopened.Close();
                }
            }
            finally
            {
                if (Environment.GetEnvironmentVariable("DINK_KEEP_CHART_FIXTURE") == "1")
                    Console.WriteLine("chart_fixture=" + xlsx);
                else
                {
                    if (File.Exists(xlsx)) File.Delete(xlsx);
                    if (File.Exists(pdf)) File.Delete(pdf);
                    if (File.Exists(native)) File.Delete(native);
                }
            }
        }
    }
}

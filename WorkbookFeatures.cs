using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed class TableDefinition
    {
        public string Name = "Table1";
        public Rectangle Range;
        public string Style = "TableStyleMedium2";
        public bool HeaderRow = true;
        public bool TotalRow;
        public bool BandedRows = true;
        public bool BandedColumns;
        public bool AutoExpand = true;
        public bool Filter = true;
        public readonly Dictionary<int, string> CalculatedColumns = new Dictionary<int, string>();
        public readonly List<FilterCriterion> Filters = new List<FilterCriterion>();

        public TableDefinition Copy()
        {
            var result = new TableDefinition { Name = Name, Range = Range, Style = Style,
                HeaderRow = HeaderRow, TotalRow = TotalRow, BandedRows = BandedRows,
                BandedColumns = BandedColumns, AutoExpand = AutoExpand, Filter = Filter };
            foreach (var entry in CalculatedColumns) result.CalculatedColumns[entry.Key] = entry.Value;
            result.Filters.AddRange(Filters.Select(f => new FilterCriterion { Column = f.Column,
                Kind = f.Kind, Operator = f.Operator, Value1 = f.Value1, Value2 = f.Value2 }));
            return result;
        }
    }

    internal sealed class PivotAxisField
    {
        public int Column;
        public string DateGroup = "None";
    }

    internal sealed class PivotValueField
    {
        public int Column;
        public string Aggregate = "Sum";
    }

    internal sealed class PivotFilterField
    {
        public int Column;
        public string Value = "";
        public string Operator = "Equals";
        public string Value2 = "";
    }

    internal sealed class PivotCalculatedField
    {
        public string Name = "";
        public string Formula = "";
    }

    internal sealed class ChartDefinition
    {
        public string Title = "Chart";
        public string Kind = "Column";
        public string PivotSource = "";
        public Rectangle Range;
        public Rectangle Placement;
        public string AxisTitleX = "";
        public string AxisTitleY = "";
        public bool Legend = true;
        public bool DataLabels;
        public bool Gridlines = true;
        public readonly List<int> SeriesColumns = new List<int>();
        public readonly Dictionary<int, string> SeriesNames = new Dictionary<int, string>();
        public readonly Dictionary<int, Color> SeriesColors = new Dictionary<int, Color>();

        public ChartDefinition Copy()
        {
            var copy = new ChartDefinition { Title = Title, Kind = Kind, Range = Range,
                Placement = Placement, PivotSource = PivotSource, AxisTitleX = AxisTitleX, AxisTitleY = AxisTitleY,
                Legend = Legend, DataLabels = DataLabels, Gridlines = Gridlines };
            copy.SeriesColumns.AddRange(SeriesColumns);
            foreach (var entry in SeriesNames) copy.SeriesNames[entry.Key] = entry.Value;
            foreach (var entry in SeriesColors) copy.SeriesColors[entry.Key] = entry.Value;
            return copy;
        }
    }

    internal sealed class SheetObject
    {
        public string Kind = "Rectangle";
        public Rectangle Placement;
        public string Text = "";
        public string ImageBase64 = "";
        public Color Fill = Color.FromArgb(73, 143, 232);
        public SheetObject Copy()
        {
            return new SheetObject { Kind = Kind, Placement = Placement, Text = Text,
                ImageBase64 = ImageBase64, Fill = Fill };
        }
    }

    internal sealed class PrintSettings
    {
        public Rectangle PrintArea;
        public bool Landscape = true;
        public string Paper = "A4";
        public int MarginLeft = 50, MarginRight = 50, MarginTop = 60, MarginBottom = 60;
        public int Scale = 100;
        public bool FitToOnePage;
        public int TitleRows;
        public string Header = "";
        public string Footer = "DinkCel · Trang &P / &N";
        public bool Gridlines = true;
        public bool PrintCharts = true;
        public readonly List<int> PageBreakRows = new List<int>();

        public PrintSettings Copy()
        {
            var copy = new PrintSettings { PrintArea = PrintArea, Landscape = Landscape,
                Paper = Paper, MarginLeft = MarginLeft, MarginRight = MarginRight,
                MarginTop = MarginTop, MarginBottom = MarginBottom, Scale = Scale,
                FitToOnePage = FitToOnePage, TitleRows = TitleRows, Header = Header,
                Footer = Footer, Gridlines = Gridlines, PrintCharts = PrintCharts };
            copy.PageBreakRows.AddRange(PageBreakRows);
            return copy;
        }
    }

    internal sealed class ValidationRule
    {
        public Rectangle Range;
        public readonly List<string> Choices = new List<string>();
        public string Kind = "List";
        public string Operator = "Between";
        public string Value1 = "";
        public string Value2 = "";
        public bool AllowBlank = true;
        public string InputTitle = "";
        public string InputMessage = "";
        public string ErrorTitle = "";
        public string ErrorMessage = "";
        public string ErrorStyle = "Stop";
    }

    internal sealed class NamedRange
    {
        public string Name;
        public string Sheet;
        public Rectangle Range;
        public string Formula = "";
        public string ScopeSheet = "";
    }

    internal sealed class PivotDefinition
    {
        public string SourceSheet;
        public Rectangle SourceRange;
        public int GroupColumn;
        public int ValueColumn;
        public string Aggregate = "Sum";
        public string TargetSheet;
        public string SourceTable = "";
        public int LastOutputRows;
        public int LastOutputColumns;
        public readonly List<PivotAxisField> Rows = new List<PivotAxisField>();
        public readonly List<PivotAxisField> Columns = new List<PivotAxisField>();
        public readonly List<PivotValueField> Values = new List<PivotValueField>();
        public readonly List<PivotFilterField> Filters = new List<PivotFilterField>();
        public readonly List<PivotCalculatedField> CalculatedFields = new List<PivotCalculatedField>();
        public readonly HashSet<string> Collapsed = new HashSet<string>();
        public bool GrandTotal = true;
        public bool Subtotal = true;
        public bool SortDescending;
        public bool SortByValue;
    }

    internal sealed partial class SpreadsheetForm
    {
        private static void SetRange(XElement element, Rectangle range)
        {
            element.SetAttributeValue("row", range.Y);
            element.SetAttributeValue("column", range.X);
            element.SetAttributeValue("width", range.Width);
            element.SetAttributeValue("height", range.Height);
        }

        private static Rectangle GetRange(XElement element)
        {
            return new Rectangle((int)element.Attribute("column"), (int)element.Attribute("row"),
                (int)element.Attribute("width"), (int)element.Attribute("height"));
        }

        private static void SerializeSheetMetadata(XElement root, SheetSnapshot sheet)
        {
            foreach (SheetObject item in sheet.Objects)
            {
                var element = new XElement("object", new XAttribute("kind", item.Kind),
                    new XAttribute("text", item.Text),
                    new XAttribute("fill", ColorTranslator.ToHtml(item.Fill)),
                    new XAttribute("image", item.ImageBase64));
                SetRange(element, item.Placement);
                root.Add(element);
            }
            foreach (TableDefinition table in sheet.Tables)
            {
                var element = new XElement("table", new XAttribute("name", table.Name),
                    new XAttribute("style", table.Style), new XAttribute("headerRow", table.HeaderRow),
                    new XAttribute("totalRow", table.TotalRow), new XAttribute("bandedRows", table.BandedRows),
                    new XAttribute("bandedColumns", table.BandedColumns), new XAttribute("autoExpand", table.AutoExpand),
                    new XAttribute("filter", table.Filter));
                SetRange(element, table.Range);
                foreach (var entry in table.CalculatedColumns)
                    element.Add(new XElement("calculatedColumn", new XAttribute("column", entry.Key), entry.Value));
                foreach (FilterCriterion criterion in table.Filters)
                    element.Add(new XElement("tableFilter", new XAttribute("column", criterion.Column),
                        new XAttribute("kind", criterion.Kind), new XAttribute("operator", criterion.Operator),
                        new XAttribute("value1", criterion.Value1), new XAttribute("value2", criterion.Value2)));
                root.Add(element);
            }
            foreach (ChartDefinition chart in sheet.Charts)
            {
                var element = new XElement("chart", new XAttribute("title", chart.Title),
                    new XAttribute("kind", chart.Kind), new XAttribute("axisX", chart.AxisTitleX),
                    new XAttribute("pivotSource", chart.PivotSource ?? ""),
                    new XAttribute("axisY", chart.AxisTitleY), new XAttribute("legend", chart.Legend),
                    new XAttribute("labels", chart.DataLabels), new XAttribute("gridlines", chart.Gridlines));
                SetRange(element, chart.Range);
                if (!chart.Placement.IsEmpty)
                    element.Add(new XElement("placement", new XAttribute("column", chart.Placement.X),
                        new XAttribute("row", chart.Placement.Y), new XAttribute("width", chart.Placement.Width),
                        new XAttribute("height", chart.Placement.Height)));
                foreach (int column in chart.SeriesColumns)
                {
                    string name;
                    Color color;
                    var series = new XElement("series", new XAttribute("column", column));
                    if (chart.SeriesNames.TryGetValue(column, out name)) series.SetAttributeValue("name", name);
                    if (chart.SeriesColors.TryGetValue(column, out color))
                        series.SetAttributeValue("color", ColorTranslator.ToHtml(color));
                    element.Add(series);
                }
                root.Add(element);
            }
            PrintSettings print = sheet.Print;
            var printElement = new XElement("print", new XAttribute("landscape", print.Landscape),
                new XAttribute("paper", print.Paper), new XAttribute("left", print.MarginLeft),
                new XAttribute("right", print.MarginRight), new XAttribute("top", print.MarginTop),
                new XAttribute("bottom", print.MarginBottom), new XAttribute("scale", print.Scale),
                new XAttribute("fit", print.FitToOnePage), new XAttribute("titleRows", print.TitleRows),
                new XAttribute("header", print.Header), new XAttribute("footer", print.Footer),
                new XAttribute("gridlines", print.Gridlines), new XAttribute("charts", print.PrintCharts));
            if (!print.PrintArea.IsEmpty)
            {
                var area = new XElement("area"); SetRange(area, print.PrintArea); printElement.Add(area);
            }
            foreach (int row in print.PageBreakRows)
                printElement.Add(new XElement("break", new XAttribute("row", row)));
            root.Add(printElement);
            foreach (ValidationRule rule in sheet.Validations)
            {
                var element = new XElement("validation", new XAttribute("kind", rule.Kind),
                    new XAttribute("operator", rule.Operator), new XAttribute("value1", rule.Value1),
                    new XAttribute("value2", rule.Value2), new XAttribute("allowBlank", rule.AllowBlank),
                    new XAttribute("inputTitle", rule.InputTitle), new XAttribute("inputMessage", rule.InputMessage),
                    new XAttribute("errorTitle", rule.ErrorTitle), new XAttribute("errorMessage", rule.ErrorMessage),
                    new XAttribute("errorStyle", rule.ErrorStyle));
                SetRange(element, rule.Range);
                foreach (string choice in rule.Choices) element.Add(new XElement("choice", choice));
                root.Add(element);
            }
        }

        private static void ReadSheetMetadata(XElement root, SheetSnapshot sheet)
        {
            foreach (XElement element in root.Elements("object"))
                sheet.Objects.Add(new SheetObject { Kind = (string)element.Attribute("kind") ?? "Rectangle",
                    Placement = GetRange(element), Text = (string)element.Attribute("text") ?? "",
                    ImageBase64 = (string)element.Attribute("image") ?? "",
                    Fill = ColorTranslator.FromHtml((string)element.Attribute("fill") ?? "#498FE8") });
            foreach (XElement element in root.Elements("table"))
            {
                var table = new TableDefinition { Name = (string)element.Attribute("name") ?? "Table",
                    Range = GetRange(element), Style = (string)element.Attribute("style") ?? "TableStyleMedium2",
                    HeaderRow = (bool?)element.Attribute("headerRow") ?? true,
                    TotalRow = (bool?)element.Attribute("totalRow") ?? false,
                    BandedRows = (bool?)element.Attribute("bandedRows") ?? true,
                    BandedColumns = (bool?)element.Attribute("bandedColumns") ?? false,
                    AutoExpand = (bool?)element.Attribute("autoExpand") ?? true,
                    Filter = (bool?)element.Attribute("filter") ?? true };
                foreach (XElement calc in element.Elements("calculatedColumn"))
                    table.CalculatedColumns[(int)calc.Attribute("column")] = calc.Value;
                foreach (XElement filter in element.Elements("tableFilter"))
                    table.Filters.Add(new FilterCriterion { Column = (int)filter.Attribute("column"),
                        Kind = (string)filter.Attribute("kind") ?? "Text",
                        Operator = (string)filter.Attribute("operator") ?? "Contains",
                        Value1 = (string)filter.Attribute("value1") ?? "",
                        Value2 = (string)filter.Attribute("value2") ?? "" });
                sheet.Tables.Add(table);
            }
            foreach (XElement element in root.Elements("chart"))
            {
                var chart = new ChartDefinition { Title = (string)element.Attribute("title") ?? "Chart",
                    Kind = (string)element.Attribute("kind") ?? "Column", Range = GetRange(element),
                    PivotSource = (string)element.Attribute("pivotSource") ?? "",
                    AxisTitleX = (string)element.Attribute("axisX") ?? "",
                    AxisTitleY = (string)element.Attribute("axisY") ?? "",
                    Legend = (bool?)element.Attribute("legend") ?? true,
                    DataLabels = (bool?)element.Attribute("labels") ?? false,
                    Gridlines = (bool?)element.Attribute("gridlines") ?? true };
                XElement placement = element.Element("placement");
                if (placement != null) chart.Placement = GetRange(placement);
                foreach (XElement series in element.Elements("series"))
                {
                    int column = (int)series.Attribute("column");
                    chart.SeriesColumns.Add(column);
                    string name = (string)series.Attribute("name");
                    if (name != null) chart.SeriesNames[column] = name;
                    string color = (string)series.Attribute("color");
                    if (!string.IsNullOrEmpty(color))
                        try { chart.SeriesColors[column] = ColorTranslator.FromHtml(color); }
                        catch (ArgumentException) { }
                }
                sheet.Charts.Add(chart);
            }
            XElement printElement = root.Element("print");
            if (printElement != null)
            {
                var print = new PrintSettings { Landscape = (bool?)printElement.Attribute("landscape") ?? true,
                    Paper = (string)printElement.Attribute("paper") ?? "A4",
                    MarginLeft = (int?)printElement.Attribute("left") ?? 50,
                    MarginRight = (int?)printElement.Attribute("right") ?? 50,
                    MarginTop = (int?)printElement.Attribute("top") ?? 60,
                    MarginBottom = (int?)printElement.Attribute("bottom") ?? 60,
                    Scale = (int?)printElement.Attribute("scale") ?? 100,
                    FitToOnePage = (bool?)printElement.Attribute("fit") ?? false,
                    TitleRows = (int?)printElement.Attribute("titleRows") ?? 0,
                    Header = (string)printElement.Attribute("header") ?? "",
                    Footer = (string)printElement.Attribute("footer") ?? "",
                    Gridlines = (bool?)printElement.Attribute("gridlines") ?? true,
                    PrintCharts = (bool?)printElement.Attribute("charts") ?? true };
                XElement area = printElement.Element("area");
                if (area != null) print.PrintArea = GetRange(area);
                foreach (XElement entry in printElement.Elements("break"))
                    print.PageBreakRows.Add((int)entry.Attribute("row"));
                sheet.Print = print;
            }
            foreach (XElement element in root.Elements("validation"))
            {
                var rule = new ValidationRule { Range = GetRange(element),
                    Kind = (string)element.Attribute("kind") ?? "List",
                    Operator = (string)element.Attribute("operator") ?? "Between",
                    Value1 = (string)element.Attribute("value1") ?? "",
                    Value2 = (string)element.Attribute("value2") ?? "",
                    AllowBlank = (bool?)element.Attribute("allowBlank") ?? true,
                    InputTitle = (string)element.Attribute("inputTitle") ?? "",
                    InputMessage = (string)element.Attribute("inputMessage") ?? "",
                    ErrorTitle = (string)element.Attribute("errorTitle") ?? "",
                    ErrorMessage = (string)element.Attribute("errorMessage") ?? "",
                    ErrorStyle = (string)element.Attribute("errorStyle") ?? "Stop" };
                foreach (XElement choice in element.Elements("choice")) rule.Choices.Add(choice.Value);
                sheet.Validations.Add(rule);
            }
        }

        private static void SerializeWorkbookMetadata(XElement root,
            IEnumerable<NamedRange> names, IEnumerable<PivotDefinition> pivotDefinitions)
        {
            foreach (NamedRange named in names)
            {
                var element = new XElement("namedRange", new XAttribute("name", named.Name),
                    new XAttribute("sheet", named.Sheet ?? ""),
                    new XAttribute("scopeSheet", named.ScopeSheet ?? ""),
                    new XAttribute("formula", named.Formula ?? ""));
                SetRange(element, named.Range);
                root.Add(element);
            }
            foreach (PivotDefinition pivot in pivotDefinitions)
            {
                var element = new XElement("pivot", new XAttribute("sourceSheet", pivot.SourceSheet),
                    new XAttribute("groupColumn", pivot.GroupColumn),
                    new XAttribute("valueColumn", pivot.ValueColumn),
                    new XAttribute("aggregate", pivot.Aggregate),
                    new XAttribute("targetSheet", pivot.TargetSheet),
                    new XAttribute("sourceTable", pivot.SourceTable),
                    new XAttribute("outputRows", pivot.LastOutputRows),
                    new XAttribute("outputColumns", pivot.LastOutputColumns),
                    new XAttribute("grandTotal", pivot.GrandTotal), new XAttribute("subtotal", pivot.Subtotal),
                    new XAttribute("sortDescending", pivot.SortDescending), new XAttribute("sortByValue", pivot.SortByValue));
                SetRange(element, pivot.SourceRange);
                foreach (PivotAxisField field in pivot.Rows)
                    element.Add(new XElement("rowField", new XAttribute("column", field.Column),
                        new XAttribute("dateGroup", field.DateGroup)));
                foreach (PivotAxisField field in pivot.Columns)
                    element.Add(new XElement("columnField", new XAttribute("column", field.Column),
                        new XAttribute("dateGroup", field.DateGroup)));
                foreach (PivotValueField field in pivot.Values)
                    element.Add(new XElement("valueField", new XAttribute("column", field.Column),
                        new XAttribute("aggregate", field.Aggregate)));
                foreach (PivotFilterField field in pivot.Filters)
                    element.Add(new XElement("filterField", new XAttribute("column", field.Column),
                        new XAttribute("value", field.Value),
                        new XAttribute("operator", field.Operator), new XAttribute("value2", field.Value2)));
                foreach (PivotCalculatedField field in pivot.CalculatedFields)
                    element.Add(new XElement("calculatedField", new XAttribute("name", field.Name),
                        new XAttribute("formula", field.Formula)));
                foreach (string group in pivot.Collapsed) element.Add(new XElement("collapsed", group));
                root.Add(element);
            }
        }

        private static void ReadWorkbookMetadata(XElement root, WorkbookSnapshot workbook)
        {
            foreach (XElement element in root.Elements("namedRange"))
                workbook.NamedRanges.Add(new NamedRange { Name = (string)element.Attribute("name"),
                    Sheet = (string)element.Attribute("sheet") ?? "",
                    ScopeSheet = (string)element.Attribute("scopeSheet") ?? "",
                    Formula = (string)element.Attribute("formula") ?? "", Range = GetRange(element) });
            foreach (XElement element in root.Elements("pivot"))
            {
                var pivot = new PivotDefinition { SourceSheet = (string)element.Attribute("sourceSheet"),
                    SourceRange = GetRange(element), GroupColumn = (int)element.Attribute("groupColumn"),
                    ValueColumn = (int)element.Attribute("valueColumn"),
                    Aggregate = (string)element.Attribute("aggregate") ?? "Sum",
                    TargetSheet = (string)element.Attribute("targetSheet"),
                    SourceTable = (string)element.Attribute("sourceTable") ?? "",
                    LastOutputRows = (int?)element.Attribute("outputRows") ?? 0,
                    LastOutputColumns = (int?)element.Attribute("outputColumns") ?? 0,
                    GrandTotal = (bool?)element.Attribute("grandTotal") ?? true,
                    Subtotal = (bool?)element.Attribute("subtotal") ?? true,
                    SortDescending = (bool?)element.Attribute("sortDescending") ?? false,
                    SortByValue = (bool?)element.Attribute("sortByValue") ?? false };
                foreach (XElement field in element.Elements("rowField")) pivot.Rows.Add(new PivotAxisField
                { Column = (int)field.Attribute("column"), DateGroup = (string)field.Attribute("dateGroup") ?? "None" });
                foreach (XElement field in element.Elements("columnField")) pivot.Columns.Add(new PivotAxisField
                { Column = (int)field.Attribute("column"), DateGroup = (string)field.Attribute("dateGroup") ?? "None" });
                foreach (XElement field in element.Elements("valueField")) pivot.Values.Add(new PivotValueField
                { Column = (int)field.Attribute("column"), Aggregate = (string)field.Attribute("aggregate") ?? "Sum" });
                foreach (XElement field in element.Elements("filterField")) pivot.Filters.Add(new PivotFilterField
                { Column = (int)field.Attribute("column"), Value = (string)field.Attribute("value") ?? "",
                    Operator = (string)field.Attribute("operator") ?? "Equals",
                    Value2 = (string)field.Attribute("value2") ?? "" });
                foreach (XElement field in element.Elements("calculatedField"))
                    pivot.CalculatedFields.Add(new PivotCalculatedField
                    { Name = (string)field.Attribute("name") ?? "",
                        Formula = (string)field.Attribute("formula") ?? "" });
                foreach (XElement collapsed in element.Elements("collapsed")) pivot.Collapsed.Add(collapsed.Value);
                workbook.Pivots.Add(pivot);
            }
        }
    }
}

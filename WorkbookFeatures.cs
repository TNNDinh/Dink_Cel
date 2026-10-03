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
    }

    internal sealed class ChartDefinition
    {
        public string Title = "Chart";
        public string Kind = "Column";
        public Rectangle Range;
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
                    new XAttribute("kind", chart.Kind));
                SetRange(element, chart.Range);
                root.Add(element);
            }
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
                sheet.Charts.Add(new ChartDefinition { Title = (string)element.Attribute("title") ?? "Chart",
                    Kind = (string)element.Attribute("kind") ?? "Column", Range = GetRange(element) });
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
                    new XAttribute("sheet", named.Sheet));
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
                        new XAttribute("value", field.Value)));
                foreach (string group in pivot.Collapsed) element.Add(new XElement("collapsed", group));
                root.Add(element);
            }
        }

        private static void ReadWorkbookMetadata(XElement root, WorkbookSnapshot workbook)
        {
            foreach (XElement element in root.Elements("namedRange"))
                workbook.NamedRanges.Add(new NamedRange { Name = (string)element.Attribute("name"),
                    Sheet = (string)element.Attribute("sheet"), Range = GetRange(element) });
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
                { Column = (int)field.Attribute("column"), Value = (string)field.Attribute("value") ?? "" });
                foreach (XElement collapsed in element.Elements("collapsed")) pivot.Collapsed.Add(collapsed.Value);
                workbook.Pivots.Add(pivot);
            }
        }
    }
}

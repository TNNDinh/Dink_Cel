using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed class TableDefinition
    {
        public string Name = "Table1";
        public Rectangle Range;
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
                var element = new XElement("table", new XAttribute("name", table.Name));
                SetRange(element, table.Range);
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
                var element = new XElement("validation");
                SetRange(element, rule.Range);
                foreach (string choice in rule.Choices) element.Add(new XElement("choice", choice));
                root.Add(element);
            }
        }

        private static void ReadSheetMetadata(XElement root, SheetSnapshot sheet)
        {
            foreach (XElement element in root.Elements("table"))
                sheet.Tables.Add(new TableDefinition { Name = (string)element.Attribute("name") ?? "Table",
                    Range = GetRange(element) });
            foreach (XElement element in root.Elements("chart"))
                sheet.Charts.Add(new ChartDefinition { Title = (string)element.Attribute("title") ?? "Chart",
                    Kind = (string)element.Attribute("kind") ?? "Column", Range = GetRange(element) });
            foreach (XElement element in root.Elements("validation"))
            {
                var rule = new ValidationRule { Range = GetRange(element) };
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
                    new XAttribute("targetSheet", pivot.TargetSheet));
                SetRange(element, pivot.SourceRange);
                root.Add(element);
            }
        }

        private static void ReadWorkbookMetadata(XElement root, WorkbookSnapshot workbook)
        {
            foreach (XElement element in root.Elements("namedRange"))
                workbook.NamedRanges.Add(new NamedRange { Name = (string)element.Attribute("name"),
                    Sheet = (string)element.Attribute("sheet"), Range = GetRange(element) });
            foreach (XElement element in root.Elements("pivot"))
                workbook.Pivots.Add(new PivotDefinition { SourceSheet = (string)element.Attribute("sourceSheet"),
                    SourceRange = GetRange(element), GroupColumn = (int)element.Attribute("groupColumn"),
                    ValueColumn = (int)element.Attribute("valueColumn"),
                    Aggregate = (string)element.Attribute("aggregate") ?? "Sum",
                    TargetSheet = (string)element.Attribute("targetSheet") });
        }
    }
}

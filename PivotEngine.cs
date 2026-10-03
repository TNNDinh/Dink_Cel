using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DinkCel
{
    internal sealed class PivotResult
    {
        public readonly List<string[]> Rows = new List<string[]>();
        public readonly Dictionary<int, string> ExpandableRows = new Dictionary<int, string>();
    }

    internal static class PivotEngine
    {
        private const string Part = "\u001F";
        private const string Pair = "\u001E";

        private sealed class Accumulator
        {
            public double Sum, Min = double.PositiveInfinity, Max = double.NegativeInfinity;
            public int Count, NumericCount;
            public void Add(string raw)
            {
                if (!string.IsNullOrEmpty(raw)) Count++;
                double number;
                if (!DataTools.Number(raw, out number)) return;
                NumericCount++; Sum += number; Min = Math.Min(Min, number); Max = Math.Max(Max, number);
            }
            public string Result(string aggregate)
            {
                if (aggregate == "Count") return Count.ToString(CultureInfo.InvariantCulture);
                if (aggregate == "Average") return NumericCount == 0 ? "" :
                    (Sum / NumericCount).ToString("0.##########", CultureInfo.InvariantCulture);
                if (aggregate == "Min") return NumericCount == 0 ? "" : Min.ToString("0.##########", CultureInfo.InvariantCulture);
                if (aggregate == "Max") return NumericCount == 0 ? "" : Max.ToString("0.##########", CultureInfo.InvariantCulture);
                return Sum.ToString("0.##########", CultureInfo.InvariantCulture);
            }
        }

        private sealed class Node
        {
            public string Label, Key;
            public int Depth;
            public readonly Dictionary<string, Node> Children = new Dictionary<string, Node>(StringComparer.CurrentCultureIgnoreCase);
        }

        private static string GroupValue(string raw, string mode)
        {
            if (string.IsNullOrEmpty(raw)) return "(Trống)";
            if (mode == "None") return raw;
            DateTime date;
            if (!DataTools.Temporal(raw, out date)) return raw;
            return mode == "Year" ? date.ToString("yyyy", CultureInfo.InvariantCulture) :
                mode == "Month" ? date.ToString("yyyy-MM", CultureInfo.InvariantCulture) :
                date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static string Key(IEnumerable<string> parts)
        {
            return string.Join(Part, parts.Select(s => s.Replace(Part, " ").Replace(Pair, " ")).ToArray());
        }

        public static PivotResult Build(PivotDefinition definition, Func<int, int, string> read,
            int maxRows, int maxColumns)
        {
            var rows = definition.Rows.Count > 0 ? definition.Rows :
                new List<PivotAxisField> { new PivotAxisField { Column = definition.GroupColumn } };
            var columns = definition.Columns;
            var values = definition.Values.Count > 0 ? definition.Values :
                new List<PivotValueField> { new PivotValueField { Column = definition.ValueColumn, Aggregate = definition.Aggregate } };
            if (rows.Count == 0 || values.Count == 0) throw new InvalidOperationException("Pivot cần Rows và Values.");
            var metrics = new Dictionary<string, Accumulator[]>();
            var columnLabels = new SortedDictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            var root = new Node { Label = "", Key = "", Depth = 0 };
            int first = definition.SourceRange.Top + 1;
            for (int row = first; row < definition.SourceRange.Bottom; row++)
            {
                bool accepted = true;
                foreach (PivotFilterField filter in definition.Filters)
                    if (!string.Equals(read(row, filter.Column) ?? "", filter.Value,
                        StringComparison.CurrentCultureIgnoreCase)) { accepted = false; break; }
                if (!accepted) continue;
                string[] rowParts = rows.Select(f => GroupValue(read(row, f.Column), f.DateGroup)).ToArray();
                string[] colParts = columns.Select(f => GroupValue(read(row, f.Column), f.DateGroup)).ToArray();
                string colKey = Key(colParts);
                if (columns.Count > 0 && !columnLabels.ContainsKey(colKey)) columnLabels[colKey] = string.Join(" / ", colParts);
                Node node = root;
                for (int i = 0; i < rowParts.Length; i++)
                {
                    Node child;
                    if (!node.Children.TryGetValue(rowParts[i], out child))
                    {
                        child = new Node { Label = rowParts[i], Depth = i + 1,
                            Key = Key(rowParts.Take(i + 1)) };
                        node.Children[rowParts[i]] = child;
                    }
                    node = child;
                }
                string[] sourceValues = values.Select(f => read(row, f.Column) ?? "").ToArray();
                for (int depth = 0; depth <= rowParts.Length; depth++)
                {
                    string rowKey = Key(rowParts.Take(depth));
                    AddMetrics(metrics, rowKey, colKey, sourceValues);
                    if (columns.Count > 0) AddMetrics(metrics, rowKey, "", sourceValues);
                }
            }

            var displayedColumns = columns.Count == 0 ? new[] { "" } : columnLabels.Keys.ToArray();
            int resultColumns = 1 + displayedColumns.Length * values.Count +
                (columns.Count > 0 && definition.GrandTotal ? values.Count : 0);
            if (resultColumns > maxColumns)
                throw new InvalidOperationException("Pivot cần " + resultColumns + " cột; bảng hiện hỗ trợ tối đa " + maxColumns + ". Hãy lọc bớt giá trị Columns.");
            var result = new PivotResult();
            var header = new string[resultColumns];
            header[0] = string.Join(" / ", rows.Select(f => read(definition.SourceRange.Top, f.Column) ?? "Rows"));
            int outputColumn = 1;
            foreach (string key in displayedColumns)
                foreach (PivotValueField value in values)
                    header[outputColumn++] = (columns.Count == 0 ? "" : columnLabels[key] + " · ") +
                        value.Aggregate + " " + (read(definition.SourceRange.Top, value.Column) ?? "Value");
            if (columns.Count > 0 && definition.GrandTotal)
                foreach (PivotValueField value in values)
                    header[outputColumn++] = "Tổng · " + value.Aggregate + " " +
                        (read(definition.SourceRange.Top, value.Column) ?? "Value");
            result.Rows.Add(header);

            Func<Node, double> sortValue = node =>
            {
                Accumulator[] accumulators;
                return metrics.TryGetValue(node.Key + Pair, out accumulators) ? accumulators[0].Sum : 0;
            };
            Action<Node> emit = null;
            emit = node =>
            {
                bool branch = node.Children.Count > 0;
                bool collapsed = branch && definition.Collapsed.Contains(node.Key);
                if (branch)
                {
                    AddOutput(result, node.Key, new string(' ', (node.Depth - 1) * 2) +
                        (collapsed ? "▶ " : "▼ ") + node.Label,
                        collapsed ? metrics : null, displayedColumns, values, columns.Count > 0,
                        definition.GrandTotal, maxRows);
                    result.ExpandableRows[result.Rows.Count - 1] = node.Key;
                    if (collapsed) return;
                    IEnumerable<Node> children = node.Children.Values;
                    children = definition.SortByValue ? children.OrderBy(sortValue) :
                        children.OrderBy(n => n.Label, StringComparer.CurrentCultureIgnoreCase);
                    if (definition.SortDescending) children = children.Reverse();
                    foreach (Node child in children) emit(child);
                    if (definition.Subtotal)
                        AddOutput(result, node.Key, new string(' ', (node.Depth - 1) * 2) +
                            "Tổng " + node.Label, metrics, displayedColumns, values, columns.Count > 0,
                            definition.GrandTotal, maxRows);
                }
                else AddOutput(result, node.Key, new string(' ', (node.Depth - 1) * 2) + node.Label,
                    metrics, displayedColumns, values, columns.Count > 0, definition.GrandTotal, maxRows);
            };
            IEnumerable<Node> firstLevel = root.Children.Values;
            firstLevel = definition.SortByValue ? firstLevel.OrderBy(sortValue) :
                firstLevel.OrderBy(n => n.Label, StringComparer.CurrentCultureIgnoreCase);
            if (definition.SortDescending) firstLevel = firstLevel.Reverse();
            foreach (Node node in firstLevel) emit(node);
            if (definition.GrandTotal)
                    AddOutput(result, "", "Tổng cộng", metrics, displayedColumns, values, columns.Count > 0,
                        definition.GrandTotal, maxRows);
            return result;
        }

        private static void AddMetrics(Dictionary<string, Accumulator[]> metrics,
            string rowKey, string colKey, string[] values)
        {
            string key = rowKey + Pair + colKey;
            Accumulator[] accumulators;
            if (!metrics.TryGetValue(key, out accumulators))
            {
                accumulators = Enumerable.Range(0, values.Length).Select(_ => new Accumulator()).ToArray();
                metrics[key] = accumulators;
            }
            for (int i = 0; i < values.Length; i++) accumulators[i].Add(values[i]);
        }

        private static void AddOutput(PivotResult result, string rowKey, string label,
            Dictionary<string, Accumulator[]> metrics, string[] columns, IList<PivotValueField> values,
            bool columnAxis, bool grandTotal, int maxRows)
        {
            if (result.Rows.Count >= maxRows) throw new InvalidOperationException("Pivot vượt quá số hàng tối đa. Hãy lọc bớt dữ liệu.");
            var row = new string[1 + columns.Length * values.Count + (columnAxis && grandTotal ? values.Count : 0)];
            row[0] = label;
            int output = 1;
            foreach (string colKey in columns)
            {
                Accumulator[] stats;
                if (metrics != null && metrics.TryGetValue(rowKey + Pair + colKey, out stats))
                    for (int i = 0; i < values.Count; i++) row[output++] = stats[i].Result(values[i].Aggregate);
                else output += values.Count;
            }
            if (columnAxis && grandTotal)
            {
                Accumulator[] stats;
                if (metrics != null && metrics.TryGetValue(rowKey + Pair, out stats))
                    for (int i = 0; i < values.Count; i++) row[output++] = stats[i].Result(values[i].Aggregate);
            }
            result.Rows.Add(row);
        }
    }
}

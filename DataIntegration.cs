using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed class TabularData
    {
        public readonly List<string> Columns = new List<string>();
        public readonly List<string[]> Rows = new List<string[]>();
        public TabularData Copy()
        {
            var copy = new TabularData();
            copy.Columns.AddRange(Columns);
            copy.Rows.AddRange(Rows.Select(row => (string[])row.Clone()));
            return copy;
        }
    }

    internal sealed class QueryStep
    {
        public string Kind = "";
        public string Column = "";
        public string Operator = "";
        public string Value = "";
        public string Value2 = "";
    }

    internal sealed class DataQuery
    {
        public string Name = "";
        public string Kind = "Csv";
        public string Source = "";
        public string Sql = "";
        public string LoadTo = "Sheet";
        public string Target = "";
        public int LastRows;
        public int LastColumns;
        public readonly List<QueryStep> Steps = new List<QueryStep>();
    }

    internal static class QueryEngine
    {
        public const int MaxRows = 100000;
        public const int MaxColumns = 100;
        public const int MaxCells = 1000000;

        private static void Limit(TabularData table)
        {
            if (table.Columns.Count == 0 || table.Columns.Count > MaxColumns ||
                table.Rows.Count > MaxRows ||
                (long)table.Columns.Count * table.Rows.Count > MaxCells)
                throw new InvalidDataException("Import exceeds 100 columns, 100,000 rows or 1,000,000 cells.");
        }

        private static void SetHeaders(TabularData table, IEnumerable<string> headers)
        {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            foreach (string header in headers)
            {
                string basis = String.IsNullOrWhiteSpace(header) ? "Column" + (++index) : header.Trim();
                string candidate = basis; int suffix = 2;
                while (!used.Add(candidate)) candidate = basis + "_" + suffix++;
                table.Columns.Add(candidate);
            }
        }

        public static TabularData Normalize(TabularData raw)
        {
            var table = new TabularData(); SetHeaders(table, raw.Columns);
            foreach (string[] source in raw.Rows)
            {
                if (source.Length > table.Columns.Count)
                    throw new InvalidDataException("Data row has more fields than the header.");
                var row = new string[table.Columns.Count];
                Array.Copy(source, row, source.Length);
                for (int c = 0; c < row.Length; c++) row[c] = row[c] ?? "";
                table.Rows.Add(row);
            }
            Limit(table); return table;
        }

        public static TabularData FromCsv(string content)
        {
            IList<string[]> rows = CsvFile.Parse(new StringReader(content ?? ""), MaxRows + 1, MaxColumns);
            if (rows.Count == 0) throw new InvalidDataException("CSV is empty.");
            var table = new TabularData();
            SetHeaders(table, rows[0]);
            for (int i = 1; i < rows.Count; i++)
            {
                if (rows[i].Length > table.Columns.Count)
                    throw new InvalidDataException("CSV row " + (i + 1) + " has more fields than the header.");
                var row = new string[table.Columns.Count];
                Array.Copy(rows[i], row, Math.Min(row.Length, rows[i].Length));
                for (int c = 0; c < row.Length; c++) row[c] = row[c] ?? "";
                table.Rows.Add(row);
            }
            Limit(table); return table;
        }

        private static string Scalar(object value, JavaScriptSerializer serializer)
        {
            if (value == null) return "";
            if (value is string) return (string)value;
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is IDictionary<string, object> || value is object[])
                return serializer.Serialize(value);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        public static TabularData FromJson(string content)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            object root = serializer.DeserializeObject(content ?? "");
            var envelope = root as IDictionary<string, object>;
            if (envelope != null)
            {
                object data;
                if (envelope.TryGetValue("data", out data) && data is object[]) root = data;
                else if (envelope.TryGetValue("items", out data) && data is object[]) root = data;
            }
            object[] items = root as object[] ?? new[] { root };
            if (items.Length == 0) throw new InvalidDataException("JSON array is empty.");
            var table = new TabularData();
            var keys = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object item in items)
            {
                var record = item as IDictionary<string, object>;
                if (record == null) { if (seen.Add("Value")) keys.Add("Value"); }
                else foreach (string key in record.Keys) if (seen.Add(key)) keys.Add(key);
                if (keys.Count > MaxColumns) throw new InvalidDataException("JSON has too many columns.");
            }
            SetHeaders(table, keys);
            foreach (object item in items)
            {
                var record = item as IDictionary<string, object>;
                table.Rows.Add(keys.Select(key =>
                {
                    if (record == null) return Scalar(item, serializer);
                    string actual = record.Keys.FirstOrDefault(k => String.Equals(k, key,
                        StringComparison.OrdinalIgnoreCase));
                    return Scalar(actual == null ? null : record[actual], serializer);
                }).ToArray());
            }
            Limit(table); return table;
        }

        public static TabularData FromXml(string content)
        {
            XDocument xml = XDocument.Parse(content ?? "", LoadOptions.None);
            if (xml.Root == null) throw new InvalidDataException("XML has no root.");
            List<XElement> items = xml.Root.Elements().ToList();
            if (items.Count == 0) items.Add(xml.Root);
            // A single container such as <root><records><record>...</record></records></root>.
            if (items.Count == 1 && items[0].HasElements && items[0].Elements().Count() > 1 &&
                items[0].Elements().Select(e => e.Name).Distinct().Count() == 1)
                items = items[0].Elements().ToList();
            var keys = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var records = new List<Dictionary<string, string>>();
            foreach (XElement item in items)
            {
                var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (XAttribute attribute in item.Attributes()) record["@" + attribute.Name.LocalName] = attribute.Value;
                foreach (XElement child in item.Elements())
                {
                    if (record.ContainsKey(child.Name.LocalName))
                        throw new InvalidDataException("XML row contains repeated field: " + child.Name.LocalName);
                    record[child.Name.LocalName] = child.Value;
                }
                if (record.Count == 0) record["Value"] = item.Value;
                foreach (string key in record.Keys) if (seen.Add(key)) keys.Add(key);
                records.Add(record);
            }
            var table = new TabularData(); SetHeaders(table, keys);
            foreach (var record in records)
                table.Rows.Add(keys.Select(key => record.ContainsKey(key) ? record[key] : "").ToArray());
            Limit(table); return table;
        }

        public static bool ReadOnlySql(string sql)
        {
            string statement = (sql ?? "").Trim();
            return Regex.IsMatch(statement, @"^SELECT\s", RegexOptions.IgnoreCase) &&
                !statement.Contains(";") && !Regex.IsMatch(statement,
                    @"\b(INTO|UPDATE|DELETE|INSERT|DROP|ALTER|EXEC|CREATE)\b",
                    RegexOptions.IgnoreCase);
        }

        private static TabularData FromOdbc(string connectionString, string sql)
        {
            if (!ReadOnlySql(sql)) throw new InvalidOperationException("ODBC import accepts one SELECT statement only.");
            var table = new TabularData();
            using (var connection = new OdbcConnection(connectionString))
            using (var command = new OdbcCommand(sql, connection))
            {
                command.CommandTimeout = 30;
                connection.Open();
                using (var reader = command.ExecuteReader())
                {
                    if (reader.FieldCount > MaxColumns) throw new InvalidDataException("SQL result has too many columns.");
                    SetHeaders(table, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
                    while (reader.Read())
                    {
                        if (table.Rows.Count >= MaxRows ||
                            (long)(table.Rows.Count + 1) * table.Columns.Count > MaxCells)
                            throw new InvalidDataException("SQL result exceeds import limits.");
                        var row = new string[table.Columns.Count];
                        for (int i = 0; i < row.Length; i++)
                            row[i] = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i), CultureInfo.InvariantCulture);
                        table.Rows.Add(row);
                    }
                }
            }
            Limit(table); return table;
        }

        private static string ReadWeb(string url, string bearer)
        {
            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri) ||
                uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
                throw new ArgumentException("Enter an HTTP or HTTPS URL.");
            if (!String.IsNullOrEmpty(bearer) && uri.Scheme != Uri.UriSchemeHttps &&
                !uri.IsLoopback) throw new InvalidOperationException("API token requires HTTPS or localhost.");
            var request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "GET"; request.Timeout = 15000; request.ReadWriteTimeout = 15000;
            request.AllowAutoRedirect = false;
            if (!String.IsNullOrEmpty(bearer)) request.Headers[HttpRequestHeader.Authorization] = "Bearer " + bearer;
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream, WebEncoding(response.CharacterSet), true))
            {
                var text = new StringBuilder(); var buffer = new char[4096]; int count;
                while ((count = reader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    text.Append(buffer, 0, count);
                    if (text.Length > 16 * 1024 * 1024)
                        throw new InvalidDataException("Web response exceeds 16 MB.");
                }
                return text.ToString();
            }
        }

        private static Encoding WebEncoding(string charset)
        {
            if (String.IsNullOrWhiteSpace(charset)) return Encoding.UTF8;
            try { return Encoding.GetEncoding(charset); }
            catch (ArgumentException) { return Encoding.UTF8; }
        }

        private static string ReadFile(string path)
        {
            try
            {
                using (var reader = new StreamReader(path, new UTF8Encoding(false, true), true))
                    return reader.ReadToEnd();
            }
            catch (DecoderFallbackException)
            {
                using (var reader = new StreamReader(path, Encoding.Default, true))
                    return reader.ReadToEnd();
            }
        }

        public static TabularData Load(DataQuery query, string temporaryCredential)
        {
            string kind = query.Kind ?? "";
            if (!new[] { "Csv", "Json", "Xml", "WebCsv", "WebJson", "WebXml", "Odbc" }
                .Contains(kind, StringComparer.Ordinal))
                throw new InvalidDataException("Unsupported query source type: " + kind);
            if (kind == "Odbc") return Apply(FromOdbc(temporaryCredential, query.Sql), query.Steps);
            if (!kind.StartsWith("Web", StringComparison.Ordinal) &&
                new FileInfo(query.Source).Length > 16 * 1024 * 1024)
                throw new InvalidDataException("Source file exceeds 16 MB.");
            string content = kind.StartsWith("Web", StringComparison.Ordinal) ?
                ReadWeb(query.Source, temporaryCredential) : ReadFile(query.Source);
            TabularData table = kind.EndsWith("Json", StringComparison.Ordinal) ? FromJson(content) :
                kind.EndsWith("Xml", StringComparison.Ordinal) ? FromXml(content) : FromCsv(content);
            return Apply(table, query.Steps);
        }

        public static TabularData Apply(TabularData source, IEnumerable<QueryStep> steps)
        {
            TabularData table = source.Copy();
            foreach (QueryStep step in steps)
            {
                int column = table.Columns.FindIndex(c => String.Equals(c, step.Column,
                    StringComparison.OrdinalIgnoreCase));
                if (step.Kind != "Distinct" && column < 0)
                    throw new InvalidDataException("Query column not found: " + step.Column);
                if (step.Kind == "Filter") table.Rows.RemoveAll(row =>
                    !DataTools.Match(row[column], step.Value2 == "Number" || step.Value2 == "Date" ?
                        step.Value2 : "Text", step.Operator, step.Value, ""));
                else if (step.Kind == "Sort")
                {
                    table.Rows.Sort((a, b) => DataTools.Compare(a[column], b[column], "Auto") *
                        (step.Operator == "Descending" ? -1 : 1));
                }
                else if (step.Kind == "Rename")
                {
                    if (String.IsNullOrWhiteSpace(step.Value) || table.Columns.Where((c, i) => i != column)
                        .Any(c => String.Equals(c, step.Value, StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidDataException("Invalid or duplicate column name.");
                    table.Columns[column] = step.Value.Trim();
                }
                else if (step.Kind == "Remove")
                {
                    if (table.Columns.Count == 1) throw new InvalidDataException("Cannot remove the last column.");
                    table.Columns.RemoveAt(column);
                    for (int r = 0; r < table.Rows.Count; r++)
                        table.Rows[r] = table.Rows[r].Where((_, i) => i != column).ToArray();
                }
                else if (step.Kind == "Type")
                {
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        string value = table.Rows[r][column];
                        if (String.IsNullOrWhiteSpace(value)) continue;
                        double number; DateTime date;
                        if (step.Value == "Number" && DataTools.Number(value, out number))
                            table.Rows[r][column] = number.ToString("G17", CultureInfo.InvariantCulture);
                        else if (step.Value == "Date" && DataTools.Temporal(value, out date))
                            table.Rows[r][column] = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        else if (step.Value != "Text")
                            throw new InvalidDataException("Cannot convert row " + (r + 2) + " in " + step.Column);
                    }
                }
                else if (step.Kind == "Distinct")
                {
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    table.Rows.RemoveAll(row => !seen.Add(String.Concat(row.Select(value =>
                        (value ?? "").Length.ToString(CultureInfo.InvariantCulture) + ":" + (value ?? "")))));
                }
                else throw new InvalidDataException("Unknown query step: " + step.Kind);
            }
            Limit(table); return table;
        }
    }

    internal sealed class ModelCalculatedColumn
    {
        public string Name = "";
        public string Formula = "";
    }

    internal sealed class ModelTable
    {
        public string Name = "";
        public readonly TabularData Data = new TabularData();
        public readonly List<ModelCalculatedColumn> Calculated = new List<ModelCalculatedColumn>();
    }

    internal sealed class ModelRelationship
    {
        public string FactTable = "", FactColumn = "", LookupTable = "", LookupColumn = "";
    }

    internal sealed class ModelMeasure
    {
        public string Name = "", Table = "", Formula = "";
        public string Aggregate = "Sum";
    }

    internal sealed class ModelPivotDefinition
    {
        public string FactTable = "", TargetSheet = "", RowField = "", ColumnField = "";
        public string Measure = "", FilterField = "", FilterValue = "";
        public string TimelineField = "", StartDate = "", EndDate = "";
        public int LastRows, LastColumns;
    }

    internal static class DataModelEngine
    {
        private static string ColumnAddress(int column, int row)
        {
            string letters = "";
            for (int c = column + 1; c > 0; c = (c - 1) / 26)
                letters = (char)('A' + (c - 1) % 26) + letters;
            return letters + (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static string Bind(string expression, IList<string> headers, int row)
        {
            return Regex.Replace(expression ?? "", @"\{([^{}]+)\}", (Match match) =>
            {
                int index = headers.ToList().FindIndex(h => String.Equals(h, match.Groups[1].Value,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0) throw new InvalidDataException("Unknown model field: " + match.Groups[1].Value);
                return ColumnAddress(index, row);
            });
        }

        public static TabularData Materialize(ModelTable source)
        {
            TabularData result = source.Data.Copy();
            foreach (ModelCalculatedColumn column in source.Calculated)
            {
                if (String.IsNullOrWhiteSpace(column.Name) ||
                    result.Columns.Contains(column.Name, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidDataException("Invalid calculated column name: " + column.Name);
                if (!column.Formula.StartsWith("=", StringComparison.Ordinal))
                    throw new InvalidDataException("Calculated column must start with =.");
                var engine = new FormulaEngine((r, c) =>
                    r < result.Rows.Count && c < result.Rows[r].Length ? result.Rows[r][c] : "",
                    Math.Max(1, result.Rows.Count), result.Columns.Count);
                for (int r = 0; r < result.Rows.Count; r++)
                {
                    string value = engine.EvaluateExpression(Bind(column.Formula, result.Columns, r));
                    if (value.StartsWith("#", StringComparison.Ordinal))
                        throw new InvalidDataException("Calculated column " + column.Name +
                            " failed on row " + (r + 1) + ": " + value);
                    result.Rows[r] = result.Rows[r].Concat(new[] { value }).ToArray();
                }
                result.Columns.Add(column.Name);
            }
            return result;
        }

        public static TabularData Projection(string factName, IList<ModelTable> tables,
            IList<ModelRelationship> relationships)
        {
            ModelTable fact = tables.FirstOrDefault(t => String.Equals(t.Name, factName,
                StringComparison.OrdinalIgnoreCase));
            if (fact == null) throw new InvalidDataException("Model fact table not found: " + factName);
            TabularData baseData = Materialize(fact);
            var result = new TabularData();
            result.Columns.AddRange(baseData.Columns.Select(c => fact.Name + "." + c));
            result.Rows.AddRange(baseData.Rows.Select(r => (string[])r.Clone()));
            foreach (ModelRelationship relation in relationships.Where(r => String.Equals(r.FactTable,
                factName, StringComparison.OrdinalIgnoreCase)))
            {
                ModelTable lookup = tables.FirstOrDefault(t => String.Equals(t.Name,
                    relation.LookupTable, StringComparison.OrdinalIgnoreCase));
                if (lookup == null) throw new InvalidDataException("Related table missing: " + relation.LookupTable);
                TabularData dimension = Materialize(lookup);
                int factKey = baseData.Columns.FindIndex(c => String.Equals(c, relation.FactColumn,
                    StringComparison.OrdinalIgnoreCase));
                int lookupKey = dimension.Columns.FindIndex(c => String.Equals(c, relation.LookupColumn,
                    StringComparison.OrdinalIgnoreCase));
                if (factKey < 0 || lookupKey < 0) throw new InvalidDataException("Relationship key not found.");
                if (result.Columns.Any(c => c.StartsWith(lookup.Name + ".",
                    StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidDataException("Lookup table is connected more than once: " + lookup.Name);
                if (result.Columns.Count + dimension.Columns.Count > QueryEngine.MaxColumns ||
                    (long)(result.Columns.Count + dimension.Columns.Count) * result.Rows.Count > QueryEngine.MaxCells)
                    throw new InvalidDataException("Model projection exceeds limits.");
                var index = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
                foreach (string[] row in dimension.Rows)
                {
                    string key = row[lookupKey] ?? "";
                    if (index.ContainsKey(key))
                        throw new InvalidDataException("Lookup key is not unique in " + lookup.Name + ": " + key);
                    index.Add(key, row);
                }
                result.Columns.AddRange(dimension.Columns.Select(c => lookup.Name + "." + c));
                for (int r = 0; r < result.Rows.Count; r++)
                {
                    string[] match;
                    if (!index.TryGetValue(baseData.Rows[r][factKey] ?? "", out match))
                        match = new string[dimension.Columns.Count];
                    result.Rows[r] = result.Rows[r].Concat(match.Select(s => s ?? "")).ToArray();
                }
            }
            if (result.Columns.Count > QueryEngine.MaxColumns ||
                (long)result.Columns.Count * result.Rows.Count > QueryEngine.MaxCells)
                throw new InvalidDataException("Model projection exceeds limits.");
            return result;
        }

        public static PivotResult BuildPivot(ModelPivotDefinition definition,
            IList<ModelTable> tables, IList<ModelRelationship> relationships,
            IList<ModelMeasure> measures, int maxRows, int maxColumns)
        {
            TabularData data = Projection(definition.FactTable, tables, relationships);
            ModelMeasure measure = measures.FirstOrDefault(m => String.Equals(m.Name, definition.Measure,
                StringComparison.OrdinalIgnoreCase) && String.Equals(m.Table, definition.FactTable,
                StringComparison.OrdinalIgnoreCase));
            if (measure == null) throw new InvalidDataException("Measure not found: " + definition.Measure);
            int rowField = data.Columns.FindIndex(c => String.Equals(c, definition.RowField,
                StringComparison.OrdinalIgnoreCase));
            int columnField = String.IsNullOrEmpty(definition.ColumnField) ? -1 :
                data.Columns.FindIndex(c => String.Equals(c, definition.ColumnField,
                    StringComparison.OrdinalIgnoreCase));
            int filterField = String.IsNullOrEmpty(definition.FilterField) ? -1 :
                data.Columns.FindIndex(c => String.Equals(c, definition.FilterField,
                    StringComparison.OrdinalIgnoreCase));
            int timelineField = String.IsNullOrEmpty(definition.TimelineField) ? -1 :
                data.Columns.FindIndex(c => String.Equals(c, definition.TimelineField,
                    StringComparison.OrdinalIgnoreCase));
            if (rowField < 0 || !String.IsNullOrEmpty(definition.ColumnField) && columnField < 0 ||
                !String.IsNullOrEmpty(definition.FilterField) && filterField < 0 ||
                !String.IsNullOrEmpty(definition.TimelineField) && timelineField < 0)
                throw new InvalidDataException("Model Pivot field not found.");
            var engine = new FormulaEngine((r, c) => r > 0 && r <= data.Rows.Count && c < data.Columns.Count ?
                data.Rows[r - 1][c] : "", Math.Max(1, data.Rows.Count + 1), data.Columns.Count);
            var pivot = new PivotDefinition { SourceRange = new System.Drawing.Rectangle(0, 0,
                data.Columns.Count + 1, data.Rows.Count + 1), GrandTotal = true };
            pivot.Rows.Add(new PivotAxisField { Column = rowField });
            if (columnField >= 0) pivot.Columns.Add(new PivotAxisField { Column = columnField });
            pivot.Values.Add(new PivotValueField { Column = data.Columns.Count, Aggregate = measure.Aggregate });
            if (filterField >= 0) pivot.Filters.Add(new PivotFilterField
            { Column = filterField, Operator = "Slicer", Value = definition.FilterValue });
            if (timelineField >= 0) pivot.Filters.Add(new PivotFilterField
            { Column = timelineField, Operator = "BetweenDate", Value = definition.StartDate,
                Value2 = definition.EndDate });
            Func<int, int, string> read = (r, c) =>
            {
                if (c < data.Columns.Count) return r == 0 ? data.Columns[c] : data.Rows[r - 1][c];
                if (r == 0) return measure.Name;
                string value = engine.EvaluateExpression(Bind(measure.Formula, data.Columns, r));
                if (value.StartsWith("#", StringComparison.Ordinal))
                    throw new InvalidDataException("Measure " + measure.Name + " failed on row " + r + ": " + value);
                return value;
            };
            return PivotEngine.Build(pivot, read, maxRows, maxColumns);
        }
    }
}

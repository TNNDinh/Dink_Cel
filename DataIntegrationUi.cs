using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private readonly List<DataQuery> dataQueries = new List<DataQuery>();
        private readonly List<ModelTable> modelTables = new List<ModelTable>();
        private readonly List<ModelRelationship> modelRelationships = new List<ModelRelationship>();
        private readonly List<ModelMeasure> modelMeasures = new List<ModelMeasure>();
        private readonly List<ModelPivotDefinition> modelPivots = new List<ModelPivotDefinition>();

        private void AddDataIntegrationMenus()
        {
            var menuItem = new ToolStripMenuItem("Data Model");
            var queries = new ToolStripMenuItem("Queries");
            AddMenuItem(queries, "Import query...", Keys.None, CreateDataQuery);
            AddMenuItem(queries, "Add query step...", Keys.None, AddDataQueryStep);
            AddMenuItem(queries, "Remove last query step...", Keys.None, RemoveLastDataQueryStep);
            AddMenuItem(queries, "Refresh query...", Keys.None, RefreshChosenDataQuery);
            AddMenuItem(queries, "Refresh all queries", Keys.None, RefreshAllDataQueries);
            AddMenuItem(queries, "Change query source...", Keys.None, ChangeDataQuerySource);
            AddMenuItem(queries, "Delete query...", Keys.None, DeleteDataQuery);
            menuItem.DropDownItems.Add(queries);
            var model = new ToolStripMenuItem("Tables and measures");
            AddMenuItem(model, "Relationship...", Keys.None, CreateModelRelationship);
            AddMenuItem(model, "Delete relationship...", Keys.None, DeleteModelRelationship);
            AddMenuItem(model, "Calculated column...", Keys.None, CreateModelCalculatedColumn);
            AddMenuItem(model, "Delete calculated column...", Keys.None, DeleteModelCalculatedColumn);
            AddMenuItem(model, "Measure...", Keys.None, CreateModelMeasure);
            AddMenuItem(model, "Delete measure...", Keys.None, DeleteModelMeasure);
            menuItem.DropDownItems.Add(model);
            var pivot = new ToolStripMenuItem("Model Pivot");
            AddMenuItem(pivot, "Create...", Keys.None, CreateModelPivot);
            AddMenuItem(pivot, "Delete...", Keys.None, DeleteModelPivot);
            AddMenuItem(pivot, "Slicer...", Keys.None, ModelPivotSlicer);
            AddMenuItem(pivot, "Timeline...", Keys.None, ModelPivotTimeline);
            AddMenuItem(pivot, "Chart...", Keys.None, ModelPivotChart);
            AddMenuItem(pivot, "Refresh all", Keys.None, RefreshAllModelPivots);
            menuItem.DropDownItems.Add(pivot);
            menu.Items.Add(menuItem);
        }

        private string PromptSecret(string title, string initial)
        {
            using (var form = new Form { Text = title, Width = 530, Height = 140,
                FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false,
                MinimizeBox = false, StartPosition = FormStartPosition.CenterParent,
                BackColor = theme.Chrome, ForeColor = theme.Text, Font = DinkDesign.Ui })
            {
                var input = new TextBox { Left = 12, Top = 12, Width = 490,
                    Text = initial ?? "", UseSystemPasswordChar = true };
                var show = new CheckBox { Left = 12, Top = 52, Width = 90, Text = "Show" };
                show.CheckedChanged += (sender, args) => input.UseSystemPasswordChar = !show.Checked;
                var ok = new Button { Left = 420, Top = 48, Width = 82,
                    Text = "OK", DialogResult = DialogResult.OK };
                form.Controls.Add(input); form.Controls.Add(show); form.Controls.Add(ok); form.AcceptButton = ok;
                return form.ShowDialog(this) == DialogResult.OK ? input.Text : null;
            }
        }

        private DataQuery ChooseDataQuery()
        {
            string name = ChooseOption("Query", dataQueries.Select(q => q.Name).ToList());
            return dataQueries.FirstOrDefault(q => q.Name == name);
        }

        private static string SafeDataSheetName(string requested, IEnumerable<SheetState> existing)
        {
            string basis = Regex.Replace(requested ?? "Import", @"[\[\]:*?/\\]", "_");
            if (basis.Length > 27) basis = basis.Substring(0, 27);
            if (basis.Length == 0) basis = "Import";
            string name = basis; int suffix = 2;
            while (existing.Any(s => String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
                name = basis + suffix++;
            return name;
        }

        private void CreateDataQuery()
        {
            string kind = ChooseOption("Source type", new[] { "Sheet", "Csv", "Json", "Xml",
                "WebCsv", "WebJson", "WebXml", "Odbc" });
            if (kind == null) return;
            string source;
            if (kind == "Sheet") source = ChooseOption("Source sheet", sheets.Select(s => s.Name).ToList());
            else if (kind == "Csv" || kind == "Json" || kind == "Xml")
            {
                using (var dialog = new OpenFileDialog { Filter = "Data files|*.csv;*.json;*.xml|All files|*.*" })
                { if (dialog.ShowDialog(this) != DialogResult.OK) return; source = dialog.FileName; }
            }
            else source = Prompt(kind == "Odbc" ? "ODBC DSN name (no password saved)" :
                "Web/API URL", kind == "Odbc" ? "MyDataSource" : "https://");
            if (String.IsNullOrWhiteSpace(source)) return;
            string sql = kind == "Odbc" ? Prompt("Read-only SELECT query", "SELECT * FROM MyTable") : "";
            if (kind == "Odbc" && (sql == null || !QueryEngine.ReadOnlySql(sql)))
            { MessageBox.Show(this, "Enter one SELECT statement without a semicolon."); return; }
            string name = Prompt("Query name", "Query" + (dataQueries.Count + 1));
            if (String.IsNullOrWhiteSpace(name)) return;
            name = name.Trim();
            if (dataQueries.Any(q => String.Equals(q.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Query name already exists."); return; }
            string destination = kind == "Sheet" ? "Model" :
                ChooseOption("Load result to", new[] { "Sheet", "Model", "Both" });
            if (destination == null) return;
            var query = new DataQuery { Name = name, Kind = kind, Source = source.Trim(),
                Sql = sql ?? "", LoadTo = destination,
                Target = destination == "Model" ? name : SafeDataSheetName(name, sheets) };
            if (!RefreshDataQuery(query, null)) return;
            dataQueries.Add(query); MarkDirty();
        }

        private string QueryCredential(DataQuery query)
        {
            if (query.Kind == "Odbc")
                return PromptSecret("ODBC connection string (not saved)", "DSN=" + query.Source + ";");
            if (query.Kind.StartsWith("Web", StringComparison.Ordinal))
                return PromptSecret("Bearer token (optional, not saved)", "");
            return "";
        }

        private TabularData LoadDataQuery(DataQuery query, string credential)
        {
            if (query.Kind != "Sheet") return QueryEngine.Load(query, credential);
            SaveActiveSheet();
            SheetState source = sheets.FirstOrDefault(s => String.Equals(s.Name, query.Source,
                StringComparison.OrdinalIgnoreCase));
            if (source == null) throw new InvalidDataException("Source sheet is missing: " + query.Source);
            var used = source.Cells.Where(item => !String.IsNullOrEmpty(
                Convert.ToString(item.Value.Value))).Select(item => item.Key).ToList();
            int lastRow = used.Count == 0 ? -1 : used.Max() / ColumnCount;
            int lastColumn = used.Count == 0 ? -1 : used.Max(key => key % ColumnCount);
            if (lastRow < 0 || lastColumn < 0) throw new InvalidDataException("Source sheet is empty.");
            if (lastRow > QueryEngine.MaxRows ||
                (long)(lastRow + 1) * (lastColumn + 1) > QueryEngine.MaxCells)
                throw new InvalidDataException("Source sheet exceeds Data Model import limits.");
            var formula = new FormulaEngine((r, c) => StateRaw(source, r, c),
                (name, r, c) =>
                {
                    SheetState other = sheets.FirstOrDefault(s => String.Equals(s.Name, name,
                        StringComparison.OrdinalIgnoreCase));
                    return other == null ? null : StateRaw(other, r, c);
                }, source.Name, MaxRowCount, ColumnCount,
                name =>
                {
                    NamedRange named = FindName(name, "");
                    return named == null || !String.IsNullOrEmpty(named.Formula) ? null :
                        new FormulaNamedRange { Sheet = named.Sheet,
                            FirstRow = named.Range.Top, LastRow = named.Range.Bottom - 1,
                            FirstColumn = named.Range.Left, LastColumn = named.Range.Right - 1 };
                });
            formula.ResolveScopedRange = (name, sheet) =>
            {
                NamedRange named = FindName(name, sheet);
                return named == null || !String.IsNullOrEmpty(named.Formula) ? null :
                    new FormulaNamedRange { Sheet = named.Sheet,
                        FirstRow = named.Range.Top, LastRow = named.Range.Bottom - 1,
                        FirstColumn = named.Range.Left, LastColumn = named.Range.Right - 1 };
            };
            formula.ResolveNamedFormula = (name, sheet) =>
            {
                NamedRange named = FindName(name, sheet);
                return named == null ? null : named.Formula;
            };
            var raw = new TabularData();
            for (int c = 0; c <= lastColumn; c++) raw.Columns.Add(StateRaw(source, 0, c));
            for (int r = 1; r <= lastRow; r++)
                raw.Rows.Add(Enumerable.Range(0, lastColumn + 1).Select(c =>
                {
                    string value = StateRaw(source, r, c);
                    return value.StartsWith("=", StringComparison.Ordinal) ? formula.Display(r, c) : value;
                }).ToArray());
            return QueryEngine.Apply(QueryEngine.Normalize(raw), query.Steps);
        }

        private bool RefreshDataQuery(DataQuery query, TabularData prepared)
        {
            try
            {
                if (prepared == null)
                {
                    string credential = QueryCredential(query);
                    if (credential == null) return false;
                    prepared = LoadDataQuery(query, credential);
                }
                ApplyDataQueryResult(query, prepared);
                status.Text = "Refreshed " + query.Name + ": " + prepared.Rows.Count + " rows";
                return true;
            }
            catch (Exception error)
            { MessageBox.Show(this, error.Message, "Data import"); return false; }
        }

        private void ApplyDataQueryResult(DataQuery query, TabularData data)
        {
            if (query.LoadTo == "Sheet" || query.LoadTo == "Both")
            {
                if (data.Columns.Count > ColumnCount || data.Rows.Count + 1 > MaxRowCount ||
                    (long)Math.Max(data.Rows.Count + 1, query.LastRows) *
                    Math.Max(data.Columns.Count, query.LastColumns) > 50000)
                    throw new InvalidDataException("Sheet import exceeds 50,000 cells, 26 columns or 50,000 rows. Load to Model instead.");
                SheetState old = sheets.FirstOrDefault(s => String.Equals(s.Name, query.Target,
                    StringComparison.OrdinalIgnoreCase));
                if (old != null && old.Protected) throw new InvalidOperationException("Target sheet is protected.");
                if (old == null && structureProtected) throw new InvalidOperationException("Workbook structure is protected.");
            }
            ModelTable replacement = null;
            if (query.LoadTo == "Model" || query.LoadTo == "Both")
            {
                ModelTable existing = modelTables.FirstOrDefault(t => String.Equals(t.Name,
                    query.Name, StringComparison.OrdinalIgnoreCase));
                replacement = new ModelTable { Name = query.Name };
                replacement.Data.Columns.AddRange(data.Columns);
                replacement.Data.Rows.AddRange(data.Rows.Select(r => (string[])r.Clone()));
                if (existing != null) replacement.Calculated.AddRange(existing.Calculated);
                DataModelEngine.Materialize(replacement);
            }
            if (query.LoadTo == "Sheet" || query.LoadTo == "Both")
            {
                int index = sheets.FindIndex(s => String.Equals(s.Name, query.Target,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                {
                    int count = sheets.Count;
                    AddSheet();
                    if (sheets.Count == count) throw new InvalidOperationException("Cannot create import sheet.");
                    sheets[activeSheetIndex].Name = query.Target;
                    RefreshSheetTabs();
                }
                else SwitchSheet(index);
                var writes = new List<ScriptCellWrite>();
                int height = Math.Max(data.Rows.Count + 1, query.LastRows);
                int width = Math.Max(data.Columns.Count, query.LastColumns);
                for (int r = 0; r < height; r++)
                    for (int c = 0; c < width; c++)
                        writes.Add(new ScriptCellWrite { Row = r, Column = c,
                            Value = r == 0 && c < data.Columns.Count ? data.Columns[c] :
                                r > 0 && r <= data.Rows.Count && c < data.Columns.Count ?
                                data.Rows[r - 1][c] : null });
                ApplyScriptWrites(writes);
                query.LastRows = data.Rows.Count + 1; query.LastColumns = data.Columns.Count;
            }
            if (replacement != null)
            {
                modelTables.RemoveAll(t => String.Equals(t.Name, query.Name, StringComparison.OrdinalIgnoreCase));
                modelTables.Add(replacement);
                MarkDirty();
            }
        }

        private void RefreshChosenDataQuery()
        {
            DataQuery query = ChooseDataQuery();
            if (query != null && RefreshDataQuery(query, null)) RefreshAllModelPivots();
        }

        private void RefreshAllDataQueries()
        {
            foreach (DataQuery query in dataQueries.ToArray())
                if (!RefreshDataQuery(query, null)) return;
            RefreshAllModelPivots();
        }

        private void ChangeDataQuerySource()
        {
            DataQuery query = ChooseDataQuery();
            if (query == null) return;
            string source = query.Kind == "Sheet" ?
                ChooseOption("Source sheet", sheets.Select(s => s.Name).ToList()) :
                Prompt("Source path, URL or DSN", query.Source);
            if (String.IsNullOrWhiteSpace(source)) return;
            string old = query.Source; query.Source = source.Trim();
            if (!RefreshDataQuery(query, null)) query.Source = old;
        }

        private void DeleteDataQuery()
        {
            DataQuery query = ChooseDataQuery();
            if (query == null) return;
            bool model = query.LoadTo == "Model" || query.LoadTo == "Both";
            if (model && (modelRelationships.Any(r => r.FactTable == query.Name || r.LookupTable == query.Name) ||
                modelMeasures.Any(m => m.Table == query.Name) || modelPivots.Any(p => p.FactTable == query.Name)))
            { MessageBox.Show(this, "Delete dependent relationships, measures and Model Pivots first."); return; }
            if (MessageBox.Show(this, "Delete query " + query.Name +
                "? Visible sheet values will remain.", "Delete query", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes) return;
            dataQueries.Remove(query);
            if (model) modelTables.RemoveAll(t => t.Name == query.Name);
            MarkDirty();
        }

        private void AddDataQueryStep()
        {
            DataQuery query = ChooseDataQuery();
            if (query == null) return;
            try
            {
                string credential = QueryCredential(query);
                if (credential == null) return;
                TabularData current = LoadDataQuery(query, credential);
                string kind = ChooseOption("Transformation", new[] { "Filter", "Sort", "Rename",
                    "Remove", "Type", "Distinct" });
                if (kind == null) return;
                var step = new QueryStep { Kind = kind };
                if (kind != "Distinct")
                {
                    step.Column = ChooseOption("Column", current.Columns);
                    if (step.Column == null) return;
                }
                if (kind == "Filter")
                {
                    step.Operator = ChooseOption("Condition", new[] { "Equals", "Contains",
                        "Begins With", "Greater", "Less", "Blank", "Nonblank" });
                    if (step.Operator == null) return;
                    step.Value2 = ChooseOption("Compare as", new[] { "Text", "Number", "Date" });
                    if (step.Value2 == null) return;
                    if (step.Operator != "Blank" && step.Operator != "Nonblank")
                    { step.Value = Prompt("Compare to", ""); if (step.Value == null) return; }
                }
                else if (kind == "Sort")
                { step.Operator = ChooseOption("Order", new[] { "Ascending", "Descending" }); if (step.Operator == null) return; }
                else if (kind == "Rename")
                { step.Value = Prompt("New column name", step.Column); if (step.Value == null) return; }
                else if (kind == "Type")
                { step.Value = ChooseOption("Type", new[] { "Text", "Number", "Date" }); if (step.Value == null) return; }
                TabularData transformed = QueryEngine.Apply(current, new[] { step });
                ApplyDataQueryResult(query, transformed);
                query.Steps.Add(step); MarkDirty();
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Query transformation"); }
        }

        private void RemoveLastDataQueryStep()
        {
            DataQuery query = ChooseDataQuery();
            if (query == null || query.Steps.Count == 0) return;
            QueryStep removed = query.Steps[query.Steps.Count - 1];
            query.Steps.RemoveAt(query.Steps.Count - 1);
            if (!RefreshDataQuery(query, null)) query.Steps.Add(removed);
            else MarkDirty();
        }

        private void CreateModelRelationship()
        {
            string factName = ChooseOption("Fact table", modelTables.Select(t => t.Name).ToList());
            if (factName == null) return;
            string lookupName = ChooseOption("Lookup table", modelTables.Where(t => t.Name != factName)
                .Select(t => t.Name).ToList());
            if (lookupName == null) return;
            ModelTable fact = modelTables.First(t => t.Name == factName);
            ModelTable lookup = modelTables.First(t => t.Name == lookupName);
            if (modelRelationships.Any(r => r.FactTable == factName && r.LookupTable == lookupName))
            { MessageBox.Show(this, "A relationship between these tables already exists."); return; }
            string factKey = ChooseOption("Fact key", DataModelEngine.Materialize(fact).Columns);
            if (factKey == null) return;
            string lookupKey = ChooseOption("Unique lookup key", DataModelEngine.Materialize(lookup).Columns);
            if (lookupKey == null) return;
            var relation = new ModelRelationship { FactTable = factName, FactColumn = factKey,
                LookupTable = lookupName, LookupColumn = lookupKey };
            modelRelationships.Add(relation);
            try { DataModelEngine.Projection(factName, modelTables, modelRelationships); MarkDirty(); }
            catch (Exception error)
            { modelRelationships.Remove(relation); MessageBox.Show(this, error.Message, "Relationship"); }
        }

        private void DeleteModelRelationship()
        {
            var labels = modelRelationships.Select(r => r.FactTable + "." + r.FactColumn +
                " -> " + r.LookupTable + "." + r.LookupColumn).ToList();
            string selected = ChooseOption("Relationship", labels);
            if (selected == null) return;
            ModelRelationship relation = modelRelationships[labels.IndexOf(selected)];
            if (modelMeasures.Any(m => m.Table == relation.FactTable &&
                m.Formula.IndexOf("{" + relation.LookupTable + ".", StringComparison.OrdinalIgnoreCase) >= 0) ||
                modelPivots.Any(p => p.FactTable == relation.FactTable && new[] {
                p.RowField, p.ColumnField, p.FilterField, p.TimelineField }.Any(f =>
                    f.StartsWith(relation.LookupTable + ".", StringComparison.OrdinalIgnoreCase))))
            { MessageBox.Show(this, "Delete Model Pivots using this related table first."); return; }
            modelRelationships.Remove(relation); MarkDirty();
        }

        private void CreateModelCalculatedColumn()
        {
            string name = ChooseOption("Model table", modelTables.Select(t => t.Name).ToList());
            if (name == null) return;
            ModelTable table = modelTables.First(t => t.Name == name);
            string columnName = Prompt("Calculated column name", "Calculated" + (table.Calculated.Count + 1));
            if (String.IsNullOrWhiteSpace(columnName)) return;
            string formula = Prompt("Formula using {Column}, e.g. ={Amount}*2", "=");
            if (formula == null) return;
            var column = new ModelCalculatedColumn { Name = columnName.Trim(), Formula = formula };
            table.Calculated.Add(column);
            try { DataModelEngine.Materialize(table); MarkDirty(); }
            catch (Exception error)
            { table.Calculated.Remove(column); MessageBox.Show(this, error.Message, "Calculated column"); }
        }

        private void DeleteModelCalculatedColumn()
        {
            var choices = modelTables.SelectMany(t => t.Calculated.Select(c => t.Name + "." + c.Name)).ToList();
            string selected = ChooseOption("Calculated column", choices);
            if (selected == null) return;
            int dot = selected.IndexOf('.');
            string tableName = selected.Substring(0, dot), columnName = selected.Substring(dot + 1);
            if (modelMeasures.Any(m => m.Formula.IndexOf("{" + tableName + "." + columnName + "}",
                StringComparison.OrdinalIgnoreCase) >= 0) ||
                modelRelationships.Any(r => r.FactTable == tableName && r.FactColumn == columnName ||
                    r.LookupTable == tableName && r.LookupColumn == columnName) ||
                modelPivots.Any(p => new[] { p.RowField, p.ColumnField, p.FilterField,
                    p.TimelineField }.Contains(selected, StringComparer.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Delete dependent measures, relationships and Pivots first."); return; }
            modelTables.First(t => t.Name == tableName).Calculated.RemoveAll(c => c.Name == columnName);
            MarkDirty();
        }

        private void CreateModelMeasure()
        {
            string table = ChooseOption("Fact table", modelTables.Select(t => t.Name).ToList());
            if (table == null) return;
            string name = Prompt("Measure name", "Measure" + (modelMeasures.Count + 1));
            if (String.IsNullOrWhiteSpace(name)) return;
            name = name.Trim();
            if (modelMeasures.Any(m => m.Table == table &&
                String.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Measure name already exists."); return; }
            TabularData projection;
            try { projection = DataModelEngine.Projection(table, modelTables, modelRelationships); }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Data Model"); return; }
            string example = projection.Columns.Count > 0 ? "={" + projection.Columns.Last() + "}" : "=1";
            string formula = Prompt("Row expression using {Table.Column}", example);
            if (formula == null) return;
            if (!formula.StartsWith("=", StringComparison.Ordinal) ||
                Regex.Matches(formula, @"\{([^{}]+)\}").Cast<Match>().Any(m =>
                    !projection.Columns.Contains(m.Groups[1].Value, StringComparer.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Invalid formula or unknown model field."); return; }
            string aggregate = ChooseOption("Aggregation", new[] { "Sum", "Count", "Average", "Min", "Max" });
            if (aggregate == null) return;
            modelMeasures.Add(new ModelMeasure { Name = name, Table = table,
                Formula = formula, Aggregate = aggregate });
            MarkDirty();
        }

        private void DeleteModelMeasure()
        {
            var labels = modelMeasures.Select(m => m.Table + "." + m.Name).ToList();
            string selected = ChooseOption("Measure", labels);
            if (selected == null) return;
            ModelMeasure measure = modelMeasures[labels.IndexOf(selected)];
            if (modelPivots.Any(p => p.FactTable == measure.Table && p.Measure == measure.Name))
            { MessageBox.Show(this, "Delete Model Pivots using this measure first."); return; }
            modelMeasures.Remove(measure); MarkDirty();
        }

        private ModelPivotDefinition ChooseModelPivot()
        {
            string target = ChooseOption("Model Pivot", modelPivots.Select(p => p.TargetSheet).ToList());
            return modelPivots.FirstOrDefault(p => p.TargetSheet == target);
        }

        private void DeleteModelPivot()
        {
            ModelPivotDefinition pivot = ChooseModelPivot();
            if (pivot == null) return;
            if (MessageBox.Show(this, "Remove Model Pivot settings for " + pivot.TargetSheet +
                "? The result sheet will remain.", "Model Pivot", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes) return;
            SaveActiveSheet();
            SheetState target = sheets.FirstOrDefault(s => s.Name == pivot.TargetSheet);
            if (target != null) foreach (ChartDefinition chart in target.Charts) chart.PivotSource = "";
            foreach (ChartDefinition chart in charts.Where(c => c.PivotSource == pivot.TargetSheet))
                chart.PivotSource = "";
            modelPivots.Remove(pivot); MarkDirty();
        }

        private void CreateModelPivot()
        {
            string fact = ChooseOption("Fact table", modelTables.Select(t => t.Name).ToList());
            if (fact == null) return;
            TabularData data;
            try { data = DataModelEngine.Projection(fact, modelTables, modelRelationships); }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Data Model"); return; }
            string row = ChooseOption("Rows field", data.Columns);
            if (row == null) return;
            string column = ChooseOption("Columns field", new[] { "(None)" }.Concat(data.Columns).ToList());
            if (column == null) return;
            string measure = ChooseOption("Values / measure", modelMeasures.Where(m => m.Table == fact)
                .Select(m => m.Name).ToList());
            if (measure == null) { MessageBox.Show(this, "Create a measure first."); return; }
            if (structureProtected) { MessageBox.Show(this, "Workbook structure is protected."); return; }
            var pivot = new ModelPivotDefinition { FactTable = fact, RowField = row,
                ColumnField = column == "(None)" ? "" : column, Measure = measure,
                TargetSheet = SafeDataSheetName("ModelPivot" + (modelPivots.Count + 1), sheets) };
            PivotResult result;
            try { result = DataModelEngine.BuildPivot(pivot, modelTables, modelRelationships,
                modelMeasures, MaxRowCount, ColumnCount); }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Model Pivot"); return; }
            int count = sheets.Count;
            AddSheet();
            if (sheets.Count == count) return;
            sheets[activeSheetIndex].Name = pivot.TargetSheet;
            RefreshSheetTabs();
            modelPivots.Add(pivot);
            if (!WriteModelPivot(pivot, result)) modelPivots.Remove(pivot);
        }

        private bool WriteModelPivot(ModelPivotDefinition pivot, PivotResult result)
        {
            try
            {
                int target = sheets.FindIndex(s => String.Equals(s.Name, pivot.TargetSheet,
                    StringComparison.OrdinalIgnoreCase));
                if (target < 0) throw new InvalidOperationException("Model Pivot sheet is missing.");
                SwitchSheet(target);
                int height = Math.Max(pivot.LastRows, result.Rows.Count);
                int width = Math.Max(pivot.LastColumns, result.Rows[0].Length);
                if (height > MaxRowCount || width > ColumnCount || (long)height * width > 100000)
                    throw new InvalidDataException("Model Pivot exceeds the 100,000-cell write limit.");
                var writes = new List<ScriptCellWrite>();
                for (int r = 0; r < height; r++)
                    for (int c = 0; c < width; c++)
                        writes.Add(new ScriptCellWrite { Row = r, Column = c,
                            Value = r < result.Rows.Count && c < result.Rows[r].Length ? result.Rows[r][c] : null });
                ApplyScriptWrites(writes);
                pivot.LastRows = result.Rows.Count; pivot.LastColumns = result.Rows[0].Length;
                foreach (ChartDefinition chart in charts.Where(c => c.PivotSource == pivot.TargetSheet))
                    chart.Range = new Rectangle(0, 0, Math.Max(2, pivot.LastColumns), Math.Max(2, pivot.LastRows));
                RefreshChartOverlays(); MarkDirty();
                status.Text = "Refreshed Model Pivot " + pivot.TargetSheet;
                return true;
            }
            catch (Exception error)
            { MessageBox.Show(this, error.Message, "Model Pivot"); return false; }
        }

        private void RefreshModelPivot(ModelPivotDefinition pivot)
        {
            try
            {
                PivotResult result = DataModelEngine.BuildPivot(pivot, modelTables, modelRelationships,
                    modelMeasures, MaxRowCount, ColumnCount);
                WriteModelPivot(pivot, result);
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Model Pivot"); }
        }

        private void RefreshAllModelPivots()
        {
            foreach (ModelPivotDefinition pivot in modelPivots.ToArray()) RefreshModelPivot(pivot);
        }

        private void ModelPivotSlicer()
        {
            ModelPivotDefinition pivot = ChooseModelPivot();
            if (pivot == null) return;
            try
            {
                TabularData data = DataModelEngine.Projection(pivot.FactTable, modelTables, modelRelationships);
                string field = ChooseOption("Slicer field", data.Columns);
                if (field == null) return;
                int column = data.Columns.IndexOf(field);
                var values = data.Rows.Select(r => r[column]).Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase).ToList();
                values.Insert(0, "(All)");
                string selected = ChooseOption("Slicer value", values);
                if (selected == null) return;
                pivot.FilterField = selected == "(All)" ? "" : field;
                pivot.FilterValue = selected == "(All)" ? "" : selected;
                RefreshModelPivot(pivot);
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Model Pivot slicer"); }
        }

        private void ModelPivotTimeline()
        {
            ModelPivotDefinition pivot = ChooseModelPivot();
            if (pivot == null) return;
            try
            {
                TabularData data = DataModelEngine.Projection(pivot.FactTable, modelTables, modelRelationships);
                string field = ChooseOption("Timeline date field", data.Columns);
                if (field == null) return;
                string startText = Prompt("From date (yyyy-MM-dd); blank both to clear", "2024-01-01");
                if (startText == null) return;
                string endText = Prompt("To date (yyyy-MM-dd); blank both to clear", DateTime.Today.ToString("yyyy-MM-dd"));
                if (endText == null) return;
                if (String.IsNullOrWhiteSpace(startText) && String.IsNullOrWhiteSpace(endText))
                { pivot.TimelineField = pivot.StartDate = pivot.EndDate = ""; RefreshModelPivot(pivot); return; }
                DateTime start, end;
                if (!DateTime.TryParse(startText, CultureInfo.InvariantCulture, DateTimeStyles.None, out start) ||
                    !DateTime.TryParse(endText, CultureInfo.InvariantCulture, DateTimeStyles.None, out end) || start > end)
                    throw new ArgumentException("Invalid date range.");
                pivot.TimelineField = field;
                pivot.StartDate = start.ToString("yyyy-MM-dd"); pivot.EndDate = end.ToString("yyyy-MM-dd");
                RefreshModelPivot(pivot);
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Model Pivot timeline"); }
        }

        private void ModelPivotChart()
        {
            ModelPivotDefinition pivot = ChooseModelPivot();
            if (pivot == null) return;
            int index = sheets.FindIndex(s => s.Name == pivot.TargetSheet);
            if (index < 0) return;
            SwitchSheet(index);
            string kind = ChooseOption("Chart type", new[] { "Column", "Line", "Bar", "Pie" });
            if (kind == null) return;
            var chart = new ChartDefinition { Title = pivot.TargetSheet, Kind = kind,
                PivotSource = pivot.TargetSheet,
                Range = new Rectangle(0, 0, Math.Max(2, pivot.LastColumns), Math.Max(2, pivot.LastRows)),
                Placement = new Rectangle(Math.Min(ColumnCount - 8, Math.Max(2, pivot.LastColumns + 1)), 0, 8, 14) };
            EnsureRowCapacity(Math.Min(MaxRowCount, chart.Placement.Bottom));
            charts.Add(chart); RecordChange(); MarkDirty(); RefreshChartOverlays(); ShowChart(chart);
        }

        private void SerializeDataIntegration(XElement root)
        {
            var catalog = new XElement("dataIntegration");
            foreach (DataQuery query in dataQueries)
            {
                var element = new XElement("query", new XAttribute("name", query.Name),
                    new XAttribute("kind", query.Kind), new XAttribute("source", query.Source),
                    new XAttribute("sql", query.Sql), new XAttribute("loadTo", query.LoadTo),
                    new XAttribute("target", query.Target), new XAttribute("lastRows", query.LastRows),
                    new XAttribute("lastColumns", query.LastColumns));
                foreach (QueryStep step in query.Steps)
                    element.Add(new XElement("step", new XAttribute("kind", step.Kind),
                        new XAttribute("column", step.Column), new XAttribute("operator", step.Operator),
                        new XAttribute("value", step.Value), new XAttribute("value2", step.Value2)));
                catalog.Add(element);
            }
            foreach (ModelTable table in modelTables)
            {
                var element = new XElement("modelTable", new XAttribute("name", table.Name));
                foreach (string column in table.Data.Columns) element.Add(new XElement("column", column));
                foreach (string[] row in table.Data.Rows)
                    element.Add(new XElement("row", row.Select(value => new XElement("value", value ?? ""))));
                foreach (ModelCalculatedColumn column in table.Calculated)
                    element.Add(new XElement("calculated", new XAttribute("name", column.Name),
                        new XAttribute("formula", column.Formula)));
                catalog.Add(element);
            }
            foreach (ModelRelationship relation in modelRelationships)
                catalog.Add(new XElement("relationship", new XAttribute("fact", relation.FactTable),
                    new XAttribute("factKey", relation.FactColumn),
                    new XAttribute("lookup", relation.LookupTable),
                    new XAttribute("lookupKey", relation.LookupColumn)));
            foreach (ModelMeasure measure in modelMeasures)
                catalog.Add(new XElement("measure", new XAttribute("name", measure.Name),
                    new XAttribute("table", measure.Table), new XAttribute("formula", measure.Formula),
                    new XAttribute("aggregate", measure.Aggregate)));
            foreach (ModelPivotDefinition pivot in modelPivots)
                catalog.Add(new XElement("modelPivot", new XAttribute("fact", pivot.FactTable),
                    new XAttribute("target", pivot.TargetSheet), new XAttribute("row", pivot.RowField),
                    new XAttribute("column", pivot.ColumnField), new XAttribute("measure", pivot.Measure),
                    new XAttribute("filterField", pivot.FilterField),
                    new XAttribute("filterValue", pivot.FilterValue),
                    new XAttribute("timelineField", pivot.TimelineField),
                    new XAttribute("start", pivot.StartDate), new XAttribute("end", pivot.EndDate),
                    new XAttribute("lastRows", pivot.LastRows),
                    new XAttribute("lastColumns", pivot.LastColumns)));
            if (catalog.HasElements) root.Add(catalog);
        }

        private static void ReadDataIntegration(XElement root, WorkbookSnapshot workbook)
        {
            XElement catalog = root.Element("dataIntegration");
            if (catalog == null) return;
            foreach (XElement element in catalog.Elements("query"))
            {
                var query = new DataQuery { Name = (string)element.Attribute("name") ?? "",
                    Kind = (string)element.Attribute("kind") ?? "Csv",
                    Source = (string)element.Attribute("source") ?? "",
                    Sql = (string)element.Attribute("sql") ?? "",
                    LoadTo = (string)element.Attribute("loadTo") ?? "Sheet",
                    Target = (string)element.Attribute("target") ?? "",
                    LastRows = (int?)element.Attribute("lastRows") ?? 0,
                    LastColumns = (int?)element.Attribute("lastColumns") ?? 0 };
                foreach (XElement step in element.Elements("step"))
                    query.Steps.Add(new QueryStep { Kind = (string)step.Attribute("kind") ?? "",
                        Column = (string)step.Attribute("column") ?? "",
                        Operator = (string)step.Attribute("operator") ?? "",
                        Value = (string)step.Attribute("value") ?? "",
                        Value2 = (string)step.Attribute("value2") ?? "" });
                if (query.Name.Length > 0) workbook.DataQueries.Add(query);
            }
            foreach (XElement element in catalog.Elements("modelTable"))
            {
                var table = new ModelTable { Name = (string)element.Attribute("name") ?? "" };
                table.Data.Columns.AddRange(element.Elements("column").Select(x => x.Value));
                if (table.Data.Columns.Count == 0 || table.Data.Columns.Count > QueryEngine.MaxColumns)
                    throw new InvalidDataException("Invalid Data Model column count.");
                foreach (XElement row in element.Elements("row"))
                {
                    if (table.Data.Rows.Count >= QueryEngine.MaxRows ||
                        (long)(table.Data.Rows.Count + 1) * table.Data.Columns.Count > QueryEngine.MaxCells)
                        throw new InvalidDataException("Data Model exceeds import limits.");
                    string[] values = row.Elements("value").Select(x => x.Value).ToArray();
                    if (values.Length != table.Data.Columns.Count)
                        throw new InvalidDataException("Data Model row width does not match its headers.");
                    table.Data.Rows.Add(values);
                }
                foreach (XElement column in element.Elements("calculated"))
                    table.Calculated.Add(new ModelCalculatedColumn
                    { Name = (string)column.Attribute("name") ?? "",
                        Formula = (string)column.Attribute("formula") ?? "" });
                workbook.ModelTables.Add(table);
            }
            foreach (XElement element in catalog.Elements("relationship"))
                workbook.ModelRelationships.Add(new ModelRelationship
                { FactTable = (string)element.Attribute("fact") ?? "",
                    FactColumn = (string)element.Attribute("factKey") ?? "",
                    LookupTable = (string)element.Attribute("lookup") ?? "",
                    LookupColumn = (string)element.Attribute("lookupKey") ?? "" });
            foreach (XElement element in catalog.Elements("measure"))
                workbook.ModelMeasures.Add(new ModelMeasure
                { Name = (string)element.Attribute("name") ?? "",
                    Table = (string)element.Attribute("table") ?? "",
                    Formula = (string)element.Attribute("formula") ?? "",
                    Aggregate = (string)element.Attribute("aggregate") ?? "Sum" });
            foreach (XElement element in catalog.Elements("modelPivot"))
                workbook.ModelPivots.Add(new ModelPivotDefinition
                { FactTable = (string)element.Attribute("fact") ?? "",
                    TargetSheet = (string)element.Attribute("target") ?? "",
                    RowField = (string)element.Attribute("row") ?? "",
                    ColumnField = (string)element.Attribute("column") ?? "",
                    Measure = (string)element.Attribute("measure") ?? "",
                    FilterField = (string)element.Attribute("filterField") ?? "",
                    FilterValue = (string)element.Attribute("filterValue") ?? "",
                    TimelineField = (string)element.Attribute("timelineField") ?? "",
                    StartDate = (string)element.Attribute("start") ?? "",
                    EndDate = (string)element.Attribute("end") ?? "",
                    LastRows = (int?)element.Attribute("lastRows") ?? 0,
                    LastColumns = (int?)element.Attribute("lastColumns") ?? 0 });
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace DinkCel
{
    internal sealed partial class FormulaEngine
    {
        private static readonly Regex dynamicSyntax = new Regex(
            @"\b(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|TRANSPOSE|LET|CHOOSECOLS|CHOOSEROWS|TAKE|DROP|VSTACK|HSTACK)\s*\(|\b\$?[A-Z]+\$?[1-9][0-9]*#",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public static bool HasDynamicArraySyntax(string formula)
        { return !String.IsNullOrEmpty(formula) && dynamicSyntax.IsMatch(formula); }
        private readonly List<Dictionary<string, Value>> letScopes = new List<Dictionary<string, Value>>();

        private bool TryLetValue(string name, out Value value)
        {
            for (int index = letScopes.Count - 1; index >= 0; index--)
                if (letScopes[index].TryGetValue(name, out value)) return true;
            value = null;
            return false;
        }

        private static Value Matrix(Value value)
        {
            return value.Kind == ValueKind.Range ? value :
                Value.Range(new List<Value> { value }, 1, 1);
        }

        private static Value At(Value value, int row, int column)
        { return value.Items[row * value.Columns + column]; }

        private static Value NumberArgument(Node node, FormulaEngine engine)
        {
            Value value = node.Evaluate(engine);
            if (value.Kind == ValueKind.Range)
                value = value.Items.Count == 1 ? value.Items[0] : Value.Error("#VALUE!");
            return AsNumber(value);
        }

        private static int CompareDynamic(Value a, Value b)
        {
            if (a.Kind == ValueKind.Blank && b.Kind != ValueKind.Blank) return 1;
            if (b.Kind == ValueKind.Blank && a.Kind != ValueKind.Blank) return -1;
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Number)
                return a.Number.CompareTo(b.Number);
            DateTime dateA, dateB;
            if (DateTime.TryParse(a.Text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateA) &&
                DateTime.TryParse(b.Text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateB))
                return dateA.CompareTo(dateB);
            string textA = a.Kind == ValueKind.Number ? a.Number.ToString(CultureInfo.InvariantCulture) : a.Text;
            string textB = b.Kind == ValueKind.Number ? b.Number.ToString(CultureInfo.InvariantCulture) : b.Text;
            return String.Compare(textA, textB, StringComparison.OrdinalIgnoreCase);
        }

        private static Value Rearrange(Value matrix, IList<int> rows, IList<int> columns)
        {
            if (rows.Count == 0 || columns.Count == 0) return Value.Error("#CALC!");
            if ((long)rows.Count * columns.Count > 100000) return Value.Error("#NUM!");
            var items = new List<Value>(rows.Count * columns.Count);
            foreach (int row in rows)
                foreach (int column in columns)
                    items.Add(At(matrix, row, column));
            return Value.Range(items, rows.Count, columns.Count);
        }

        private static int[] All(int count)
        { return Enumerable.Range(0, count).ToArray(); }

        private Value EvaluateDynamic(string name, List<Node> arguments)
        {
            if (name == "LET")
            {
                if (arguments.Count < 3 || arguments.Count % 2 == 0) return Value.Error("#VALUE!");
                var scope = new Dictionary<string, Value>(StringComparer.OrdinalIgnoreCase);
                letScopes.Add(scope);
                try
                {
                    for (int i = 0; i < arguments.Count - 1; i += 2)
                    {
                        NameNode identifier = arguments[i] as NameNode;
                        if (identifier == null) return Value.Error("#NAME?");
                        scope[identifier.Name] = arguments[i + 1].Evaluate(this);
                    }
                    return arguments[arguments.Count - 1].Evaluate(this);
                }
                finally { letScopes.RemoveAt(letScopes.Count - 1); }
            }
            if (name != "FILTER" && name != "SORT" && name != "SORTBY" &&
                name != "UNIQUE" && name != "SEQUENCE" && name != "TRANSPOSE" &&
                name != "CHOOSECOLS" && name != "CHOOSEROWS" && name != "TAKE" &&
                name != "DROP" && name != "VSTACK" && name != "HSTACK") return null;

            if (name == "SEQUENCE")
            {
                if (arguments.Count < 1 || arguments.Count > 4) return Value.Error("#VALUE!");
                Value r = NumberArgument(arguments[0], this);
                Value c = arguments.Count > 1 ? NumberArgument(arguments[1], this) : Value.Numeric(1);
                Value start = arguments.Count > 2 ? NumberArgument(arguments[2], this) : Value.Numeric(1);
                Value step = arguments.Count > 3 ? NumberArgument(arguments[3], this) : Value.Numeric(1);
                if (r.Kind == ValueKind.Error) return r;
                if (c.Kind == ValueKind.Error) return c;
                if (start.Kind == ValueKind.Error) return start;
                if (step.Kind == ValueKind.Error) return step;
                int rows = (int)r.Number, columns = (int)c.Number;
                if (rows < 1 || columns < 1 || (long)rows * columns > 100000) return Value.Error("#NUM!");
                var values = new List<Value>(rows * columns);
                for (int i = 0; i < rows * columns; i++) values.Add(Value.Numeric(start.Number + i * step.Number));
                return Value.Range(values, rows, columns);
            }
            if (arguments.Count == 0) return Value.Error("#VALUE!");
            Value source = Matrix(arguments[0].Evaluate(this));
            if (source.Kind == ValueKind.Error) return source;
            if (name == "TRANSPOSE")
            {
                if (arguments.Count != 1) return Value.Error("#VALUE!");
                var values = new List<Value>();
                for (int col = 0; col < source.Columns; col++)
                    for (int row = 0; row < source.Rows; row++) values.Add(At(source, row, col));
                return Value.Range(values, source.Columns, source.Rows);
            }
            if (name == "FILTER")
            {
                if (arguments.Count < 2 || arguments.Count > 3) return Value.Error("#VALUE!");
                Value include = Matrix(arguments[1].Evaluate(this));
                if (include.Kind == ValueKind.Error) return include;
                bool rows = include.Rows == source.Rows && include.Columns == 1;
                bool columns = include.Columns == source.Columns && include.Rows == 1;
                if (!rows && !columns) return Value.Error("#VALUE!");
                var selected = new List<int>();
                for (int i = 0; i < (rows ? source.Rows : source.Columns); i++)
                {
                    Value flag = include.Items[i];
                    if (flag.Kind == ValueKind.Error) return flag;
                    Value number = AsNumber(flag);
                    if (number.Kind == ValueKind.Number && number.Number != 0) selected.Add(i);
                }
                if (selected.Count == 0) return arguments.Count == 3 ?
                    arguments[2].Evaluate(this) : Value.Error("#CALC!");
                return rows ? Rearrange(source, selected, All(source.Columns)) :
                    Rearrange(source, All(source.Rows), selected);
            }
            if (name == "SORT" || name == "SORTBY")
            {
                if (name == "SORT" && (arguments.Count > 4 || arguments.Count < 1) ||
                    name == "SORTBY" && (arguments.Count < 2 || arguments.Count > 9))
                    return Value.Error("#VALUE!");
                bool byColumn = false;
                int field = 0, order = 1;
                if (name == "SORT")
                {
                    if (arguments.Count > 1) { Value n = NumberArgument(arguments[1], this); if (n.Kind == ValueKind.Error) return n; field = (int)n.Number - 1; }
                    if (arguments.Count > 2) { Value n = NumberArgument(arguments[2], this); if (n.Kind == ValueKind.Error) return n; order = (int)n.Number; }
                    if (arguments.Count > 3) { Value n = NumberArgument(arguments[3], this); if (n.Kind == ValueKind.Error) return n; byColumn = n.Number != 0; }
                    if (field < 0 || field >= (byColumn ? source.Rows : source.Columns) || Math.Abs(order) != 1)
                        return Value.Error("#VALUE!");
                }
                int count = byColumn ? source.Columns : source.Rows;
                var indexes = All(count).ToList();
                var keys = new List<Value>();
                var orders = new List<int>();
                if (name == "SORTBY")
                {
                    for (int i = 1; i < arguments.Count; i++)
                    {
                        Value key = Matrix(arguments[i].Evaluate(this));
                        if (key.Kind == ValueKind.Error) return key;
                        if (key.Rows != source.Rows || key.Columns != 1) return Value.Error("#VALUE!");
                        keys.Add(key); orders.Add(1);
                        if (i + 1 < arguments.Count)
                        {
                            Value maybeOrder = NumberArgument(arguments[i + 1], this);
                            if (maybeOrder.Kind == ValueKind.Number && Math.Abs(maybeOrder.Number) == 1)
                            { orders[orders.Count - 1] = (int)maybeOrder.Number; i++; }
                        }
                    }
                }
                indexes.Sort((a, b) =>
                {
                    if (name == "SORT") return order * CompareDynamic(byColumn ? At(source, field, a) : At(source, a, field),
                        byColumn ? At(source, field, b) : At(source, b, field));
                    for (int k = 0; k < keys.Count; k++)
                    {
                        int comparison = orders[k] * CompareDynamic(At(keys[k], a, 0), At(keys[k], b, 0));
                        if (comparison != 0) return comparison;
                    }
                    return a.CompareTo(b);
                });
                return byColumn ? Rearrange(source, All(source.Rows), indexes) :
                    Rearrange(source, indexes, All(source.Columns));
            }
            if (name == "UNIQUE")
            {
                if (arguments.Count > 3) return Value.Error("#VALUE!");
                bool byColumn = arguments.Count > 1 && NumberArgument(arguments[1], this).Number != 0;
                bool once = arguments.Count > 2 && NumberArgument(arguments[2], this).Number != 0;
                int count = byColumn ? source.Columns : source.Rows;
                var groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < count; i++)
                {
                    var parts = new List<string>();
                    for (int j = 0; j < (byColumn ? source.Rows : source.Columns); j++)
                    {
                        Value item = byColumn ? At(source, j, i) : At(source, i, j);
                        parts.Add(item.Kind + ":" + (item.Kind == ValueKind.Number ?
                            item.Number.ToString("R", CultureInfo.InvariantCulture) : item.Text));
                    }
                    string key = String.Join("\u001f", parts.ToArray());
                    List<int> matches;
                    if (!groups.TryGetValue(key, out matches)) groups[key] = matches = new List<int>();
                    matches.Add(i);
                }
                var selected = groups.Values.Where(x => !once || x.Count == 1).Select(x => x[0]).OrderBy(x => x).ToList();
                return byColumn ? Rearrange(source, All(source.Rows), selected) :
                    Rearrange(source, selected, All(source.Columns));
            }
            if (name == "CHOOSECOLS" || name == "CHOOSEROWS")
            {
                if (arguments.Count < 2) return Value.Error("#VALUE!");
                int size = name == "CHOOSECOLS" ? source.Columns : source.Rows;
                var selected = new List<int>();
                for (int i = 1; i < arguments.Count; i++)
                {
                    Value number = NumberArgument(arguments[i], this);
                    if (number.Kind == ValueKind.Error) return number;
                    int index = (int)number.Number;
                    if (index == 0 || Math.Abs(index) > size) return Value.Error("#VALUE!");
                    selected.Add(index > 0 ? index - 1 : size + index);
                }
                return name == "CHOOSECOLS" ? Rearrange(source, All(source.Rows), selected) :
                    Rearrange(source, selected, All(source.Columns));
            }
            if (name == "TAKE" || name == "DROP")
            {
                if (arguments.Count < 2 || arguments.Count > 3) return Value.Error("#VALUE!");
                Value r = NumberArgument(arguments[1], this);
                Value c = arguments.Count == 3 ? NumberArgument(arguments[2], this) : Value.Numeric(source.Columns);
                if (r.Kind == ValueKind.Error) return r;
                if (c.Kind == ValueKind.Error) return c;
                int rows = (int)r.Number, columns = (int)c.Number;
                if (name == "TAKE")
                {
                    if (rows == 0 || columns == 0) return Value.Error("#CALC!");
                    int[] rr = rows > 0 ? All(Math.Min(rows, source.Rows)) :
                        Enumerable.Range(source.Rows - Math.Min(-rows, source.Rows), Math.Min(-rows, source.Rows)).ToArray();
                    int[] cc = columns > 0 ? All(Math.Min(columns, source.Columns)) :
                        Enumerable.Range(source.Columns - Math.Min(-columns, source.Columns), Math.Min(-columns, source.Columns)).ToArray();
                    return Rearrange(source, rr, cc);
                }
                int firstRow = rows >= 0 ? Math.Min(rows, source.Rows) : 0;
                int firstColumn = columns >= 0 ? Math.Min(columns, source.Columns) : 0;
                int remainingRows = Math.Max(0, source.Rows - Math.Abs(rows));
                int remainingColumns = Math.Max(0, source.Columns - Math.Abs(columns));
                return Rearrange(source, Enumerable.Range(firstRow, remainingRows).ToArray(),
                    Enumerable.Range(firstColumn, remainingColumns).ToArray());
            }
            if (name == "VSTACK" || name == "HSTACK")
            {
                var arrays = new List<Value>();
                foreach (Node argument in arguments)
                {
                    Value value = Matrix(argument.Evaluate(this));
                    if (value.Kind == ValueKind.Error) return value;
                    arrays.Add(value);
                }
                int rows = name == "VSTACK" ? arrays.Sum(x => x.Rows) : arrays.Max(x => x.Rows);
                int columns = name == "VSTACK" ? arrays.Max(x => x.Columns) : arrays.Sum(x => x.Columns);
                if ((long)rows * columns > 100000) return Value.Error("#NUM!");
                var result = Enumerable.Repeat(Value.Error("#N/A"), rows * columns).ToList();
                int offset = 0;
                foreach (Value array in arrays)
                {
                    for (int r = 0; r < array.Rows; r++)
                        for (int c = 0; c < array.Columns; c++)
                            result[(name == "VSTACK" ? offset + r : r) * columns +
                                (name == "VSTACK" ? c : offset + c)] = At(array, r, c);
                    offset += name == "VSTACK" ? array.Rows : array.Columns;
                }
                return Value.Range(result, rows, columns);
            }
            return null;
        }
    }
}

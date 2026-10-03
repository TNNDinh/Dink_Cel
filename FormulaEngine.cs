using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DinkCel
{
    internal sealed class FormulaNamedRange
    {
        public string Sheet;
        public int FirstRow;
        public int FirstColumn;
        public int LastRow;
        public int LastColumn;
    }

    internal sealed partial class FormulaEngine
    {
        private readonly Func<int, int, string> readCell;
        private readonly Func<string, int, int, string> readOtherSheet;
        private readonly string currentSheet;
        private readonly Func<string, FormulaNamedRange> resolveName;
        private readonly int rowCount;
        private readonly int columnCount;
        private readonly Dictionary<string, Value> cache = new Dictionary<string, Value>();
        private readonly HashSet<string> active = new HashSet<string>();
        private readonly Dictionary<string, HashSet<string>> dependencies =
            new Dictionary<string, HashSet<string>>();
        private readonly Dictionary<string, HashSet<string>> dependents =
            new Dictionary<string, HashSet<string>>();
        private readonly List<string> calculationChain = new List<string>();
        private readonly Stack<string> evaluationStack = new Stack<string>();
        private readonly HashSet<string> volatileCells = new HashSet<string>();
        private readonly Func<DateTime> nowProvider;

        public FormulaEngine(Func<int, int, string> readCell, int rowCount, int columnCount)
            : this(readCell, null, "", rowCount, columnCount) { }

        public FormulaEngine(Func<int, int, string> readCell,
            Func<string, int, int, string> readOtherSheet, string currentSheet,
            int rowCount, int columnCount)
            : this(readCell, readOtherSheet, currentSheet, rowCount, columnCount, null) { }

        public FormulaEngine(Func<int, int, string> readCell,
            Func<string, int, int, string> readOtherSheet, string currentSheet,
            int rowCount, int columnCount, Func<string, FormulaNamedRange> resolveName)
            : this(readCell, readOtherSheet, currentSheet, rowCount, columnCount,
                resolveName, null) { }

        public FormulaEngine(Func<int, int, string> readCell,
            Func<string, int, int, string> readOtherSheet, string currentSheet,
            int rowCount, int columnCount, Func<string, FormulaNamedRange> resolveName,
            Func<DateTime> nowProvider)
        {
            this.readCell = readCell;
            this.readOtherSheet = readOtherSheet;
            this.currentSheet = currentSheet ?? "";
            this.resolveName = resolveName;
            this.rowCount = rowCount;
            this.columnCount = columnCount;
            this.nowProvider = nowProvider ?? delegate { return DateTime.Now; };
        }

        public IList<string> CalculationChain
        { get { return calculationChain.AsReadOnly(); } }

        public IList<string> DependenciesFor(string sheet, int row, int column)
        {
            HashSet<string> items;
            return dependencies.TryGetValue(CellKey(sheet, row, column), out items) ?
                new List<string>(items).AsReadOnly() : new List<string>().AsReadOnly();
        }

        public void InvalidateAll()
        {
            cache.Clear(); active.Clear(); dependencies.Clear(); dependents.Clear();
            calculationChain.Clear(); evaluationStack.Clear(); volatileCells.Clear();
        }

        public void Invalidate(string sheet, int row, int column)
        {
            var pending = new Stack<string>();
            var affected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pending.Push(CellKey(sheet, row, column));
            foreach (string key in volatileCells) pending.Push(key);
            while (pending.Count > 0)
            {
                string key = pending.Pop();
                if (!affected.Add(key)) continue;
                HashSet<string> users;
                if (dependents.TryGetValue(key, out users))
                    foreach (string user in users) pending.Push(user);
            }
            foreach (string key in affected)
            {
                cache.Remove(key); volatileCells.Remove(key);
                HashSet<string> inputs;
                if (dependencies.TryGetValue(key, out inputs))
                    foreach (string input in inputs)
                    {
                        HashSet<string> users;
                        if (dependents.TryGetValue(input, out users)) users.Remove(key);
                    }
                dependencies.Remove(key);
                dependents.Remove(key);
            }
            calculationChain.RemoveAll(affected.Contains);
        }

        private static string CellKey(string sheet, int row, int column)
        {
            return (sheet ?? "").ToUpperInvariant() + "!" + ColumnName(column + 1) +
                (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        private void MarkVolatile()
        {
            if (evaluationStack.Count > 0) volatileCells.Add(evaluationStack.Peek());
        }

        public string Display(int row, int column)
        {
            Value result = EvaluateCell(row, column);
            if (result.Kind == ValueKind.Number)
            {
                if (Double.IsNaN(result.Number) || Double.IsInfinity(result.Number))
                    return "#NUM!";
                return result.Number.ToString("0.##########", CultureInfo.InvariantCulture);
            }
            if (result.Kind == ValueKind.Blank)
                return "0";
            if (result.Kind == ValueKind.Range)
                return "#VALUE!";
            return result.Text;
        }

        public string EvaluateExpression(string expression)
        {
            try
            {
                Value result = new Parser(this, (expression ?? "").TrimStart('='), currentSheet).Parse().Evaluate(this);
                if (result.Kind == ValueKind.Number) return result.Number.ToString("0.##########", CultureInfo.InvariantCulture);
                if (result.Kind == ValueKind.Blank) return "0";
                if (result.Kind == ValueKind.Range) return "#VALUE!";
                return result.Text;
            }
            catch (Exception) { return "#VALUE!"; }
        }

        private Value EvaluateCell(int row, int column)
        {
            return EvaluateCell(currentSheet, row, column);
        }

        private Value EvaluateCell(string sheet, int row, int column)
        {
            if (row < 0 || row >= rowCount || column < 0 || column >= columnCount)
                return Value.Error("#REF!");
            string key = CellKey(sheet, row, column);
            if (evaluationStack.Count > 0)
            {
                string parent = evaluationStack.Peek();
                HashSet<string> inputs;
                if (!dependencies.TryGetValue(parent, out inputs))
                    dependencies[parent] = inputs = new HashSet<string>();
                inputs.Add(key);
                HashSet<string> users;
                if (!dependents.TryGetValue(key, out users))
                    dependents[key] = users = new HashSet<string>();
                users.Add(parent);
            }
            Value cached;
            if (cache.TryGetValue(key, out cached))
                return cached;
            if (active.Contains(key))
                return Value.Error("#CYCLE!");

            active.Add(key);
            evaluationStack.Push(key);
            Value result;
            bool formula = false;
            try
            {
                string raw = string.Equals(sheet, currentSheet, StringComparison.OrdinalIgnoreCase) ?
                    readCell(row, column) : readOtherSheet == null ? null : readOtherSheet(sheet, row, column);
                if (raw == null) result = Value.Error("#REF!");
                else if (raw.StartsWith("=", StringComparison.Ordinal))
                {
                    formula = true;
                    result = new Parser(this, raw.Substring(1), sheet).Parse().Evaluate(this);
                }
                else if (raw.Length == 0)
                    result = Value.Blank();
                else if (IsKnownError(raw)) result = Value.Error(raw.ToUpperInvariant());
                else
                {
                    double number;
                    result = TryParseNumber(raw, out number)
                        ? Value.Numeric(number) : Value.String(raw);
                }
            }
            catch (FormatException)
            {
                result = Value.Error("#VALUE!");
            }
            catch (OverflowException)
            {
                result = Value.Error("#NUM!");
            }
            finally
            {
                evaluationStack.Pop();
                active.Remove(key);
            }
            cache[key] = result;
            if (formula) calculationChain.Add(key);
            return result;
        }

        private static bool IsKnownError(string text)
        {
            return text == "#N/A" || text == "#VALUE!" || text == "#REF!" ||
                text == "#DIV/0!" || text == "#NAME?" || text == "#NUM!" ||
                text == "#CYCLE!";
        }

        private static bool TryParseNumber(string text, out double number)
        {
            return Double.TryParse(text, NumberStyles.Float,
                CultureInfo.InvariantCulture, out number) ||
                Double.TryParse(text, NumberStyles.Float,
                    CultureInfo.CurrentCulture, out number);
        }

        private static Value AsNumber(Value value)
        {
            if (value.Kind == ValueKind.Error)
                return value;
            if (value.Kind == ValueKind.Number)
                return value;
            if (value.Kind == ValueKind.Blank)
                return Value.Numeric(0);
            double number;
            return value.Kind == ValueKind.Text && TryParseNumber(value.Text, out number)
                ? Value.Numeric(number) : Value.Error("#VALUE!");
        }

        private enum ValueKind { Blank, Number, Text, Error, Range }

        private sealed class Value
        {
            public ValueKind Kind;
            public double Number;
            public string Text = "";
            public List<Value> Items;
            public int Rows;
            public int Columns;

            public static Value Blank() { return new Value { Kind = ValueKind.Blank }; }
            public static Value Numeric(double number)
            {
                if (Double.IsNaN(number) || Double.IsInfinity(number))
                    return Error("#NUM!");
                return new Value { Kind = ValueKind.Number, Number = number };
            }
            public static Value String(string text)
            {
                return new Value { Kind = ValueKind.Text, Text = text };
            }
            public static Value Error(string text)
            {
                return new Value { Kind = ValueKind.Error, Text = text };
            }
            public static Value Range(List<Value> items, int rows, int columns)
            {
                return new Value { Kind = ValueKind.Range, Items = items,
                    Rows = rows, Columns = columns };
            }
        }

        private abstract class Node
        {
            public abstract Value Evaluate(FormulaEngine engine);
        }

        private sealed class LiteralNode : Node
        {
            private readonly Value value;
            public LiteralNode(Value value) { this.value = value; }
            public override Value Evaluate(FormulaEngine engine) { return value; }
        }

        private sealed class ReferenceNode : Node
        {
            private readonly string sheet;
            private readonly int row;
            private readonly int column;
            public ReferenceNode(string sheet, int row, int column)
            {
                this.sheet = sheet;
                this.row = row;
                this.column = column;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                return engine.EvaluateCell(sheet, row, column);
            }
        }

        private sealed class NameNode : Node
        {
            private readonly string name;
            public NameNode(string name) { this.name = name; }
            public override Value Evaluate(FormulaEngine engine)
            {
                FormulaNamedRange range = engine.resolveName == null ? null : engine.resolveName(name);
                if (range == null) return Value.Error("#NAME?");
                if (range.FirstRow == range.LastRow && range.FirstColumn == range.LastColumn)
                    return engine.EvaluateCell(range.Sheet, range.FirstRow, range.FirstColumn);
                return new RangeNode(range.Sheet, range.FirstRow, range.FirstColumn,
                    range.LastRow, range.LastColumn).Evaluate(engine);
            }
        }

        private sealed class RangeNode : Node
        {
            private readonly string sheet;
            private readonly int firstRow;
            private readonly int firstColumn;
            private readonly int lastRow;
            private readonly int lastColumn;
            public RangeNode(string sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
            {
                this.sheet = sheet;
                this.firstRow = firstRow;
                this.firstColumn = firstColumn;
                this.lastRow = lastRow;
                this.lastColumn = lastColumn;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                if (firstRow < 0 || lastRow < 0 || firstColumn < 0 || lastColumn < 0 ||
                    firstRow >= engine.rowCount || lastRow >= engine.rowCount ||
                    firstColumn >= engine.columnCount || lastColumn >= engine.columnCount)
                    return Value.Error("#REF!");
                var items = new List<Value>();
                for (int row = Math.Min(firstRow, lastRow);
                    row <= Math.Max(firstRow, lastRow); row++)
                {
                    for (int column = Math.Min(firstColumn, lastColumn);
                        column <= Math.Max(firstColumn, lastColumn); column++)
                        items.Add(engine.EvaluateCell(sheet, row, column));
                }
                return Value.Range(items, Math.Abs(lastRow - firstRow) + 1,
                    Math.Abs(lastColumn - firstColumn) + 1);
            }
        }

        private sealed class UnaryNode : Node
        {
            private readonly char operation;
            private readonly Node child;
            public UnaryNode(char operation, Node child)
            {
                this.operation = operation;
                this.child = child;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                Value number = AsNumber(child.Evaluate(engine));
                if (number.Kind == ValueKind.Error)
                    return number;
                return Value.Numeric(operation == '-' ? -number.Number : number.Number);
            }
        }

        private sealed class BinaryNode : Node
        {
            private readonly string operation;
            private readonly Node left;
            private readonly Node right;
            public BinaryNode(string operation, Node left, Node right)
            {
                this.operation = operation;
                this.left = left;
                this.right = right;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                Value a = left.Evaluate(engine);
                if (a.Kind == ValueKind.Error)
                    return a;
                Value b = right.Evaluate(engine);
                if (b.Kind == ValueKind.Error)
                    return b;
                if (a.Kind == ValueKind.Range || b.Kind == ValueKind.Range)
                    return Value.Error("#VALUE!");

                if (operation == "=" || operation == "<>" || operation == "<" ||
                    operation == "<=" || operation == ">" || operation == ">=")
                {
                    int comparison;
                    Value numericA = AsNumber(a);
                    Value numericB = AsNumber(b);
                    if (numericA.Kind == ValueKind.Number &&
                        numericB.Kind == ValueKind.Number)
                        comparison = numericA.Number.CompareTo(numericB.Number);
                    else
                        comparison = String.Compare(a.Text, b.Text,
                            StringComparison.OrdinalIgnoreCase);
                    bool matches = operation == "=" ? comparison == 0 :
                        operation == "<>" ? comparison != 0 :
                        operation == "<" ? comparison < 0 :
                        operation == "<=" ? comparison <= 0 :
                        operation == ">" ? comparison > 0 : comparison >= 0;
                    return Value.Numeric(matches ? 1 : 0);
                }

                a = AsNumber(a);
                b = AsNumber(b);
                if (a.Kind == ValueKind.Error)
                    return a;
                if (b.Kind == ValueKind.Error)
                    return b;
                if (operation == "+")
                    return Value.Numeric(a.Number + b.Number);
                if (operation == "-")
                    return Value.Numeric(a.Number - b.Number);
                if (operation == "*")
                    return Value.Numeric(a.Number * b.Number);
                if (operation == "/")
                    return b.Number == 0 ? Value.Error("#DIV/0!")
                        : Value.Numeric(a.Number / b.Number);
                if (operation == "^")
                    return Value.Numeric(Math.Pow(a.Number, b.Number));
                return Value.Error("#VALUE!");
            }
        }

        private sealed class FunctionNode : Node
        {
            private readonly string name;
            private readonly List<Node> arguments;
            public FunctionNode(string name, List<Node> arguments)
            {
                this.name = name.ToUpperInvariant();
                this.arguments = arguments;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                Value advanced = engine.EvaluateAdvanced(name, arguments);
                if (advanced != null) return advanced;
                if (name == "IF")
                {
                    if (arguments.Count != 3)
                        return Value.Error("#VALUE!");
                    Value condition = arguments[0].Evaluate(engine);
                    if (condition.Kind == ValueKind.Error)
                        return condition;
                    Value numeric = AsNumber(condition);
                    bool yes = numeric.Kind == ValueKind.Number
                        ? numeric.Number != 0 : condition.Text.Length != 0;
                    return arguments[yes ? 1 : 2].Evaluate(engine);
                }
                if (name == "NOT" && arguments.Count == 1)
                {
                    Value value = AsNumber(arguments[0].Evaluate(engine));
                    return value.Kind == ValueKind.Error ? value : Value.Numeric(value.Number == 0 ? 1 : 0);
                }
                if ((name == "AND" || name == "OR") && arguments.Count > 0)
                {
                    bool answer = name == "AND";
                    foreach (Node argument in arguments)
                    {
                        Value value = AsNumber(argument.Evaluate(engine));
                        if (value.Kind == ValueKind.Error) return value;
                        if (name == "AND") answer &= value.Number != 0;
                        else answer |= value.Number != 0;
                    }
                    return Value.Numeric(answer ? 1 : 0);
                }
                if (name == "ABS" || name == "SQRT" || name == "INT" ||
                    name == "ROUND" || name == "ROUNDUP" || name == "ROUNDDOWN" ||
                    name == "POWER" || name == "MOD")
                {
                    int required = name == "ROUND" || name == "ROUNDUP" || name == "ROUNDDOWN" || name == "POWER" || name == "MOD" ? 2 : 1;
                    if (arguments.Count != required) return Value.Error("#VALUE!");
                    Value a = AsNumber(arguments[0].Evaluate(engine));
                    if (a.Kind == ValueKind.Error) return a;
                    Value b = required == 2 ? AsNumber(arguments[1].Evaluate(engine)) : Value.Numeric(0);
                    if (b.Kind == ValueKind.Error) return b;
                    if (name == "ABS") return Value.Numeric(Math.Abs(a.Number));
                    if (name == "SQRT") return a.Number < 0 ? Value.Error("#NUM!") : Value.Numeric(Math.Sqrt(a.Number));
                    if (name == "INT") return Value.Numeric(Math.Floor(a.Number));
                    if (name == "POWER") return Value.Numeric(Math.Pow(a.Number, b.Number));
                    if (name == "MOD") return b.Number == 0 ? Value.Error("#DIV/0!") : Value.Numeric(a.Number - b.Number * Math.Floor(a.Number / b.Number));
                    int digits = (int)b.Number;
                    if (digits < -15 || digits > 15) return Value.Error("#NUM!");
                    double scale = Math.Pow(10, digits);
                    double scaled = a.Number * scale;
                    if (name == "ROUNDUP") return Value.Numeric(Math.Sign(scaled) * Math.Ceiling(Math.Abs(scaled)) / scale);
                    if (name == "ROUNDDOWN") return Value.Numeric(Math.Truncate(scaled) / scale);
                    return Value.Numeric(Math.Round(scaled, 0, MidpointRounding.AwayFromZero) / scale);
                }
                if (name == "LEN" || name == "UPPER" || name == "LOWER" || name == "TRIM" || name == "LEFT" || name == "RIGHT")
                {
                    int required = name == "LEFT" || name == "RIGHT" ? 2 : 1;
                    if (arguments.Count != required) return Value.Error("#VALUE!");
                    Value a = arguments[0].Evaluate(engine);
                    if (a.Kind == ValueKind.Error) return a;
                    string value = a.Kind == ValueKind.Number ? a.Number.ToString(CultureInfo.InvariantCulture) : a.Text;
                    if (name == "LEN") return Value.Numeric(value.Length);
                    if (name == "UPPER") return Value.String(value.ToUpperInvariant());
                    if (name == "LOWER") return Value.String(value.ToLowerInvariant());
                    if (name == "TRIM") return Value.String(Regex.Replace(value.Trim(), @"\s+", " "));
                    Value length = AsNumber(arguments[1].Evaluate(engine));
                    if (length.Kind == ValueKind.Error) return length;
                    int textLength = (int)length.Number;
                    if (textLength < 0) return Value.Error("#VALUE!");
                    textLength = Math.Min(textLength, value.Length);
                    return Value.String(name == "LEFT" ? value.Substring(0, textLength) : value.Substring(value.Length - textLength));
                }
                if (name == "CONCAT" || name == "CONCATENATE")
                {
                    var builder = new StringBuilder();
                    foreach (Node argument in arguments)
                    {
                        Value value = argument.Evaluate(engine);
                        if (value.Kind == ValueKind.Error) return value;
                        foreach (Value item in value.Kind == ValueKind.Range ? value.Items : new List<Value> { value })
                            builder.Append(item.Kind == ValueKind.Number ? item.Number.ToString(CultureInfo.InvariantCulture) : item.Text);
                    }
                    return Value.String(builder.ToString());
                }
                if (name == "COUNTIF" || name == "SUMIF")
                {
                    if (arguments.Count < 2 || arguments.Count > (name == "SUMIF" ? 3 : 2)) return Value.Error("#VALUE!");
                    Value source = arguments[0].Evaluate(engine);
                    Value criterion = arguments[1].Evaluate(engine);
                    if (source.Kind == ValueKind.Error) return source;
                    if (criterion.Kind == ValueKind.Error) return criterion;
                    var items = source.Kind == ValueKind.Range ? source.Items : new List<Value> { source };
                    Value sums = name == "SUMIF" && arguments.Count == 3 ? arguments[2].Evaluate(engine) : source;
                    if (sums.Kind == ValueKind.Error) return sums;
                    var sumItems = sums.Kind == ValueKind.Range ? sums.Items : new List<Value> { sums };
                    if (sumItems.Count != items.Count) return Value.Error("#VALUE!");
                    double matchedSum = 0;
                    for (int i = 0; i < items.Count; i++)
                    {
                        if (items[i].Kind == ValueKind.Error) return items[i];
                        if (!engine.CriteriaMatches(items[i], criterion)) continue;
                        if (name == "COUNTIF") matchedSum++;
                        else if (sumItems[i].Kind == ValueKind.Number) matchedSum += sumItems[i].Number;
                    }
                    return Value.Numeric(matchedSum);
                }
                if (name != "SUM" && name != "AVERAGE" &&
                    name != "MIN" && name != "MAX" && name != "COUNT" &&
                    name != "COUNTA" && name != "MEDIAN")
                    return Value.Error("#NAME?");

                var values = new List<Value>();
                foreach (Node argument in arguments)
                {
                    Value value = argument.Evaluate(engine);
                    if (value.Kind == ValueKind.Range)
                        values.AddRange(value.Items);
                    else
                        values.Add(value);
                }
                int count = 0;
                double total = 0;
                double minimum = Double.PositiveInfinity;
                double maximum = Double.NegativeInfinity;
                var numericValues = new List<double>();
                int nonempty = 0;
                foreach (Value value in values)
                {
                    if (value.Kind == ValueKind.Error)
                        return value;
                    if (value.Kind != ValueKind.Blank && (value.Kind != ValueKind.Text || value.Text.Length > 0)) nonempty++;
                    if (value.Kind != ValueKind.Number)
                        continue;
                    count++;
                    numericValues.Add(value.Number);
                    total += value.Number;
                    minimum = Math.Min(minimum, value.Number);
                    maximum = Math.Max(maximum, value.Number);
                }
                if (name == "COUNT")
                    return Value.Numeric(count);
                if (name == "COUNTA") return Value.Numeric(nonempty);
                if (name == "MEDIAN")
                {
                    if (count == 0) return Value.Error("#NUM!");
                    numericValues.Sort();
                    return Value.Numeric(count % 2 == 1 ? numericValues[count / 2] :
                        (numericValues[count / 2 - 1] + numericValues[count / 2]) / 2);
                }
                if (name == "SUM")
                    return Value.Numeric(total);
                if (name == "AVERAGE")
                    return count == 0 ? Value.Error("#DIV/0!")
                        : Value.Numeric(total / count);
                if (name == "MIN")
                    return Value.Numeric(count == 0 ? 0 : minimum);
                return Value.Numeric(count == 0 ? 0 : maximum);
            }

            private static bool MatchesCriterion(Value candidate, Value criterion)
            {
                string text = criterion.Kind == ValueKind.Number ? criterion.Number.ToString(CultureInfo.InvariantCulture) : criterion.Text;
                string operation = "=";
                foreach (string prefix in new[] { ">=", "<=", "<>", ">", "<", "=" })
                    if (text.StartsWith(prefix, StringComparison.Ordinal))
                    { operation = prefix; text = text.Substring(prefix.Length); break; }
                double expected;
                Value number = AsNumber(candidate);
                int comparison = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out expected) && number.Kind == ValueKind.Number
                    ? number.Number.CompareTo(expected)
                    : string.Compare(candidate.Kind == ValueKind.Number ? candidate.Number.ToString(CultureInfo.InvariantCulture) : candidate.Text,
                        text, StringComparison.OrdinalIgnoreCase);
                return operation == "=" ? comparison == 0 : operation == "<>" ? comparison != 0 :
                    operation == ">" ? comparison > 0 : operation == "<" ? comparison < 0 :
                    operation == ">=" ? comparison >= 0 : comparison <= 0;
            }
        }

        private sealed class Parser
        {
            private readonly FormulaEngine engine;
            private readonly string source;
            private readonly string sheet;
            private int position;

            public Parser(FormulaEngine engine, string source, string sheet)
            {
                this.engine = engine;
                this.source = source;
                this.sheet = sheet;
            }

            public Node Parse()
            {
                Node result = Comparison();
                SkipSpaces();
                if (position != source.Length)
                    throw new FormatException();
                return result;
            }

            private Node Comparison()
            {
                Node left = Additive();
                string operation = null;
                if (Take("<>")) operation = "<>";
                else if (Take("<=")) operation = "<=";
                else if (Take(">=")) operation = ">=";
                else if (Take("=")) operation = "=";
                else if (Take("<")) operation = "<";
                else if (Take(">")) operation = ">";
                return operation == null ? left :
                    new BinaryNode(operation, left, Additive());
            }

            private Node Additive()
            {
                Node left = Multiplicative();
                while (true)
                {
                    if (Take("+")) left = new BinaryNode("+", left, Multiplicative());
                    else if (Take("-")) left = new BinaryNode("-", left, Multiplicative());
                    else return left;
                }
            }

            private Node Multiplicative()
            {
                Node left = Power();
                while (true)
                {
                    if (Take("*")) left = new BinaryNode("*", left, Power());
                    else if (Take("/")) left = new BinaryNode("/", left, Power());
                    else return left;
                }
            }

            private Node Power()
            {
                Node left = Unary();
                return Take("^") ? new BinaryNode("^", left, Power()) : left;
            }

            private Node Unary()
            {
                if (Take("+")) return new UnaryNode('+', Unary());
                if (Take("-")) return new UnaryNode('-', Unary());
                return Primary();
            }

            private Node Primary()
            {
                SkipSpaces();
                if (Take("("))
                {
                    Node expression = Comparison();
                    Require(")");
                    return expression;
                }
                if (position >= source.Length)
                    throw new FormatException();
                if (source[position] == '"')
                    return new LiteralNode(Value.String(ReadString()));
                if (source[position] == '#')
                    return new LiteralNode(Value.Error(ReadError()));
                if (Char.IsDigit(source[position]) || source[position] == '.')
                    return new LiteralNode(Value.Numeric(ReadNumber()));

                string explicitSheet = null;
                if (source[position] == '\'')
                {
                    explicitSheet = ReadSheetName();
                    Require("!");
                }
                string word = ReadWord();
                if (word.Length == 0)
                    throw new FormatException();
                if (explicitSheet == null && Take("!"))
                {
                    explicitSheet = word;
                    word = ReadWord();
                }
                if (Take("("))
                {
                    var arguments = new List<Node>();
                    if (!Take(")"))
                    {
                        do { arguments.Add(Comparison()); }
                        while (Take(",") || Take(";"));
                        Require(")");
                    }
                    return new FunctionNode(word, arguments);
                }
                if (String.Equals(word, "TRUE", StringComparison.OrdinalIgnoreCase))
                    return new LiteralNode(Value.Numeric(1));
                if (String.Equals(word, "FALSE", StringComparison.OrdinalIgnoreCase))
                    return new LiteralNode(Value.Numeric(0));

                int row, column;
                if (!TryAddress(word, out row, out column))
                    return new NameNode(word);
                if (Take(":"))
                {
                    SkipSpaces();
                    string endSheet = explicitSheet ?? sheet;
                    string end;
                    if (position < source.Length && source[position] == '\'')
                    {
                        endSheet = ReadSheetName();
                        Require("!");
                        end = ReadWord();
                    }
                    else
                    {
                        end = ReadWord();
                        if (Take("!"))
                        { endSheet = end; end = ReadWord(); }
                    }
                    int lastRow, lastColumn;
                    if (!TryAddress(end, out lastRow, out lastColumn))
                        throw new FormatException();
                    if (!string.Equals(endSheet, explicitSheet ?? sheet,
                        StringComparison.OrdinalIgnoreCase))
                        return new LiteralNode(Value.Error("#REF!"));
                    return new RangeNode(explicitSheet ?? sheet, row, column, lastRow, lastColumn);
                }
                return new ReferenceNode(explicitSheet ?? sheet, row, column);
            }

            private string ReadSheetName()
            {
                position++;
                var name = new StringBuilder();
                while (position < source.Length)
                {
                    char c = source[position++];
                    if (c == '\'')
                    {
                        if (position < source.Length && source[position] == '\'')
                        { name.Append('\''); position++; }
                        else return name.ToString();
                    }
                    else name.Append(c);
                }
                throw new FormatException();
            }

            private string ReadWord()
            {
                SkipSpaces();
                int start = position;
                while (position < source.Length &&
                    (Char.IsLetterOrDigit(source[position]) ||
                     source[position] == '$' || source[position] == '_'))
                    position++;
                return source.Substring(start, position - start);
            }

            private string ReadString()
            {
                position++;
                var text = new StringBuilder();
                while (position < source.Length)
                {
                    char current = source[position++];
                    if (current == '"')
                    {
                        if (position < source.Length && source[position] == '"')
                        {
                            text.Append('"');
                            position++;
                        }
                        else
                            return text.ToString();
                    }
                    else
                        text.Append(current);
                }
                throw new FormatException();
            }

            private string ReadError()
            {
                int start = position;
                while (position < source.Length &&
                    !Char.IsWhiteSpace(source[position]) &&
                    "+-*^(),;<>=".IndexOf(source[position]) < 0)
                    position++;
                return source.Substring(start, position - start);
            }

            private double ReadNumber()
            {
                int start = position;
                while (position < source.Length &&
                    (Char.IsDigit(source[position]) || source[position] == '.'))
                    position++;
                double number;
                if (!Double.TryParse(source.Substring(start, position - start),
                    NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                    throw new FormatException();
                return number;
            }

            private bool Take(string token)
            {
                SkipSpaces();
                if (position + token.Length > source.Length ||
                    String.Compare(source, position, token, 0, token.Length,
                        StringComparison.Ordinal) != 0)
                    return false;
                position += token.Length;
                return true;
            }

            private void Require(string token)
            {
                if (!Take(token))
                    throw new FormatException();
            }

            private void SkipSpaces()
            {
                while (position < source.Length && Char.IsWhiteSpace(source[position]))
                    position++;
            }

            private bool TryAddress(string text, out int row, out int column)
            {
                row = -1;
                column = -1;
                int index = 0;
                if (index < text.Length && text[index] == '$') index++;
                int letters = index;
                int value = 0;
                while (index < text.Length && Char.IsLetter(text[index]))
                {
                    value = value * 26 + (Char.ToUpperInvariant(text[index]) - 'A' + 1);
                    index++;
                }
                if (index == letters)
                    return false;
                if (index < text.Length && text[index] == '$') index++;
                int digits = index;
                int number;
                if (digits == text.Length ||
                    !Int32.TryParse(text.Substring(digits), out number))
                    return false;
                column = value - 1;
                row = number - 1;
                return column >= 0;
            }
        }

        public static string ShiftReferences(string formula, int rowOffset,
            int columnOffset, int maxRows, int maxColumns)
        {
            if (String.IsNullOrEmpty(formula) || formula[0] != '=')
                return formula;
            var result = new StringBuilder();
            int index = 0;
            while (index < formula.Length)
            {
                if (formula[index] == '"')
                {
                    result.Append(formula[index++]);
                    while (index < formula.Length)
                    {
                        char current = formula[index++];
                        result.Append(current);
                        if (current == '"')
                        {
                            if (index < formula.Length && formula[index] == '"')
                                result.Append(formula[index++]);
                            else
                                break;
                        }
                    }
                    continue;
                }
                bool boundary = index == 0 ||
                    !(Char.IsLetterOrDigit(formula[index - 1]) ||
                      formula[index - 1] == '_');
                int start = index;
                int cursor = index;
                bool fixedColumn = cursor < formula.Length && formula[cursor] == '$';
                if (fixedColumn) cursor++;
                int letters = cursor;
                int column = 0;
                while (cursor < formula.Length && Char.IsLetter(formula[cursor]))
                {
                    column = column * 26 +
                        (Char.ToUpperInvariant(formula[cursor]) - 'A' + 1);
                    cursor++;
                }
                bool fixedRow = cursor < formula.Length && formula[cursor] == '$';
                if (fixedRow) cursor++;
                int digits = cursor;
                int row = 0;
                while (cursor < formula.Length && Char.IsDigit(formula[cursor]))
                    cursor++;
                bool valid = boundary && letters < digits &&
                    digits < cursor &&
                    Int32.TryParse(formula.Substring(digits, cursor - digits), out row) &&
                    (cursor == formula.Length ||
                     !(Char.IsLetterOrDigit(formula[cursor]) ||
                       formula[cursor] == '_'));
                if (valid)
                {
                    int nextColumn = column + (fixedColumn ? 0 : columnOffset);
                    int nextRow = row + (fixedRow ? 0 : rowOffset);
                    if (nextColumn < 1 || nextColumn > maxColumns ||
                        nextRow < 1 || nextRow > maxRows)
                        result.Append("#REF!");
                    else
                        result.Append(fixedColumn ? "$" : "")
                            .Append(ColumnName(nextColumn))
                            .Append(fixedRow ? "$" : "")
                            .Append(nextRow);
                    index = cursor;
                }
                else
                {
                    result.Append(formula[start]);
                    index++;
                }
            }
            return result.ToString();
        }

        private static readonly Regex StructureReference = new Regex(
            @"(?<![A-Za-z0-9_])(\$?[A-Za-z]+\$?[1-9][0-9]*)(?::(\$?[A-Za-z]+\$?[1-9][0-9]*))?(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        public static string ShiftStructureReferences(string formula, bool rows,
            int index, bool insert, int maxRows, int maxColumns)
        {
            if (String.IsNullOrEmpty(formula) || formula[0] != '=')
                return formula;
            var result = new StringBuilder();
            int start = 0;
            int cursor = 0;
            while (cursor < formula.Length)
            {
                if (formula[cursor] != '"')
                {
                    cursor++;
                    continue;
                }
                AppendShiftedSegment(result, formula.Substring(start, cursor - start),
                    rows, index, insert, maxRows, maxColumns);
                start = cursor++;
                while (cursor < formula.Length)
                {
                    if (formula[cursor++] != '"')
                        continue;
                    if (cursor < formula.Length && formula[cursor] == '"')
                        cursor++;
                    else
                        break;
                }
                result.Append(formula, start, cursor - start);
                start = cursor;
            }
            AppendShiftedSegment(result, formula.Substring(start),
                rows, index, insert, maxRows, maxColumns);
            return result.ToString();
        }

        private static void AppendShiftedSegment(StringBuilder result, string segment,
            bool rows, int index, bool insert, int maxRows, int maxColumns)
        {
            result.Append(StructureReference.Replace(segment, delegate(Match match)
            {
                CellReference first;
                if (!TryParseReference(match.Groups[1].Value, out first))
                    return match.Value;
                bool range = match.Groups[2].Success;
                CellReference last = new CellReference();
                if (range && !TryParseReference(match.Groups[2].Value, out last))
                    return match.Value;
                int axis = rows ? first.Row : first.Column;
                int other = rows ? last.Row : last.Column;
                if (!insert && range && axis == index + 1 && other == index + 1)
                    return "#REF!";
                int firstNext = MoveStructureIndex(axis, range ? other : axis,
                    index + 1, insert, range);
                int lastNext = range ? MoveStructureIndex(other, axis,
                    index + 1, insert, true) : 0;
                if (firstNext < 1 || firstNext > (rows ? maxRows : maxColumns) ||
                    (range && (lastNext < 1 ||
                    lastNext > (rows ? maxRows : maxColumns))))
                    return "#REF!";
                if (rows)
                {
                    first.Row = firstNext;
                    last.Row = lastNext;
                }
                else
                {
                    first.Column = firstNext;
                    last.Column = lastNext;
                }
                return FormatReference(first) +
                    (range ? ":" + FormatReference(last) : "");
            }));
        }

        private static int MoveStructureIndex(int value, int other, int at,
            bool insert, bool range)
        {
            if (insert)
                return value >= at ? value + 1 : value;
            if (value > at)
                return value - 1;
            if (value < at)
                return value;
            if (!range)
                return 0;
            return other > at ? at : at - 1;
        }

        public static string MoveStructureReferences(string formula, bool rows,
            int sourceIndex, int targetIndex)
        {
            if (String.IsNullOrEmpty(formula) || formula[0] != '=' ||
                sourceIndex == targetIndex)
                return formula;
            var result = new StringBuilder();
            int start = 0;
            int cursor = 0;
            while (cursor < formula.Length)
            {
                if (formula[cursor] != '"')
                {
                    cursor++;
                    continue;
                }
                AppendMovedSegment(result, formula.Substring(start, cursor - start),
                    rows, sourceIndex, targetIndex);
                start = cursor++;
                while (cursor < formula.Length)
                {
                    if (formula[cursor++] != '"')
                        continue;
                    if (cursor < formula.Length && formula[cursor] == '"')
                        cursor++;
                    else
                        break;
                }
                result.Append(formula, start, cursor - start);
                start = cursor;
            }
            AppendMovedSegment(result, formula.Substring(start),
                rows, sourceIndex, targetIndex);
            return result.ToString();
        }

        private static void AppendMovedSegment(StringBuilder result, string segment,
            bool rows, int sourceIndex, int targetIndex)
        {
            result.Append(StructureReference.Replace(segment, delegate(Match match)
            {
                CellReference first;
                if (!TryParseReference(match.Groups[1].Value, out first))
                    return match.Value;
                bool range = match.Groups[2].Success;
                CellReference last = new CellReference();
                if (range && !TryParseReference(match.Groups[2].Value, out last))
                    return match.Value;
                int firstAxis = rows ? first.Row : first.Column;
                int lastAxis = rows ? last.Row : last.Column;
                int firstNext = MapMovedIndex(firstAxis, sourceIndex, targetIndex);
                int lastNext = 0;
                if (range)
                {
                    int minimum = Int32.MaxValue;
                    int maximum = 0;
                    for (int value = Math.Min(firstAxis, lastAxis);
                        value <= Math.Max(firstAxis, lastAxis); value++)
                    {
                        int moved = MapMovedIndex(value, sourceIndex, targetIndex);
                        minimum = Math.Min(minimum, moved);
                        maximum = Math.Max(maximum, moved);
                    }
                    firstNext = firstAxis <= lastAxis ? minimum : maximum;
                    lastNext = firstAxis <= lastAxis ? maximum : minimum;
                }
                if (rows)
                {
                    first.Row = firstNext;
                    last.Row = lastNext;
                }
                else
                {
                    first.Column = firstNext;
                    last.Column = lastNext;
                }
                return FormatReference(first) +
                    (range ? ":" + FormatReference(last) : "");
            }));
        }

        private static int MapMovedIndex(int oneBased, int sourceIndex, int targetIndex)
        {
            int source = sourceIndex + 1;
            int target = targetIndex + 1;
            if (oneBased == source)
                return target;
            if (source < target && oneBased > source && oneBased <= target)
                return oneBased - 1;
            if (source > target && oneBased >= target && oneBased < source)
                return oneBased + 1;
            return oneBased;
        }

        private struct CellReference
        {
            public int Row;
            public int Column;
            public bool FixedRow;
            public bool FixedColumn;
        }

        private static bool TryParseReference(string source, out CellReference reference)
        {
            reference = new CellReference();
            int cursor = 0;
            if (source[cursor] == '$')
            {
                reference.FixedColumn = true;
                cursor++;
            }
            int letters = cursor;
            while (cursor < source.Length && Char.IsLetter(source[cursor]))
            {
                reference.Column = reference.Column * 26 +
                    (Char.ToUpperInvariant(source[cursor]) - 'A' + 1);
                cursor++;
            }
            if (cursor == letters)
                return false;
            if (cursor < source.Length && source[cursor] == '$')
            {
                reference.FixedRow = true;
                cursor++;
            }
            return Int32.TryParse(source.Substring(cursor), out reference.Row);
        }

        private static string FormatReference(CellReference reference)
        {
            return (reference.FixedColumn ? "$" : "") +
                ColumnName(reference.Column) +
                (reference.FixedRow ? "$" : "") + reference.Row;
        }

        private static string ColumnName(int column)
        {
            var result = new StringBuilder();
            while (column > 0)
            {
                column--;
                result.Insert(0, (char)('A' + column % 26));
                column /= 26;
            }
            return result.ToString();
        }
    }
}

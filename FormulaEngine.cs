using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DinkCel
{
    internal sealed class FormulaEngine
    {
        private readonly Func<int, int, string> readCell;
        private readonly int rowCount;
        private readonly int columnCount;
        private readonly Dictionary<int, Value> cache = new Dictionary<int, Value>();
        private readonly HashSet<int> active = new HashSet<int>();

        public FormulaEngine(Func<int, int, string> readCell, int rowCount, int columnCount)
        {
            this.readCell = readCell;
            this.rowCount = rowCount;
            this.columnCount = columnCount;
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
            return result.Text;
        }

        private Value EvaluateCell(int row, int column)
        {
            if (row < 0 || row >= rowCount || column < 0 || column >= columnCount)
                return Value.Error("#REF!");
            int key = row * columnCount + column;
            Value cached;
            if (cache.TryGetValue(key, out cached))
                return cached;
            if (active.Contains(key))
                return Value.Error("#CYCLE!");

            active.Add(key);
            Value result;
            try
            {
                string raw = readCell(row, column) ?? "";
                if (raw.StartsWith("=", StringComparison.Ordinal))
                    result = new Parser(this, raw.Substring(1)).Parse().Evaluate(this);
                else if (raw.Length == 0)
                    result = Value.Blank();
                else
                {
                    double number;
                    result = TryParseNumber(raw, out number)
                        ? Value.Numeric(number) : Value.String(raw);
                }
            }
            catch (FormatException)
            {
                result = Value.Error("#ERROR!");
            }
            catch (OverflowException)
            {
                result = Value.Error("#NUM!");
            }
            finally
            {
                active.Remove(key);
            }
            cache[key] = result;
            return result;
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

            public static Value Blank() { return new Value { Kind = ValueKind.Blank }; }
            public static Value Numeric(double number)
            {
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
            public static Value Range(List<Value> items)
            {
                return new Value { Kind = ValueKind.Range, Items = items };
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
            private readonly int row;
            private readonly int column;
            public ReferenceNode(int row, int column)
            {
                this.row = row;
                this.column = column;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                return engine.EvaluateCell(row, column);
            }
        }

        private sealed class RangeNode : Node
        {
            private readonly int firstRow;
            private readonly int firstColumn;
            private readonly int lastRow;
            private readonly int lastColumn;
            public RangeNode(int firstRow, int firstColumn, int lastRow, int lastColumn)
            {
                this.firstRow = firstRow;
                this.firstColumn = firstColumn;
                this.lastRow = lastRow;
                this.lastColumn = lastColumn;
            }
            public override Value Evaluate(FormulaEngine engine)
            {
                var items = new List<Value>();
                for (int row = Math.Min(firstRow, lastRow);
                    row <= Math.Max(firstRow, lastRow); row++)
                {
                    for (int column = Math.Min(firstColumn, lastColumn);
                        column <= Math.Max(firstColumn, lastColumn); column++)
                        items.Add(engine.EvaluateCell(row, column));
                }
                return Value.Range(items);
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
                return Value.Error("#ERROR!");
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
                if (name == "IF")
                {
                    if (arguments.Count != 3)
                        return Value.Error("#ERROR!");
                    Value condition = arguments[0].Evaluate(engine);
                    if (condition.Kind == ValueKind.Error)
                        return condition;
                    Value numeric = AsNumber(condition);
                    bool yes = numeric.Kind == ValueKind.Number
                        ? numeric.Number != 0 : condition.Text.Length != 0;
                    return arguments[yes ? 1 : 2].Evaluate(engine);
                }
                if (name != "SUM" && name != "AVERAGE" &&
                    name != "MIN" && name != "MAX" && name != "COUNT")
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
                foreach (Value value in values)
                {
                    if (value.Kind == ValueKind.Error)
                        return value;
                    if (value.Kind != ValueKind.Number)
                        continue;
                    count++;
                    total += value.Number;
                    minimum = Math.Min(minimum, value.Number);
                    maximum = Math.Max(maximum, value.Number);
                }
                if (name == "COUNT")
                    return Value.Numeric(count);
                if (name == "SUM")
                    return Value.Numeric(total);
                if (name == "AVERAGE")
                    return count == 0 ? Value.Error("#DIV/0!")
                        : Value.Numeric(total / count);
                if (name == "MIN")
                    return Value.Numeric(count == 0 ? 0 : minimum);
                return Value.Numeric(count == 0 ? 0 : maximum);
            }
        }

        private sealed class Parser
        {
            private readonly FormulaEngine engine;
            private readonly string source;
            private int position;

            public Parser(FormulaEngine engine, string source)
            {
                this.engine = engine;
                this.source = source;
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

                string word = ReadWord();
                if (word.Length == 0)
                    throw new FormatException();
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
                    return new LiteralNode(Value.Error("#NAME?"));
                if (Take(":"))
                {
                    string end = ReadWord();
                    int lastRow, lastColumn;
                    if (!TryAddress(end, out lastRow, out lastColumn))
                        throw new FormatException();
                    return new RangeNode(row, column, lastRow, lastColumn);
                }
                return new ReferenceNode(row, column);
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
                return column >= 0 && column < engine.columnCount &&
                    row >= 0 && row < engine.rowCount;
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

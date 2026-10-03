using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DinkCel
{
    internal sealed partial class FormulaEngine
    {
        private static Value Scalar(Value value)
        {
            return value.Kind == ValueKind.Range ?
                value.Items.Count == 1 ? value.Items[0] : Value.Error("#VALUE!") : value;
        }

        private static List<Value> Items(Value value)
        {
            return value.Kind == ValueKind.Range ? value.Items : new List<Value> { value };
        }

        private static int Rows(Value value) { return value.Kind == ValueKind.Range ? value.Rows : 1; }
        private static int Columns(Value value) { return value.Kind == ValueKind.Range ? value.Columns : 1; }

        private static string TextOf(Value value)
        {
            return value.Kind == ValueKind.Number ? value.Number.ToString("G15", CultureInfo.InvariantCulture) :
                value.Kind == ValueKind.Blank ? "" : value.Text;
        }

        private static Value NumberOf(Node node, FormulaEngine engine)
        {
            return AsNumber(Scalar(node.Evaluate(engine)));
        }

        private static bool SameShape(Value first, Value second)
        { return Rows(first) == Rows(second) && Columns(first) == Columns(second); }

        private static int CompareLookup(Value a, Value b)
        {
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Number)
                return a.Number.CompareTo(b.Number);
            return string.Compare(TextOf(a), TextOf(b), StringComparison.OrdinalIgnoreCase);
        }

        private static bool LookupEqual(Value a, Value b)
        {
            if (a.Kind == ValueKind.Number || b.Kind == ValueKind.Number)
                return a.Kind == ValueKind.Number && b.Kind == ValueKind.Number && a.Number == b.Number;
            if (a.Kind == ValueKind.Blank || b.Kind == ValueKind.Blank)
                return a.Kind == b.Kind || TextOf(a).Length == 0 && TextOf(b).Length == 0;
            return string.Equals(TextOf(a), TextOf(b), StringComparison.OrdinalIgnoreCase);
        }

        private static string WildcardPattern(string pattern)
        {
            var result = new StringBuilder("^");
            for (int index = 0; index < pattern.Length; index++)
            {
                char current = pattern[index];
                if (current == '~' && index + 1 < pattern.Length)
                    result.Append(Regex.Escape(pattern[++index].ToString()));
                else if (current == '*') result.Append(".*");
                else if (current == '?') result.Append('.');
                else result.Append(Regex.Escape(current.ToString()));
            }
            return result.Append('$').ToString();
        }

        private static bool WildcardMatch(string value, string pattern)
        {
            return Regex.IsMatch(value, WildcardPattern(pattern), RegexOptions.IgnoreCase |
                RegexOptions.Singleline | RegexOptions.CultureInvariant);
        }

        private static int FindLookup(Value lookup, List<Value> values, int matchMode, int searchMode)
        {
            bool backward = searchMode == -1 || searchMode == -2;
            int start = backward ? values.Count - 1 : 0;
            int end = backward ? -1 : values.Count;
            int step = backward ? -1 : 1;
            int nearest = -1;
            for (int index = start; index != end; index += step)
            {
                Value candidate = values[index];
                if (candidate.Kind == ValueKind.Error) continue;
                if (matchMode == 2)
                {
                    if (WildcardMatch(TextOf(candidate), TextOf(lookup))) return index;
                    continue;
                }
                if (LookupEqual(candidate, lookup)) return index;
                if (matchMode != -1 && matchMode != 1) continue;
                if (candidate.Kind != lookup.Kind || candidate.Kind == ValueKind.Blank) continue;
                int comparison = CompareLookup(candidate, lookup);
                if (matchMode == -1 && comparison <= 0 || matchMode == 1 && comparison >= 0)
                {
                    if (nearest < 0 || matchMode == -1 && CompareLookup(candidate, values[nearest]) > 0 ||
                        matchMode == 1 && CompareLookup(candidate, values[nearest]) < 0)
                        nearest = index;
                }
            }
            return nearest;
        }

        private Value EvaluateLookup(string name, List<Node> arguments)
        {
            if (name == "INDEX")
            {
                if (arguments.Count < 2 || arguments.Count > 3) return Value.Error("#VALUE!");
                Value array = arguments[0].Evaluate(this);
                if (array.Kind == ValueKind.Error) return array;
                Value rowValue = NumberOf(arguments[1], this);
                if (rowValue.Kind == ValueKind.Error) return rowValue;
                Value columnValue = arguments.Count == 3 ? NumberOf(arguments[2], this) : Value.Numeric(1);
                if (columnValue.Kind == ValueKind.Error) return columnValue;
                int row = (int)rowValue.Number, column = (int)columnValue.Number;
                if (arguments.Count == 2 && Rows(array) == 1)
                { column = row; row = 1; }
                if (row < 1 || column < 1 || row > Rows(array) || column > Columns(array))
                    return Value.Error("#REF!");
                return Items(array)[(row - 1) * Columns(array) + column - 1];
            }
            if (name == "MATCH")
            {
                if (arguments.Count < 2 || arguments.Count > 3) return Value.Error("#VALUE!");
                Value lookup = Scalar(arguments[0].Evaluate(this));
                Value array = arguments[1].Evaluate(this);
                if (lookup.Kind == ValueKind.Error) return lookup;
                if (array.Kind == ValueKind.Error) return array;
                if (Rows(array) > 1 && Columns(array) > 1) return Value.Error("#VALUE!");
                Value type = arguments.Count == 3 ? NumberOf(arguments[2], this) : Value.Numeric(1);
                if (type.Kind == ValueKind.Error) return type;
                int mode = (int)type.Number;
                if (mode != -1 && mode != 0 && mode != 1) return Value.Error("#VALUE!");
                int index = FindLookup(lookup, Items(array), mode == 1 ? -1 : mode == -1 ? 1 : 0, 1);
                return index < 0 ? Value.Error("#N/A") : Value.Numeric(index + 1);
            }
            if (name == "VLOOKUP" || name == "HLOOKUP")
            {
                if (arguments.Count < 3 || arguments.Count > 4) return Value.Error("#VALUE!");
                Value lookup = Scalar(arguments[0].Evaluate(this));
                Value table = arguments[1].Evaluate(this);
                Value indexValue = NumberOf(arguments[2], this);
                if (lookup.Kind == ValueKind.Error) return lookup;
                if (table.Kind == ValueKind.Error) return table;
                if (indexValue.Kind == ValueKind.Error) return indexValue;
                bool vertical = name == "VLOOKUP";
                int index = (int)indexValue.Number;
                if (index < 1 || index > (vertical ? Columns(table) : Rows(table)))
                    return Value.Error("#REF!");
                Value approximate = arguments.Count == 4 ? NumberOf(arguments[3], this) : Value.Numeric(1);
                if (approximate.Kind == ValueKind.Error) return approximate;
                var lookupCells = new List<Value>();
                List<Value> tableCells = Items(table);
                for (int i = 0; i < (vertical ? Rows(table) : Columns(table)); i++)
                    lookupCells.Add(tableCells[vertical ? i * Columns(table) : i]);
                int found = FindLookup(lookup, lookupCells, approximate.Number == 0 ? 0 : -1, 1);
                if (found < 0) return Value.Error("#N/A");
                return tableCells[vertical ? found * Columns(table) + index - 1 :
                    (index - 1) * Columns(table) + found];
            }
            if (arguments.Count < 3 || arguments.Count > 6) return Value.Error("#VALUE!");
            Value find = Scalar(arguments[0].Evaluate(this));
            Value lookupArray = arguments[1].Evaluate(this);
            Value returnArray = arguments[2].Evaluate(this);
            if (find.Kind == ValueKind.Error) return find;
            if (lookupArray.Kind == ValueKind.Error) return lookupArray;
            if (returnArray.Kind == ValueKind.Error) return returnArray;
            if (Rows(lookupArray) > 1 && Columns(lookupArray) > 1 ||
                !SameShape(lookupArray, returnArray)) return Value.Error("#VALUE!");
            Value matchValue = arguments.Count >= 5 ? NumberOf(arguments[4], this) : Value.Numeric(0);
            Value searchValue = arguments.Count >= 6 ? NumberOf(arguments[5], this) : Value.Numeric(1);
            if (matchValue.Kind == ValueKind.Error) return matchValue;
            if (searchValue.Kind == ValueKind.Error) return searchValue;
            int matchMode = (int)matchValue.Number, searchMode = (int)searchValue.Number;
            if (matchMode != 0 && matchMode != -1 && matchMode != 1 && matchMode != 2 ||
                searchMode != 1 && searchMode != -1 && searchMode != 2 && searchMode != -2)
                return Value.Error("#VALUE!");
            int position = FindLookup(find, Items(lookupArray), matchMode, searchMode);
            return position >= 0 ? Items(returnArray)[position] :
                arguments.Count >= 4 ? arguments[3].Evaluate(this) : Value.Error("#N/A");
        }

        private bool CriteriaMatches(Value candidate, Value criterion)
        {
            string text = TextOf(criterion);
            string operation = "=";
            foreach (string prefix in new[] { ">=", "<=", "<>", ">", "<", "=" })
                if (text.StartsWith(prefix, StringComparison.Ordinal))
                { operation = prefix; text = text.Substring(prefix.Length); break; }
            if (text.Length == 0 && (operation == "=" || operation == "<>"))
            {
                bool blank = candidate.Kind == ValueKind.Blank ||
                    candidate.Kind == ValueKind.Text && candidate.Text.Length == 0;
                return operation == "=" ? blank : !blank;
            }
            if ((text.Contains("*") || text.Contains("?")) &&
                (operation == "=" || operation == "<>"))
            {
                bool found = WildcardMatch(TextOf(candidate), text);
                return operation == "=" ? found : !found;
            }
            double expected;
            int comparison;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out expected) &&
                candidate.Kind == ValueKind.Number)
                comparison = candidate.Number.CompareTo(expected);
            else comparison = string.Compare(TextOf(candidate), text, StringComparison.OrdinalIgnoreCase);
            return operation == "=" ? comparison == 0 : operation == "<>" ? comparison != 0 :
                operation == ">" ? comparison > 0 : operation == "<" ? comparison < 0 :
                operation == ">=" ? comparison >= 0 : comparison <= 0;
        }

        private Value EvaluateConditionalAggregate(string name, List<Node> arguments)
        {
            bool count = name == "COUNTIF" || name == "COUNTIFS";
            bool single = name == "COUNTIF" || name == "SUMIF" || name == "AVERAGEIF";
            if (single && (arguments.Count < 2 || arguments.Count > (count ? 2 : 3)) ||
                !single && (arguments.Count < (count ? 2 : 3) ||
                    arguments.Count % 2 != (count ? 0 : 1)))
                return Value.Error("#VALUE!");
            Value values = count ? null : arguments[single && arguments.Count == 2 ? 0 :
                single && name != "COUNTIF" ? 2 : 0].Evaluate(this);
            if (values != null && values.Kind == ValueKind.Error) return values;
            var criteriaRanges = new List<List<Value>>();
            var criteria = new List<Value>();
            int start = single ? 0 : count ? 0 : 1;
            int pairCount = single ? 1 : (arguments.Count - start) / 2;
            int targetRows = values == null ? -1 : Rows(values);
            int targetColumns = values == null ? -1 : Columns(values);
            for (int pair = 0; pair < pairCount; pair++)
            {
                Value range = arguments[start + pair * 2].Evaluate(this);
                Value criterion = Scalar(arguments[start + pair * 2 + 1].Evaluate(this));
                if (range.Kind == ValueKind.Error) return range;
                if (criterion.Kind == ValueKind.Error) return criterion;
                if (targetRows < 0) { targetRows = Rows(range); targetColumns = Columns(range); }
                if (Rows(range) != targetRows || Columns(range) != targetColumns)
                    return Value.Error("#VALUE!");
                criteriaRanges.Add(Items(range)); criteria.Add(criterion);
            }
            if (values == null) values = Value.Range(criteriaRanges[0], targetRows, targetColumns);
            if (!SameShape(values, Value.Range(criteriaRanges[0], targetRows, targetColumns)))
                return Value.Error("#VALUE!");
            List<Value> data = Items(values);
            double total = 0, minimum = Double.PositiveInfinity, maximum = Double.NegativeInfinity;
            int matches = 0, numericCount = 0;
            for (int index = 0; index < data.Count; index++)
            {
                bool accepted = true;
                for (int pair = 0; pair < pairCount; pair++)
                    if (!CriteriaMatches(criteriaRanges[pair][index], criteria[pair]))
                    { accepted = false; break; }
                if (!accepted) continue;
                matches++;
                if (count) continue;
                Value item = data[index];
                if (item.Kind == ValueKind.Error) return item;
                if (item.Kind != ValueKind.Number) continue;
                numericCount++; total += item.Number;
                minimum = Math.Min(minimum, item.Number);
                maximum = Math.Max(maximum, item.Number);
            }
            if (count) return Value.Numeric(matches);
            if (name == "AVERAGEIF" || name == "AVERAGEIFS")
                return numericCount == 0 ? Value.Error("#DIV/0!") : Value.Numeric(total / numericCount);
            if (name == "MAXIFS") return Value.Numeric(numericCount == 0 ? 0 : maximum);
            if (name == "MINIFS") return Value.Numeric(numericCount == 0 ? 0 : minimum);
            return Value.Numeric(total);
        }

        private static Value DateSerial(Value source)
        {
            Value value = Scalar(source);
            if (value.Kind == ValueKind.Error) return value;
            if (value.Kind == ValueKind.Number) return value;
            if (value.Kind == ValueKind.Blank) return Value.Numeric(0);
            DateTime parsed;
            string[] patterns = { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "dd/MM/yyyy", "d/M/yyyy" };
            if (DateTime.TryParseExact(value.Text, patterns, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out parsed)) return Value.Numeric(parsed.ToOADate());
            return Value.Error("#VALUE!");
        }

        private static Value DateFrom(Value source, out DateTime date)
        {
            date = DateTime.MinValue;
            Value serial = DateSerial(source);
            if (serial.Kind == ValueKind.Error) return serial;
            try { date = DateTime.FromOADate(serial.Number); return null; }
            catch (ArgumentException) { return Value.Error("#NUM!"); }
        }

        private static int IsoWeekNumber(DateTime date)
        {
            int weekday = ((int)date.DayOfWeek + 6) % 7;
            DateTime thursday = date.AddDays(3 - weekday);
            DateTime jan4 = new DateTime(thursday.Year, 1, 4);
            DateTime firstThursday = jan4.AddDays(3 - ((int)jan4.DayOfWeek + 6) % 7);
            return 1 + (thursday - firstThursday).Days / 7;
        }

        private Value EvaluateDate(string name, List<Node> arguments)
        {
            if (name == "TODAY" || name == "NOW")
            {
                if (arguments.Count != 0) return Value.Error("#VALUE!");
                MarkVolatile();
                DateTime now = nowProvider();
                return Value.Numeric((name == "TODAY" ? now.Date : now).ToOADate());
            }
            if (name == "DATE" || name == "TIME")
            {
                if (arguments.Count != 3) return Value.Error("#VALUE!");
                var parts = new int[3];
                for (int i = 0; i < 3; i++)
                {
                    Value number = NumberOf(arguments[i], this);
                    if (number.Kind == ValueKind.Error) return number;
                    if (number.Number < Int32.MinValue || number.Number > Int32.MaxValue)
                        return Value.Error("#NUM!");
                    parts[i] = (int)number.Number;
                }
                if (name == "TIME")
                {
                    if (parts[0] < 0 || parts[1] < 0 || parts[2] < 0) return Value.Error("#NUM!");
                    double seconds = (double)parts[0] * 3600 + (double)parts[1] * 60 + parts[2];
                    return Value.Numeric((seconds % 86400) / 86400);
                }
                int year = parts[0] < 1900 && parts[0] >= 0 ? parts[0] + 1900 : parts[0];
                try
                {
                    return Value.Numeric(new DateTime(year, 1, 1).AddMonths(parts[1] - 1)
                        .AddDays(parts[2] - 1).ToOADate());
                }
                catch (ArgumentException) { return Value.Error("#NUM!"); }
            }
            int expected = name == "EOMONTH" ? 2 : 1;
            if ((name == "WEEKDAY" || name == "WEEKNUM") && arguments.Count == 2) expected = 2;
            if (arguments.Count != expected) return Value.Error("#VALUE!");
            DateTime date;
            Value dateError = DateFrom(arguments[0].Evaluate(this), out date);
            if (dateError != null) return dateError;
            if (name == "YEAR") return Value.Numeric(date.Year);
            if (name == "MONTH") return Value.Numeric(date.Month);
            if (name == "DAY") return Value.Numeric(date.Day);
            Value option = arguments.Count == 2 ? NumberOf(arguments[1], this) : Value.Numeric(1);
            if (option.Kind == ValueKind.Error) return option;
            int mode = (int)option.Number;
            if (name == "WEEKDAY")
            {
                if (mode != 1 && mode != 2 && mode != 3) return Value.Error("#NUM!");
                int monday = ((int)date.DayOfWeek + 6) % 7;
                return Value.Numeric(mode == 1 ? (int)date.DayOfWeek + 1 : mode == 2 ? monday + 1 : monday);
            }
            if (name == "WEEKNUM")
            {
                if (mode != 1 && mode != 2 && mode != 21) return Value.Error("#NUM!");
                if (mode == 21) return Value.Numeric(IsoWeekNumber(date));
                return Value.Numeric(CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(date,
                    CalendarWeekRule.FirstDay, mode == 1 ? DayOfWeek.Sunday : DayOfWeek.Monday));
            }
            try
            {
                return Value.Numeric(new DateTime(date.Year, date.Month, 1).AddMonths(mode + 1)
                    .AddDays(-1).ToOADate());
            }
            catch (ArgumentException) { return Value.Error("#NUM!"); }
        }

        private Value EvaluateAdvanced(string name, List<Node> arguments)
        {
            if (name == "IFERROR" || name == "IFNA")
            {
                if (arguments.Count != 2) return Value.Error("#VALUE!");
                Value answer = arguments[0].Evaluate(this);
                return answer.Kind == ValueKind.Error && answer.Text != "#CYCLE!" &&
                    (name == "IFERROR" || answer.Text == "#N/A") ?
                    arguments[1].Evaluate(this) : answer;
            }
            if (name == "ISBLANK" || name == "ISNUMBER" || name == "ISTEXT" ||
                name == "ISERROR" || name == "ISNA")
            {
                if (arguments.Count != 1) return Value.Error("#VALUE!");
                Value value = Scalar(arguments[0].Evaluate(this));
                bool result = name == "ISBLANK" ? value.Kind == ValueKind.Blank :
                    name == "ISNUMBER" ? value.Kind == ValueKind.Number :
                    name == "ISTEXT" ? value.Kind == ValueKind.Text :
                    name == "ISERROR" ? value.Kind == ValueKind.Error :
                    value.Kind == ValueKind.Error && value.Text == "#N/A";
                return Value.Numeric(result ? 1 : 0);
            }
            if (name == "XLOOKUP" || name == "VLOOKUP" || name == "HLOOKUP" ||
                name == "INDEX" || name == "MATCH") return EvaluateLookup(name, arguments);
            if (name == "COUNTIF" || name == "SUMIF" || name == "AVERAGEIF" ||
                name == "COUNTIFS" || name == "SUMIFS" || name == "AVERAGEIFS" ||
                name == "MAXIFS" || name == "MINIFS")
                return EvaluateConditionalAggregate(name, arguments);
            if (name == "DATE" || name == "TIME" || name == "TODAY" || name == "NOW" ||
                name == "YEAR" || name == "MONTH" || name == "DAY" ||
                name == "WEEKDAY" || name == "WEEKNUM" || name == "EOMONTH")
                return EvaluateDate(name, arguments);
            return null;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DinkCel
{
    internal sealed partial class FormulaEngine
    {
        private static List<double> NumericItems(IEnumerable<Value> values)
        {
            var result = new List<double>();
            foreach (Value value in values)
                if (value.Kind == ValueKind.Number) result.Add(value.Number);
            return result;
        }

        private static double FinancialValue(double rate, int periods, double payment, double present, int due)
        {
            if (Math.Abs(rate) < 1e-12) return present + payment * periods;
            double growth = Math.Pow(1 + rate, periods);
            return present * growth + payment * (1 + rate * due) * (growth - 1) / rate;
        }

        private static DateTime DateOf(Value value)
        {
            if (value.Kind == ValueKind.Number) return DateTime.FromOADate(value.Number);
            DateTime date;
            if (DateTime.TryParse(value.Text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
                DateTime.TryParse(value.Text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                return date;
            throw new FormatException();
        }

        private static double YearFraction(DateTime first, DateTime second, int basis)
        {
            if (first > second) return -YearFraction(second, first, basis);
            if (basis == 2) return (second - first).TotalDays / 360.0;
            if (basis == 3) return (second - first).TotalDays / 365.0;
            if (basis == 1)
            {
                if (first.Year == second.Year)
                    return (second - first).TotalDays / (DateTime.IsLeapYear(first.Year) ? 366.0 : 365.0);
                double result = (new DateTime(first.Year + 1, 1, 1) - first).TotalDays /
                    (DateTime.IsLeapYear(first.Year) ? 366.0 : 365.0);
                for (int year = first.Year + 1; year < second.Year; year++) result++;
                return result + (second - new DateTime(second.Year, 1, 1)).TotalDays /
                    (DateTime.IsLeapYear(second.Year) ? 366.0 : 365.0);
            }
            int day1 = first.Day, day2 = second.Day;
            if (basis == 0)
            {
                if (first.Month == 2 && first.Day == DateTime.DaysInMonth(first.Year, 2)) day1 = 30;
                if (second.Month == 2 && second.Day == DateTime.DaysInMonth(second.Year, 2) && day1 >= 30)
                    day2 = 30;
                if (day1 == 31) day1 = 30;
                if (day2 == 31 && day1 >= 30) day2 = 30;
            }
            else { day1 = Math.Min(day1, 30); day2 = Math.Min(day2, 30); }
            return ((second.Year - first.Year) * 360 + (second.Month - first.Month) * 30 + day2 - day1) / 360.0;
        }

        private Value EvaluateFurther(string name, List<Node> arguments)
        {
            if (name == "CHOOSE")
            {
                if (arguments.Count < 2) return Value.Error("#VALUE!");
                Value index = NumberArgument(arguments[0], this);
                if (index.Kind == ValueKind.Error) return index;
                int selected = (int)index.Number;
                return selected < 1 || selected >= arguments.Count ? Value.Error("#VALUE!") :
                    arguments[selected].Evaluate(this);
            }
            if (name == "XMATCH" || name == "LOOKUP")
            {
                if (arguments.Count < 2 || arguments.Count > (name == "XMATCH" ? 4 : 3))
                    return Value.Error("#VALUE!");
                Value lookup = Scalar(arguments[0].Evaluate(this));
                Value vector = arguments[1].Evaluate(this);
                if (lookup.Kind == ValueKind.Error) return lookup;
                if (vector.Kind == ValueKind.Error) return vector;
                List<Value> values = Items(vector);
                if (vector.Kind == ValueKind.Range && vector.Rows > 1 && vector.Columns > 1)
                    return Value.Error("#VALUE!");
                if (name == "XMATCH")
                {
                    int mode = 0, search = 1;
                    if (arguments.Count > 2) { Value n = NumberArgument(arguments[2], this); if (n.Kind == ValueKind.Error) return n; mode = (int)n.Number; }
                    if (arguments.Count > 3) { Value n = NumberArgument(arguments[3], this); if (n.Kind == ValueKind.Error) return n; search = (int)n.Number; }
                    int index = FindLookup(lookup, values, mode, search);
                    return index < 0 ? Value.Error("#N/A") : Value.Numeric(index + 1);
                }
                List<Value> outputs = arguments.Count == 3 ? Items(arguments[2].Evaluate(this)) : values;
                if (outputs.Count != values.Count) return Value.Error("#VALUE!");
                int match = -1;
                for (int i = 0; i < values.Count; i++)
                    if (CompareLookup(values[i], lookup) <= 0 &&
                        (match < 0 || CompareLookup(values[i], values[match]) >= 0)) match = i;
                return match < 0 ? Value.Error("#N/A") : outputs[match];
            }
            if (name == "STDEV.S" || name == "STDEV.P" || name == "STDEV" ||
                name == "STDEVP" || name == "VAR.S" || name == "VAR.P" ||
                name == "VAR" || name == "VARP" || name == "MODE.SNGL" ||
                name == "PERCENTILE.INC" || name == "QUARTILE.INC" || name == "CORREL")
            {
                if (arguments.Count == 0) return Value.Error("#VALUE!");
                if (name == "CORREL")
                {
                    if (arguments.Count != 2) return Value.Error("#VALUE!");
                    List<Value> x = Items(arguments[0].Evaluate(this));
                    List<Value> y = Items(arguments[1].Evaluate(this));
                    if (x.Count != y.Count || x.Count < 2) return Value.Error("#N/A");
                    var pairs = x.Zip(y, (a, b) => new { a, b })
                        .Where(p => p.a.Kind == ValueKind.Number && p.b.Kind == ValueKind.Number).ToList();
                    if (pairs.Count < 2) return Value.Error("#DIV/0!");
                    double ax = pairs.Average(p => p.a.Number), ay = pairs.Average(p => p.b.Number);
                    double numerator = pairs.Sum(p => (p.a.Number - ax) * (p.b.Number - ay));
                    double dx = pairs.Sum(p => Math.Pow(p.a.Number - ax, 2));
                    double dy = pairs.Sum(p => Math.Pow(p.b.Number - ay, 2));
                    return dx == 0 || dy == 0 ? Value.Error("#DIV/0!") :
                        Value.Numeric(numerator / Math.Sqrt(dx * dy));
                }
                var values = new List<double>();
                int count = name == "PERCENTILE.INC" || name == "QUARTILE.INC" ? 1 : arguments.Count;
                for (int i = 0; i < count; i++)
                {
                    Value value = arguments[i].Evaluate(this);
                    if (value.Kind == ValueKind.Error) return value;
                    values.AddRange(NumericItems(Items(value)));
                }
                if (values.Count == 0) return Value.Error("#NUM!");
                if (name == "MODE.SNGL")
                {
                    var group = values.GroupBy(x => x).OrderByDescending(x => x.Count())
                        .ThenBy(x => x.Key).First();
                    return group.Count() < 2 ? Value.Error("#N/A") : Value.Numeric(group.Key);
                }
                if (name == "PERCENTILE.INC" || name == "QUARTILE.INC")
                {
                    if (arguments.Count != 2) return Value.Error("#VALUE!");
                    Value fraction = NumberArgument(arguments[1], this);
                    if (fraction.Kind == ValueKind.Error) return fraction;
                    double p = name == "QUARTILE.INC" ? fraction.Number / 4.0 : fraction.Number;
                    if (p < 0 || p > 1 || name == "QUARTILE.INC" && fraction.Number != Math.Truncate(fraction.Number))
                        return Value.Error("#NUM!");
                    values.Sort();
                    double position = p * (values.Count - 1);
                    int low = (int)Math.Floor(position), high = (int)Math.Ceiling(position);
                    return Value.Numeric(values[low] + (values[high] - values[low]) * (position - low));
                }
                bool sample = name == "STDEV.S" || name == "STDEV" || name == "VAR.S" || name == "VAR";
                if (values.Count < (sample ? 2 : 1)) return Value.Error("#DIV/0!");
                double mean = values.Average();
                double variance = values.Sum(x => (x - mean) * (x - mean)) / (values.Count - (sample ? 1 : 0));
                return Value.Numeric(name.StartsWith("STDEV", StringComparison.Ordinal) ? Math.Sqrt(variance) : variance);
            }
            if (name == "PMT" || name == "FV" || name == "PV" || name == "NPV" ||
                name == "IRR" || name == "RATE")
            {
                if (name == "NPV")
                {
                    if (arguments.Count < 2) return Value.Error("#VALUE!");
                    Value rate = NumberArgument(arguments[0], this);
                    if (rate.Kind == ValueKind.Error) return rate;
                    if (rate.Number <= -1) return Value.Error("#NUM!");
                    double sum = 0; int period = 0;
                    for (int i = 1; i < arguments.Count; i++)
                        foreach (Value value in Items(arguments[i].Evaluate(this)))
                            if (value.Kind == ValueKind.Number)
                                sum += value.Number / Math.Pow(1 + rate.Number, ++period);
                    return Value.Numeric(sum);
                }
                if (name == "IRR")
                {
                    if (arguments.Count < 1 || arguments.Count > 2) return Value.Error("#VALUE!");
                    List<double> flows = NumericItems(Items(arguments[0].Evaluate(this)));
                    if (!flows.Any(x => x < 0) || !flows.Any(x => x > 0)) return Value.Error("#NUM!");
                    double guess = arguments.Count == 2 ? NumberArgument(arguments[1], this).Number : 0.1;
                    for (int iteration = 0; iteration < 100; iteration++)
                    {
                        double cashflowSum = 0, slope = 0;
                        for (int i = 0; i < flows.Count; i++)
                        { cashflowSum += flows[i] / Math.Pow(1 + guess, i); if (i > 0) slope -= i * flows[i] / Math.Pow(1 + guess, i + 1); }
                        if (Math.Abs(cashflowSum) < 1e-9) return Value.Numeric(guess);
                        if (Math.Abs(slope) < 1e-12) break;
                        double next = guess - cashflowSum / slope;
                        if (next <= -1 || Double.IsNaN(next) || Double.IsInfinity(next)) break;
                        guess = next;
                    }
                    return Value.Error("#NUM!");
                }
                if (arguments.Count < 3 || arguments.Count > (name == "RATE" ? 6 : 5))
                    return Value.Error("#VALUE!");
                var numbers = new List<double>();
                foreach (Node argument in arguments)
                {
                    Value value = NumberArgument(argument, this);
                    if (value.Kind == ValueKind.Error) return value;
                    numbers.Add(value.Number);
                }
                if (name == "RATE")
                {
                    int periods = (int)numbers[0];
                    double payment = numbers[1], present = numbers[2], future = numbers.Count > 3 ? numbers[3] : 0;
                    int due = numbers.Count > 4 ? (int)numbers[4] : 0;
                    double rate = numbers.Count > 5 ? numbers[5] : 0.1;
                    if (periods <= 0) return Value.Error("#NUM!");
                    for (int i = 0; i < 100; i++)
                    {
                        double function = FinancialValue(rate, periods, payment, present, due) + future;
                        if (Math.Abs(function) < 1e-8) return Value.Numeric(rate);
                        double epsilon = 1e-6;
                        double slope = (FinancialValue(rate + epsilon, periods, payment, present, due) -
                            FinancialValue(rate - epsilon, periods, payment, present, due)) / (2 * epsilon);
                        if (Math.Abs(slope) < 1e-12) break;
                        rate -= function / slope;
                        if (rate <= -1 || Double.IsNaN(rate) || Double.IsInfinity(rate)) break;
                    }
                    return Value.Error("#NUM!");
                }
                double r = numbers[0], n = numbers[1], amount = numbers[2];
                double optional = numbers.Count > 3 ? numbers[3] : 0;
                int timing = numbers.Count > 4 ? (int)numbers[4] : 0;
                if (n <= 0 || r <= -1 || timing < 0 || timing > 1) return Value.Error("#NUM!");
                if (name == "FV") return Value.Numeric(-FinancialValue(r, (int)n, amount, optional, timing));
                if (name == "PV")
                {
                    double growth = Math.Pow(1 + r, n);
                    return Value.Numeric(r == 0 ? -(amount * n + optional) :
                        -(optional + amount * (1 + r * timing) * (growth - 1) / r) / growth);
                }
                double factor = Math.Pow(1 + r, n);
                return Value.Numeric(r == 0 ? -(amount + optional) / n :
                    -(amount * factor + optional) * r / ((1 + r * timing) * (factor - 1)));
            }
            if (name == "WORKDAY" || name == "NETWORKDAYS" || name == "EDATE" ||
                name == "DAYS" || name == "DATEDIF" || name == "YEARFRAC")
            {
                if (arguments.Count < 2 || arguments.Count > 3) return Value.Error("#VALUE!");
                try
                {
                    DateTime first = DateOf(Scalar(arguments[0].Evaluate(this)));
                    Value secondValue = Scalar(arguments[1].Evaluate(this));
                    if (name == "EDATE") return Value.Numeric(first.AddMonths((int)AsNumber(secondValue).Number).ToOADate());
                    if (name == "WORKDAY")
                    {
                        int days = (int)AsNumber(secondValue).Number, direction = Math.Sign(days);
                        HashSet<DateTime> holidays = arguments.Count > 2 ?
                            new HashSet<DateTime>(Items(arguments[2].Evaluate(this)).Select(DateOf).Select(d => d.Date)) :
                            new HashSet<DateTime>();
                        DateTime result = first.Date;
                        for (int i = 0; i < Math.Abs(days);)
                        {
                            result = result.AddDays(direction);
                            if (result.DayOfWeek != DayOfWeek.Saturday && result.DayOfWeek != DayOfWeek.Sunday &&
                                !holidays.Contains(result)) i++;
                        }
                        return Value.Numeric(result.ToOADate());
                    }
                    DateTime second = DateOf(secondValue);
                    if (name == "DAYS") return Value.Numeric((first.Date - second.Date).TotalDays);
                    if (name == "NETWORKDAYS")
                    {
                        HashSet<DateTime> holidays = arguments.Count > 2 ?
                            new HashSet<DateTime>(Items(arguments[2].Evaluate(this)).Select(DateOf).Select(d => d.Date)) :
                            new HashSet<DateTime>();
                        int direction = first <= second ? 1 : -1, count = 0;
                        for (DateTime date = first.Date; direction > 0 ? date <= second.Date : date >= second.Date;
                            date = date.AddDays(direction))
                            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday &&
                                !holidays.Contains(date)) count += direction;
                        return Value.Numeric(count);
                    }
                    if (name == "YEARFRAC")
                    {
                        int basis = arguments.Count > 2 ? (int)AsNumber(Scalar(arguments[2].Evaluate(this))).Number : 0;
                        return basis < 0 || basis > 4 ? Value.Error("#NUM!") :
                            Value.Numeric(YearFraction(first.Date, second.Date, basis));
                    }
                    string unit = arguments.Count > 2 ? TextOf(Scalar(arguments[2].Evaluate(this))).ToUpperInvariant() : "D";
                    if (first > second) return Value.Error("#NUM!");
                    if (unit == "D") return Value.Numeric((second.Date - first.Date).TotalDays);
                    int months = (second.Year - first.Year) * 12 + second.Month - first.Month - (second.Day < first.Day ? 1 : 0);
                    if (unit == "M") return Value.Numeric(months);
                    if (unit == "Y") return Value.Numeric(months / 12);
                    if (unit == "YM") return Value.Numeric(months % 12);
                    if (unit == "MD") return Value.Numeric((second - first.AddMonths(months)).TotalDays);
                    if (unit == "YD") return Value.Numeric((second - first.AddYears(months / 12)).TotalDays);
                    return Value.Error("#NUM!");
                }
                catch { return Value.Error("#VALUE!"); }
            }
            if (name == "CONVERT" || name == "DEC2BIN" || name == "BIN2DEC" ||
                name == "DEC2HEX" || name == "HEX2DEC")
            {
                try
                {
                    if (name == "CONVERT")
                    {
                        if (arguments.Count != 3) return Value.Error("#VALUE!");
                        Value number = NumberArgument(arguments[0], this);
                        if (number.Kind == ValueKind.Error) return number;
                        string from = TextOf(Scalar(arguments[1].Evaluate(this)));
                        string to = TextOf(Scalar(arguments[2].Evaluate(this)));
                        var units = new Dictionary<string, Tuple<string, double>>(StringComparer.OrdinalIgnoreCase)
                        {
                            { "m", Tuple.Create("length", 1.0) }, { "km", Tuple.Create("length", 1000.0) },
                            { "cm", Tuple.Create("length", 0.01) }, { "in", Tuple.Create("length", 0.0254) },
                            { "ft", Tuple.Create("length", 0.3048) }, { "mi", Tuple.Create("length", 1609.344) },
                            { "g", Tuple.Create("mass", 1.0) }, { "kg", Tuple.Create("mass", 1000.0) },
                            { "lbm", Tuple.Create("mass", 453.59237) },
                            { "s", Tuple.Create("time", 1.0) }, { "min", Tuple.Create("time", 60.0) },
                            { "hr", Tuple.Create("time", 3600.0) }, { "day", Tuple.Create("time", 86400.0) }
                        };
                        if (from == "C" || from == "F" || from == "K")
                        {
                            if (to != "C" && to != "F" && to != "K") return Value.Error("#N/A");
                            double celsius = from == "F" ? (number.Number - 32) * 5 / 9 :
                                from == "K" ? number.Number - 273.15 : number.Number;
                            return Value.Numeric(to == "F" ? celsius * 9 / 5 + 32 :
                                to == "K" ? celsius + 273.15 : celsius);
                        }
                        Tuple<string, double> a, b;
                        if (!units.TryGetValue(from, out a) || !units.TryGetValue(to, out b) || a.Item1 != b.Item1)
                            return Value.Error("#N/A");
                        return Value.Numeric(number.Number * a.Item2 / b.Item2);
                    }
                    if (arguments.Count < 1 || arguments.Count > 2) return Value.Error("#VALUE!");
                    string raw = TextOf(Scalar(arguments[0].Evaluate(this))).Trim();
                    if (name == "BIN2DEC") return Value.Numeric(Convert.ToInt64(raw, 2));
                    if (name == "HEX2DEC") return Value.Numeric(Convert.ToInt64(raw, 16));
                    Value decimalValue = NumberArgument(arguments[0], this);
                    if (decimalValue.Kind == ValueKind.Error) return decimalValue;
                    long integer = (long)decimalValue.Number;
                    if (name == "DEC2BIN" && (integer < -512 || integer > 511)) return Value.Error("#NUM!");
                    string output = Convert.ToString(integer, name == "DEC2BIN" ? 2 : 16).ToUpperInvariant();
                    if (arguments.Count == 2)
                    {
                        int places = (int)NumberArgument(arguments[1], this).Number;
                        if (places < output.Length) return Value.Error("#NUM!");
                        output = output.PadLeft(places, '0');
                    }
                    return Value.String(output);
                }
                catch { return Value.Error("#NUM!"); }
            }
            return null;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DinkCel
{
    internal static class GoalSeekSolver
    {
        public static bool Solve(Func<double, double> evaluate, double target,
            double initial, out double result)
        {
            double x = initial;
            for (int iteration = 0; iteration < 80; iteration++)
            {
                double value = evaluate(x) - target;
                if (Double.IsNaN(value) || Double.IsInfinity(value)) break;
                if (Math.Abs(value) <= 1e-8 * Math.Max(1, Math.Abs(target)))
                { result = x; return true; }
                double step = Math.Max(1e-5, Math.Abs(x) * 1e-5);
                double left = evaluate(x - step), right = evaluate(x + step);
                double slope = (right - left) / (2 * step);
                if (Double.IsNaN(slope) || Math.Abs(slope) < 1e-12)
                { x += iteration % 2 == 0 ? step * (iteration + 1) : -step * (iteration + 1); continue; }
                double update = value / slope;
                if (Math.Abs(update) > 1e8) update = Math.Sign(update) * 1e8;
                x -= update;
            }
            result = x;
            return false;
        }
    }

    internal static class AdvancedData
    {
        public static string[] SplitLine(string line, char delimiter)
        {
            var parts = new List<string>();
            var value = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < (line ?? "").Length; i++)
            {
                char current = line[i];
                if (current == '"')
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"')
                    { value.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (current == delimiter && !quoted)
                { parts.Add(value.ToString()); value.Clear(); }
                else value.Append(current);
            }
            parts.Add(value.ToString());
            return parts.ToArray();
        }

        public static List<int> UniqueRows(string[][] rows, bool header)
        {
            var keep = new List<int>();
            var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            for (int i = 0; i < rows.Length; i++)
            {
                if (i == 0 && header) { keep.Add(i); continue; }
                string key = String.Concat(rows[i].Select(x =>
                {
                    string value = x ?? "";
                    return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
                }).ToArray());
                if (seen.Add(key)) keep.Add(i);
            }
            return keep;
        }

        public static Func<string, string, string> InferFlashFill(
            IList<Tuple<string, string, string>> examples)
        {
            var patterns = new List<Func<string, string, string>>
            {
                (left, right) => left.Trim(),
                (left, right) => left.ToUpperInvariant(),
                (left, right) => left.ToLowerInvariant(),
                (left, right) => left.Split(' ')[0],
                (left, right) => left.Trim().Split(' ').Last(),
                (left, right) => left.Contains("@") ? left.Substring(0, left.IndexOf('@')) : left,
                (left, right) => left.Contains("@") ? left.Substring(left.IndexOf('@') + 1) : left
            };
            foreach (string separator in new[] { " ", "-", ", ", "", " / " })
            {
                string captured = separator;
                patterns.Add((left, right) => left + captured + right);
                patterns.Add((left, right) => right + captured + left);
            }
            return patterns.FirstOrDefault(pattern => examples.All(item =>
                String.Equals(pattern(item.Item1 ?? "", item.Item2 ?? ""), item.Item3 ?? "",
                    StringComparison.CurrentCultureIgnoreCase)));
        }

        public static bool MatchesCriteria(string[] row, string[] headers,
            IList<Dictionary<string, string>> alternatives)
        {
            if (alternatives == null || alternatives.Count == 0) return true;
            return alternatives.Any(criteria => criteria.All(item =>
            {
                int column = Array.FindIndex(headers, h =>
                    String.Equals(h, item.Key, StringComparison.CurrentCultureIgnoreCase));
                if (column < 0 || column >= row.Length) return false;
                string value = row[column] ?? "", test = item.Value ?? "";
                if (test.StartsWith(">=", StringComparison.Ordinal) || test.StartsWith("<=", StringComparison.Ordinal) ||
                    test.StartsWith("<>", StringComparison.Ordinal))
                {
                    double a, b;
                    string kind = DataTools.Number(value, out a) && DataTools.Number(test.Substring(2), out b) ?
                        "Number" : "Text";
                    int comparison = DataTools.Compare(value, test.Substring(2), kind);
                    return test.StartsWith(">=") ? comparison >= 0 : test.StartsWith("<=") ?
                        comparison <= 0 : comparison != 0;
                }
                if (test.StartsWith(">", StringComparison.Ordinal) || test.StartsWith("<", StringComparison.Ordinal))
                {
                    double a, b;
                    string kind = DataTools.Number(value, out a) && DataTools.Number(test.Substring(1), out b) ?
                        "Number" : "Text";
                    int comparison = DataTools.Compare(value, test.Substring(1), kind);
                    return test[0] == '>' ? comparison > 0 : comparison < 0;
                }
                if (test.IndexOfAny(new[] { '*', '?' }) >= 0)
                    return System.Text.RegularExpressions.Regex.IsMatch(value,
                        "^" + System.Text.RegularExpressions.Regex.Escape(test)
                            .Replace("\\*", ".*").Replace("\\?", ".") + "$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                return String.Equals(value, test.TrimStart('='), StringComparison.CurrentCultureIgnoreCase);
            }));
        }
    }

    internal sealed class ScenarioDefinition
    {
        public string Name = "";
        public string Sheet = "";
        public readonly Dictionary<int, string> Values = new Dictionary<int, string>();
    }
}

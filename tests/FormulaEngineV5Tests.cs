using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DinkCel
{
    internal static class FormulaEngineV5Tests
    {
        private static readonly Dictionary<int, string> cells = new Dictionary<int, string>();
        private static FormulaEngine engine;
        private static DateTime clock = new DateTime(2026, 10, 3, 14, 30, 0);
        private static readonly Dictionary<int, int> reads = new Dictionary<int, int>();
        private const int ResultIndex = 199 * 26 + 25;

        private static void Set(int row, int column, string value)
        { cells[row * 26 + column] = value; }

        private static string Read(int row, int column)
        {
            int index = row * 26 + column;
            int count;
            reads[index] = reads.TryGetValue(index, out count) ? count + 1 : 1;
            string value;
            return cells.TryGetValue(index, out value) ? value : "";
        }

        private static string Number(double value)
        { return value.ToString("0.##########", CultureInfo.InvariantCulture); }

        private static void Equal(string name, string expected, string actual)
        {
            if (expected != actual) throw new Exception(name + ": expected " + expected + ", got " + actual);
        }

        private static void Check(bool condition, string message)
        { if (!condition) throw new Exception(message); }

        private static string Calculate(string expression)
        {
            cells[ResultIndex] = expression;
            engine.Invalidate("Main", 199, 25);
            return engine.Display(199, 25);
        }

        private static void Formula(string name, string expected, string expression)
        { Equal(name, expected, Calculate(expression)); }

        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            string[] keys = { "Apple", "Banana", "apple", "Cherry" };
            string[] groups = { "North", "South", "North", "South" };
            for (int row = 0; row < 4; row++)
            {
                Set(row, 0, keys[row]);
                Set(row, 1, ((row + 1) * 10).ToString(CultureInfo.InvariantCulture));
                Set(row, 2, groups[row]);
                Set(row, 3, (row + 1).ToString(CultureInfo.InvariantCulture));
                Set(row, 10, (40 - row * 10).ToString(CultureInfo.InvariantCulture));
            }
            Set(0, 5, "Jan"); Set(0, 6, "Feb"); Set(0, 7, "Mar");
            Set(1, 5, "5"); Set(1, 6, "10"); Set(1, 7, "15");
            Set(0, 12, "2"); Set(0, 13, "=M1*3"); Set(0, 14, "=N1+1");
            Set(0, 16, "=100"); Set(0, 17, "=S1"); Set(0, 18, "=R1");
            Set(0, 19, "=TODAY()"); Set(0, 20, "=T1+1");
            Set(0, 21, "=IF(M1>0,N1,Q1)");
            engine = new FormulaEngine(Read, delegate(string sheet, int row, int column)
            {
                if (sheet != "Sales Data") return null;
                return (row < 2 && column < 2) ? (row * 2 + column + 1).ToString() : "";
            }, "Main", 200, 26, delegate(string name)
            {
                return name == "Revenue" ? new FormulaNamedRange
                { Sheet = "Main", FirstRow = 0, LastRow = 3, FirstColumn = 1, LastColumn = 1 } : null;
            }, delegate { return clock; });

            Formula("XLOOKUP exact", "20", "=XLOOKUP(\"Banana\",A1:A4,B1:B4)");
            Formula("XLOOKUP missing", "#N/A", "=XLOOKUP(\"Missing\",A1:A4,B1:B4)");
            Formula("XLOOKUP fallback", "none", "=XLOOKUP(\"Missing\",A1:A4,B1:B4,\"none\")");
            Formula("XLOOKUP wildcard", "20", "=XLOOKUP(\"B*\",A1:A4,B1:B4,0,2)");
            Formula("XLOOKUP reverse", "30", "=XLOOKUP(\"Apple\",A1:A4,B1:B4,0,0,-1)");
            Formula("XLOOKUP smaller", "Banana", "=XLOOKUP(25,B1:B4,A1:A4,\"\",-1)");
            Formula("XLOOKUP larger", "apple", "=XLOOKUP(25,B1:B4,A1:A4,\"\",1)");
            Formula("VLOOKUP exact", "20", "=VLOOKUP(\"Banana\",A1:B4,2,FALSE)");
            Formula("VLOOKUP approximate", "South", "=VLOOKUP(25,B1:C4,2,TRUE)");
            Formula("HLOOKUP", "10", "=HLOOKUP(\"Feb\",F1:H2,2,FALSE)");
            Formula("INDEX two dimensional", "30", "=INDEX(B1:D4,3,1)");
            Formula("INDEX one row", "Feb", "=INDEX(F1:H1,2)");
            Formula("MATCH exact", "4", "=MATCH(\"Cherry\",A1:A4,0)");
            Formula("MATCH ascending", "2", "=MATCH(25,B1:B4,1)");
            Formula("MATCH descending", "2", "=MATCH(25,K1:K4,-1)");
            Formula("INDEX MATCH", "40", "=INDEX(B1:B4,MATCH(\"Cherry\",A1:A4,0))");

            Formula("IFERROR fallback", "fallback", "=IFERROR(1/0,\"fallback\")");
            Formula("IFERROR lazy", "42", "=IFERROR(42,1/0)");
            Formula("IFNA catches N/A", "missing", "=IFNA(XLOOKUP(\"X\",A1:A4,B1:B4),\"missing\")");
            Formula("IFNA preserves other errors", "#DIV/0!", "=IFNA(1/0,0)");
            Formula("SUMIFS", "40", "=SUMIFS(B1:B4,C1:C4,\"North\")");
            Formula("COUNTIFS", "1", "=COUNTIFS(C1:C4,\"South\",B1:B4,\">25\")");
            Formula("AVERAGEIF", "20", "=AVERAGEIF(C1:C4,\"North\",B1:B4)");
            Formula("AVERAGEIFS", "40", "=AVERAGEIFS(B1:B4,C1:C4,\"South\",B1:B4,\">20\")");
            Formula("MAXIFS", "30", "=MAXIFS(B1:B4,C1:C4,\"North\")");
            Formula("MINIFS", "20", "=MINIFS(B1:B4,C1:C4,\"South\")");
            Formula("COUNTIF wildcard", "2", "=COUNTIF(A1:A4,\"A*\")");
            Formula("Named range", "100", "=SUM(Revenue)");

            Formula("DATE", Number(new DateTime(2024, 2, 29).ToOADate()), "=DATE(2024,2,29)");
            Formula("DATE overflow", Number(new DateTime(2025, 1, 1).ToOADate()), "=DATE(2024,13,1)");
            Formula("TIME", Number(12.5 / 24), "=TIME(12,30,0)");
            Formula("TODAY", Number(clock.Date.ToOADate()), "=TODAY()");
            Formula("NOW", Number(clock.ToOADate()), "=NOW()");
            Formula("YEAR", "2024", "=YEAR(DATE(2024,2,29))");
            Formula("MONTH", "2", "=MONTH(DATE(2024,2,29))");
            Formula("DAY", "29", "=DAY(DATE(2024,2,29))");
            Formula("YEAR text", "2026", "=YEAR(\"2026-10-03\")");
            Formula("WEEKDAY Monday", "4", "=WEEKDAY(DATE(2024,2,29),2)");
            Formula("WEEKNUM Monday", "1", "=WEEKNUM(DATE(2024,1,1),2)");
            Formula("WEEKNUM ISO", "53", "=WEEKNUM(DATE(2021,1,1),21)");
            Formula("EOMONTH", Number(new DateTime(2024, 2, 29).ToOADate()),
                "=EOMONTH(DATE(2024,2,10),0)");

            Formula("ISBLANK", "1", "=ISBLANK(E1)");
            Formula("ISBLANK formula empty", "0", "=ISBLANK(\"\")");
            Formula("ISNUMBER", "1", "=ISNUMBER(B1)");
            Formula("ISTEXT", "1", "=ISTEXT(A1)");
            Formula("ISERROR", "1", "=ISERROR(1/0)");
            Formula("ISNA", "1", "=ISNA(XLOOKUP(\"X\",A1:A4,B1:B4))");
            Formula("ISNA other error", "0", "=ISNA(1/0)");
            Formula("literal N/A", "#N/A", "=#N/A");
            Formula("value error", "#VALUE!", "=1+\"x\"");
            Formula("reference error", "#REF!", "=AA1");
            Formula("division error", "#DIV/0!", "=1/0");
            Formula("name error", "#NAME?", "=UNKNOWN(1)");
            Formula("number error", "#NUM!", "=SQRT(-1)");
            Formula("unknown sheet", "#REF!", "=Missing!A1");
            Formula("cross sheet range", "10", "=SUM('Sales Data'!A1:'Sales Data'!B2)");
            Formula("cross sheet mixed range", "10", "=SUM('Sales Data'!$A$1:$B2)");
            Formula("cross sheet invalid range", "#REF!", "=SUM(Main!A1:'Sales Data'!B2)");
            Equal("mixed reference fill", "='Sales Data'!$A2+B$2",
                FormulaEngine.ShiftReferences("='Sales Data'!$A1+A$2", 1, 1, 200, 26));

            Equal("dependency result", "7", engine.Display(0, 14));
            Check(engine.DependenciesFor("Main", 0, 14).Contains("MAIN!N1"), "O1 should depend on N1");
            Check(engine.CalculationChain.IndexOf("MAIN!N1") < engine.CalculationChain.IndexOf("MAIN!O1"),
                "Calculation chain should order dependencies first");
            Equal("independent", "100", engine.Display(0, 16));
            int independentReads = reads[16];
            Set(0, 12, "5"); engine.Invalidate("Main", 0, 12);
            Equal("dependent recalculation", "16", engine.Display(0, 14));
            Equal("independent cached", "100", engine.Display(0, 16));
            Check(reads[16] == independentReads, "Unrelated formula should stay cached");
            Equal("cycle", "#CYCLE!", engine.Display(0, 17));
            Set(0, 17, "=IFERROR(S1,0)"); engine.Invalidate("Main", 0, 17);
            Equal("cycle cannot be hidden", "#CYCLE!", engine.Display(0, 17));
            Set(0, 18, "3"); engine.Invalidate("Main", 0, 18);
            Equal("cycle resolved", "3", engine.Display(0, 17));
            Equal("volatile initial", Number(clock.Date.ToOADate() + 1), engine.Display(0, 20));
            clock = clock.AddDays(1);
            Set(0, 12, "-1"); engine.Invalidate("Main", 0, 12);
            Equal("volatile recalculation", Number(clock.Date.ToOADate() + 1), engine.Display(0, 20));
            Equal("dynamic branch", "100", engine.Display(0, 21));
            Check(engine.DependenciesFor("Main", 0, 21).Contains("MAIN!Q1"), "IF should depend on selected branch");
            Check(!engine.DependenciesFor("Main", 0, 21).Contains("MAIN!N1"),
                "IF should release inactive dependency");
            Console.WriteLine("FormulaEngine v0.5: lookups, criteria, dates, errors, graph and chain passed.");
        }
    }
}

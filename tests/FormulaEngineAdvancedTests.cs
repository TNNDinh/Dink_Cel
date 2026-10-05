using System;
using System.Globalization;

namespace DinkCel
{
    internal static class FormulaEngineAdvancedTests
    {
        private static readonly string[,] data = {
            { "10", "A", "100", "-100" },
            { "20", "B", "110", "110" },
            { "30", "C", "120", "" }
        };
        private static readonly FormulaEngine engine = new FormulaEngine((row, col) =>
            row < 3 && col < 4 ? data[row, col] : "", 200, 26);
        private static void Equal(string name, string expected, string expression)
        {
            string actual = engine.EvaluateExpression(expression);
            if (actual != expected) throw new Exception(name + ": expected " + expected + ", got " + actual);
        }
        private static void Near(string name, double expected, string expression)
        {
            double actual = Double.Parse(engine.EvaluateExpression(expression), CultureInfo.InvariantCulture);
            if (Math.Abs(actual - expected) > 1e-6) throw new Exception(name + ": expected " + expected + ", got " + actual);
        }
        private static void Date(string name, int year, int month, int day, string expression)
        {
            double value = Double.Parse(engine.EvaluateExpression(expression), CultureInfo.InvariantCulture);
            System.DateTime actual = System.DateTime.FromOADate(value);
            if (actual.Year != year || actual.Month != month || actual.Day != day)
                throw new Exception(name + ": got " + actual.ToString("yyyy-MM-dd"));
        }
        private static void Main()
        {
            try { Run(); Console.WriteLine("Advanced formulas: statistics, finance, dates, lookups and engineering passed."); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }
        private static void Run()
        {
            Near("stdev sample", 10, "=STDEV.S(A1:A3)");
            Near("variance population", 66.6666666667, "=VAR.P(A1:A3)");
            Equal("mode none", "#N/A", "=MODE.SNGL(A1:A3)");
            Near("percentile", 15, "=PERCENTILE.INC(A1:A3,0.25)");
            Near("quartile", 25, "=QUARTILE.INC(A1:A3,3)");
            Near("correl", 1, "=CORREL(A1:A3,C1:C3)");
            Near("fv zero rate", 1200, "=FV(0,12,-100)");
            Near("pv zero rate", 1200, "=PV(0,12,-100)");
            Near("pmt", -88.84878868, "=PMT(0.01,12,1000)");
            Near("npv", 173.553719008, "=NPV(0.1,100,100)");
            Near("irr", 0.1, "=IRR(D1:D2)");
            Date("edate", 2024, 2, 29, "=EDATE(DATE(2024,1,31),1)");
            Date("workday", 2024, 1, 8, "=WORKDAY(DATE(2024,1,5),1)");
            Equal("networkdays", "5", "=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))");
            Equal("days", "4", "=DAYS(DATE(2024,1,5),DATE(2024,1,1))");
            Equal("datedif", "1", "=DATEDIF(DATE(2024,1,1),DATE(2025,2,1),\"Y\")");
            Near("yearfrac actual leap", 1.0 / 366,
                "=YEARFRAC(DATE(2024,1,1),DATE(2024,1,2),1)");
            Near("yearfrac US 30/360", 0.5,
                "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)");
            Equal("yearfrac invalid basis", "#NUM!",
                "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),5)");
            Equal("xmatch", "2", "=XMATCH(\"B\",B1:B3)");
            Equal("lookup", "B", "=LOOKUP(25,A1:A3,B1:B3)");
            Equal("choose", "C", "=CHOOSE(3,\"A\",\"B\",\"C\")");
            Near("convert", 2000, "=CONVERT(2,\"km\",\"m\")");
            Near("temperature", 0, "=CONVERT(32,\"F\",\"C\")");
            Equal("dec2bin", "1010", "=DEC2BIN(10)");
            Equal("bin2dec", "10", "=BIN2DEC(\"1010\")");
            Equal("dec2hex", "FF", "=DEC2HEX(255)");
            Equal("hex2dec", "255", "=HEX2DEC(\"FF\")");
        }
    }
}

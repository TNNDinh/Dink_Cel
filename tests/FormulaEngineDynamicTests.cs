using System;
using System.Collections.Generic;
using System.Linq;

namespace DinkCel
{
    internal static class FormulaEngineDynamicTests
    {
        private const int Columns = 26;
        private static readonly Dictionary<int, string> cells = new Dictionary<int, string>();
        private static FormulaEngine engine;
        private static void Put(int row, int column, string raw)
        { cells[row * Columns + column] = raw; }
        private static string Read(int row, int column)
        { string raw; return cells.TryGetValue(row * Columns + column, out raw) ? raw : ""; }
        private static void Equal(string name, string expected, string actual)
        { if (expected != actual) throw new Exception(name + ": expected " + expected + ", got " + actual); }
        private static void Prepare()
        {
            engine = new FormulaEngine(Read, 200, Columns);
            engine.PrepareSpills(cells.Where(x => x.Value.StartsWith("=")).Select(x => x.Key),
                (r, c) => Read(r, c).Length > 0);
        }
        private static string Spill(int row, int column)
        {
            string answer;
            return engine.SpillDisplays().TryGetValue(row * Columns + column, out answer) ? answer : "";
        }
        private static void Main()
        {
            try { Run(); Console.WriteLine("Dynamic arrays: functions, spill and references passed."); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }
        private static void Run()
        {
            Put(1, 0, "Anh"); Put(1, 1, "Hanoi"); Put(1, 2, "10");
            Put(2, 0, "Binh"); Put(2, 1, "Hue"); Put(2, 2, "20");
            Put(3, 0, "Chi"); Put(3, 1, "Hanoi"); Put(3, 2, "30");
            Put(0, 4, "=FILTER(A2:C4,B2:B4=\"Hanoi\")");
            Put(0, 9, "=SUM(E1#)");
            Prepare();
            Equal("filter first", "Anh", engine.Display(0, 4));
            Equal("filter row 2", "Chi", Spill(1, 4));
            Equal("filter column", "30", Spill(1, 6));
            Equal("spill ref sum", "40", engine.Display(0, 9));
            Put(0, 14, "=SEQUENCE(3,2,5,2)");
            Prepare();
            Equal("sequence", "5", engine.Display(0, 14));
            Equal("sequence spill", "15", Spill(2, 15));
            Put(0, 18, "=TRANSPOSE(A2:C2)");
            Prepare();
            Equal("transpose", "Hanoi", Spill(1, 18));
            Put(5, 4, "=UNIQUE(B2:B4)");
            Put(5, 9, "=SORT(C2:C4,1,-1)");
            Prepare();
            Equal("unique", "Hanoi", engine.Display(5, 4));
            Equal("unique spill", "Hue", Spill(6, 4));
            Equal("sort", "30", engine.Display(5, 9));
            Equal("sort spill", "10", Spill(7, 9));
            Put(10, 0, "=LET(x,SEQUENCE(2,1),SUM(x))");
            Put(10, 3, "=CHOOSECOLS(A2:C4,3,1)");
            Put(10, 7, "=CHOOSEROWS(A2:C4,-1,1)");
            Put(10, 11, "=TAKE(A2:C4,-2,2)");
            Put(10, 15, "=DROP(A2:C4,1,1)");
            Put(15, 0, "=VSTACK(A2:A3,A4:A4)");
            Put(15, 3, "=HSTACK(A2:A3,C2:C3)");
            Put(15, 7, "=SORTBY(A2:C4,C2:C4,-1)");
            Prepare();
            Equal("let", "3", engine.Display(10, 0));
            Equal("choosecols", "10", engine.Display(10, 3));
            Equal("chooserows", "Chi", engine.Display(10, 7));
            Equal("take", "Binh", engine.Display(10, 11));
            Equal("drop", "Hue", engine.Display(10, 15));
            Equal("vstack", "Chi", Spill(17, 0));
            Equal("hstack", "20", Spill(16, 4));
            Equal("sortby", "Chi", engine.Display(15, 7));
            Put(2, 3, "=@A2:A4");
            Put(18, 0, "=FILTER(A2:A4,B2:B4=\"Missing\")");
            Put(18, 3, "=UNIQUE(B2:B4,0,1)");
            Prepare();
            Equal("implicit intersect", "Binh", engine.Display(2, 3));
            Equal("filter empty", "#CALC!", engine.Display(18, 0));
            Equal("unique exactly once", "Hue", engine.Display(18, 3));
            Put(20, 0, "=SEQUENCE(2,2)"); Put(21, 1, "occupied");
            Prepare();
            Equal("blocked spill", "#SPILL!", engine.Display(20, 0));
            Put(25, 0, "=@C2:C4");
            Prepare();
            Equal("implicit", "#VALUE!", engine.Display(25, 0));
            Equal("no spill ref", "#REF!", engine.EvaluateExpression("=A2#"));
        }
    }
}

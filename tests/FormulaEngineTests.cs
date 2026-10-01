using System;
using System.Collections.Generic;

namespace DinkCel
{
    internal static class FormulaEngineTests
    {
        private static void Check(string name, string expected, string actual)
        {
            if (expected != actual)
                throw new Exception(name + ": expected " + expected + ", got " + actual);
        }

        private static void Main()
        {
            var cells = new Dictionary<int, string>();
            cells[0] = "10";                    // A1
            cells[26] = "20";                   // A2
            cells[52] = "=SUM(A1:A2)";          // A3
            cells[1] = "=AVERAGE(A1:A2)";       // B1
            cells[27] = "=MIN(A1:A2)";          // B2
            cells[53] = "=MAX(A1:A2)";          // B3
            cells[2] = "=COUNT(A1:A2)";         // C1
            cells[28] = "=IF(A1>5,100,1/0)";   // C2
            cells[54] = "=A1*2+5";              // C3
            cells[3] = "=A1/0";                 // D1
            cells[29] = "=Z200";                // D2
            Func<int, int, string> read = delegate(int row, int column)
            {
                string result;
                return cells.TryGetValue(row * 26 + column, out result) ? result : "";
            };

            var engine = new FormulaEngine(read, 200, 26);
            Check("SUM", "30", engine.Display(2, 0));
            Check("AVERAGE", "15", engine.Display(0, 1));
            Check("MIN", "10", engine.Display(1, 1));
            Check("MAX", "20", engine.Display(2, 1));
            Check("COUNT", "2", engine.Display(0, 2));
            Check("IF", "100", engine.Display(1, 2));
            Check("Arithmetic", "25", engine.Display(2, 2));
            Check("Division by zero", "#DIV/0!", engine.Display(0, 3));
            Check("Blank reference", "0", engine.Display(1, 3));
            Check("Relative fill", "=SUM(B2:$B$2)+D4",
                FormulaEngine.ShiftReferences("=SUM(A1:$B$2)+C3",
                    1, 1, 200, 26));
            Check("Quoted text in fill", "=IF(B2>0,\"A1\",0)",
                FormulaEngine.ShiftReferences("=IF(A1>0,\"A1\",0)",
                    1, 1, 200, 26));
            Check("Out of bounds fill", "=#REF!",
                FormulaEngine.ShiftReferences("=A1", -1, 0, 200, 26));
            Check("Insert row references", "=SUM(A1:A4)+$B$3+\"A2\"",
                FormulaEngine.ShiftStructureReferences(
                    "=SUM(A1:A3)+$B$2+\"A2\"", true, 1, true, 200, 26));
            Check("Delete row range edge", "=SUM(A1:A2)",
                FormulaEngine.ShiftStructureReferences(
                    "=SUM(A1:A3)", true, 0, false, 200, 26));
            Check("Delete direct reference", "=#REF!",
                FormulaEngine.ShiftStructureReferences(
                    "=A2", true, 1, false, 200, 26));
            Check("Insert column absolute", "=$C$2+D4",
                FormulaEngine.ShiftStructureReferences(
                    "=$B$2+C4", false, 1, true, 200, 26));
            Check("Delete column range", "=SUM(A1:B1)",
                FormulaEngine.ShiftStructureReferences(
                    "=SUM(A1:C1)", false, 1, false, 200, 26));

            cells[0] = "=B1";
            cells[1] = "=A1";
            var cycle = new FormulaEngine(read, 200, 26);
            Check("Cycle", "#CYCLE!", cycle.Display(0, 0));
            Console.WriteLine("FormulaEngine tests: OK");
        }
    }
}

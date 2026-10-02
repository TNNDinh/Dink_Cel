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
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
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
            cells[4] = "=ROUND(12.345,2)";
            cells[5] = "=MOD(-5,3)";
            cells[6] = "=MEDIAN(A1:A2,30)";
            cells[7] = "=COUNTA(A1:A2,Z200)";
            cells[8] = "=LEFT(UPPER(\"abc\"),2)";
            cells[9] = "=CONCAT(\"Hi\",\" \",RIGHT(\"there\",3))";
            cells[10] = "=AND(A1>0,A2>0)";
            cells[11] = "=SQRT(9)+ABS(-2)";
            cells[14] = "=COUNTIF(A1:A2,\">15\")";
            cells[15] = "=SUMIF(A1:A2,\">15\")";
            cells[16] = "=SUM(Revenue)";
            var expanded = new FormulaEngine(read, 200, 26);
            Check("ROUND", "12.35", expanded.Display(0, 4));
            Check("MOD", "1", expanded.Display(0, 5));
            Check("MEDIAN", "20", expanded.Display(0, 6));
            Check("COUNTA", "2", expanded.Display(0, 7));
            Check("LEFT UPPER", "AB", expanded.Display(0, 8));
            Check("CONCAT RIGHT", "Hi ere", expanded.Display(0, 9));
            Check("AND", "1", expanded.Display(0, 10));
            Check("SQRT ABS", "5", expanded.Display(0, 11));
            Check("COUNTIF", "1", expanded.Display(0, 14));
            Check("SUMIF", "20", expanded.Display(0, 15));
            var names = new FormulaEngine(read, null, "Main", 200, 26,
                delegate(string name)
                {
                    return name == "Revenue" ? new FormulaNamedRange
                    { Sheet = "Main", FirstRow = 0, FirstColumn = 0, LastRow = 1, LastColumn = 0 } : null;
                });
            Check("Named range", "30", names.Display(0, 16));
            cells[12] = "=SUM(Other!A1:A2)+'My Sheet'!B1";
            cells[13] = "=Other!C1";
            var otherCells = new Dictionary<int, string>();
            otherCells[0] = "3";
            otherCells[26] = "4";
            otherCells[1] = "5";
            otherCells[2] = "=Main!N1";
            var multi = new FormulaEngine(read, delegate(string sheet, int row, int column)
            {
                string value;
                if (sheet == "Other") return otherCells.TryGetValue(row * 26 + column, out value) ? value : "";
                if (sheet == "My Sheet") return column == 1 && row == 0 ? "5" : "";
                return null;
            }, "Main", 200, 26);
            Check("Cross-sheet references", "12", multi.Display(0, 12));
            Check("Cross-sheet cycle", "#CYCLE!", multi.Display(0, 13));
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
            Check("Move row references and range", "=A5+$B$3+SUM(A1:A5)+\"A2\"",
                FormulaEngine.MoveStructureReferences(
                    "=A2+$B$4+SUM(A1:A3)+\"A2\"", true, 1, 4));
            Check("Move column references", "=B4+$C$1+SUM(B1:D1)",
                FormulaEngine.MoveStructureReferences(
                    "=D4+$B$1+SUM(B1:D1)", false, 3, 1));
            Check("Move row upward", "=A2+A3+A5",
                FormulaEngine.MoveStructureReferences(
                    "=A5+A2+A4", true, 4, 1));

            cells[0] = "=B1";
            cells[1] = "=A1";
            var cycle = new FormulaEngine(read, 200, 26);
            Check("Cycle", "#CYCLE!", cycle.Display(0, 0));
            Console.WriteLine("FormulaEngine tests: OK");
        }
    }
}

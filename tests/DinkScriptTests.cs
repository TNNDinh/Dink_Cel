using System;
using System.Collections.Generic;
using System.Threading;

namespace DinkCel
{
    internal static class DinkScriptTests
    {
        private static void Check(bool condition, string message)
        { if (!condition) throw new Exception(message); }

        private static void Main()
        {
            EmbeddedDependencies.Install();
            TestRuntime();
            Console.WriteLine("DinkCel Script runtime: OK");
        }

        private static void TestRuntime()
        {
            var data = new Dictionary<string, object[][]> {
                { "A1:B2", new[] { new object[] { 2, 3 }, new object[] { 4, 5 } } }
            };
            object[][] written = null;
            var logs = new List<string>();
            string code = "function main() { const s = DinkCel.getActiveWorkbook().getActiveSheet(); " +
                "const x = s.getRange('A1:B2').getValues(); s.getRange('C1:C2').setValues(x.map(r => [r[0] * r[1]])); " +
                "Logger.log('updated'); return x[1][0]; }\n" +
                "/** @customfunction */ function DOUBLE(x) { return x * 2; }";
            var runtime = new DinkScript();
            object result = runtime.Invoke(code, "main", null, "Sheet1",
                (sheet, address) => data[address],
                (sheet, address, values) => { Check(address == "C1:C2", "address"); written = values; },
                logs.Add, null);
            Check(Convert.ToDouble(result) == 4, "return value");
            Check(Convert.ToDouble(written[0][0]) == 6 && Convert.ToDouble(written[1][0]) == 20, "setValues");
            Check(logs.Count == 1, "logging");
            Check(DinkScript.CustomFunctions(code).Contains("DOUBLE"), "function marker");
            Check(DinkScript.CustomFunctions("/** @customfunction @network */ function AI_FN(x) { return x; }")
                .Contains("AI_FN"), "network function marker");
            Check((string)runtime.Invoke("function main(){ return AI.generate('prompt'); }", "main",
                null, "Sheet1", (sheet, address) => data[address], null, null,
                prompt => { Check(prompt == "prompt", "AI prompt only"); return "answer"; }) == "answer",
                "AI bridge");
            Check(Convert.ToDouble(runtime.Invoke(code, "DOUBLE", new object[] { 7 }, "Sheet1",
                (sheet, address) => data[address], null, null, null)) == 14, "custom function");
            bool rejected = false;
            try { runtime.Invoke("function main(){DinkCel.getActiveWorkbook().getActiveSheet().getRange('A1').setValue(2)}",
                "main", null, "Sheet1", (sheet, address) => data[address], null, null, null); }
            catch { rejected = true; }
            Check(rejected, "read only function");
            rejected = false;
            try { runtime.Invoke("function main(){ while(true){} }", "main", null, "Sheet1",
                (sheet, address) => data[address], null, null, null, new CancellationToken()); }
            catch { rejected = true; }
            Check(rejected, "statement limit");
        }
    }
}

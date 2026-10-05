using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using Jint;

namespace DinkCel
{
    // The JavaScript bridge exposes only JSON strings and delegates. No workbook or CLR
    // object is passed to the interpreter.
    internal sealed class DinkScript
    {
        private static readonly Regex FunctionMarker = new Regex(
            @"/\*\*(?:(?!\*/)[\s\S])*?@customfunction\b(?:(?!\*/)[\s\S])*?\*/\s*function\s+([A-Za-z_$][\w$]*)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();

        public static HashSet<string> CustomFunctions(string code)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in FunctionMarker.Matches(code ?? "")) names.Add(match.Groups[1].Value);
            return names;
        }

        public object Invoke(string code, string function, object[] args,
            string activeSheet, Func<string, string, object[][]> read,
            Action<string, string, object[][]> write, Action<string> log,
            Func<string, string> ai, CancellationToken cancellation = default(CancellationToken))
        {
            if (String.IsNullOrWhiteSpace(function) || !Regex.IsMatch(function, @"^[A-Za-z_$][\w$]*$"))
                throw new ArgumentException("Invalid function name.");
            var engine = new Engine(options => options.TimeoutInterval(TimeSpan.FromSeconds(5))
                .MaxStatements(100000).LimitMemory(32 * 1024 * 1024)
                .CancellationToken(cancellation));
            engine.SetValue("__activeSheet", activeSheet ?? "Sheet1");
            engine.SetValue("__read", new Func<string, string, string>((sheet, address) =>
                json.Serialize(read(sheet, address))));
            engine.SetValue("__write", new Action<string, string, string>((sheet, address, values) =>
                {
                    if (write == null) throw new InvalidOperationException("Custom functions cannot edit cells.");
                    write(sheet, address, json.Deserialize<object[][]>(values));
                }));
            engine.SetValue("__log", new Action<string>(message => { if (log != null) log(message); }));
            engine.SetValue("__ai", new Func<string, string>(prompt =>
                {
                    if (ai == null) throw new InvalidOperationException("AI is not enabled for this session.");
                    return ai(prompt);
                }));
            engine.Execute(@"
                var Logger = Object.freeze({log: function(x) { __log(String(x)); }});
                var AI = Object.freeze({generate: function(prompt) { return __ai(String(prompt)); }});
                var DinkCel = Object.freeze({getActiveWorkbook: function() { return Object.freeze({
                    getActiveSheet: function() { return this.getSheetByName(__activeSheet); },
                    getSheetByName: function(name) { return Object.freeze({
                        getRange: function(address) { return Object.freeze({
                            getValues: function() { return JSON.parse(__read(String(name), String(address))); },
                            getValue: function() { return this.getValues()[0][0]; },
                            setValues: function(values) { __write(String(name), String(address), JSON.stringify(values)); },
                            setValue: function(value) { this.setValues([[value]]); }
                        }); }
                    }); }
                }); }});
            ");
            engine.Execute(code ?? "");
            object answer = engine.Invoke(function, args ?? new object[0]).ToObject();
            if (answer == null) return null;
            if (answer is bool || answer is string || answer is double || answer is int || answer is long)
                return answer;
            throw new InvalidOperationException("Functions must return one number, text, boolean or blank value.");
        }
    }
}

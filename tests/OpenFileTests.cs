using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class OpenFileTests
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string directory = Path.Combine(Path.GetTempPath(),
                "DinkCelOpenTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string csv = Path.Combine(directory, "sample.csv");
                File.WriteAllText(csv, "name,value\r\n\"A,B\",42\r\n", new UTF8Encoding(true));
                using (var form = new SpreadsheetForm(csv))
                {
                    form.Show();
                    Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    Equal("name", grid[0, 0].Value);
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(csv, Field(form, "currentPath"));
                    grid.ClearSelection();
                    grid[0, 1].Selected = true;
                    grid[1, 1].Selected = true;
                    Invoke(form, "GridKeyDown", grid,
                        new KeyEventArgs(Keys.Delete));
                    Equal(null, grid[0, 1].Value);
                    Equal(null, grid[1, 1].Value);
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(false, Field(form, "dirty"));
                    Call(form, "Redo");
                    Equal(null, grid[0, 1].Value);
                    Equal(null, grid[1, 1].Value);
                    Equal(true, Field(form, "dirty"));
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(false, Field(form, "dirty"));
                    grid[4, 0].Value = "new";
                    Call(form, "Undo");
                    Equal(null, grid[4, 0].Value);
                    Call(form, "Redo");
                    Equal("new", grid[4, 0].Value);
                    Call(form, "Undo");
                    grid[2, 1].Value = "=B2";
                    grid[0, 1].Style.BackColor = Color.Yellow;
                    Call(form, "SelectHeader", 1, true);
                    Equal(26, grid.SelectedCells.Count);
                    Call(form, "InsertRow");
                    Equal(null, grid[0, 1].Value);
                    Equal("A,B", grid[0, 2].Value);
                    Equal("=B3", grid[2, 2].Value);
                    Equal(Color.Yellow, grid[0, 2].Style.BackColor);
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("=B2", grid[2, 1].Value);
                    Call(form, "Redo");
                    Equal("A,B", grid[0, 2].Value);
                    Equal("=B3", grid[2, 2].Value);
                    Call(form, "DeleteRow");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("=B2", grid[2, 1].Value);

                    grid[3, 0].Value = "=$B$2";
                    Call(form, "SelectHeader", 1, false);
                    Equal(200, grid.SelectedCells.Count);
                    Call(form, "InsertColumn");
                    Equal(null, grid[1, 1].Value);
                    Equal("42", grid[2, 1].Value);
                    Equal("=$C$2", grid[4, 0].Value);
                    Call(form, "DeleteColumn");
                    Equal("42", grid[1, 1].Value);
                    Equal("=$B$2", grid[3, 0].Value);
                    grid[0, 1].Value = "A,B\r\n\"Z\"";
                    Equal(true, Invoke(form, "SaveDocument"));
                    Equal(csv, Field(form, "currentPath"));
                    Equal(false, Field(form, "dirty"));
                    grid[0, 0].Value = "changed";
                    Call(form, "Undo");
                    Equal("name", grid[0, 0].Value);
                    Equal(false, Field(form, "dirty"));
                    var csvRows = CsvFile.Read(csv, 200, 26);
                    Equal("A,B\r\n\"Z\"", csvRows[1][0]);
                    Equal("=B2", csvRows[1][2]);
                    Equal("=$B$2", csvRows[0][3]);
                    byte[] bytes = File.ReadAllBytes(csv);
                    Equal((byte)0xEF, bytes[0]);
                    Equal((byte)0xBB, bytes[1]);
                    Equal((byte)0xBF, bytes[2]);
                    string edited = Path.Combine(directory, "edited.dinkcel");
                    Equal(true, Invoke(form, "WriteWorkbook", edited));
                    form.Close();
                    using (var reopened = new SpreadsheetForm(edited))
                    {
                        reopened.Show();
                        Application.DoEvents();
                        var reopenedGrid = (DataGridView)Field(reopened, "grid");
                        Equal("A,B\n\"Z\"", reopenedGrid[0, 1].Value);
                        Equal("=$B$2", reopenedGrid[3, 0].Value);
                        Equal(Color.Yellow, reopenedGrid[0, 1].Style.BackColor);
                        reopened.Close();
                    }
                }

                string native = Path.Combine(directory, "old.dinkcel");
                File.WriteAllText(native,
                    "<workbook rows=\"20\" columns=\"10\"><cell row=\"2\" column=\"3\">old</cell></workbook>");
                using (var form = new SpreadsheetForm(native))
                {
                    form.Show();
                    Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    Equal("old", grid[2, 1].Value);
                    Equal(native, Field(form, "currentPath"));
                    form.Close();
                }
                Console.WriteLine("Open file: CSV and older .dinkcel passed.");
            }
            finally
            {
                Directory.Delete(directory, true);
            }
        }

        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name,
                BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        }

        private static void Call(object target, string name, params object[] args)
        {
            Invoke(target, name, args);
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name,
                BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args);
        }

        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual))
                throw new Exception("Expected " + expected + ", got " + actual);
        }
    }
}

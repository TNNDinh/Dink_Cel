using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private readonly Dictionary<ChartDefinition, Panel> chartOverlays =
            new Dictionary<ChartDefinition, Panel>();
        private bool chartOverlayEvents;
        private bool chartRefreshQueued;

        private void QueueChartRefresh(int row, int column)
        {
            if (loading || chartRefreshQueued || !charts.Any(c => c.Range.Contains(column, row))) return;
            chartRefreshQueued = true;
            BeginInvoke((Action)delegate
            {
                chartRefreshQueued = false;
                if (!IsDisposed) RefreshChartOverlays();
            });
        }

        private Rectangle ChartBounds(ChartDefinition definition)
        {
            Rectangle place = definition.Placement.IsEmpty ?
                new Rectangle(Math.Min(ColumnCount - 8, definition.Range.Right + 1),
                    definition.Range.Top, 8, 14) : definition.Placement;
            int left = Math.Max(0, Math.Min(ColumnCount - 1, place.Left));
            int top = Math.Max(0, Math.Min(RowCount - 1, place.Top));
            int right = Math.Max(left, Math.Min(ColumnCount - 1, place.Right - 1));
            int bottom = Math.Max(top, Math.Min(RowCount - 1, place.Bottom - 1));
            Rectangle first = grid.GetCellDisplayRectangle(left, top, false);
            Rectangle last = grid.GetCellDisplayRectangle(right, bottom, false);
            return Rectangle.FromLTRB(first.Left, first.Top, Math.Max(first.Left + 160, last.Right),
                Math.Max(first.Top + 110, last.Bottom));
        }

        private void PositionChartOverlays()
        {
            foreach (var entry in chartOverlays)
            {
                Rectangle bounds = ChartBounds(entry.Key);
                entry.Value.Bounds = bounds;
                entry.Value.Visible = bounds.IntersectsWith(grid.ClientRectangle);
            }
        }

        private void RefreshChartOverlays()
        {
            if (!chartOverlayEvents)
            {
                chartOverlayEvents = true;
                grid.Scroll += delegate { PositionChartOverlays(); };
                grid.ColumnWidthChanged += delegate { PositionChartOverlays(); };
                grid.RowHeightChanged += delegate { PositionChartOverlays(); };
                grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
                { if (e.RowIndex >= 0 && e.ColumnIndex >= 0) QueueChartRefresh(e.RowIndex, e.ColumnIndex); };
            }
            foreach (Panel old in chartOverlays.Values)
            { grid.Controls.Remove(old); old.Dispose(); }
            chartOverlays.Clear();
            foreach (ChartDefinition definition in charts)
            {
                var panel = new Panel { BackColor = theme.Sheet, BorderStyle = BorderStyle.FixedSingle };
                var title = new Label { Text = "☰  " + definition.Title + "   (kéo để di chuyển)",
                    Dock = DockStyle.Top, Height = 25, BackColor = theme.AccentSoft,
                    ForeColor = theme.Text, TextAlign = ContentAlignment.MiddleLeft, Cursor = Cursors.SizeAll };
                var grip = new Label { Text = "◢", Dock = DockStyle.Bottom, Height = 18,
                    TextAlign = ContentAlignment.MiddleRight, BackColor = theme.Chrome,
                    ForeColor = theme.Text, Cursor = Cursors.SizeNWSE };
                var chart = BuildChart(definition);
                chart.DoubleClick += delegate { ShowChart(definition); };
                title.DoubleClick += delegate { ShowChart(definition); };
                Point last = Point.Empty;
                title.MouseDown += delegate(object sender, MouseEventArgs e)
                { if (e.Button == MouseButtons.Left) last = Control.MousePosition; };
                title.MouseMove += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left || last.IsEmpty) return;
                    Point now = Control.MousePosition;
                    panel.Left += now.X - last.X; panel.Top += now.Y - last.Y; last = now;
                };
                title.MouseUp += delegate
                {
                    if (last.IsEmpty) return;
                    last = Point.Empty;
                    DataGridView.HitTestInfo hit = grid.HitTest(panel.Left + 5, panel.Top + 5);
                    if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
                    {
                        Rectangle place = definition.Placement.IsEmpty ?
                            new Rectangle(0, 0, 8, 14) : definition.Placement;
                        definition.Placement = new Rectangle(Math.Min(ColumnCount - place.Width,
                            hit.ColumnIndex), hit.RowIndex,
                            place.Width, place.Height);
                        RecordChange(); MarkDirty();
                    }
                    PositionChartOverlays();
                };
                grip.MouseDown += delegate(object sender, MouseEventArgs e)
                { if (e.Button == MouseButtons.Left) last = Control.MousePosition; };
                grip.MouseMove += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left || last.IsEmpty) return;
                    Point now = Control.MousePosition;
                    panel.Width = Math.Max(160, panel.Width + now.X - last.X);
                    panel.Height = Math.Max(110, panel.Height + now.Y - last.Y);
                    last = now;
                };
                grip.MouseUp += delegate
                {
                    if (last.IsEmpty) return;
                    last = Point.Empty;
                    DataGridView.HitTestInfo hit = grid.HitTest(panel.Right - 5, panel.Bottom - 5);
                    if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
                    {
                        Rectangle place = definition.Placement.IsEmpty ?
                            new Rectangle(0, 0, 8, 14) : definition.Placement;
                        definition.Placement = Rectangle.FromLTRB(place.Left, place.Top,
                            Math.Min(ColumnCount, Math.Max(place.Left + 2, hit.ColumnIndex + 1)),
                            Math.Max(place.Top + 3, hit.RowIndex + 1));
                        RecordChange(); MarkDirty();
                    }
                    PositionChartOverlays();
                };
                panel.Controls.Add(chart); panel.Controls.Add(title); panel.Controls.Add(grip);
                grid.Controls.Add(panel); chartOverlays[definition] = panel;
                panel.BringToFront();
            }
            PositionChartOverlays();
        }

        private readonly Dictionary<SheetObject, Panel> objectOverlays =
            new Dictionary<SheetObject, Panel>();
        private bool objectOverlayEvents;

        private Rectangle ObjectBounds(SheetObject item)
        {
            Rectangle place = item.Placement;
            int left = Math.Max(0, Math.Min(ColumnCount - 1, place.Left));
            int top = Math.Max(0, Math.Min(RowCount - 1, place.Top));
            int right = Math.Max(left, Math.Min(ColumnCount - 1, place.Right - 1));
            int bottom = Math.Max(top, Math.Min(RowCount - 1, place.Bottom - 1));
            Rectangle a = grid.GetCellDisplayRectangle(left, top, false);
            Rectangle b = grid.GetCellDisplayRectangle(right, bottom, false);
            return Rectangle.FromLTRB(a.Left, a.Top, Math.Max(a.Left + 100, b.Right),
                Math.Max(a.Top + 70, b.Bottom));
        }

        private void PositionObjectOverlays()
        {
            foreach (var pair in objectOverlays)
            {
                pair.Value.Bounds = ObjectBounds(pair.Key);
                pair.Value.Visible = pair.Value.Bounds.IntersectsWith(grid.ClientRectangle);
            }
        }

        private void RefreshObjectOverlays()
        {
            if (!objectOverlayEvents)
            {
                objectOverlayEvents = true;
                grid.Scroll += delegate { PositionObjectOverlays(); };
                grid.ColumnWidthChanged += delegate { PositionObjectOverlays(); };
                grid.RowHeightChanged += delegate { PositionObjectOverlays(); };
            }
            foreach (Panel old in objectOverlays.Values)
            { grid.Controls.Remove(old); old.Dispose(); }
            objectOverlays.Clear();
            foreach (SheetObject item in objects)
            {
                var panel = new Panel { BackColor = theme.Sheet, BorderStyle = BorderStyle.FixedSingle };
                if (item.Kind == "Image")
                {
                    try
                    {
                        byte[] bytes = Convert.FromBase64String(item.ImageBase64);
                        using (var stream = new MemoryStream(bytes))
                        using (Image decoded = Image.FromStream(stream))
                        {
                            var picture = new PictureBox { Dock = DockStyle.Fill,
                                Image = new Bitmap(decoded), SizeMode = PictureBoxSizeMode.Zoom };
                            panel.Controls.Add(picture);
                            panel.Disposed += delegate { picture.Image.Dispose(); };
                        }
                    }
                    catch (Exception) { panel.Controls.Add(new Label { Text = "Image unavailable", Dock = DockStyle.Fill }); }
                }
                else
                {
                    var shape = new Panel { Dock = DockStyle.Fill, BackColor = item.Kind == "Ellipse" ?
                        Color.Transparent : item.Fill };
                    shape.Paint += delegate(object sender, PaintEventArgs e)
                    {
                        if (item.Kind != "Ellipse") return;
                        using (var brush = new SolidBrush(item.Fill))
                            e.Graphics.FillEllipse(brush, 3, 3, Math.Max(1, shape.Width - 7),
                                Math.Max(1, shape.Height - 7));
                    };
                    shape.Controls.Add(new Label { Text = item.Text, Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.Transparent,
                        ForeColor = Color.White });
                    panel.Controls.Add(shape);
                }
                var title = new Label { Text = item.Kind + "  ·  drag", Dock = DockStyle.Top,
                    Height = 22, BackColor = theme.AccentSoft, ForeColor = theme.Text,
                    Cursor = Cursors.SizeAll };
                var grip = new Label { Text = "◢", Dock = DockStyle.Bottom, Height = 16,
                    TextAlign = ContentAlignment.MiddleRight, BackColor = theme.Chrome,
                    ForeColor = theme.Text, Cursor = Cursors.SizeNWSE };
                Point last = Point.Empty;
                title.MouseDown += delegate(object sender, MouseEventArgs e)
                { if (e.Button == MouseButtons.Left) last = Control.MousePosition; };
                title.MouseMove += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left || last.IsEmpty || grid.ReadOnly) return;
                    Point now = Control.MousePosition;
                    panel.Left += now.X - last.X; panel.Top += now.Y - last.Y; last = now;
                };
                title.MouseUp += delegate
                {
                    if (last.IsEmpty) return;
                    last = Point.Empty;
                    DataGridView.HitTestInfo hit = grid.HitTest(panel.Left + 4, panel.Top + 4);
                    if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0 && !grid.ReadOnly)
                    {
                        item.Placement = new Rectangle(Math.Min(ColumnCount - item.Placement.Width,
                            hit.ColumnIndex), hit.RowIndex, item.Placement.Width, item.Placement.Height);
                        RecordChange(); MarkDirty();
                    }
                    PositionObjectOverlays();
                };
                grip.MouseDown += delegate(object sender, MouseEventArgs e)
                { if (e.Button == MouseButtons.Left) last = Control.MousePosition; };
                grip.MouseMove += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left || last.IsEmpty || grid.ReadOnly) return;
                    Point now = Control.MousePosition;
                    panel.Width = Math.Max(100, panel.Width + now.X - last.X);
                    panel.Height = Math.Max(70, panel.Height + now.Y - last.Y); last = now;
                };
                grip.MouseUp += delegate
                {
                    if (last.IsEmpty) return;
                    last = Point.Empty;
                    DataGridView.HitTestInfo hit = grid.HitTest(panel.Right - 4, panel.Bottom - 4);
                    if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0 && !grid.ReadOnly)
                    {
                        item.Placement = Rectangle.FromLTRB(item.Placement.Left, item.Placement.Top,
                            Math.Min(ColumnCount, Math.Max(item.Placement.Left + 1, hit.ColumnIndex + 1)),
                            Math.Max(item.Placement.Top + 1, hit.RowIndex + 1));
                        RecordChange(); MarkDirty();
                    }
                    PositionObjectOverlays();
                };
                var context = new ContextMenuStrip();
                context.Items.Add("Delete object", null, delegate
                {
                    if (grid.ReadOnly) return;
                    objects.Remove(item); RefreshObjectOverlays(); RecordChange(); MarkDirty();
                });
                panel.ContextMenuStrip = title.ContextMenuStrip = context;
                panel.Controls.Add(title); panel.Controls.Add(grip);
                grid.Controls.Add(panel); objectOverlays[item] = panel;
                panel.BringToFront();
            }
            PositionObjectOverlays();
        }

        private void InsertImageObject()
        {
            if (grid.ReadOnly || grid.CurrentCell == null) return;
            using (var dialog = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                byte[] bytes = File.ReadAllBytes(dialog.FileName);
                if (bytes.Length > 10 * 1024 * 1024)
                { MessageBox.Show(this, "Image must be 10 MB or smaller."); return; }
                objects.Add(new SheetObject { Kind = "Image", ImageBase64 = Convert.ToBase64String(bytes),
                    Placement = new Rectangle(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex, 5, 9) });
                RefreshObjectOverlays(); RecordChange(); MarkDirty();
            }
        }

        private void InsertShapeObject(string kind)
        {
            if (grid.ReadOnly || grid.CurrentCell == null) return;
            string label = Prompt("Shape text", "");
            if (label == null) return;
            objects.Add(new SheetObject { Kind = kind, Text = label,
                Placement = new Rectangle(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex, 4, 6) });
            RefreshObjectOverlays(); RecordChange(); MarkDirty();
        }
    }
}

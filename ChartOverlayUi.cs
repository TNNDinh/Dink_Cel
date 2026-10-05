using System;
using System.Collections.Generic;
using System.Drawing;
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
    }
}

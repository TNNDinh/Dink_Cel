using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class DinkDesign
    {
        public const int Small = 4;
        public const int Medium = 8;
        public const int Large = 12;
        public static readonly Font Ui = new Font("Segoe UI", 9F);
        public static readonly Font UiBold = new Font("Segoe UI", 9F, FontStyle.Bold);
        public static readonly Font Heading = new Font("Segoe UI", 11F, FontStyle.Bold);

        public static Button Button(string label, EventHandler click)
        {
            var button = new DinkButton { Text = label, Height = 28, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(7, 1, 7, 1),
                Margin = new Padding(3, 2, 3, 2), Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat, Font = Ui };
            button.FlatAppearance.BorderSize = 0;
            button.Click += click;
            return button;
        }
    }

    internal sealed class DinkButton : Button
    {
        public bool AccentUnderline;
        private bool hover;
        public DinkButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.UserPaint, true);
        }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new Rectangle(1, 1, Width - 3, Height - 3);
            using (var path = new GraphicsPath())
            {
                int radius = 7, d = radius * 2;
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                using (var brush = new SolidBrush(hover ? ControlPaint.Light(BackColor, 0.08F) : BackColor))
                    e.Graphics.FillPath(brush, path);
            }
            if (AccentUnderline)
                using (var pen = new Pen(ForeColor, 2F))
                    e.Graphics.DrawLine(pen, 7, Height - 2, Width - 7, Height - 2);
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis);
        }
    }

    internal sealed class DinkLogo : Label
    {
        public DinkLogo()
        {
            AccessibleName = "DinkCel";
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.UserPaint, true);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(Parent == null ? Color.Transparent : Parent.BackColor);
            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = new GraphicsPath())
            {
                path.AddArc(rect.Left, rect.Top, 9, 9, 180, 90);
                path.AddArc(rect.Right - 9, rect.Top, 9, 9, 270, 90);
                path.AddArc(rect.Right - 9, rect.Bottom - 9, 9, 9, 0, 90);
                path.AddArc(rect.Left, rect.Bottom - 9, 9, 9, 90, 90);
                path.CloseFigure();
                using (var brush = new SolidBrush(BackColor)) e.Graphics.FillPath(brush, path);
            }
            using (var pen = new Pen(Color.White, 2F))
            {
                pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round;
                e.Graphics.DrawLine(pen, 9, 7, 9, Height - 8);
                e.Graphics.DrawLine(pen, 9, 7, 16, 7);
                e.Graphics.DrawLine(pen, 9, Height - 8, 16, Height - 8);
                e.Graphics.DrawBezier(pen, 16, 7, 29, 7, 29, Height - 8, 16, Height - 8);
                e.Graphics.DrawLine(pen, 13, Height / 2, 22, Height / 2);
            }
        }
    }

    internal sealed class DinkMenuColors : ProfessionalColorTable
    {
        private readonly ThemePalette palette;
        public DinkMenuColors(ThemePalette palette) { this.palette = palette; UseSystemColors = false; }
        public override Color ToolStripDropDownBackground { get { return palette.Chrome; } }
        public override Color MenuBorder { get { return palette.Border; } }
        public override Color MenuItemBorder { get { return palette.AccentSoft; } }
        public override Color MenuItemSelected { get { return palette.Hover; } }
        public override Color MenuItemSelectedGradientBegin { get { return palette.Hover; } }
        public override Color MenuItemSelectedGradientEnd { get { return palette.Hover; } }
        public override Color MenuItemPressedGradientBegin { get { return palette.AccentSoft; } }
        public override Color MenuItemPressedGradientEnd { get { return palette.AccentSoft; } }
        public override Color ImageMarginGradientBegin { get { return palette.Chrome; } }
        public override Color ImageMarginGradientMiddle { get { return palette.Chrome; } }
        public override Color ImageMarginGradientEnd { get { return palette.Chrome; } }
        public override Color SeparatorDark { get { return palette.Border; } }
        public override Color SeparatorLight { get { return palette.Chrome; } }
        public override Color ButtonSelectedHighlight { get { return palette.Hover; } }
        public override Color ButtonSelectedBorder { get { return palette.Border; } }
    }

    internal sealed partial class SpreadsheetForm
    {
        private readonly Panel inspector = new Panel();
        private readonly FlowLayoutPanel inspectorContent = new FlowLayoutPanel();
        private readonly FlowLayoutPanel inspectorNavigation = new FlowLayoutPanel();
        private readonly Label propertyDetails = new Label();
        private readonly Label inspectorTitle = new Label();
        private readonly Label summaryLabel = new Label();
        private readonly FlowLayoutPanel floatingActions = new FlowLayoutPanel();
        private bool lastInputKeyboard;
        private readonly TrackBar zoomBar = new TrackBar();
        private readonly Label zoomLabel = new Label();
        private readonly Timer summaryTimer = new Timer();
        private bool minimalMode;
        private bool inspectorOpen;
        private int zoomPercent = 100;
        private bool clipboardFallback;
        private uint clipboardFallbackSequence;
        [DllImport("user32.dll")]
        private static extern uint GetClipboardSequenceNumber();
        private readonly ToolStripLabel contextLabel = new ToolStripLabel();
        private readonly ToolStripButton contextPrimary = new ToolStripButton();
        private readonly ToolStripButton contextSecondary = new ToolStripButton();
        private Action primaryCommand;
        private Action secondaryCommand;
        private readonly RichTextBox formulaPreview = new RichTextBox();
        private readonly Panel welcomePanel = new Panel();
        private readonly Panel welcomeCard = new Panel();
        private readonly ListBox recentList = new ListBox();

        private void BuildFormulaInput(TableLayoutPanel formula)
        {
            var host = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            host.Controls.Add(contentBox);
            formulaPreview.Dock = DockStyle.Fill;
            formulaPreview.BorderStyle = BorderStyle.FixedSingle;
            formulaPreview.ReadOnly = true;
            formulaPreview.Multiline = false;
            formulaPreview.ScrollBars = RichTextBoxScrollBars.None;
            formulaPreview.TabStop = false;
            formulaPreview.Font = new Font("Consolas", 10F);
            formulaPreview.Cursor = Cursors.IBeam;
            formulaPreview.Click += delegate { contentBox.Focus(); contentBox.SelectionStart = contentBox.TextLength; };
            host.Controls.Add(formulaPreview);
            formula.Controls.Add(host, 2, 0);
        }

        private void UpdateFormulaPreview()
        {
            if (formulaPreview.IsDisposed) return;
            string value = contentBox.Text;
            bool show = !contentBox.Focused && value.StartsWith("=", StringComparison.Ordinal);
            formulaPreview.Visible = show;
            if (!show || formulaPreview.Text == value) return;
            formulaPreview.Text = value;
            formulaPreview.SelectAll(); formulaPreview.SelectionColor = theme.Text;
            foreach (Match match in Regex.Matches(value,
                @"\b[A-Za-z_][A-Za-z0-9_]*(?=\s*\()"))
            { formulaPreview.Select(match.Index, match.Length); formulaPreview.SelectionColor = theme.Accent; }
            foreach (Match match in Regex.Matches(value,
                @"(?<![A-Za-z0-9_])\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?",
                RegexOptions.IgnoreCase))
            { formulaPreview.Select(match.Index, match.Length); formulaPreview.SelectionColor = theme.Dark ? Color.Turquoise : Color.Teal; }
            formulaPreview.Select(0, 0);
        }

        private void ConfigureCompactCommands()
        {
            toolbar.Items.Clear();
            AddToolbarButton(toolbar, "↶", "Hoàn tác (Ctrl+Z)", Undo);
            AddToolbarButton(toolbar, "↷", "Làm lại (Ctrl+Y)", Redo);
            AddToolbarButton(toolbar, "Lưu", "Lưu bảng tính (Ctrl+S)", delegate { SaveDocument(); });
            toolbar.Items.Add(new ToolStripSeparator());
            AddToolbarButton(toolbar, "Sao chép", "Sao chép ô đã chọn", CopySelected);
            AddToolbarButton(toolbar, "Dán", "Dán từ bộ nhớ tạm", PasteSelected);
            toolbar.Items.Add(new ToolStripSeparator());
            AddToolbarButton(toolbar, "Định dạng", "Mở bảng định dạng", ToggleInspector);
            AddToolbarButton(toolbar, "Sắp xếp", "Sắp xếp vùng chọn", delegate { SortRows(false); });
            AddToolbarButton(toolbar, "Lọc", "Lọc dữ liệu", SetFilter);
            AddToolbarButton(toolbar, "Biểu đồ", "Tạo biểu đồ", CreateChart);
            AddToolbarButton(toolbar, "⌕  Lệnh", "Tìm lệnh (Ctrl+K)", ShowCommandPalette);
            AddToolbarButton(toolbar, "⋯", "Thêm lệnh", ShowCommandPalette);
            toolbar.Items.Add(new ToolStripSeparator());
            contextLabel.ForeColor = theme.Muted;
            toolbar.Items.Add(contextLabel);
            contextPrimary.Click += delegate { if (primaryCommand != null) primaryCommand(); };
            contextSecondary.Click += delegate { if (secondaryCommand != null) secondaryCommand(); };
            toolbar.Items.Add(contextPrimary);
            toolbar.Items.Add(contextSecondary);
        }

        private void UpdateContextCommands()
        {
            if (grid.CurrentCell == null || sheets.Count <= activeSheetIndex) return;
            int row = grid.CurrentCell.RowIndex, col = grid.CurrentCell.ColumnIndex;
            PivotDefinition pivot = pivots.FirstOrDefault(p =>
                string.Equals(p.TargetSheet, sheets[activeSheetIndex].Name, StringComparison.OrdinalIgnoreCase));
            TableDefinition table = tables.FirstOrDefault(t => t.Range.Contains(col, row));
            ChartDefinition chart = charts.FirstOrDefault(c => c.Range.Contains(col, row));
            if (pivot != null)
            {
                contextLabel.Text = "Pivot";
                contextPrimary.Text = "Làm mới"; primaryCommand = delegate { RefreshPivot(pivot); };
                contextSecondary.Text = "Định dạng"; secondaryCommand = ToggleInspector;
            }
            else if (table != null)
            {
                contextLabel.Text = "Table";
                contextPrimary.Text = "Lọc"; primaryCommand = SetFilter;
                contextSecondary.Text = "Sắp xếp"; secondaryCommand = delegate { SortRows(false); };
            }
            else if (chart != null)
            {
                contextLabel.Text = "Biểu đồ";
                contextPrimary.Text = "Xem"; primaryCommand = delegate { ShowChart(chart); };
                contextSecondary.Text = "Định dạng"; secondaryCommand = ToggleInspector;
            }
            else if (grid.SelectedCells.Count > 1)
            {
                contextLabel.Text = "Vùng chọn";
                contextPrimary.Text = "Table"; primaryCommand = CreateTable;
                contextSecondary.Text = "Định dạng"; secondaryCommand = ToggleInspector;
            }
            else
            {
                contextLabel.Text = "Ô";
                contextPrimary.Text = "Màu ô"; primaryCommand = delegate { ChooseColor(true); };
                contextSecondary.Text = "Định dạng"; secondaryCommand = ToggleInspector;
            }
        }

        private void BuildVisualChrome(Panel header, Panel footer)
        {
            var actions = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 205,
                FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Padding = new Padding(0, 8, 3, 0) };
            actions.Controls.Add(DinkDesign.Button("⌂", delegate { ShowWelcome(); }));
            actions.Controls.Add(DinkDesign.Button("⌕", delegate { ShowQuickNavigator(); }));
            actions.Controls.Add(DinkDesign.Button("◐", delegate { ChooseTheme(); }));
            actions.Controls.Add(DinkDesign.Button("☰", delegate { ToggleMinimalMode(); }));
            header.Controls.Add(actions);

            summaryLabel.Text = "Sẵn sàng";
            summaryLabel.Font = DinkDesign.Ui;
            summaryLabel.TextAlign = ContentAlignment.MiddleLeft;
            summaryLabel.Bounds = new Rectangle(700, 3, 450, 28);
            footer.Controls.Add(summaryLabel);
            zoomBar.Minimum = 60;
            zoomBar.Maximum = 160;
            zoomBar.Value = 100;
            zoomBar.TickStyle = TickStyle.None;
            zoomBar.Width = 100;
            zoomBar.Height = 28;
            zoomBar.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            zoomBar.ValueChanged += delegate { ApplyZoom(zoomBar.Value); };
            footer.Controls.Add(zoomBar);
            zoomLabel.Text = "100%";
            zoomLabel.TextAlign = ContentAlignment.MiddleCenter;
            zoomLabel.Bounds = new Rectangle(0, 4, 45, 25);
            zoomLabel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            footer.Controls.Add(zoomLabel);
            footer.Resize += delegate
            {
                sheetTabs.Width = Math.Max(180, Math.Min(640, footer.ClientSize.Width / 2));
                summaryLabel.Left = sheetTabs.Right + 10;
                summaryLabel.Width = Math.Max(60, status.Left - summaryLabel.Left - 10);
                zoomBar.Left = footer.ClientSize.Width - 105;
                zoomLabel.Left = zoomBar.Left - 43;
            };
            summaryTimer.Interval = 110;
            summaryTimer.Tick += delegate { summaryTimer.Stop(); UpdateSelectionSummary(); UpdateFloatingActions(); };
        }

        private void BuildFloatingActions(Panel workspace)
        {
            floatingActions.Size = new Size(225, 34);
            floatingActions.FlowDirection = FlowDirection.LeftToRight;
            floatingActions.WrapContents = false;
            floatingActions.Padding = new Padding(3, 1, 3, 1);
            floatingActions.Visible = false;
            floatingActions.BorderStyle = BorderStyle.FixedSingle;
            floatingActions.Controls.Add(DinkDesign.Button("B", delegate { ToggleFontStyle(FontStyle.Bold); }));
            floatingActions.Controls.Add(DinkDesign.Button("I", delegate { ToggleFontStyle(FontStyle.Italic); }));
            floatingActions.Controls.Add(DinkDesign.Button("U", delegate { ToggleFontStyle(FontStyle.Underline); }));
            floatingActions.Controls.Add(DinkDesign.Button("Màu", delegate { ChooseColor(true); }));
            floatingActions.Controls.Add(DinkDesign.Button("⋯", delegate { ToggleInspector(); }));
            workspace.Controls.Add(floatingActions);
            grid.MouseDown += delegate { lastInputKeyboard = false; };
            grid.KeyDown += delegate { lastInputKeyboard = true; floatingActions.Visible = false; };
            grid.Scroll += delegate { floatingActions.Visible = false; };
            grid.CellBeginEdit += delegate { floatingActions.Visible = false; };
        }

        private void UpdateFloatingActions()
        {
            if (lastInputKeyboard || minimalMode || grid.CurrentCell == null || grid.SelectedCells.Count < 2 ||
                grid.IsCurrentCellInEditMode || !grid.Visible || welcomePanel.Visible)
            { floatingActions.Visible = false; return; }
            Rectangle range = SelectedRectangle(grid, false);
            Rectangle first = grid.GetCellDisplayRectangle(range.Left, range.Top, true);
            Rectangle last = grid.GetCellDisplayRectangle(range.Right - 1, range.Bottom - 1, true);
            if (first.IsEmpty || last.IsEmpty)
            { floatingActions.Visible = false; return; }
            floatingActions.Left = Math.Max(4, Math.Min(grid.Right - floatingActions.Width - 8, last.Right + 6));
            floatingActions.Top = first.Top > 40 ? first.Top - floatingActions.Height - 3 :
                Math.Min(grid.Bottom - floatingActions.Height - 4, last.Bottom + 4);
            floatingActions.Visible = true;
            floatingActions.BringToFront();
        }

        private void BuildInspector(Panel workspace)
        {
            inspector.Dock = DockStyle.Right;
            inspector.Width = 250;
            inspector.Visible = false;
            inspector.Padding = new Padding(12, 10, 12, 8);
            inspectorContent.Dock = DockStyle.Fill;
            inspectorContent.AutoScroll = true;
            inspectorContent.FlowDirection = FlowDirection.TopDown;
            inspectorContent.WrapContents = false;
            inspector.Controls.Add(inspectorContent);
            var heading = new Panel { Dock = DockStyle.Top, Height = 38 };
            var title = inspectorTitle;
            title.Text = "Định dạng";
            title.Dock = DockStyle.Fill;
            title.Font = DinkDesign.Heading;
            title.TextAlign = ContentAlignment.MiddleLeft;
            var close = DinkDesign.Button("×", delegate { ToggleInspector(); });
            close.Dock = DockStyle.Right;
            heading.Controls.Add(title);
            heading.Controls.Add(close);
            inspector.Controls.Add(heading);
            workspace.Controls.Add(inspector);
            workspace.Resize += delegate
            {
                if (workspace.ClientSize.Width < 1050)
                {
                    inspector.Dock = DockStyle.None;
                    inspector.Bounds = new Rectangle(Math.Max(0, workspace.ClientSize.Width - 250),
                        0, 250, workspace.ClientSize.Height);
                    inspector.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
                }
                else
                {
                    inspector.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
                    inspector.Dock = DockStyle.Right;
                }
                LayoutWorkspace(workspace);
                if (inspector.Visible) inspector.BringToFront();
            };
            inspectorNavigation.Dock = DockStyle.Top;
            inspectorNavigation.Height = 76;
            inspectorNavigation.WrapContents = true;
            inspectorNavigation.FlowDirection = FlowDirection.LeftToRight;
            inspector.Controls.Add(inspectorNavigation);
            heading.BringToFront();
            foreach (string page in new[] { "Ô", "Format", "Table", "Chart", "Dropdown", "Điều kiện" })
            {
                string selected = page;
                inspectorNavigation.Controls.Add(DinkDesign.Button(page,
                    delegate { ShowInspectorPage(selected); }));
            }
            ShowInspectorPage("Format");
        }

        private void ShowInspectorPage(string page)
        {
            inspectorTitle.Text = page == "Ô" ? "Thuộc tính" : page;
            foreach (Control child in inspectorContent.Controls.Cast<Control>().ToArray())
            {
                inspectorContent.Controls.Remove(child);
                if (child != propertyDetails) child.Dispose();
            }
            if (page == "Ô")
            {
                InspectorSection("Ô hiện tại");
                propertyDetails.Width = 215;
                propertyDetails.Height = 95;
                propertyDetails.Font = DinkDesign.Ui;
                inspectorContent.Controls.Add(propertyDetails);
                UpdateInspectorProperties();
                InspectorRow(new[] { "Đi tới", "Tìm" }, new Action[] { ShowQuickNavigator, FindReplace });
                return;
            }
            if (page == "Table")
            {
                InspectorSection("Table");
                InspectorRow(new[] { "Tạo Table" }, new Action[] { CreateTable });
                InspectorRow(new[] { "Sắp xếp", "Lọc" }, new Action[] {
                    delegate { SortRows(false); }, SetFilter });
                InspectorRow(new[] { "Bỏ lọc" }, new Action[] {
                    delegate { filterColumn = -1; filterValue = ""; ApplyFreezeAndFilter(); } });
                return;
            }
            if (page == "Chart")
            {
                InspectorSection("Biểu đồ");
                InspectorRow(new[] { "Tạo biểu đồ", "Xem" }, new Action[] { CreateChart, OpenChart });
                return;
            }
            if (page == "Dropdown")
            {
                InspectorSection("Data validation");
                InspectorRow(new[] { "Thêm danh sách chọn" }, new Action[] { AddDropdown });
                return;
            }
            if (page == "Điều kiện")
            {
                InspectorSection("Conditional formatting");
                InspectorRow(new[] { "Tô màu theo điều kiện" }, new Action[] { ConditionalColor });
                return;
            }
            InspectorSection("Chữ");
            InspectorRow(new[] { "B", "I", "U", "S" }, new Action[] {
                delegate { ToggleFontStyle(FontStyle.Bold); },
                delegate { ToggleFontStyle(FontStyle.Italic); },
                delegate { ToggleFontStyle(FontStyle.Underline); },
                delegate { ToggleFontStyle(FontStyle.Strikeout); } });
            InspectorRow(new[] { "Phông chữ", "Cỡ chữ" }, new Action[] { ChooseFontFamily,
                delegate { ChooseFontSize(); } });
            InspectorSection("Màu và căn chỉnh");
            InspectorRow(new[] { "Màu chữ", "Màu ô" }, new Action[] {
                delegate { ChooseColor(false); }, delegate { ChooseColor(true); } });
            InspectorRow(new[] { "Trái", "Giữa", "Phải" }, new Action[] {
                delegate { ApplyAlignment(DataGridViewContentAlignment.MiddleLeft); },
                delegate { ApplyAlignment(DataGridViewContentAlignment.MiddleCenter); },
                delegate { ApplyAlignment(DataGridViewContentAlignment.MiddleRight); } });
            InspectorRow(new[] { "Xuống dòng", "Viền ô" }, new Action[] {
                ToggleWrap, delegate { SetBorders("Tất cả"); } });
            InspectorSection("Định dạng số");
            InspectorRow(new[] { "Số", "Tiền", "%", "Ngày" }, new Action[] {
                delegate { ApplyNumberFormat("Number"); },
                delegate { ApplyNumberFormat("Currency"); },
                delegate { ApplyNumberFormat("Percentage"); },
                delegate { ApplyNumberFormat("Date"); } });
            InspectorRow(new[] { "Xóa định dạng" }, new Action[] { ClearFormats });
        }

        private void UpdateInspectorProperties()
        {
            if (grid.CurrentCell == null) return;
            propertyDetails.Text = "Địa chỉ  " + CellAddress(grid.CurrentCell.ColumnIndex,
                grid.CurrentCell.RowIndex) + "\r\nNội dung  " +
                (Convert.ToString(grid.CurrentCell.Value) ?? "");
        }

        private void BuildWelcome(Panel workspace)
        {
            welcomePanel.Dock = DockStyle.Fill;
            welcomePanel.Visible = false;
            welcomeCard.Size = new Size(540, 370);
            welcomePanel.Controls.Add(welcomeCard);
            welcomePanel.Resize += delegate
            {
                welcomeCard.Left = Math.Max(16, (welcomePanel.Width - welcomeCard.Width) / 2);
                welcomeCard.Top = Math.Max(12, (welcomePanel.Height - welcomeCard.Height) / 2);
            };
            var heading = new Label { Text = "DinkCel", Font = new Font("Segoe UI", 22F, FontStyle.Bold),
                Bounds = new Rectangle(22, 18, 490, 50) };
            var subtitle = new Label { Text = "Bắt đầu với dữ liệu của bạn", Font = DinkDesign.Heading,
                Bounds = new Rectangle(24, 75, 490, 29) };
            var newButton = DinkDesign.Button("＋  Bảng tính mới", delegate { NewDocument(); });
            newButton.Bounds = new Rectangle(22, 125, 220, 44);
            newButton.AutoSize = false;
            var openButton = DinkDesign.Button("▱  Mở tệp", delegate { OpenDocument(); });
            openButton.Bounds = new Rectangle(253, 125, 220, 44);
            openButton.AutoSize = false;
            var recentHeading = new Label { Text = "GẦN ĐÂY", Font = DinkDesign.UiBold,
                Bounds = new Rectangle(24, 195, 450, 24) };
            recentList.Bounds = new Rectangle(22, 224, 490, 115);
            recentList.BorderStyle = BorderStyle.None;
            recentList.Font = DinkDesign.Ui;
            recentList.DoubleClick += delegate
            {
                string path = recentList.SelectedItem as string;
                if (path != null) OpenPath(path);
            };
            welcomeCard.Controls.Add(heading);
            welcomeCard.Controls.Add(subtitle);
            welcomeCard.Controls.Add(newButton);
            welcomeCard.Controls.Add(openButton);
            welcomeCard.Controls.Add(recentHeading);
            welcomeCard.Controls.Add(recentList);
            workspace.Controls.Add(welcomePanel);
        }

        private string RecentFilesPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DinkCel", "recent.txt");
        }

        private void RecordRecentFile(string path)
        {
            try
            {
                string location = RecentFilesPath();
                var paths = File.Exists(location) ? File.ReadAllLines(location).ToList() : new List<string>();
                paths.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
                paths.Insert(0, path);
                Directory.CreateDirectory(Path.GetDirectoryName(location));
                File.WriteAllLines(location, paths.Take(10).ToArray());
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void ShowWelcome()
        {
            floatingActions.Visible = false;
            recentList.Items.Clear();
            try
            {
                string location = RecentFilesPath();
                if (File.Exists(location))
                    foreach (string path in File.ReadAllLines(location).Where(File.Exists).Take(8))
                        recentList.Items.Add(path);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            welcomePanel.Visible = true;
            welcomePanel.BringToFront();
            ApplyVisualTheme();
        }

        private void HideWelcome()
        {
            welcomePanel.Visible = false;
            grid.Focus();
        }

        private void ChooseFontSize()
        {
            string input = Prompt("Cỡ chữ", grid.CurrentCell == null ? "10" :
                (grid.CurrentCell.InheritedStyle.Font ?? grid.Font).Size.ToString(CultureInfo.InvariantCulture));
            float size;
            if (float.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out size) && size >= 6 && size <= 72)
                ApplyFontSize(size);
        }

        private void InspectorSection(string title)
        {
            var label = new Label { Text = title.ToUpperInvariant(), AutoSize = false,
                Width = 210, Height = 31, TextAlign = ContentAlignment.BottomLeft,
                Font = DinkDesign.UiBold };
            inspectorContent.Controls.Add(label);
        }

        private void InspectorRow(string[] labels, Action[] actions)
        {
            var row = new FlowLayoutPanel { Width = 218, Height = 35, WrapContents = false,
                Margin = Padding.Empty };
            for (int i = 0; i < labels.Length; i++)
            {
                Action action = actions[i];
                row.Controls.Add(DinkDesign.Button(labels[i], delegate { action(); }));
            }
            inspectorContent.Controls.Add(row);
        }

        private void ToggleInspector()
        {
            floatingActions.Visible = false;
            inspectorOpen = !inspectorOpen;
            inspector.Visible = inspectorOpen && !minimalMode;
            if (inspector.Visible) inspector.BringToFront();
            LayoutWorkspace(inspector.Parent as Panel);
            grid.Focus();
        }

        private void LayoutWorkspace(Panel workspace)
        {
            if (workspace == null) return;
            if (inspector.Visible && workspace.ClientSize.Width >= 1050)
            {
                grid.Dock = DockStyle.None;
                grid.Bounds = new Rectangle(0, 0, Math.Max(100, workspace.ClientSize.Width - inspector.Width),
                    workspace.ClientSize.Height);
            }
            else grid.Dock = DockStyle.Fill;
        }

        private void ToggleMinimalMode()
        {
            floatingActions.Visible = false;
            minimalMode = !minimalMode;
            var layout = headerPanel.Parent as TableLayoutPanel;
            if (layout != null)
            {
                layout.RowStyles[0].Height = minimalMode ? 0 : 48;
                layout.RowStyles[1].Height = minimalMode ? 0 : 28;
                layout.RowStyles[2].Height = minimalMode ? 0 : 38;
            }
            menu.Visible = toolbar.Visible = footerPanel.Visible = !minimalMode;
            headerPanel.Visible = !minimalMode;
            inspector.Visible = inspectorOpen && !minimalMode;
            LayoutWorkspace(inspector.Parent as Panel);
            // Sheet tabs remain available while the status controls disappear.
            footerPanel.Visible = true;
            sheetTabs.Visible = true;
            status.Visible = summaryLabel.Visible = zoomBar.Visible = zoomLabel.Visible = !minimalMode;
            if (layout != null) layout.RowStyles[5].Height = minimalMode ? 32 : 36;
            grid.Focus();
        }

        private void ApplyVisualTheme()
        {
            var renderer = new ToolStripProfessionalRenderer(new DinkMenuColors(theme));
            menu.Renderer = renderer;
            toolbar.Renderer = renderer;
            rowContext.Renderer = renderer;
            columnContext.Renderer = renderer;
            inspector.BackColor = theme.Chrome;
            inspectorContent.BackColor = theme.Chrome;
            foreach (Control child in inspector.Controls)
            {
                child.BackColor = theme.Chrome;
                child.ForeColor = theme.Text;
            }
            inspectorNavigation.BackColor = theme.Chrome;
            foreach (Control button in inspectorNavigation.Controls)
            { button.BackColor = theme.Hover; button.ForeColor = theme.Text; }
            propertyDetails.ForeColor = theme.Text;
            foreach (Control child in inspectorContent.Controls)
            {
                child.BackColor = theme.Chrome;
                child.ForeColor = theme.Muted;
                foreach (Control button in child.Controls)
                {
                    button.BackColor = theme.Hover;
                    button.ForeColor = theme.Text;
                }
            }
            summaryLabel.ForeColor = theme.Muted;
            floatingActions.BackColor = theme.Chrome;
            foreach (Control button in floatingActions.Controls)
            { button.BackColor = theme.Hover; button.ForeColor = theme.Text; }
            zoomLabel.ForeColor = theme.Muted;
            zoomBar.BackColor = theme.Chrome;
            formulaPreview.BackColor = theme.Sheet;
            formulaPreview.ForeColor = theme.Text;
            formulaPreview.Text = "";
            UpdateFormulaPreview();
            welcomePanel.BackColor = theme.Surface;
            welcomeCard.BackColor = theme.Chrome;
            foreach (Control control in welcomeCard.Controls)
            {
                control.ForeColor = theme.Text;
                if (control is Button) control.BackColor = theme.AccentSoft;
            }
            recentList.BackColor = theme.Chrome;
            recentList.ForeColor = theme.Text;
            foreach (Control control in headerPanel.Controls)
                if (control is FlowLayoutPanel)
                {
                    control.BackColor = theme.Chrome;
                    foreach (Control action in control.Controls)
                    { action.BackColor = theme.Hover; action.ForeColor = theme.Text; }
                }
            RefreshSheetTabs();
        }

        private void ScheduleSelectionSummary()
        {
            summaryTimer.Stop();
            summaryTimer.Start();
        }

        private void UpdateSelectionSummary()
        {
            int count = grid.SelectedCells.Count;
            if (count <= 1) { summaryLabel.Text = "Sẵn sàng"; return; }
            double sum = 0, min = double.MaxValue, max = double.MinValue;
            int numbers = 0;
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                double value;
                string display = cell.FormattedValue == null ? "" : cell.FormattedValue.ToString();
                if (!double.TryParse(display, NumberStyles.Any, CultureInfo.CurrentCulture, out value) &&
                    !double.TryParse(Convert.ToString(cell.Value), NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                    continue;
                sum += value; min = Math.Min(min, value); max = Math.Max(max, value); numbers++;
            }
            summaryLabel.Text = count + " ô" + (numbers == 0 ? "" :
                "   Tổng " + sum.ToString("N2") + "   TB " + (sum / numbers).ToString("N2") +
                "   Min " + min.ToString("N2") + "   Max " + max.ToString("N2"));
        }

        private void ApplyZoom(int percent)
        {
            if (percent == zoomPercent) return;
            float scale = percent / (float)zoomPercent;
            zoomPercent = percent;
            zoomLabel.Text = percent + "%";
            bool wasLoading = loading;
            loading = true;
            try
            {
                grid.Font = new Font("Segoe UI", Math.Max(7F, 9F * percent / 100F));
                foreach (DataGridViewColumn column in grid.Columns)
                    column.Width = Math.Max(10, (int)Math.Round(column.Width * scale));
                foreach (DataGridViewRow row in grid.Rows)
                    row.Height = Math.Max(5, (int)Math.Round(row.Height * scale));
                grid.ColumnHeadersHeight = Math.Max(22, (int)Math.Round(28 * percent / 100F));
            }
            finally { loading = wasLoading; }
        }

        private sealed class PaletteEntry
        {
            public string Label;
            public Action Run;
            public override string ToString() { return Label; }
        }

        private List<PaletteEntry> PaletteEntries(bool navigation)
        {
            var entries = new List<PaletteEntry>();
            if (!navigation)
            {
                entries.Add(new PaletteEntry { Label = "Tệp · Mở", Run = OpenDocument });
                entries.Add(new PaletteEntry { Label = "Tệp · Lưu", Run = delegate { SaveDocument(); } });
                entries.Add(new PaletteEntry { Label = "Xem · Đổi giao diện", Run = ChooseTheme });
                entries.Add(new PaletteEntry { Label = "Xem · Bảng định dạng", Run = ToggleInspector });
                entries.Add(new PaletteEntry { Label = "Xem · Chế độ tối giản", Run = ToggleMinimalMode });
                entries.Add(new PaletteEntry { Label = "Dữ liệu · Tạo Table", Run = CreateTable });
                entries.Add(new PaletteEntry { Label = "Dữ liệu · Tạo biểu đồ", Run = CreateChart });
                entries.Add(new PaletteEntry { Label = "Dữ liệu · Lọc", Run = SetFilter });
                entries.Add(new PaletteEntry { Label = "Ô · Cố định", Run = FreezeAtCell });
                entries.Add(new PaletteEntry { Label = "Ô · Bỏ cố định", Run = delegate { freezeRow = freezeColumn = 0; ApplyFreezeAndFilter(); } });
                entries.Add(new PaletteEntry { Label = "Ô · Định dạng số", Run = SetNumberFormat });
                entries.Add(new PaletteEntry { Label = "Ô · Tô màu có điều kiện", Run = ConditionalColor });
                entries.Add(new PaletteEntry { Label = "Sửa · Tìm và thay thế", Run = FindReplace });
                foreach (string function in new[] { "SUM", "AVERAGE", "MIN", "MAX", "COUNT",
                    "IF", "IFERROR", "SUMIFS", "COUNTIFS", "XLOOKUP", "VLOOKUP",
                    "HLOOKUP", "INDEX", "MATCH", "TODAY", "NOW", "DATE", "TEXT" })
                {
                    string name = function;
                    entries.Add(new PaletteEntry { Label = "Hàm · " + name,
                        Run = delegate { contentBox.Text = "=" + name + "(";
                            contentBox.Focus(); contentBox.SelectionStart = contentBox.TextLength; } });
                }
            }
            for (int i = 0; i < sheets.Count; i++)
            {
                int index = i;
                entries.Add(new PaletteEntry { Label = "Trang tính · " + sheets[i].Name,
                    Run = delegate { SwitchSheet(index); } });
            }
            foreach (NamedRange named in namedRanges)
            {
                NamedRange target = named;
                entries.Add(new PaletteEntry { Label = "Vùng · " + named.Name,
                    Run = delegate { GoToAddress(target.Name); } });
            }
            for (int i = 0; i < sheets.Count; i++)
            {
                int index = i;
                IEnumerable<TableDefinition> currentTables = i == activeSheetIndex ? tables : sheets[i].Tables;
                foreach (TableDefinition table in currentTables)
                {
                    Rectangle range = table.Range;
                    entries.Add(new PaletteEntry { Label = "Table · " + table.Name,
                        Run = delegate { SwitchSheet(index); GoToAddress(CellAddress(range.Left, range.Top)); } });
                }
                IEnumerable<ChartDefinition> currentCharts = i == activeSheetIndex ? charts : sheets[i].Charts;
                foreach (ChartDefinition chart in currentCharts)
                {
                    ChartDefinition target = chart;
                    entries.Add(new PaletteEntry { Label = "Biểu đồ · " + chart.Title,
                        Run = delegate { SwitchSheet(index); ShowChart(target); } });
                }
            }
            return entries;
        }

        private IEnumerable<PaletteEntry> SearchCells(string query)
        {
            if (query.Length < 2) yield break;
            int found = 0;
            for (int i = 0; i < sheets.Count && found < 12; i++)
            {
                int index = i;
                if (index == activeSheetIndex)
                {
                    for (int row = 0; row < RowCount && found < 12; row++)
                        for (int col = 0; col < ColumnCount && found < 12; col++)
                        {
                            string value = Convert.ToString(grid[col, row].Value);
                            if (string.IsNullOrEmpty(value) || value.IndexOf(query,
                                StringComparison.OrdinalIgnoreCase) < 0) continue;
                            string address = CellAddress(col, row);
                            found++;
                            yield return new PaletteEntry { Label = "Ô · " + sheets[i].Name + "!" + address + "  " + value,
                                Run = delegate { SwitchSheet(index); GoToAddress(address); } };
                        }
                }
                else foreach (var cell in sheets[i].Cells)
                {
                    if (found >= 12) break;
                    string value = Convert.ToString(cell.Value.Value);
                    if (string.IsNullOrEmpty(value) || value.IndexOf(query,
                        StringComparison.OrdinalIgnoreCase) < 0) continue;
                    string address = CellAddress(cell.Key % ColumnCount, cell.Key / ColumnCount);
                    found++;
                    yield return new PaletteEntry { Label = "Ô · " + sheets[i].Name + "!" + address + "  " + value,
                        Run = delegate { SwitchSheet(index); GoToAddress(address); } };
                }
            }
        }

        private void ShowCommandPalette() { ShowPalette(false); }
        private void ShowQuickNavigator() { ShowPalette(true); }

        private void ShowPalette(bool navigation)
        {
            using (var dialog = new Form { Text = navigation ? "Đi nhanh · DinkCel" : "Lệnh · DinkCel",
                StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false, MaximizeBox = false, ShowIcon = false,
                ClientSize = new Size(510, 390), BackColor = theme.Chrome, ForeColor = theme.Text,
                Font = DinkDesign.Ui })
            {
                var input = new TextBox { Bounds = new Rectangle(15, 15, 480, 30),
                    Font = new Font("Segoe UI", 12F), BackColor = theme.Sheet, ForeColor = theme.Text,
                    BorderStyle = BorderStyle.FixedSingle };
                var list = new ListBox { Bounds = new Rectangle(15, 54, 480, 320),
                    BorderStyle = BorderStyle.None, BackColor = theme.Chrome, ForeColor = theme.Text,
                    Font = new Font("Segoe UI", 10F), ItemHeight = 27 };
                List<PaletteEntry> entries = PaletteEntries(navigation);
                Action filter = delegate
                {
                    list.BeginUpdate(); list.Items.Clear();
                    foreach (PaletteEntry item in entries.Where(e => e.Label.IndexOf(input.Text,
                        StringComparison.OrdinalIgnoreCase) >= 0).Take(35)) list.Items.Add(item);
                    if (navigation)
                        foreach (PaletteEntry item in SearchCells(input.Text)) list.Items.Add(item);
                    list.EndUpdate();
                    if (list.Items.Count > 0) list.SelectedIndex = 0;
                };
                Action run = delegate
                {
                    if (navigation && GoToAddress(input.Text)) { dialog.Close(); return; }
                    PaletteEntry item = list.SelectedItem as PaletteEntry;
                    if (item == null) return;
                    dialog.Close();
                    BeginInvoke((Action)delegate { item.Run(); });
                };
                input.TextChanged += delegate { filter(); };
                input.KeyDown += delegate(object sender, KeyEventArgs e)
                {
                    if (e.KeyCode == Keys.Down && list.Items.Count > 0)
                    { list.SelectedIndex = Math.Min(list.Items.Count - 1, list.SelectedIndex + 1); e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.Up && list.Items.Count > 0)
                    { list.SelectedIndex = Math.Max(0, list.SelectedIndex - 1); e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.Enter) { run(); e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.Escape) { dialog.Close(); e.SuppressKeyPress = true; }
                };
                list.DoubleClick += delegate { run(); };
                dialog.Controls.Add(input); dialog.Controls.Add(list);
                filter();
                dialog.Shown += delegate { input.Focus(); };
                dialog.ShowDialog(this);
            }
        }
    }
}

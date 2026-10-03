using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            EmbeddedDependencies.Install();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SpreadsheetForm(args.Length > 0 ? args[0] : null));
        }
    }

    internal sealed class SmoothGrid : DataGridView
    {
        public Func<Keys, bool> ShortcutHandler;
        public SmoothGrid()
        {
            DoubleBuffered = true;
        }
        protected override bool ProcessDialogKey(Keys keyData)
        {
            return ShortcutHandler != null && ShortcutHandler(keyData) ||
                base.ProcessDialogKey(keyData);
        }
        protected override bool ProcessDataGridViewKey(KeyEventArgs e)
        {
            return ShortcutHandler != null && ShortcutHandler(e.KeyData) ||
                base.ProcessDataGridViewKey(e);
        }
    }

    internal sealed class CellSnapshot
    {
        public string Text = "";
        public string FontName = "Arial";
        public FontStyle FontStyle = FontStyle.Regular;
        public float FontSize = 10F;
        public Color ForeColor = Color.Empty;
        public Color BackColor = Color.Empty;
        public DataGridViewContentAlignment Alignment = DataGridViewContentAlignment.NotSet;
        public bool HasFont;
        public string NumberFormat = "";
        public CellExtras Extras;
    }

    internal sealed class ConditionalRule
    {
        public Rectangle Range;
        public double Threshold;
        public Color Color;
    }

    internal sealed class WorkbookSnapshot
    {
        public readonly List<SheetSnapshot> Sheets = new List<SheetSnapshot>();
        public readonly List<NamedRange> NamedRanges = new List<NamedRange>();
        public readonly List<PivotDefinition> Pivots = new List<PivotDefinition>();
        public Dictionary<int, CellSnapshot> Cells { get { return Sheets[0].Cells; } }
        public Color Background = Color.FromArgb(232, 240, 248);
        public bool HasBackground;
        public string ThemeId;
        public WorkbookSnapshot() { Sheets.Add(new SheetSnapshot()); }
    }

    internal sealed class SheetSnapshot
    {
        public string Name = "Sheet1";
        public bool Hidden;
        public Color TabColor = Color.Empty;
        public Color Background = Color.Empty;
        public string ThemeId;
        public readonly Dictionary<int, CellSnapshot> Cells = new Dictionary<int, CellSnapshot>();
        public readonly Dictionary<int, int> RowHeights = new Dictionary<int, int>();
        public readonly Dictionary<int, int> ColumnWidths = new Dictionary<int, int>();
        public readonly HashSet<int> HiddenRows = new HashSet<int>();
        public readonly HashSet<int> HiddenColumns = new HashSet<int>();
        public readonly List<Rectangle> Merges = new List<Rectangle>();
        public readonly List<ConditionalRule> Rules = new List<ConditionalRule>();
        public readonly List<TableDefinition> Tables = new List<TableDefinition>();
        public readonly List<ChartDefinition> Charts = new List<ChartDefinition>();
        public readonly List<ValidationRule> Validations = new List<ValidationRule>();
        public int FreezeRow;
        public int FreezeColumn;
        public int FilterColumn = -1;
        public string FilterValue = "";
    }

    internal sealed class SheetState
    {
        public string Name = "Sheet1";
        public bool Hidden;
        public Color TabColor = Color.Empty;
        public readonly List<Rectangle> Merges = new List<Rectangle>();
        public readonly List<ConditionalRule> Rules = new List<ConditionalRule>();
        public readonly List<TableDefinition> Tables = new List<TableDefinition>();
        public readonly List<ChartDefinition> Charts = new List<ChartDefinition>();
        public readonly List<ValidationRule> Validations = new List<ValidationRule>();
        public int FreezeRow;
        public int FreezeColumn;
        public int FilterColumn = -1;
        public string FilterValue = "";
        public readonly Dictionary<int, CellState> Cells =
            new Dictionary<int, CellState>();
        public readonly int[] RowHeights = new int[200];
        public readonly int[] ColumnWidths = new int[26];
        public readonly bool[] HiddenRows = new bool[200];
        public readonly bool[] HiddenColumns = new bool[26];
        public Color Background;
        public string ThemeId;
        public int CsvRows;
        public int CsvColumns;
        public long RevisionId;
    }

    internal sealed class SheetHistory
    {
        public readonly List<SheetState> Undo = new List<SheetState>();
        public readonly List<SheetState> Redo = new List<SheetState>();
        public SheetState Last;
        public long NextRevision;
        public long SavedRevision;
    }

    internal sealed class CellState
    {
        public object Value;
        public DataGridViewCellStyle Style;
        public CellExtras Extras;
    }

    internal sealed partial class SpreadsheetForm : Form
    {
        private const int RowCount = 200;
        private const int ColumnCount = 26;
        private static readonly Font HeaderFont = new Font("Segoe UI", 9F);

        private readonly SmoothGrid grid = new SmoothGrid();
        private readonly Panel headerPanel = new Panel();
        private readonly Label logo = new DinkLogo();
        private readonly Label subtitle = new Label();
        private readonly MenuStrip menu = new MenuStrip();
        private readonly ContextMenuStrip rowContext = new ContextMenuStrip();
        private readonly ContextMenuStrip columnContext = new ContextMenuStrip();
        private readonly ToolStrip toolbar = new ToolStrip();
        private readonly TableLayoutPanel formulaPanel = new TableLayoutPanel();
        private readonly Label fx = new Label();
        private readonly Panel footerPanel = new Panel();
        private readonly Label sheetTab = new Label();
        private readonly FlowLayoutPanel sheetTabs = new FlowLayoutPanel();
        private readonly List<SheetState> sheets = new List<SheetState>();
        private int activeSheetIndex;
        private readonly List<Rectangle> merges = new List<Rectangle>();
        private CellState[,] copiedCells;
        private string[,] copiedDisplays;
        private string copiedClipboardText;
        private readonly List<ConditionalRule> conditionalRules = new List<ConditionalRule>();
        private readonly List<TableDefinition> tables = new List<TableDefinition>();
        private readonly List<ChartDefinition> charts = new List<ChartDefinition>();
        private readonly List<ValidationRule> validations = new List<ValidationRule>();
        private readonly List<NamedRange> namedRanges = new List<NamedRange>();
        private readonly List<PivotDefinition> pivots = new List<PivotDefinition>();
        private int freezeRow;
        private int freezeColumn;
        private int filterColumn = -1;
        private string filterValue = "";
        private readonly ToolStripLabel themeSwatch = new ToolStripLabel("●");
        private readonly TextBox addressBox = new TextBox();
        private readonly TextBox contentBox = new TextBox();
        private readonly Label documentTitle = new Label();
        private readonly Label saveIndicator = new Label();
        private readonly Label status = new Label();
        private readonly ToolStripComboBox sizeCombo = new ToolStripComboBox();
        private readonly ToolStripButton boldButton = new ToolStripButton("B");
        private readonly ToolStripButton italicButton = new ToolStripButton("I");
        private readonly ToolStripButton underlineButton = new ToolStripButton("U");
        private readonly Dictionary<int, string> calculated =
            new Dictionary<int, string>();
        private FormulaEngine formulaEngine;
        private readonly List<SheetState> undoHistory = new List<SheetState>();
        private readonly List<SheetState> redoHistory = new List<SheetState>();
        private readonly List<SheetHistory> sheetHistories = new List<SheetHistory>();
        private SheetState lastState;
        private long nextRevision;
        private long savedRevision;

        private string currentPath;
        private CsvDocument csvDocument;
        private ThemePalette theme;
        private Color sheetBackground;
        private bool dirty;
        private bool otherSheetsDirty;
        private bool loading;
        private bool syncingContent;
        private bool syncingToolbar;
        private bool selectingHeader;
        private bool fillDragging;
        private bool headerDragPending;
        private bool headerDragging;
        private bool headerDragRow;
        private bool suppressHeaderClick;
        private int headerSource;
        private int headerTarget;
        private Point headerStartPoint;
        private int fillSourceRow;
        private int fillSourceColumn;
        private int fillTargetRow;
        private int fillTargetColumn;

        public SpreadsheetForm(string startupPath)
        {
            Text = "DinkCel";
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            Width = 1380;
            Height = 860;
            MinimumSize = new Size(800, 500);
            WindowState = FormWindowState.Maximized;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 10F);
            theme = LoadThemePreference();
            sheetBackground = theme.Sheet;
            BackColor = theme.Surface;

            loading = true;
            BuildInterface();
            BuildGrid();
            sheets.Add(new SheetState { Name = "Sheet1" });
            RefreshSheetTabs();
            loading = false;
            dirty = false;
            ApplyTheme(theme, false, false);
            Recalculate();
            UpdateTitle();
            UpdateSelection();
            ResetHistory();

            FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (!ConfirmDiscardChanges())
                    e.Cancel = true;
            };
            Shown += delegate
            {
                if (!string.IsNullOrEmpty(startupPath))
                    OpenPath(startupPath);
                else ShowWelcome();
                if (!welcomePanel.Visible) grid.Focus();
            };
        }

        private void BuildInterface()
        {
            var layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Margin = Padding.Empty;
            layout.Padding = Padding.Empty;
            layout.ColumnCount = 1;
            layout.RowCount = 6;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            Controls.Add(layout);

            Panel header = headerPanel;
            header.Dock = DockStyle.Fill;
            header.BackColor = theme.Chrome;
            logo.Text = "";
            logo.Font = new Font("Segoe UI", 17F, FontStyle.Bold);
            logo.ForeColor = Color.White;
            logo.BackColor = theme.Logo;
            logo.TextAlign = ContentAlignment.MiddleCenter;
            logo.Bounds = new Rectangle(12, 7, 32, 32);
            header.Controls.Add(logo);
            documentTitle.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            documentTitle.ForeColor = theme.Text;
            documentTitle.Location = new Point(56, 4);
            documentTitle.Size = new Size(440, 23);
            header.Controls.Add(documentTitle);
            subtitle.Text = "DinkCel  •  Bảng tính trên máy";
            subtitle.Font = new Font("Segoe UI", 8.5F);
            subtitle.ForeColor = theme.Muted;
            subtitle.Location = new Point(57, 25);
            subtitle.Size = new Size(400, 20);
            header.Controls.Add(subtitle);
            saveIndicator.TextAlign = ContentAlignment.MiddleRight;
            saveIndicator.ForeColor = theme.Muted;
            saveIndicator.Bounds = new Rectangle(850, 11, 125, 25);
            header.Controls.Add(saveIndicator);
            header.Resize += delegate
            {
                saveIndicator.Left = header.ClientSize.Width - saveIndicator.Width - 174;
                documentTitle.Width = Math.Max(200, saveIndicator.Left - documentTitle.Left - 12);
            };
            layout.Controls.Add(header, 0, 0);

            menu.Dock = DockStyle.Fill;
            menu.GripStyle = ToolStripGripStyle.Hidden;
            menu.BackColor = theme.Chrome;
            menu.ForeColor = theme.Text;
            menu.Font = new Font("Segoe UI", 10F);
            menu.Padding = new Padding(12, 0, 0, 0);
            var fileMenu = new ToolStripMenuItem("Tệp");
            AddMenuItem(fileMenu, "Mới", Keys.Control | Keys.N, NewDocument);
            AddMenuItem(fileMenu, "Mở...", Keys.Control | Keys.O, OpenDocument);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            AddMenuItem(fileMenu, "Lưu", Keys.Control | Keys.S, delegate { SaveDocument(); });
            AddMenuItem(fileMenu, "Lưu thành...", Keys.None, delegate { SaveDocumentAs(); });
            menu.Items.Add(fileMenu);
            var editMenu = new ToolStripMenuItem("Chỉnh sửa");
            AddMenuItem(editMenu, "Sao chép", Keys.Control | Keys.C, CopySelected);
            AddMenuItem(editMenu, "Cắt", Keys.Control | Keys.X, CutSelected);
            AddMenuItem(editMenu, "Dán", Keys.Control | Keys.V, PasteSelected);
            AddMenuItem(editMenu, "Dán giá trị", Keys.Control | Keys.Shift | Keys.V, PasteValues);
            AddMenuItem(editMenu, "Dán công thức", Keys.None, PasteFormulas);
            AddMenuItem(editMenu, "Dán định dạng", Keys.None, PasteFormats);
            AddMenuItem(editMenu, "Dán chuyển vị", Keys.None, PasteTranspose);
            AddMenuItem(editMenu, "Hoàn tác", Keys.Control | Keys.Z, Undo);
            AddMenuItem(editMenu, "Làm lại", Keys.Control | Keys.Y, Redo);
            AddMenuItem(editMenu, "Xóa nội dung ô đã chọn", Keys.None, ClearSelectedCells);
            editMenu.DropDownItems.Add(new ToolStripSeparator());
            AddMenuItem(editMenu, "Điền xuống", Keys.Control | Keys.D, FillDown);
            AddMenuItem(editMenu, "Điền sang phải", Keys.Control | Keys.R, FillRight);
            AddMenuItem(editMenu, "Điền chuỗi tăng", Keys.None, FillSeries);
            editMenu.DropDownItems.Add(new ToolStripSeparator());
            AddMenuItem(editMenu, "Chèn hàng phía trên", Keys.None, InsertRow);
            AddMenuItem(editMenu, "Xóa hàng", Keys.None, DeleteRow);
            AddMenuItem(editMenu, "Chèn cột bên trái", Keys.None, InsertColumn);
            AddMenuItem(editMenu, "Xóa cột", Keys.None, DeleteColumn);
            menu.Items.Add(editMenu);
            var formatMenu = new ToolStripMenuItem("Định dạng");
            AddMenuItem(formatMenu, "In đậm", Keys.Control | Keys.B,
                delegate { ToggleFontStyle(FontStyle.Bold); });
            AddMenuItem(formatMenu, "In nghiêng", Keys.Control | Keys.I,
                delegate { ToggleFontStyle(FontStyle.Italic); });
            AddMenuItem(formatMenu, "Gạch chân", Keys.Control | Keys.U,
                delegate { ToggleFontStyle(FontStyle.Underline); });
            AddFormattingMenus(formatMenu);
            menu.Items.Add(formatMenu);
            var viewMenu = new ToolStripMenuItem("Xem");
            AddMenuItem(viewMenu, "Đổi giao diện...", Keys.None, ChooseTheme);
            AddMenuItem(viewMenu, "Tìm lệnh...", Keys.Control | Keys.K, ShowCommandPalette);
            AddMenuItem(viewMenu, "Đi nhanh...", Keys.Control | Keys.Shift | Keys.P, ShowQuickNavigator);
            AddMenuItem(viewMenu, "Bảng Inspector", Keys.Control | Keys.Shift | Keys.I, ToggleInspector);
            AddMenuItem(viewMenu, "Chế độ tối giản", Keys.Control | Keys.Shift | Keys.M, ToggleMinimalMode);
            AddMenuItem(viewMenu, "Màn hình bắt đầu", Keys.None, ShowWelcome);
            menu.Items.Add(viewMenu);
            AddSpreadsheetMenus();
            var helpMenu = new ToolStripMenuItem("Trợ giúp");
            AddMenuItem(helpMenu, "Giấy phép thư viện...", Keys.None, ShowThirdPartyLicenses);
            menu.Items.Add(helpMenu);
            MainMenuStrip = menu;
            layout.Controls.Add(menu, 0, 1);

            toolbar.Dock = DockStyle.Fill;
            toolbar.GripStyle = ToolStripGripStyle.Hidden;
            toolbar.BackColor = theme.Chrome;
            toolbar.ForeColor = theme.Text;
            toolbar.Font = new Font("Segoe UI", 10F);
            toolbar.Padding = new Padding(10, 3, 8, 3);
            toolbar.RenderMode = ToolStripRenderMode.System;
            AddToolbarButton(toolbar, "Mở", "Mở bảng tính", OpenDocument);
            AddToolbarButton(toolbar, "Lưu", "Lưu bảng tính", delegate { SaveDocument(); });
            toolbar.Items.Add(new ToolStripSeparator());
            AddToolbarButton(toolbar, "Sao chép", "Sao chép ô đã chọn", CopySelected);
            AddToolbarButton(toolbar, "Dán", "Dán từ bộ nhớ tạm", PasteSelected);
            AddToolbarButton(toolbar, "Chổi", "Sao chép định dạng rồi nhấp ô đích", StartFormatPainter);
            toolbar.Items.Add(new ToolStripSeparator());
            toolbar.Items.Add(new ToolStripLabel("Cỡ chữ"));
            AddFontPicker();
            sizeCombo.AutoSize = false;
            sizeCombo.Width = 55;
            sizeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            sizeCombo.Items.AddRange(new object[] { "9", "10", "11", "12", "14", "16", "18", "24" });
            sizeCombo.SelectedIndexChanged += delegate
            {
                if (syncingToolbar)
                    return;
                float size;
                if (float.TryParse(sizeCombo.Text, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out size))
                    ApplyFontSize(size);
            };
            toolbar.Items.Add(sizeCombo);
            toolbar.Items.Add(new ToolStripSeparator());
            boldButton.Font = new Font("Arial", 10F, FontStyle.Bold);
            boldButton.ToolTipText = "In đậm";
            boldButton.Click += delegate { ToggleFontStyle(FontStyle.Bold); };
            toolbar.Items.Add(boldButton);
            italicButton.Font = new Font("Arial", 10F, FontStyle.Italic);
            italicButton.ToolTipText = "In nghiêng";
            italicButton.Click += delegate { ToggleFontStyle(FontStyle.Italic); };
            toolbar.Items.Add(italicButton);
            underlineButton.Font = new Font("Arial", 10F, FontStyle.Underline);
            underlineButton.ToolTipText = "Gạch chân";
            underlineButton.Click += delegate { ToggleFontStyle(FontStyle.Underline); };
            toolbar.Items.Add(underlineButton);
            toolbar.Items.Add(new ToolStripSeparator());
            AddToolbarButton(toolbar, "Màu chữ", "Chọn màu chữ", delegate { ChooseColor(false); });
            AddToolbarButton(toolbar, "Màu nền", "Chọn màu nền ô", delegate { ChooseColor(true); });
            toolbar.Items.Add(new ToolStripSeparator());
            AddToolbarButton(toolbar, "Trái", "Căn trái", delegate
            {
                ApplyAlignment(DataGridViewContentAlignment.MiddleLeft);
            });
            AddToolbarButton(toolbar, "Giữa", "Căn giữa", delegate
            {
                ApplyAlignment(DataGridViewContentAlignment.MiddleCenter);
            });
            AddToolbarButton(toolbar, "Phải", "Căn phải", delegate
            {
                ApplyAlignment(DataGridViewContentAlignment.MiddleRight);
            });
            toolbar.Items.Add(new ToolStripSeparator());
            themeSwatch.ForeColor = theme.Accent;
            themeSwatch.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
            toolbar.Items.Add(themeSwatch);
            AddToolbarButton(toolbar, "Giao diện", "Chọn màu toàn bộ giao diện", ChooseTheme);
            ConfigureCompactCommands();
            layout.Controls.Add(toolbar, 0, 2);

            TableLayoutPanel formula = formulaPanel;
            formula.Dock = DockStyle.Fill;
            formula.Margin = Padding.Empty;
            formula.Padding = new Padding(10, 4, 10, 4);
            formula.ColumnCount = 3;
            formula.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            formula.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            formula.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            formula.BackColor = theme.Sheet;
            addressBox.ReadOnly = false;
            addressBox.TabStop = true;
            addressBox.BorderStyle = BorderStyle.FixedSingle;
            addressBox.BackColor = theme.Header;
            addressBox.ForeColor = theme.Text;
            addressBox.TextAlign = HorizontalAlignment.Center;
            addressBox.Dock = DockStyle.Fill;
            addressBox.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            var formulaTips = new ToolTip();
            formulaTips.SetToolTip(addressBox, "Nhập A1, A1:C5 hoặc tên vùng rồi nhấn Enter");
            addressBox.KeyDown += AddressBoxKeyDown;
            addressBox.Enter += delegate { addressBox.SelectAll(); };
            formula.Controls.Add(addressBox, 0, 0);
            fx.Text = "fx";
            fx.ForeColor = theme.Muted;
            fx.Font = new Font("Georgia", 12F, FontStyle.Italic);
            fx.TextAlign = ContentAlignment.MiddleCenter;
            fx.Dock = DockStyle.Fill;
            formula.Controls.Add(fx, 1, 0);
            contentBox.BorderStyle = BorderStyle.FixedSingle;
            contentBox.Font = new Font("Consolas", 10F);
            contentBox.ForeColor = theme.Text;
            contentBox.BackColor = theme.Sheet;
            contentBox.Dock = DockStyle.Fill;
            formulaTips.SetToolTip(contentBox, "Sửa nội dung hoặc công thức; Enter để lưu, Esc để hủy");
            contentBox.TextChanged += delegate { if (!syncingContent) formulaBarChanged = true; UpdateFormulaPreview(); };
            contentBox.KeyDown += FormulaBarKeyDown;
            contentBox.Leave += delegate { CommitFormulaBar(); UpdateFormulaPreview(); };
            contentBox.Enter += delegate { UpdateFormulaPreview(); };
            BuildFormulaInput(formula);
            formula.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (var pen = new Pen(theme.Border))
                    e.Graphics.DrawLine(pen, 0, formula.Height - 1, formula.Width, formula.Height - 1);
            };
            layout.Controls.Add(formula, 0, 3);

            grid.Dock = DockStyle.Fill;
            grid.Margin = Padding.Empty;
            grid.BorderStyle = BorderStyle.None;
            grid.BackgroundColor = theme.Sheet;
            grid.GridColor = theme.GridLine;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToOrderColumns = false;
            grid.AllowUserToResizeRows = true;
            grid.MultiSelect = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.RowHeadersWidth = 50;
            grid.ColumnHeadersHeight = 28;
            grid.RowTemplate.Height = 27;
            grid.EnableHeadersVisualStyles = false;
            grid.Font = new Font("Segoe UI", 9F);
            grid.DefaultCellStyle.ForeColor = theme.Text;
            grid.DefaultCellStyle.BackColor = theme.Sheet;
            grid.DefaultCellStyle.SelectionForeColor = theme.Text;
            grid.DefaultCellStyle.SelectionBackColor = theme.Selection;
            grid.ColumnHeadersDefaultCellStyle.BackColor = theme.Header;
            grid.RowHeadersDefaultCellStyle.BackColor = theme.Header;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            grid.ShortcutHandler = HandleEditingShortcut;
            grid.CellPainting += GridCellPainting;
            grid.Paint += PaintMergedCells;
            grid.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                foreach (Rectangle merge in merges)
                    if (merge.Contains(e.ColumnIndex, e.RowIndex) &&
                        (e.ColumnIndex != merge.X || e.RowIndex != merge.Y))
                    { grid.CurrentCell = grid[merge.X, merge.Y]; break; }
            };
            grid.CellFormatting += GridCellFormatting;
            grid.CellPainting += PaintCellExtras;
            grid.EditingControlShowing += GridEditingControlShowing;
            grid.CellBeginEdit += delegate(object sender, DataGridViewCellCancelEventArgs e)
            {
                if (ValidationFor(e.RowIndex, e.ColumnIndex) == null) return;
                e.Cancel = true;
                int row = e.RowIndex, column = e.ColumnIndex;
                BeginInvoke((Action)delegate { ShowValidationDropdown(row, column); });
            };
            grid.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (painterSource != null && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    ApplyFormatPainter();
                    return;
                }
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                    ValidationFor(e.RowIndex, e.ColumnIndex) != null)
                    ShowValidationDropdown(e.RowIndex, e.ColumnIndex);
            };
            grid.SelectionChanged += delegate
            {
                if (selectingHeader)
                    return;
                if (!selectingByKeyboard && grid.CurrentCell != null)
                {
                    selectionAnchorRow = grid.CurrentCell.RowIndex;
                    selectionAnchorColumn = grid.CurrentCell.ColumnIndex;
                }
                UpdateSelection();
                ScheduleSelectionSummary();
            };
            grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (!loading)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                        !ValidateCellChange(e.RowIndex, e.ColumnIndex)) return;
                    RecordChange();
                    Recalculate(e.RowIndex, e.ColumnIndex);
                    MarkDirty();
                    UpdateSelection();
                }
            };
            grid.RowHeightChanged += delegate
            {
                if (!loading && lastState != null)
                {
                    RecordChange();
                    MarkDirty();
                }
            };
            grid.ColumnWidthChanged += delegate
            {
                if (!loading && lastState != null)
                {
                    RecordChange();
                    MarkDirty();
                }
            };
            grid.KeyDown += GridKeyDown;
            grid.CellEndEdit += delegate { if (pendingEditMove) FinishEditMove(); };
            grid.CellMouseDown += GridCellMouseDown;
            grid.RowHeaderMouseClick += GridCellMouseClick;
            grid.ColumnHeaderMouseClick += GridCellMouseClick;
            grid.CellMouseMove += GridCellMouseMove;
            grid.MouseMove += GridMouseMove;
            grid.MouseUp += GridMouseUp;
            rowContext.Items.Add("Chèn hàng phía trên", null,
                delegate { InsertRow(); });
            rowContext.Items.Add("Xóa hàng", null, delegate { DeleteRow(); });
            rowContext.Items.Add("Ẩn hàng", null, delegate { SetHidden(true, true); });
            rowContext.Items.Add("Hiện hàng đã ẩn", null, delegate { SetHidden(true, false); });
            columnContext.Items.Add("Chèn cột bên trái", null,
                delegate { InsertColumn(); });
            columnContext.Items.Add("Xóa cột", null, delegate { DeleteColumn(); });
            columnContext.Items.Add("Ẩn cột", null, delegate { SetHidden(false, true); });
            columnContext.Items.Add("Hiện cột đã ẩn", null, delegate { SetHidden(false, false); });
            var workspace = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            workspace.Controls.Add(grid);
            BuildInspector(workspace);
            BuildFloatingActions(workspace);
            BuildWelcome(workspace);
            layout.Controls.Add(workspace, 0, 4);

            Panel footer = footerPanel;
            footer.Dock = DockStyle.Fill;
            footer.BackColor = theme.Chrome;
            footer.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (var pen = new Pen(theme.Border))
                    e.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
                using (var pen = new Pen(theme.Accent, 3))
                    e.Graphics.DrawLine(pen, 52, footer.Height - 2, 172, footer.Height - 2);
            };
            Label tab = sheetTab;
            tab.Text = "▦  Trang tính 1";
            tab.ForeColor = theme.Accent;
            tab.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            tab.TextAlign = ContentAlignment.MiddleCenter;
            tab.Bounds = new Rectangle(50, 5, 124, 32);
            sheetTabs.Bounds = new Rectangle(8, 3, 680, 36);
            sheetTabs.AutoScroll = true;
            sheetTabs.WrapContents = false;
            sheetTabs.FlowDirection = FlowDirection.LeftToRight;
            footer.Controls.Add(sheetTabs);
            status.ForeColor = theme.Muted;
            status.TextAlign = ContentAlignment.MiddleRight;
            status.Bounds = new Rectangle(850, 3, 135, 28);
            footer.Controls.Add(status);
            footer.Resize += delegate
            {
                status.Left = footer.ClientSize.Width - status.Width - 150;
            };
            layout.Controls.Add(footer, 0, 5);
            BuildVisualChrome(header, footer);
            RefreshSheetTabs();
        }

        private void BuildGrid()
        {
            for (int column = 0; column < ColumnCount; column++)
            {
                var item = new DataGridViewTextBoxColumn();
                item.Name = ((char)('A' + column)).ToString();
                item.HeaderText = item.Name;
                item.Width = 120;
                item.SortMode = DataGridViewColumnSortMode.NotSortable;
                grid.Columns.Add(item);
            }

            grid.Rows.Add(RowCount);
            for (int row = 0; row < RowCount; row++)
                grid.Rows[row].HeaderCell.Value = (row + 1).ToString();
            grid.CurrentCell = grid[0, 0];
        }

        private void Recalculate()
        {
            formulaEngine = null;
            Recalculate(-1, -1);
        }

        private void Recalculate(int changedRow, int changedColumn)
        {
            calculated.Clear();
            if (formulaEngine == null)
                formulaEngine = new FormulaEngine(delegate(int row, int column)
            {
                return Convert.ToString(grid[column, row].Value) ?? "";
            }, delegate(string name, int row, int column)
            {
                foreach (SheetState sheet in sheets)
                    if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        CellState cell;
                        return sheet.Cells.TryGetValue(row * ColumnCount + column, out cell) ?
                            Convert.ToString(cell.Value) ?? "" : "";
                    }
                return null;
            }, sheets.Count > activeSheetIndex ? sheets[activeSheetIndex].Name : "Sheet1",
                RowCount, ColumnCount, delegate(string name)
                {
                    NamedRange named = namedRanges.FirstOrDefault(n =>
                        string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
                    return named == null ? null : new FormulaNamedRange
                    {
                        Sheet = named.Sheet,
                        FirstRow = named.Range.Top,
                        FirstColumn = named.Range.Left,
                        LastRow = named.Range.Bottom - 1,
                        LastColumn = named.Range.Right - 1
                    };
                });
            else if (changedRow >= 0 && changedColumn >= 0)
                formulaEngine.Invalidate(sheets[activeSheetIndex].Name,
                    changedRow, changedColumn);
            for (int row = 0; row < RowCount; row++)
            {
                for (int column = 0; column < ColumnCount; column++)
                {
                    string raw = Convert.ToString(grid[column, row].Value) ?? "";
                    if (raw.StartsWith("=", StringComparison.Ordinal))
                        calculated[row * ColumnCount + column] =
                            formulaEngine.Display(row, column);
                }
            }
            grid.Invalidate();
        }

        private void GridCellFormatting(object sender,
            DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;
            string result;
            if (calculated.TryGetValue(e.RowIndex * ColumnCount + e.ColumnIndex,
                out result))
            {
                e.Value = result;
                e.FormattingApplied = true;
            }
            string raw = result ?? Convert.ToString(grid[e.ColumnIndex, e.RowIndex].Value) ?? "";
            string format = grid[e.ColumnIndex, e.RowIndex].Style.Format;
            double numeric;
            if (!string.IsNullOrEmpty(format) &&
                double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out numeric))
            {
                try { e.Value = FormatNumeric(numeric, format); e.FormattingApplied = true; }
                catch (FormatException) { }
            }
            foreach (ConditionalRule rule in conditionalRules)
                if (rule.Range.Contains(e.ColumnIndex, e.RowIndex) &&
                    double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out numeric) &&
                    numeric > rule.Threshold)
                    e.CellStyle.BackColor = rule.Color;
        }

        private void GridEditingControlShowing(object sender,
            DataGridViewEditingControlShowingEventArgs e)
        {
            if (grid.CurrentCell == null)
                return;
            var editor = e.Control as TextBox;
            string raw = Convert.ToString(grid.CurrentCell.Value) ?? "";
            if (editor != null && raw.StartsWith("=", StringComparison.Ordinal))
                editor.Text = raw;
        }

        private static string AppearanceSettingsPath()
        {
            return Path.Combine(
                Path.GetDirectoryName(typeof(SpreadsheetForm).Assembly.Location),
                "appearance.xml");
        }

        private static ThemePalette LoadThemePreference()
        {
            try
            {
                string path = AppearanceSettingsPath();
                if (File.Exists(path))
                {
                    var document = XDocument.Load(path);
                    XAttribute id = document.Root == null
                        ? null : document.Root.Attribute("theme");
                    ThemePalette selected = id == null
                        ? null : ThemePalette.Find(id.Value);
                    if (selected != null)
                        return selected;
                }
            }
            catch
            {
                // A missing or damaged preference uses the default theme.
            }
            return ThemePalette.All[0];
        }

        private void SaveThemePreference()
        {
            try
            {
                string path = AppearanceSettingsPath();
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                new XDocument(new XElement("appearance",
                    new XAttribute("theme", theme.Id))).Save(path);
            }
            catch
            {
                status.Text = "Đã đổi giao diện; chưa lưu được lựa chọn cho lần sau";
            }
        }

        private void ApplyTheme(ThemePalette palette, bool markDirty, bool persist)
        {
            theme = palette;
            BackColor = theme.Surface;
            headerPanel.BackColor = theme.Chrome;
            logo.BackColor = theme.Logo;
            documentTitle.ForeColor = theme.Text;
            subtitle.ForeColor = theme.Muted;
            saveIndicator.ForeColor = theme.Muted;
            menu.BackColor = theme.Chrome;
            menu.ForeColor = theme.Text;
            foreach (ToolStripItem item in menu.Items)
            {
                item.ForeColor = theme.Text;
                var parent = item as ToolStripMenuItem;
                if (parent == null)
                    continue;
                parent.DropDown.BackColor = theme.Chrome;
                foreach (ToolStripItem child in parent.DropDownItems)
                {
                    child.BackColor = theme.Chrome;
                    child.ForeColor = theme.Text;
                }
            }
            foreach (ContextMenuStrip context in new[] { rowContext, columnContext })
            {
                context.BackColor = theme.Chrome;
                context.ForeColor = theme.Text;
                foreach (ToolStripItem item in context.Items)
                {
                    item.BackColor = theme.Chrome;
                    item.ForeColor = theme.Text;
                }
            }
            toolbar.BackColor = theme.Chrome;
            toolbar.ForeColor = theme.Text;
            foreach (ToolStripItem item in toolbar.Items)
                item.ForeColor = theme.Text;
            themeSwatch.ForeColor = theme.Accent;
            sizeCombo.BackColor = theme.Sheet;
            sizeCombo.ForeColor = theme.Text;
            fx.ForeColor = theme.Muted;
            footerPanel.BackColor = theme.Chrome;
            sheetTab.ForeColor = theme.Accent;
            status.ForeColor = theme.Muted;
            ApplySheetBackground(theme.Sheet, false);
            ApplyVisualTheme();
            headerPanel.Invalidate();
            formulaPanel.Invalidate();
            footerPanel.Invalidate();
            if (persist)
                SaveThemePreference();
            if (markDirty)
            {
                RecordChange();
                MarkDirty();
            }
        }

        private void ApplySheetBackground(Color color, bool markDirty)
        {
            sheetBackground = color;
            bool dark = (color.R * 299 + color.G * 587 + color.B * 114) / 1000 < 150;
            Color foreground = color.ToArgb() == theme.Sheet.ToArgb()
                ? theme.Text : dark ? Color.FromArgb(235, 242, 249)
                    : Color.FromArgb(32, 41, 51);
            grid.BackgroundColor = color;
            grid.DefaultCellStyle.BackColor = color;
            grid.DefaultCellStyle.ForeColor = foreground;
            grid.DefaultCellStyle.SelectionForeColor = foreground;
            grid.DefaultCellStyle.SelectionBackColor =
                color.ToArgb() == theme.Sheet.ToArgb() ? theme.Selection
                : dark ? Color.FromArgb(60, 80, 105)
                    : Color.FromArgb(211, 227, 244);
            grid.GridColor = color.ToArgb() == theme.Sheet.ToArgb()
                ? theme.GridLine : dark ? Color.FromArgb(78, 92, 108)
                    : Color.FromArgb(198, 209, 221);
            grid.ColumnHeadersDefaultCellStyle.BackColor = theme.Header;
            grid.RowHeadersDefaultCellStyle.BackColor = theme.Header;
            formulaPanel.BackColor = color;
            addressBox.BackColor = theme.Header;
            addressBox.ForeColor = foreground;
            contentBox.BackColor = color;
            contentBox.ForeColor = foreground;
            grid.Invalidate();
            if (markDirty)
            {
                RecordChange();
                MarkDirty();
            }
        }

        private void ChooseTheme()
        {
            using (var picker = new ThemePickerForm(theme))
            {
                if (picker.ShowDialog(this) == DialogResult.OK &&
                    picker.SelectedTheme != null && picker.SelectedTheme != theme)
                    ApplyTheme(picker.SelectedTheme, true, true);
            }
        }

        private void GridCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                bool active = grid.CurrentCell != null &&
                    ((e.RowIndex == -1 && e.ColumnIndex == grid.CurrentCell.ColumnIndex) ||
                     (e.ColumnIndex == -1 && e.RowIndex == grid.CurrentCell.RowIndex));
                using (var brush = new SolidBrush(
                    active ? theme.HeaderActive : theme.Header))
                    e.Graphics.FillRectangle(brush, e.CellBounds);
                using (var pen = new Pen(theme.Border))
                {
                    e.Graphics.DrawLine(pen, e.CellBounds.Right - 1, e.CellBounds.Top,
                        e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    e.Graphics.DrawLine(pen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
                        e.CellBounds.Right, e.CellBounds.Bottom - 1);
                }
                string text = "";
                if (e.RowIndex >= 0)
                    text = (e.RowIndex + 1).ToString();
                else if (e.ColumnIndex >= 0)
                    text = grid.Columns[e.ColumnIndex].HeaderText;
                TextRenderer.DrawText(e.Graphics, text, HeaderFont, e.CellBounds,
                    active ? theme.Accent : theme.Muted,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                if (headerDragging &&
                    ((headerDragRow && e.ColumnIndex == -1 &&
                      e.RowIndex == headerTarget) ||
                     (!headerDragRow && e.RowIndex == -1 &&
                      e.ColumnIndex == headerTarget)))
                {
                    using (var marker = new Pen(theme.Accent, 3F))
                    {
                        if (headerDragRow)
                            e.Graphics.DrawLine(marker, e.CellBounds.Left,
                                headerSource < headerTarget ?
                                    e.CellBounds.Bottom - 2 : e.CellBounds.Top + 1,
                                e.CellBounds.Right,
                                headerSource < headerTarget ?
                                    e.CellBounds.Bottom - 2 : e.CellBounds.Top + 1);
                        else
                            e.Graphics.DrawLine(marker,
                                headerSource < headerTarget ?
                                    e.CellBounds.Right - 2 : e.CellBounds.Left + 1,
                                e.CellBounds.Top,
                                headerSource < headerTarget ?
                                    e.CellBounds.Right - 2 : e.CellBounds.Left + 1,
                                e.CellBounds.Bottom);
                    }
                }
                e.Handled = true;
                return;
            }

            foreach (Rectangle merge in merges)
                if (merge.Contains(e.ColumnIndex, e.RowIndex))
                { e.Handled = true; return; }

            if (grid.CurrentCell != null &&
                e.RowIndex == grid.CurrentCell.RowIndex &&
                e.ColumnIndex == grid.CurrentCell.ColumnIndex)
            {
                e.Paint(e.CellBounds, e.PaintParts & ~DataGridViewPaintParts.Focus);
                using (var pen = new Pen(theme.Accent, 2F))
                {
                    var rectangle = e.CellBounds;
                    rectangle.Width -= 1;
                    rectangle.Height -= 1;
                    e.Graphics.DrawRectangle(pen, rectangle);
                }
                if (!fillSelection.IsEmpty && e.RowIndex == fillSelection.Bottom - 1 &&
                    e.ColumnIndex == fillSelection.Right - 1)
                    using (var brush = new SolidBrush(theme.Accent))
                        e.Graphics.FillRectangle(brush, e.CellBounds.Right - 7,
                            e.CellBounds.Bottom - 7, 7, 7);
                e.Handled = true;
                return;
            }

            if (!fillSelection.IsEmpty && e.RowIndex == fillSelection.Bottom - 1 &&
                e.ColumnIndex == fillSelection.Right - 1)
            {
                e.Paint(e.CellBounds, e.PaintParts & ~DataGridViewPaintParts.Focus);
                using (var brush = new SolidBrush(theme.Accent))
                    e.Graphics.FillRectangle(brush, e.CellBounds.Right - 7,
                        e.CellBounds.Bottom - 7, 7, 7);
                e.Handled = true;
                return;
            }

            if (fillDragging &&
                e.RowIndex >= Math.Min(fillSourceRow, fillTargetRow) &&
                e.RowIndex <= Math.Max(fillSourceRow, fillTargetRow) &&
                e.ColumnIndex >= Math.Min(fillSourceColumn, fillTargetColumn) &&
                e.ColumnIndex <= Math.Max(fillSourceColumn, fillTargetColumn))
            {
                e.Paint(e.CellBounds, e.PaintParts & ~DataGridViewPaintParts.Focus);
                using (var brush = new SolidBrush(Color.FromArgb(55, theme.Accent)))
                    e.Graphics.FillRectangle(brush, e.CellBounds);
                e.Handled = true;
            }
        }

        private static void AddMenuItem(ToolStripMenuItem parent, string text,
            Keys shortcut, Action action)
        {
            var item = new ToolStripMenuItem(text);
            item.ShortcutKeys = shortcut;
            item.Click += delegate { action(); };
            parent.DropDownItems.Add(item);
        }

        private static void AddToolbarButton(ToolStrip toolbar, string text,
            string tooltip, Action action)
        {
            var button = new ToolStripButton(text);
            button.ToolTipText = tooltip;
            button.DisplayStyle = ToolStripItemDisplayStyle.Text;
            button.Padding = new Padding(4, 2, 4, 2);
            button.Click += delegate { action(); };
            toolbar.Items.Add(button);
        }

        private void UpdateSelection()
        {
            if (grid.CurrentCell == null)
                return;
            fillSelection = SelectedRectangle(grid, true);
            if (!addressBox.Focused)
                addressBox.Text = CellAddress(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex);
            if (!contentBox.Focused || !formulaBarChanged)
            {
                syncingContent = true;
                contentBox.Text = Convert.ToString(grid.CurrentCell.Value) ?? "";
                syncingContent = false;
            }

            Font font = grid.CurrentCell.InheritedStyle.Font ?? grid.Font;
            syncingToolbar = true;
            sizeCombo.Text = font.Size.ToString("0", CultureInfo.InvariantCulture);
            fontCombo.Text = font.Name;
            boldButton.Checked = (font.Style & FontStyle.Bold) != 0;
            italicButton.Checked = (font.Style & FontStyle.Italic) != 0;
            underlineButton.Checked = (font.Style & FontStyle.Underline) != 0;
            syncingToolbar = false;
            UpdateContextCommands();
            UpdateInspectorProperties();
        }

        private void MarkDirty()
        {
            dirty = true;
            status.Text = "Có thay đổi chưa lưu";
            saveIndicator.Text = "Chưa lưu";
            UpdateTitle();
        }

        private SheetState CaptureSheet()
        {
            var state = new SheetState();
            state.Name = sheets.Count > activeSheetIndex ? sheets[activeSheetIndex].Name : "Sheet1";
            if (sheets.Count > activeSheetIndex)
            { state.Hidden = sheets[activeSheetIndex].Hidden; state.TabColor = sheets[activeSheetIndex].TabColor; }
            state.Merges.AddRange(merges);
            state.Rules.AddRange(conditionalRules);
            state.Tables.AddRange(tables);
            state.Charts.AddRange(charts);
            state.Validations.AddRange(validations);
            state.FreezeRow = freezeRow;
            state.FreezeColumn = freezeColumn;
            state.FilterColumn = filterColumn;
            state.FilterValue = filterValue;
            state.Background = sheetBackground;
            state.ThemeId = theme.Id;
            state.CsvRows = csvDocument == null ? 0 : csvDocument.DataRows;
            state.CsvColumns = csvDocument == null ? 0 : csvDocument.DataColumns;
            for (int row = 0; row < RowCount; row++)
            {
                state.RowHeights[row] = Math.Max(1, (int)Math.Round(grid.Rows[row].Height * 100.0 / zoomPercent));
                for (int column = 0; column < ColumnCount; column++)
                {
                    DataGridViewCell cell = grid[column, row];
                    DataGridViewCellStyle style = cell.HasStyle ? cell.Style : null;
                    if (cell.Value == null && !HasMeaningfulStyle(style) && cell.Tag == null)
                        continue;
                    state.Cells[row * ColumnCount + column] = new CellState
                    {
                        Value = cell.Value,
                        Style = HasMeaningfulStyle(style) ?
                            new DataGridViewCellStyle(style) : null,
                        Extras = CellExtras.Copy(cell.Tag as CellExtras)
                    };
                }
            }
            for (int column = 0; column < ColumnCount; column++)
                state.ColumnWidths[column] = Math.Max(1, (int)Math.Round(grid.Columns[column].Width * 100.0 / zoomPercent));
            Array.Copy(manualHiddenRows, state.HiddenRows, RowCount);
            Array.Copy(manualHiddenColumns, state.HiddenColumns, ColumnCount);
            return state;
        }

        private static bool HasMeaningfulStyle(DataGridViewCellStyle style)
        {
            return style != null && (style.Font != null ||
                !style.ForeColor.IsEmpty || !style.BackColor.IsEmpty ||
                style.Alignment != DataGridViewContentAlignment.NotSet ||
                !string.IsNullOrEmpty(style.Format));
        }

        private void ResetHistory()
        {
            crossSheetMoves.Clear();
            undoHistory.Clear();
            redoHistory.Clear();
            nextRevision = 0;
            savedRevision = 0;
            lastState = CaptureSheet();
            lastState.RevisionId = 0;
            sheetHistories.Clear();
            for (int i = 0; i < sheets.Count; i++) sheetHistories.Add(new SheetHistory());
            if (sheetHistories.Count > activeSheetIndex)
                sheetHistories[activeSheetIndex].Last = lastState;
        }

        private void RecordChange()
        {
            if (lastState == null)
            {
                lastState = CaptureSheet();
                return;
            }
            undoHistory.Add(lastState);
            if (undoHistory.Count > 50)
                undoHistory.RemoveAt(0);
            redoHistory.Clear();
            lastState = CaptureSheet();
            lastState.RevisionId = ++nextRevision;
        }

        private void RestoreSheet(SheetState state)
        {
            if (validationEditor != null)
            { grid.Controls.Remove(validationEditor); validationEditor.Dispose(); validationEditor = null; }
            loading = true;
            grid.SuspendLayout();
            try
            {
                foreach (DataGridViewRow row in grid.Rows)
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null || cell.HasStyle || cell.Tag != null)
                        {
                            cell.Value = null;
                            cell.Style = new DataGridViewCellStyle();
                            cell.Tag = null;
                        }
                    }
                foreach (KeyValuePair<int, CellState> item in state.Cells)
                {
                    DataGridViewCell cell = grid[item.Key % ColumnCount,
                        item.Key / ColumnCount];
                    cell.Value = item.Value.Value;
                    if (item.Value.Style != null)
                        cell.Style = new DataGridViewCellStyle(item.Value.Style);
                    cell.Tag = CellExtras.Copy(item.Value.Extras);
                }
                for (int row = 0; row < RowCount; row++)
                    grid.Rows[row].Height = Math.Max(5, (int)Math.Round(state.RowHeights[row] * zoomPercent / 100.0));
                for (int column = 0; column < ColumnCount; column++)
                    grid.Columns[column].Width = Math.Max(10, (int)Math.Round(state.ColumnWidths[column] * zoomPercent / 100.0));
                RestoreHidden(state);
                merges.Clear();
                merges.AddRange(state.Merges);
                UpdateMergedReadOnly();
                conditionalRules.Clear();
                conditionalRules.AddRange(state.Rules);
                tables.Clear(); tables.AddRange(state.Tables);
                charts.Clear(); charts.AddRange(state.Charts);
                validations.Clear(); validations.AddRange(state.Validations);
                freezeRow = state.FreezeRow;
                freezeColumn = state.FreezeColumn;
                filterColumn = state.FilterColumn;
                filterValue = state.FilterValue;
                ApplyFreezeAndFilter();
                if (csvDocument != null)
                {
                    csvDocument.DataRows = state.CsvRows;
                    csvDocument.DataColumns = state.CsvColumns;
                }
                ThemePalette selected = ThemePalette.Find(state.ThemeId);
                if (selected != null && selected != theme)
                    ApplyTheme(selected, false, true);
                ApplySheetBackground(state.Background, false);
            }
            finally
            {
                grid.ResumeLayout();
                loading = false;
            }
            Recalculate();
            dirty = state.RevisionId != savedRevision || otherSheetsDirty;
            saveIndicator.Text = dirty ? "Chưa lưu" :
                currentPath == null ? "" : "Đã lưu trên máy";
            UpdateTitle();
            UpdateSelection();
            grid.Invalidate();
        }

        private void Undo()
        {
            grid.EndEdit();
            if (undoHistory.Count == 0)
                return;
            BeforeUndoCrossSheetMove();
            int last = undoHistory.Count - 1;
            SheetState previous = undoHistory[last];
            undoHistory.RemoveAt(last);
            redoHistory.Add(lastState);
            RestoreSheet(previous);
            lastState = CaptureSheet();
            lastState.RevisionId = previous.RevisionId;
            status.Text = "Đã hoàn tác";
        }

        private void Redo()
        {
            grid.EndEdit();
            if (redoHistory.Count == 0)
                return;
            int last = redoHistory.Count - 1;
            SheetState next = redoHistory[last];
            redoHistory.RemoveAt(last);
            undoHistory.Add(lastState);
            RestoreSheet(next);
            lastState = CaptureSheet();
            AfterRedoCrossSheetMove(next.RevisionId);
            lastState.RevisionId = next.RevisionId;
            status.Text = "Đã làm lại";
        }

        private void UpdateTitle()
        {
            string name = currentPath == null ? "Bảng tính chưa đặt tên" :
                Path.GetFileNameWithoutExtension(currentPath);
            Text = name + (dirty ? " *" : "") + " - DinkCel";
            documentTitle.Text = name;
            if (!dirty)
                saveIndicator.Text = currentPath == null ? "" : "Đã lưu trên máy";
        }

        private bool ConfirmDiscardChanges()
        {
            if (!CommitFormulaBar()) return false;
            if (!dirty)
                return true;
            DialogResult answer = MessageBox.Show(this,
                "Bảng tính có thay đổi chưa lưu. Bạn muốn lưu không?",
                "DinkCel", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (answer == DialogResult.Cancel)
                return false;
            if (answer == DialogResult.Yes)
                return SaveDocument();
            return true;
        }

        private void ClearGrid()
        {
            loading = true;
            grid.SuspendLayout();
            foreach (DataGridViewRow row in grid.Rows)
            {
                row.Visible = true;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null || cell.HasStyle || cell.Tag != null)
                    {
                        cell.Value = null;
                        cell.Style = new DataGridViewCellStyle();
                        cell.Tag = null;
                    }
                }
            }
            foreach (DataGridViewColumn column in grid.Columns) column.Visible = true;
            Array.Clear(manualHiddenRows, 0, RowCount);
            Array.Clear(manualHiddenColumns, 0, ColumnCount);
            grid.ResumeLayout();
            loading = false;
            grid.ClearSelection();
            grid.CurrentCell = grid[0, 0];
            grid[0, 0].Selected = true;
            calculated.Clear();
            UpdateSelection();
        }

        private void NewDocument()
        {
            if (!ConfirmDiscardChanges())
                return;
            ClearGrid();
            loading = true;
            for (int row = 0; row < RowCount; row++)
                grid.Rows[row].Height = Math.Max(5, (int)Math.Round(27 * zoomPercent / 100.0));
            for (int column = 0; column < ColumnCount; column++)
                grid.Columns[column].Width = Math.Max(10, (int)Math.Round(120 * zoomPercent / 100.0));
            loading = false;
            sheets.Clear();
            sheets.Add(new SheetState { Name = "Sheet1" });
            activeSheetIndex = 0;
            merges.Clear();
            UpdateMergedReadOnly();
            conditionalRules.Clear();
            tables.Clear(); charts.Clear(); validations.Clear();
            namedRanges.Clear(); pivots.Clear();
            freezeRow = freezeColumn = 0;
            filterColumn = -1;
            filterValue = "";
            RefreshSheetTabs();
            ApplySheetBackground(theme.Sheet, false);
            currentPath = null;
            csvDocument = null;
            dirty = false;
            otherSheetsDirty = false;
            status.Text = "Bảng tính mới";
            Recalculate();
            UpdateTitle();
            ResetHistory();
            HideWelcome();
        }

        private void OpenDocument()
        {
            if (!ConfirmDiscardChanges())
                return;
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Bảng tính (*.dinkcel;*.xlsx;*.xls;*.ods;*.csv)|*.dinkcel;*.xlsx;*.xls;*.ods;*.csv|DinkCel (*.dinkcel)|*.dinkcel|Excel (*.xlsx;*.xls)|*.xlsx;*.xls|OpenDocument (*.ods)|*.ods|CSV (*.csv)|*.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                OpenPath(dialog.FileName);
            }
        }

        private void OpenPath(string path)
        {
                try
                {
                    bool isCsv = string.Equals(Path.GetExtension(path), ".csv",
                        StringComparison.OrdinalIgnoreCase);
                    string extension = Path.GetExtension(path).ToLowerInvariant();
                    CsvDocument csv = null;
                    WorkbookSnapshot workbook = isCsv ? ReadCsvWorkbook(path, out csv) :
                        extension == ".xlsx" ? XlsxFile.Read(path, RowCount, ColumnCount) :
                        extension == ".xls" ? XlsFile.Read(path, RowCount, ColumnCount) :
                        extension == ".ods" ? OdsFile.Read(path, RowCount, ColumnCount) : ReadWorkbook(path);
                    ClearGrid();
                    ThemePalette savedTheme = ThemePalette.Find(workbook.ThemeId);
                    if (savedTheme != null)
                        ApplyTheme(savedTheme, false, true);
                    if (workbook.HasBackground)
                        ApplySheetBackground(workbook.Background, false);
                    sheets.Clear();
                    namedRanges.Clear(); namedRanges.AddRange(workbook.NamedRanges);
                    pivots.Clear(); pivots.AddRange(workbook.Pivots);
                    foreach (SheetSnapshot snapshot in workbook.Sheets)
                    {
                        SheetState state = StateFromSnapshot(snapshot);
                        state.Background = snapshot.Background.IsEmpty ?
                            workbook.HasBackground ? workbook.Background : theme.Sheet : snapshot.Background;
                        state.ThemeId = snapshot.ThemeId ?? workbook.ThemeId ?? theme.Id;
                        if (csv != null) { state.CsvRows = csv.DataRows; state.CsvColumns = csv.DataColumns; }
                        sheets.Add(state);
                    }
                    if (sheets.Count == 0)
                        sheets.Add(new SheetState { Name = "Sheet1" });
                    activeSheetIndex = sheets.FindIndex(s => !s.Hidden);
                    if (activeSheetIndex < 0) { sheets[0].Hidden = false; activeSheetIndex = 0; }
                    csvDocument = csv;
                    RestoreSheet(sheets[activeSheetIndex]);
                    RefreshSheetTabs();
                    currentPath = path;
                    dirty = false;
                    otherSheetsDirty = false;
                    Recalculate();
                    UpdateSelection();
                    UpdateTitle();
                    status.Text = "Đã mở " + Path.GetFileName(path);
                    ResetHistory();
                    RecordRecentFile(path);
                    HideWelcome();
                }
                catch (Exception error)
                {
                    loading = false;
                    MessageBox.Show(this, "Không mở được tệp: " + error.Message,
                        "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }

        private static WorkbookSnapshot ReadCsvWorkbook(string path, out CsvDocument csv)
        {
            var workbook = new WorkbookSnapshot();
            workbook.Sheets[0].Name = Path.GetFileNameWithoutExtension(path);
            csv = CsvFile.ReadDocument(path, RowCount, ColumnCount);
            IList<string[]> rows = csv.Rows;
            for (int row = 0; row < rows.Count; row++)
                for (int column = 0; column < rows[row].Length; column++)
                    if (rows[row][column].Length > 0)
                        workbook.Cells[row * ColumnCount + column] =
                            new CellSnapshot { Text = rows[row][column] };
            return workbook;
        }

        private static WorkbookSnapshot ReadWorkbook(string path)
        {
            var document = XDocument.Load(path);
            if (document.Root == null || document.Root.Name != "workbook")
                throw new InvalidDataException("Định dạng bảng tính không hợp lệ.");
            var workbook = new WorkbookSnapshot();
            XAttribute themeAttribute = document.Root.Attribute("theme");
            if (themeAttribute != null)
                workbook.ThemeId = themeAttribute.Value;
            XAttribute background = document.Root.Attribute("background");
            if (background != null)
            {
                workbook.Background = ColorTranslator.FromHtml(background.Value);
                workbook.HasBackground = true;
            }
            ReadWorkbookMetadata(document.Root, workbook);
            List<XElement> sheetElements = new List<XElement>(document.Root.Elements("sheet"));
            if (sheetElements.Count > 0)
            {
                workbook.Sheets.Clear();
                foreach (XElement sheetElement in sheetElements)
                {
                    var sheet = new SheetSnapshot();
                    sheet.Name = (string)sheetElement.Attribute("name") ?? "Sheet" + (workbook.Sheets.Count + 1);
                    sheet.Hidden = (bool?)sheetElement.Attribute("hidden") ?? false;
                    string tabColor = (string)sheetElement.Attribute("tabColor");
                    if (!string.IsNullOrEmpty(tabColor)) sheet.TabColor = ColorTranslator.FromHtml(tabColor);
                    sheet.ThemeId = (string)sheetElement.Attribute("theme");
                    XAttribute sheetBackground = sheetElement.Attribute("background");
                    if (sheetBackground != null)
                        sheet.Background = ColorTranslator.FromHtml(sheetBackground.Value);
                    sheet.FreezeRow = (int?)sheetElement.Attribute("freezeRow") ?? 0;
                    sheet.FreezeColumn = (int?)sheetElement.Attribute("freezeColumn") ?? 0;
                    sheet.FilterColumn = (int?)sheetElement.Attribute("filterColumn") ?? -1;
                    sheet.FilterValue = (string)sheetElement.Attribute("filterValue") ?? "";
                    foreach (XElement merge in sheetElement.Elements("merge"))
                        sheet.Merges.Add(new Rectangle((int)merge.Attribute("column"),
                            (int)merge.Attribute("row"), (int)merge.Attribute("width"),
                            (int)merge.Attribute("height")));
                    foreach (XElement rule in sheetElement.Elements("conditional"))
                        sheet.Rules.Add(new ConditionalRule { Range = new Rectangle((int)rule.Attribute("column"),
                            (int)rule.Attribute("row"), (int)rule.Attribute("width"), (int)rule.Attribute("height")),
                            Threshold = (double)rule.Attribute("threshold"), Color = ColorTranslator.FromHtml((string)rule.Attribute("color")) });
                    ReadSheetMetadata(sheetElement, sheet);
                    foreach (XElement dimension in sheetElement.Elements("row"))
                        sheet.RowHeights[(int)dimension.Attribute("index")] = (int)dimension.Attribute("height");
                    foreach (XElement dimension in sheetElement.Elements("column"))
                        sheet.ColumnWidths[(int)dimension.Attribute("index")] = (int)dimension.Attribute("width");
                    foreach (XElement dimension in sheetElement.Elements("hiddenRow"))
                        sheet.HiddenRows.Add((int)dimension.Attribute("index"));
                    foreach (XElement dimension in sheetElement.Elements("hiddenColumn"))
                        sheet.HiddenColumns.Add((int)dimension.Attribute("index"));
                    ReadCells(sheetElement.Elements("cell"), sheet.Cells);
                    workbook.Sheets.Add(sheet);
                }
                return workbook;
            }
            ReadCells(document.Root.Elements("cell"), workbook.Cells);
            return workbook;
        }

        private static void ReadCells(IEnumerable<XElement> elements, Dictionary<int, CellSnapshot> cells)
        {
            foreach (XElement element in elements)
            {
                int row = int.Parse(element.Attribute("row").Value) - 1;
                int column = int.Parse(element.Attribute("column").Value) - 1;
                if (row < 0 || row >= RowCount || column < 0 || column >= ColumnCount)
                    throw new InvalidDataException("Tệp chứa ô nằm ngoài bảng.");
                var snapshot = new CellSnapshot();
                snapshot.Text = element.Value;
                XAttribute style = element.Attribute("fontStyle");
                XAttribute size = element.Attribute("fontSize");
                if (style != null || size != null)
                {
                    snapshot.HasFont = true;
                    if (style != null)
                    {
                        int flags = int.Parse(style.Value);
                        if (flags < 0 || flags > 15)
                            throw new InvalidDataException("Định dạng chữ không hợp lệ.");
                        snapshot.FontStyle = (FontStyle)flags;
                    }
                    if (size != null)
                    {
                        snapshot.FontSize = float.Parse(size.Value,
                            CultureInfo.InvariantCulture);
                        if (float.IsNaN(snapshot.FontSize) ||
                            float.IsInfinity(snapshot.FontSize) ||
                            snapshot.FontSize < 6F || snapshot.FontSize > 72F)
                            throw new InvalidDataException("Cỡ chữ không hợp lệ.");
                    }
                }
                XAttribute fore = element.Attribute("fore");
                XAttribute back = element.Attribute("back");
                XAttribute align = element.Attribute("align");
                if (fore != null)
                    snapshot.ForeColor = ColorTranslator.FromHtml(fore.Value);
                if (back != null)
                    snapshot.BackColor = ColorTranslator.FromHtml(back.Value);
                if (align != null)
                    snapshot.Alignment = (DataGridViewContentAlignment)Enum.Parse(
                        typeof(DataGridViewContentAlignment), align.Value);
                snapshot.NumberFormat = (string)element.Attribute("numberFormat") ?? "";
                snapshot.FontName = (string)element.Attribute("fontName") ?? "Arial";
                snapshot.Extras = CellExtras.ReadXml(element);
                cells[row * ColumnCount + column] = snapshot;
            }
        }

        private bool SaveDocument()
        {
            if (!CommitFormulaBar()) return false;
            grid.EndEdit();
            if (currentPath == null)
                return SaveDocumentAs();
            if (!dirty)
                return true;
            return WriteDocument(currentPath);
        }

        private bool SaveDocumentAs()
        {
            if (!CommitFormulaBar()) return false;
            using (var dialog = new SaveFileDialog())
            {
                bool csv = IsCsvPath(currentPath);
                dialog.Filter = "DinkCel (*.dinkcel)|*.dinkcel|Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|OpenDocument (*.ods)|*.ods|CSV (*.csv)|*.csv";
                dialog.DefaultExt = csv ? "csv" : "dinkcel";
                dialog.FilterIndex = csv ? 5 : 1;
                dialog.AddExtension = false;
                dialog.OverwritePrompt = false;
                dialog.FileName = currentPath == null ? "BangTinh" :
                    Path.GetFileNameWithoutExtension(currentPath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return false;
                string path = Path.ChangeExtension(dialog.FileName,
                    dialog.FilterIndex == 5 ? ".csv" : dialog.FilterIndex == 4 ? ".ods" :
                    dialog.FilterIndex == 3 ? ".xls" : dialog.FilterIndex == 2 ? ".xlsx" : ".dinkcel");
                if (File.Exists(path) && MessageBox.Show(this,
                    "Tệp đã tồn tại. Bạn muốn ghi đè không?", "DinkCel",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return false;
                return WriteDocument(path);
            }
        }

        private static bool IsCsvPath(string path)
        {
            return path != null && string.Equals(Path.GetExtension(path), ".csv",
                StringComparison.OrdinalIgnoreCase);
        }

        private bool WriteDocument(string path)
        {
            bool saved = IsCsvPath(path) ? WriteCsv(path) :
                string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase) ?
                WriteXlsx(path) :
                string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase) ?
                WriteXls(path) :
                string.Equals(Path.GetExtension(path), ".ods", StringComparison.OrdinalIgnoreCase) ?
                WriteOds(path) : WriteWorkbook(path);
            if (saved) RecordRecentFile(path);
            return saved;
        }

        private bool WriteCsv(string path)
        {
            try
            {
                if (sheets.Count > 1)
                {
                    MessageBox.Show(this, "CSV chỉ lưu được một sheet. Hãy dùng Lưu thành và chọn .xlsx hoặc .dinkcel.",
                        "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                grid.EndEdit();
                int rows = csvDocument == null ? 0 : csvDocument.DataRows;
                int columns = csvDocument == null ? 0 : csvDocument.DataColumns;
                for (int row = 0; row < RowCount; row++)
                    for (int column = 0; column < ColumnCount; column++)
                    {
                        string value = Convert.ToString(grid[column, row].Value) ?? "";
                        if (value.Length == 0)
                            continue;
                        rows = Math.Max(rows, row + 1);
                        columns = Math.Max(columns, column + 1);
                    }
                var values = new List<string[]>();
                for (int row = 0; row < rows; row++)
                {
                    var fields = new string[columns];
                    for (int column = 0; column < columns; column++)
                        fields[column] = Convert.ToString(grid[column, row].Value) ?? "";
                    values.Add(fields);
                }
                CsvFile.Write(path, values, csvDocument);
                Encoding encoding = csvDocument == null ? new UTF8Encoding(true) : csvDocument.Encoding;
                string newline = csvDocument == null ? Environment.NewLine : csvDocument.NewLine;
                bool trailing = csvDocument == null || csvDocument.EndsWithNewLine;
                csvDocument = new CsvDocument
                {
                    Rows = values,
                    Encoding = encoding,
                    NewLine = newline,
                    EndsWithNewLine = trailing,
                    DataRows = rows,
                    DataColumns = columns
                };
                currentPath = path;
                dirty = false;
                otherSheetsDirty = false;
                lastState = CaptureSheet();
                lastState.RevisionId = nextRevision;
                savedRevision = nextRevision;
                MarkAllHistoriesSaved();
                UpdateTitle();
                status.Text = "Đã lưu " + Path.GetFileName(path) + " (CSV chỉ lưu dữ liệu ô)";
                return true;
            }
            catch (Exception error)
            {
                MessageBox.Show(this, "Không lưu được CSV: " + error.Message,
                    "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool WriteWorkbook(string path)
        {
            try
            {
                grid.EndEdit();
                SaveActiveSheet();
                var root = new XElement("workbook",
                    new XAttribute("rows", RowCount),
                    new XAttribute("columns", ColumnCount),
                    new XAttribute("theme", theme.Id),
                    new XAttribute("background",
                        ColorTranslator.ToHtml(sheetBackground)));
                foreach (SheetState state in sheets)
                    root.Add(SerializeSheet(SnapshotFromState(state)));
                SerializeWorkbookMetadata(root, namedRanges, pivots);
                new XDocument(root).Save(path);
                currentPath = path;
                csvDocument = null;
                dirty = false;
                otherSheetsDirty = false;
                lastState = CaptureSheet();
                lastState.RevisionId = nextRevision;
                savedRevision = nextRevision;
                MarkAllHistoriesSaved();
                UpdateTitle();
                status.Text = "Đã lưu " + Path.GetFileName(path);
                return true;
            }
            catch (Exception error)
            {
                MessageBox.Show(this, "Không lưu được tệp: " + error.Message,
                    "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void ApplyToSelection(Action<DataGridViewCell> action)
        {
            if (grid.CurrentCell == null)
                return;
            if (grid.SelectedCells.Count == 0)
                action(grid.CurrentCell);
            else
                foreach (DataGridViewCell cell in grid.SelectedCells)
                    action(cell);
            RecordChange();
            MarkDirty();
            UpdateSelection();
            grid.Invalidate();
        }

        private void ToggleFontStyle(FontStyle flag)
        {
            if (grid.CurrentCell == null)
                return;
            Font current = grid.CurrentCell.InheritedStyle.Font ?? grid.Font;
            bool enable = (current.Style & flag) == 0;
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                Font oldFont = cell.InheritedStyle.Font ?? grid.Font;
                FontStyle next = enable ? oldFont.Style | flag : oldFont.Style & ~flag;
                cell.Style.Font = new Font(oldFont.FontFamily, oldFont.Size, next);
            });
        }

        private void ApplyFontSize(float size)
        {
            if (size < 6F || size > 72F)
                return;
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                Font oldFont = cell.InheritedStyle.Font ?? grid.Font;
                cell.Style.Font = new Font(oldFont.FontFamily, size, oldFont.Style);
            });
        }

        private void ChooseColor(bool background)
        {
            using (var dialog = new ColorDialog())
            {
                dialog.FullOpen = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                ApplyToSelection(delegate(DataGridViewCell cell)
                {
                    if (background)
                        cell.Style.BackColor = dialog.Color;
                    else
                        cell.Style.ForeColor = dialog.Color;
                });
            }
        }

        private void ApplyAlignment(DataGridViewContentAlignment alignment)
        {
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                cell.Style.Alignment = alignment;
            });
        }

        private void CopySelected()
        {
            CopySelection(false);
        }

        private void ClearSelectedCells()
        {
            if (grid.CurrentCell == null)
                return;
            grid.EndEdit();
            var targets = new List<DataGridViewCell>();
            if (grid.SelectedCells.Count == 0)
                targets.Add(grid.CurrentCell);
            else
                foreach (DataGridViewCell cell in grid.SelectedCells)
                    targets.Add(cell);
            bool changed = false;
            loading = true;
            try
            {
                foreach (DataGridViewCell cell in targets)
                {
                    if (cell.Value == null)
                        continue;
                    cell.Value = null;
                    changed = true;
                }
            }
            finally
            {
                loading = false;
            }
            if (!changed)
                return;
            Recalculate();
            RecordChange();
            MarkDirty();
            UpdateSelection();
            status.Text = "Đã xóa nội dung " + targets.Count + " ô";
        }

        private void PasteSelected()
        {
            PasteClipboard(PasteKind.All);
        }

        private void GridCellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (suppressHeaderClick)
                return;
            if (e.RowIndex >= 0 && e.ColumnIndex == -1)
            {
                SelectHeader(e.RowIndex, true);
                if (e.Button == MouseButtons.Right)
                    rowContext.Show(grid, grid.PointToClient(Cursor.Position));
            }
            else if (e.ColumnIndex >= 0 && e.RowIndex == -1)
            {
                SelectHeader(e.ColumnIndex, false);
                if (e.Button == MouseButtons.Right)
                    columnContext.Show(grid, grid.PointToClient(Cursor.Position));
            }
        }

        private void SelectHeader(int index, bool row)
        {
            selectingHeader = true;
            try
            {
                grid.CurrentCell = row ? grid[0, index] : grid[index, 0];
                grid.ClearSelection();
                if (row)
                    for (int column = 0; column < ColumnCount; column++)
                        grid[column, index].Selected = true;
                else
                    for (int line = 0; line < RowCount; line++)
                        grid[index, line].Selected = true;
            }
            finally
            {
                selectingHeader = false;
            }
            UpdateSelection();
            grid.Invalidate();
        }

        private void InsertRow() { ChangeStructure(true, true); }
        private void DeleteRow() { ChangeStructure(true, false); }
        private void InsertColumn() { ChangeStructure(false, true); }
        private void DeleteColumn() { ChangeStructure(false, false); }

        private void ChangeStructure(bool row, bool insert)
        {
            if (grid.CurrentCell == null)
                return;
            grid.EndEdit();
            int index = row ? grid.CurrentCell.RowIndex : grid.CurrentCell.ColumnIndex;
            if (insert && index == (row ? RowCount : ColumnCount) - 1)
            {
                MessageBox.Show(this,
                    "Không thể chèn tại mép cuối của bảng 200 hàng × 26 cột.",
                    "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (insert && EdgeHasContent(row))
            {
                MessageBox.Show(this,
                    row ? "Không thể chèn: hàng 200 đang có dữ liệu hoặc định dạng."
                        : "Không thể chèn: cột Z đang có dữ liệu hoặc định dạng.",
                    "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            loading = true;
            grid.SuspendLayout();
            try
            {
                if (row)
                {
                    if (insert)
                    {
                        for (int line = RowCount - 1; line > index; line--)
                            CopyRow(line - 1, line);
                        ClearRow(index);
                    }
                    else
                    {
                        for (int line = index; line < RowCount - 1; line++)
                            CopyRow(line + 1, line);
                        ClearRow(RowCount - 1);
                    }
                }
                else
                {
                    if (insert)
                    {
                        for (int column = ColumnCount - 1; column > index; column--)
                            CopyColumn(column - 1, column);
                        ClearColumn(index);
                    }
                    else
                    {
                        for (int column = index; column < ColumnCount - 1; column++)
                            CopyColumn(column + 1, column);
                        ClearColumn(ColumnCount - 1);
                    }
                }
                RebaseFormulas(row, index, insert);
                if (csvDocument != null)
                {
                    if (row && (insert ? index <= csvDocument.DataRows :
                        index < csvDocument.DataRows))
                        csvDocument.DataRows = Math.Max(0, Math.Min(RowCount,
                            csvDocument.DataRows + (insert ? 1 : -1)));
                    if (!row && (insert ? index <= csvDocument.DataColumns :
                        index < csvDocument.DataColumns))
                        csvDocument.DataColumns = Math.Max(0, Math.Min(ColumnCount,
                            csvDocument.DataColumns + (insert ? 1 : -1)));
                }
            }
            finally
            {
                grid.ResumeLayout();
                loading = false;
            }
            SelectHeader(index, row);
            Recalculate();
            RecordChange();
            MarkDirty();
            status.Text = (insert ? "Đã chèn " : "Đã xóa ") +
                (row ? "hàng " + (index + 1) : "cột " + (char)('A' + index));
        }

        private bool EdgeHasContent(bool row)
        {
            if (row && grid.Rows[RowCount - 1].Height != grid.RowTemplate.Height)
                return true;
            if (!row && grid.Columns[ColumnCount - 1].Width != 120)
                return true;
            int count = row ? ColumnCount : RowCount;
            for (int i = 0; i < count; i++)
            {
                DataGridViewCell cell = row ? grid[i, RowCount - 1]
                    : grid[ColumnCount - 1, i];
                DataGridViewCellStyle style = cell.HasStyle ? cell.Style : null;
                if (cell.Value != null && Convert.ToString(cell.Value).Length > 0)
                    return true;
                if (style != null && (style.Font != null ||
                    !style.ForeColor.IsEmpty || !style.BackColor.IsEmpty ||
                    style.Alignment != DataGridViewContentAlignment.NotSet))
                    return true;
            }
            return false;
        }

        private void CopyCell(int sourceColumn, int sourceRow,
            int targetColumn, int targetRow)
        {
            DataGridViewCell source = grid[sourceColumn, sourceRow];
            DataGridViewCell target = grid[targetColumn, targetRow];
            target.Value = source.Value;
            target.Style = source.HasStyle ? new DataGridViewCellStyle(source.Style)
                : new DataGridViewCellStyle();
            target.Tag = CellExtras.Copy(source.Tag as CellExtras);
        }

        private void ClearCell(int column, int row)
        {
            DataGridViewCell cell = grid[column, row];
            cell.Value = null;
            cell.Style = new DataGridViewCellStyle();
            cell.Tag = null;
        }

        private void CopyRow(int source, int target)
        {
            for (int column = 0; column < ColumnCount; column++)
                CopyCell(column, source, column, target);
            grid.Rows[target].Height = grid.Rows[source].Height;
        }

        private void ClearRow(int row)
        {
            for (int column = 0; column < ColumnCount; column++)
                ClearCell(column, row);
            grid.Rows[row].Height = grid.RowTemplate.Height;
        }

        private void CopyColumn(int source, int target)
        {
            for (int row = 0; row < RowCount; row++)
                CopyCell(source, row, target, row);
            grid.Columns[target].Width = grid.Columns[source].Width;
        }

        private CellState CaptureCell(int column, int row)
        {
            DataGridViewCell cell = grid[column, row];
            return new CellState
            {
                Value = cell.Value,
                Style = cell.HasStyle ? new DataGridViewCellStyle(cell.Style) : null,
                Extras = CellExtras.Copy(cell.Tag as CellExtras)
            };
        }

        private void RestoreCell(int column, int row, CellState saved)
        {
            DataGridViewCell cell = grid[column, row];
            cell.Value = saved.Value;
            cell.Style = saved.Style == null ? new DataGridViewCellStyle() :
                new DataGridViewCellStyle(saved.Style);
            cell.Tag = CellExtras.Copy(saved.Extras);
        }

        private void MoveHeader(bool row, int source, int target)
        {
            int limit = row ? RowCount : ColumnCount;
            if (source < 0 || source >= limit || target < 0 || target >= limit ||
                source == target)
                return;
            grid.EndEdit();
            int count = row ? ColumnCount : RowCount;
            var held = new CellState[count];
            for (int i = 0; i < count; i++)
                held[i] = row ? CaptureCell(i, source) : CaptureCell(source, i);
            int heldSize = row ? grid.Rows[source].Height :
                grid.Columns[source].Width;

            loading = true;
            grid.SuspendLayout();
            try
            {
                if (source < target)
                    for (int index = source; index < target; index++)
                    {
                        if (row) CopyRow(index + 1, index);
                        else CopyColumn(index + 1, index);
                    }
                else
                    for (int index = source; index > target; index--)
                    {
                        if (row) CopyRow(index - 1, index);
                        else CopyColumn(index - 1, index);
                    }
                for (int i = 0; i < count; i++)
                {
                    if (row) RestoreCell(i, target, held[i]);
                    else RestoreCell(target, i, held[i]);
                }
                if (row) grid.Rows[target].Height = heldSize;
                else grid.Columns[target].Width = heldSize;
                for (int line = 0; line < RowCount; line++)
                    for (int column = 0; column < ColumnCount; column++)
                    {
                        DataGridViewCell cell = grid[column, line];
                        string raw = Convert.ToString(cell.Value) ?? "";
                        if (raw.StartsWith("=", StringComparison.Ordinal))
                            cell.Value = FormulaEngine.MoveStructureReferences(
                                raw, row, source, target);
                    }
            }
            finally
            {
                grid.ResumeLayout();
                loading = false;
            }
            SelectHeader(target, row);
            Recalculate();
            RecordChange();
            MarkDirty();
            status.Text = row ?
                "Đã chuyển hàng " + (source + 1) + " đến hàng " + (target + 1) :
                "Đã chuyển cột " + (char)('A' + source) + " đến cột " +
                    (char)('A' + target);
        }

        private void ClearColumn(int column)
        {
            for (int row = 0; row < RowCount; row++)
                ClearCell(column, row);
            grid.Columns[column].Width = 120;
        }

        private void RebaseFormulas(bool row, int index, bool insert)
        {
            for (int line = 0; line < RowCount; line++)
                for (int column = 0; column < ColumnCount; column++)
                {
                    DataGridViewCell cell = grid[column, line];
                    string raw = Convert.ToString(cell.Value) ?? "";
                    if (raw.StartsWith("=", StringComparison.Ordinal))
                        cell.Value = FormulaEngine.ShiftStructureReferences(raw,
                            row, index, insert, RowCount, ColumnCount);
                }
        }

        private void GridCellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (headerDragPending || fillDragging) return;
            if (grid.CurrentCell == null ||
                fillSelection.IsEmpty ||
                e.RowIndex != fillSelection.Bottom - 1 ||
                e.ColumnIndex != fillSelection.Right - 1)
            { grid.Cursor = Cursors.Default; return; }
            Rectangle rectangle = grid.GetCellDisplayRectangle(
                e.ColumnIndex, e.RowIndex, false);
            grid.Cursor = e.X >= rectangle.Width - 11 &&
                e.Y >= rectangle.Height - 11 ? Cursors.Cross : Cursors.Default;
        }

        private void GridCellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left &&
                ((e.RowIndex >= 0 && e.ColumnIndex == -1) ||
                 (e.ColumnIndex >= 0 && e.RowIndex == -1)))
            {
                if (e.ColumnIndex == -1 &&
                    e.Y >= grid.Rows[e.RowIndex].Height - 6)
                    return;
                if (e.RowIndex == -1 &&
                    e.X >= grid.Columns[e.ColumnIndex].Width - 6)
                    return;
                headerDragPending = true;
                headerDragging = false;
                headerDragRow = e.ColumnIndex == -1;
                headerSource = headerDragRow ? e.RowIndex : e.ColumnIndex;
                headerTarget = headerSource;
                headerStartPoint = grid.PointToClient(Control.MousePosition);
                grid.Capture = true;
                return;
            }
            if (e.Button != MouseButtons.Left || grid.CurrentCell == null ||
                fillSelection.IsEmpty ||
                e.RowIndex != fillSelection.Bottom - 1 ||
                e.ColumnIndex != fillSelection.Right - 1)
                return;
            Rectangle rectangle = grid.GetCellDisplayRectangle(
                e.ColumnIndex, e.RowIndex, false);
            if (e.X < rectangle.Width - 11 || e.Y < rectangle.Height - 11)
                return;
            fillDragging = true;
            fillDragSource = fillSelection;
            fillSourceRow = fillSelection.Bottom - 1;
            fillSourceColumn = fillSelection.Right - 1;
            fillTargetRow = e.RowIndex;
            fillTargetColumn = e.ColumnIndex;
            grid.Capture = true;
            grid.Cursor = Cursors.Cross;
        }

        private void GridMouseMove(object sender, MouseEventArgs e)
        {
            if (headerDragPending)
            {
                if ((e.Button & MouseButtons.Left) == 0)
                    return;
                if (!headerDragging &&
                    Math.Abs(e.X - headerStartPoint.X) < 5 &&
                    Math.Abs(e.Y - headerStartPoint.Y) < 5)
                    return;
                headerDragging = true;
                grid.Cursor = headerDragRow ? Cursors.SizeNS : Cursors.SizeWE;
                DataGridView.HitTestInfo targetHit = grid.HitTest(e.X, e.Y);
                int target = headerDragRow ? targetHit.RowIndex :
                    targetHit.ColumnIndex;
                if (target >= 0 && target != headerTarget)
                {
                    headerTarget = target;
                    status.Text = headerDragRow ?
                        "Thả để chuyển đến hàng " + (target + 1) :
                        "Thả để chuyển đến cột " + (char)('A' + target);
                    grid.Invalidate();
                }
                return;
            }
            if (!fillDragging)
                return;
            DataGridView.HitTestInfo hit = grid.HitTest(e.X, e.Y);
            if (hit.RowIndex < 0 || hit.ColumnIndex < 0 ||
                (hit.RowIndex == fillTargetRow && hit.ColumnIndex == fillTargetColumn))
                return;
            fillTargetRow = hit.RowIndex;
            fillTargetColumn = hit.ColumnIndex;
            status.Text = "Kéo để sao chép đến " +
                ((char)('A' + fillTargetColumn)) + (fillTargetRow + 1);
            grid.Invalidate();
        }

        private void GridMouseUp(object sender, MouseEventArgs e)
        {
            if (headerDragPending)
            {
                DataGridView.HitTestInfo targetHit = grid.HitTest(e.X, e.Y);
                int target = headerDragRow ? targetHit.RowIndex :
                    targetHit.ColumnIndex;
                bool move = headerDragging && target >= 0 &&
                    target != headerSource;
                int source = headerSource;
                bool row = headerDragRow;
                headerDragPending = false;
                headerDragging = false;
                grid.Capture = false;
                grid.Cursor = Cursors.Default;
                grid.Invalidate();
                if (move)
                {
                    suppressHeaderClick = true;
                    MoveHeader(row, source, target);
                    BeginInvoke((MethodInvoker)delegate
                    {
                        suppressHeaderClick = false;
                    });
                }
                return;
            }
            if (!fillDragging)
                return;
            DataGridView.HitTestInfo hit = grid.HitTest(e.X, e.Y);
            if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
            {
                fillTargetRow = hit.RowIndex;
                fillTargetColumn = hit.ColumnIndex;
            }
            fillDragging = false;
            grid.Capture = false;
            grid.Cursor = Cursors.Default;
            AutoFillSelection(fillDragSource, fillTargetRow, fillTargetColumn);
            grid.Invalidate();
        }

        private void FillRange(int sourceRow, int sourceColumn,
            int targetRow, int targetColumn)
        {
            if (sourceRow == targetRow && sourceColumn == targetColumn)
                return;
            DataGridViewCell source = grid[sourceColumn, sourceRow];
            string raw = Convert.ToString(source.Value) ?? "";
            DataGridViewCellStyle sourceStyle = source.HasStyle
                ? new DataGridViewCellStyle(source.Style)
                : new DataGridViewCellStyle();
            CellExtras sourceExtras = CellExtras.Copy(source.Tag as CellExtras);
            loading = true;
            grid.SuspendLayout();
            try
            {
                for (int row = Math.Min(sourceRow, targetRow);
                    row <= Math.Max(sourceRow, targetRow); row++)
                {
                    for (int column = Math.Min(sourceColumn, targetColumn);
                        column <= Math.Max(sourceColumn, targetColumn); column++)
                    {
                        if (row == sourceRow && column == sourceColumn)
                            continue;
                        DataGridViewCell cell = grid[column, row];
                        cell.Value = FormulaEngine.ShiftReferences(raw,
                            row - sourceRow, column - sourceColumn,
                            RowCount, ColumnCount);
                        cell.Style = new DataGridViewCellStyle(sourceStyle);
                        cell.Tag = CellExtras.Copy(sourceExtras);
                    }
                }
            }
            finally
            {
                grid.ResumeLayout();
                loading = false;
            }
            Recalculate();
            RecordChange();
            MarkDirty();
            UpdateSelection();
        }

        private void GridKeyDown(object sender, KeyEventArgs e)
        {
            if (HandleEditingShortcut(e.KeyData))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Delete && !grid.IsCurrentCellInEditMode)
            {
                ClearSelectedCells();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        protected override bool ProcessCmdKey(ref Message message, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.K)) { ShowCommandPalette(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.P)) { ShowQuickNavigator(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.I)) { ToggleInspector(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.M)) { ToggleMinimalMode(); return true; }
            if (addressBox.Focused || contentBox.Focused)
                return base.ProcessCmdKey(ref message, keyData);
            if (HandleEditingShortcut(keyData)) return true;
            if (keyData == (Keys.Control | Keys.Z))
            {
                Undo();
                return true;
            }
            if (keyData == (Keys.Control | Keys.Y))
            {
                Redo();
                return true;
            }
            return base.ProcessCmdKey(ref message, keyData);
        }
    }
}

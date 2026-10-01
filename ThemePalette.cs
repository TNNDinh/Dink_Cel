using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed class ThemePalette
    {
        public readonly string Id;
        public readonly string Name;
        public readonly string Description;
        public readonly Color Surface;
        public readonly Color Chrome;
        public readonly Color Sheet;
        public readonly Color Header;
        public readonly Color HeaderActive;
        public readonly Color Border;
        public readonly Color Text;
        public readonly Color Muted;
        public readonly Color Accent;
        public readonly Color Selection;
        public readonly Color Logo;
        public readonly bool Dark;

        private ThemePalette(string id, string name, string description,
            string surface, string chrome, string sheet, string header,
            string headerActive, string border, string text, string muted,
            string accent, string selection, string logo, bool dark)
        {
            Id = id;
            Name = name;
            Description = description;
            Surface = ColorTranslator.FromHtml(surface);
            Chrome = ColorTranslator.FromHtml(chrome);
            Sheet = ColorTranslator.FromHtml(sheet);
            Header = ColorTranslator.FromHtml(header);
            HeaderActive = ColorTranslator.FromHtml(headerActive);
            Border = ColorTranslator.FromHtml(border);
            Text = ColorTranslator.FromHtml(text);
            Muted = ColorTranslator.FromHtml(muted);
            Accent = ColorTranslator.FromHtml(accent);
            Selection = ColorTranslator.FromHtml(selection);
            Logo = ColorTranslator.FromHtml(logo);
            Dark = dark;
        }

        public static readonly ThemePalette[] All = new ThemePalette[]
        {
            new ThemePalette("ocean", "Xanh dịu", "Mát mắt, dễ đọc",
                "#DCE7F3", "#E6EEF8", "#E8F0F8", "#DDE8F4",
                "#BED5F4", "#C2D1E0", "#1E2B3A", "#5C6B7A",
                "#2563EB", "#CFE3FF", "#168A5A", false),
            new ThemePalette("mint", "Bạc hà", "Sáng và nhẹ",
                "#D9EBDF", "#DCEFE5", "#E6F5EC", "#D0E8D8",
                "#ACDCC2", "#B5D5C1", "#1B3328", "#557166",
                "#168A57", "#C4EBD5", "#168A57", false),
            new ThemePalette("sand", "Kem ấm", "Dịu vào buổi tối",
                "#EEE2D1", "#F3E8D8", "#F8EEDF", "#EBDDC9",
                "#E7C59B", "#D7C6AE", "#392C20", "#776A5A",
                "#BB6B22", "#F3DEBF", "#BB6B22", false),
            new ThemePalette("lavender", "Tím sương", "Nhẹ và mềm",
                "#E4DCEE", "#EBE6F4", "#F1EBFA", "#E2D9F1",
                "#D3BFF2", "#D2C5E2", "#30273D", "#6D617A",
                "#7851B8", "#E0D0F5", "#7851B8", false),
            new ThemePalette("slate", "Xám xanh", "Trung tính",
                "#D4DEE7", "#DCE3EA", "#E9EEF3", "#D3DDE6",
                "#B9CDDD", "#BCCBD8", "#1D2A33", "#5D6B75",
                "#3B6F91", "#D1E1ED", "#3B6F91", false),
            new ThemePalette("night", "Tối dịu", "Giảm chói",
                "#202A36", "#263241", "#2D3B4C", "#344456",
                "#3B5B7A", "#4B5D70", "#E8EFF7", "#B4C1CF",
                "#81B6FF", "#3A5068", "#2EA975", true)
        };

        public static ThemePalette Find(string id)
        {
            foreach (ThemePalette item in All)
                if (String.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
                    return item;
            return null;
        }
    }

    internal sealed class ThemeCardButton : Button
    {
        private readonly ThemePalette palette;
        private readonly bool selected;

        public ThemeCardButton(ThemePalette palette, bool selected)
        {
            this.palette = palette;
            this.selected = selected;
            Size = new Size(296, 98);
            Margin = new Padding(7);
            Cursor = Cursors.Hand;
            TabStop = true;
            AccessibleName = "Giao diện " + palette.Name;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle card = new Rectangle(2, 2, Width - 5, Height - 5);
            using (GraphicsPath path = Rounded(card, 14))
            using (var brush = new SolidBrush(palette.Chrome))
            using (var pen = new Pen(selected || Focused
                ? palette.Accent : palette.Border,
                selected ? 3F : Focused ? 2F : 1F))
            {
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            }

            Rectangle preview = new Rectangle(168, 14, 108, 68);
            using (var brush = new SolidBrush(palette.Sheet))
                e.Graphics.FillRectangle(brush, preview);
            using (var brush = new SolidBrush(palette.Header))
            {
                e.Graphics.FillRectangle(brush, preview.X, preview.Y,
                    preview.Width, 15);
                e.Graphics.FillRectangle(brush, preview.X, preview.Y, 19,
                    preview.Height);
            }
            using (var pen = new Pen(palette.Border))
            {
                e.Graphics.DrawRectangle(pen, preview);
                e.Graphics.DrawLine(pen, preview.X + 19, preview.Y,
                    preview.X + 19, preview.Bottom);
                e.Graphics.DrawLine(pen, preview.X + 62, preview.Y,
                    preview.X + 62, preview.Bottom);
                e.Graphics.DrawLine(pen, preview.X, preview.Y + 15,
                    preview.Right, preview.Y + 15);
                e.Graphics.DrawLine(pen, preview.X, preview.Y + 42,
                    preview.Right, preview.Y + 42);
            }
            using (var pen = new Pen(palette.Accent, 2F))
                e.Graphics.DrawRectangle(pen, preview.X + 20, preview.Y + 16,
                    41, 25);

            using (var titleFont = new Font("Segoe UI", 11F, FontStyle.Bold))
            using (var bodyFont = new Font("Segoe UI", 9F))
            {
                TextRenderer.DrawText(e.Graphics, palette.Name, titleFont,
                    new Rectangle(16, 18, 145, 28), palette.Text,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                TextRenderer.DrawText(e.Graphics, palette.Description, bodyFont,
                    new Rectangle(16, 47, 145, 30), palette.Muted,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
        }

        private static GraphicsPath Rounded(Rectangle rectangle, int radius)
        {
            int diameter = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Top,
                diameter, diameter, 270, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter,
                diameter, diameter, 0, 90);
            path.AddArc(rectangle.Left, rectangle.Bottom - diameter,
                diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal sealed class ThemePickerForm : Form
    {
        public ThemePalette SelectedTheme { get; private set; }

        public ThemePickerForm(ThemePalette current)
        {
            Text = "Chọn giao diện - DinkCel";
            ClientSize = new Size(656, 472);
            MinimumSize = Size;
            MaximumSize = Size;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(241, 245, 249);
            Font = new Font("Segoe UI", 10F);

            var title = new Label();
            title.Text = "Chọn màu giao diện";
            title.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            title.ForeColor = Color.FromArgb(30, 41, 59);
            title.Bounds = new Rectangle(18, 15, 500, 35);
            Controls.Add(title);

            var hint = new Label();
            hint.Text = "Chọn một mẫu để đổi cả cửa sổ và bảng tính.";
            hint.ForeColor = Color.FromArgb(86, 102, 119);
            hint.Bounds = new Rectangle(20, 53, 550, 25);
            Controls.Add(hint);

            var cards = new FlowLayoutPanel();
            cards.Bounds = new Rectangle(10, 84, 636, 352);
            cards.WrapContents = true;
            cards.FlowDirection = FlowDirection.LeftToRight;
            cards.BackColor = BackColor;
            Controls.Add(cards);
            foreach (ThemePalette palette in ThemePalette.All)
            {
                var card = new ThemeCardButton(palette, palette == current);
                card.Click += delegate
                {
                    SelectedTheme = palette;
                    DialogResult = DialogResult.OK;
                    Close();
                };
                cards.Controls.Add(card);
            }

            var cancel = new Button();
            cancel.Text = "Đóng";
            cancel.Size = new Size(84, 28);
            cancel.Location = new Point(550, 440);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
        }
    }
}

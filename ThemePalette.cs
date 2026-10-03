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
        public Color Hover { get { return Mix(Chrome, Accent, Dark ? 0.16F : 0.07F); } }
        public Color AccentSoft { get { return Mix(Sheet, Accent, Dark ? 0.24F : 0.11F); } }
        public Color GridLine { get { return Mix(Sheet, Border, Dark ? 0.46F : 0.55F); } }

        private static Color Mix(Color first, Color second, float weight)
        {
            return Color.FromArgb((int)(first.R * (1 - weight) + second.R * weight),
                (int)(first.G * (1 - weight) + second.G * weight),
                (int)(first.B * (1 - weight) + second.B * weight));
        }

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
            new ThemePalette("light", "DinkCel Light", "Sáng dịu, tập trung dữ liệu",
                "#F4F5F8", "#FBFBFD", "#FFFFFF", "#F8F9FC",
                "#E8E7FF", "#DCE0E9", "#222437", "#777C90",
                "#6D5EF7", "#ECEAFF", "#6957EA", false),
            new ThemePalette("dark", "DinkCel Dark", "Than chì, tím dịu",
                "#1A1C25", "#242633", "#20222D", "#292C39",
                "#3B385E", "#3A3D4E", "#EFF0F7", "#A4A7B8",
                "#A497FF", "#3D385B", "#8C7CFF", true),
            new ThemePalette("midnight", "DinkCel Midnight", "Đen sâu, điểm nhấn cyan",
                "#11131B", "#1A1D29", "#171A24", "#202431",
                "#2A3A53", "#303647", "#F1F4FC", "#9DA7BA",
                "#73CEEB", "#254655", "#6E9FFF", true),
            new ThemePalette("paper", "DinkCel Paper", "Trắng ấm, mực nâu",
                "#F3F0E8", "#FBF9F3", "#FFFEFA", "#F7F3EA",
                "#EAE5D2", "#E2DCCD", "#3B362F", "#827C70",
                "#86724D", "#F0EAD7", "#816D4F", false),
            new ThemePalette("solar", "DinkCel Solar", "Ấm áp, cam tinh tế",
                "#F9F4EC", "#FFFDF8", "#FFFFFF", "#FCF7EF",
                "#FFE9D2", "#E9E0D4", "#352B29", "#8A7972",
                "#E57E36", "#FFF0DD", "#E0783E", false),
            new ThemePalette("mint", "DinkCel Mint", "Xanh mát, gọn gàng",
                "#EFF6F3", "#FAFDFC", "#FFFFFF", "#F4FAF7",
                "#DCF5E9", "#D6E6DF", "#21372F", "#6C8278",
                "#279C79", "#E2F5ED", "#31A981", false)
        };

        public static ThemePalette Find(string id)
        {
            if (String.Equals(id, "ocean", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(id, "lavender", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(id, "slate", StringComparison.OrdinalIgnoreCase)) id = "light";
            else if (String.Equals(id, "night", StringComparison.OrdinalIgnoreCase)) id = "dark";
            else if (String.Equals(id, "sand", StringComparison.OrdinalIgnoreCase)) id = "paper";
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
            BackColor = current.Surface;
            Font = new Font("Segoe UI", 10F);

            var title = new Label();
            title.Text = "Chọn màu giao diện";
            title.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            title.ForeColor = current.Text;
            title.Bounds = new Rectangle(18, 15, 500, 35);
            Controls.Add(title);

            var hint = new Label();
            hint.Text = "Chọn một mẫu để đổi cả cửa sổ và bảng tính.";
            hint.ForeColor = current.Muted;
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
            cancel.FlatStyle = FlatStyle.Flat;
            cancel.BackColor = current.Hover;
            cancel.ForeColor = current.Text;
            Controls.Add(cancel);
            CancelButton = cancel;
        }
    }
}

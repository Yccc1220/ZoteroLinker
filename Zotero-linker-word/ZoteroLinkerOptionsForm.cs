using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace Zotero_linker
{
    public partial class ZoteroLinkerOptionsForm : Form
    {
        private ComboBox colorComboBox;
        private NumericUpDown fontSizeUpDown;
        private Button applyColorButton;
        private Button applyFontSizeButton;

        public ZoteroLinkerOptionsForm()
            : this(ZoteroLinkerOptions.Load())
        {
        }

        internal ZoteroLinkerOptionsForm(ZoteroLinkerOptions options)
        {
            const int actionButtonWidth = 128;
            const int actionButtonHeight = 30;

            if (options == null)
            {
                options = new ZoteroLinkerOptions();
            }

            SuspendLayout();

            Text = "Zotero Linker Options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = SystemFonts.MessageBoxFont;
            Padding = new Padding(16, 14, 16, 14);

            Label colorLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Text = "Citation color"
            };

            colorComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DrawMode = DrawMode.OwnerDrawFixed,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                ItemHeight = Math.Max(Font.Height + 8, 24),
                Margin = new Padding(0, 0, 8, 0),
                Width = Math.Max(220, TextRenderer.MeasureText("Green (#00FF00)", Font).Width + 56)
            };
            colorComboBox.DropDownWidth = colorComboBox.Width;
            colorComboBox.DrawItem += ColorComboBox_DrawItem;
            colorComboBox.Items.Add(new ColorOption("Red", "#FF0000", Color.FromArgb(255, 0, 0)));
            colorComboBox.Items.Add(new ColorOption("Blue", "#0000FF", Color.FromArgb(0, 0, 255)));
            colorComboBox.Items.Add(new ColorOption("Green", "#00FF00", Color.FromArgb(0, 255, 0)));
            colorComboBox.Items.Add(new ColorOption("Black", "#000000", Color.FromArgb(0, 0, 0)));
            colorComboBox.SelectedItem = FindColorOption(options.ColorHex ?? "#FF0000");

            Label fontLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Text = "Citation font size (pt)"
            };

            fontSizeUpDown = new NumericUpDown
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                DecimalPlaces = 1,
                Minimum = (decimal)ZoteroLinkerOptions.MinFontSize,
                Maximum = (decimal)ZoteroLinkerOptions.MaxFontSize,
                Margin = new Padding(0, 0, 8, 0),
                Increment = 0.5M,
                Value = (decimal)ZoteroLinkerOptions.ClampFontSize(options.FontSize),
                Width = Math.Max(100, TextRenderer.MeasureText("00.0", Font).Width + 44)
            };

            applyColorButton = new Button
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                MinimumSize = new Size(actionButtonWidth, actionButtonHeight),
                Size = new Size(actionButtonWidth, actionButtonHeight),
                Text = "Apply Color"
            };
            applyColorButton.Click += ApplyColorButton_Click;

            applyFontSizeButton = new Button
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                MinimumSize = new Size(actionButtonWidth, actionButtonHeight),
                Size = new Size(actionButtonWidth, actionButtonHeight),
                Text = "Apply Font Size"
            };
            applyFontSizeButton.Click += ApplyFontSizeButton_Click;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                RowCount = 4
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, colorComboBox.Width));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, actionButtonWidth));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(colorLabel, 0, 0);
            layout.SetColumnSpan(colorLabel, 2);
            layout.Controls.Add(colorComboBox, 0, 1);
            layout.Controls.Add(applyColorButton, 1, 1);
            layout.Controls.Add(fontLabel, 0, 2);
            layout.SetColumnSpan(fontLabel, 2);
            layout.Controls.Add(fontSizeUpDown, 0, 3);
            layout.Controls.Add(applyFontSizeButton, 1, 3);
            Controls.Add(layout);

            ResumeLayout(false);
            PerformLayout();
        }

        internal event EventHandler ApplyColorRequested;

        internal event EventHandler ApplyFontSizeRequested;

        internal string ColorHex
        {
            get
            {
                return colorComboBox.SelectedItem is ColorOption option
                    ? option.Hex
                    : "#FF0000";
            }
        }

        internal float FontSize
        {
            get { return (float)fontSizeUpDown.Value; }
        }

        private void ApplyColorButton_Click(object sender, EventArgs e)
        {
            EventHandler handler = ApplyColorRequested;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void ApplyFontSizeButton_Click(object sender, EventArgs e)
        {
            EventHandler handler = ApplyFontSizeRequested;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private ColorOption FindColorOption(string colorHex)
        {
            string normalized = ZoteroLinkerOptions.NormalizeColorHex(colorHex);
            foreach (ColorOption option in colorComboBox.Items)
            {
                if (string.Equals(option.Hex, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return option;
                }
            }

            return (ColorOption)colorComboBox.Items[0];
        }

        private void ColorComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0)
            {
                return;
            }

            ColorOption option = (ColorOption)colorComboBox.Items[e.Index];
            int swatchWidth = Math.Max(18, e.Font.Height);
            int swatchHeight = Math.Max(12, e.Font.Height - 4);
            int swatchTop = e.Bounds.Top + Math.Max(2, (e.Bounds.Height - swatchHeight) / 2);
            Rectangle swatch = new Rectangle(e.Bounds.Left + 6, swatchTop, swatchWidth, swatchHeight);
            using (SolidBrush brush = new SolidBrush(option.Color))
            {
                e.Graphics.FillRectangle(brush, swatch);
            }

            e.Graphics.DrawRectangle(SystemPens.ControlDark, swatch);
            Rectangle textBounds = new Rectangle(
                swatch.Right + 8,
                e.Bounds.Top,
                Math.Max(0, e.Bounds.Right - swatch.Right - 12),
                e.Bounds.Height);
            TextRenderer.DrawText(
                e.Graphics,
                string.Format(CultureInfo.InvariantCulture, "{0} ({1})", option.Name, option.Hex),
                e.Font,
                textBounds,
                e.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);

            e.DrawFocusRectangle();
        }

        private class ColorOption
        {
            internal ColorOption(string name, string hex, Color color)
            {
                Name = name;
                Hex = hex;
                Color = color;
            }

            internal string Name { get; private set; }
            internal string Hex { get; private set; }
            internal Color Color { get; private set; }

            public override string ToString()
            {
                return Name;
            }
        }
    }
}

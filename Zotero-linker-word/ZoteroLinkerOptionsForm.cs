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
        private Button saveButton;
        private Button cancelButton;

        public ZoteroLinkerOptionsForm()
            : this(ZoteroLinkerOptions.Load())
        {
        }

        internal ZoteroLinkerOptionsForm(ZoteroLinkerOptions options)
        {
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
                DrawMode = DrawMode.OwnerDrawFixed,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                ItemHeight = Math.Max(Font.Height + 8, 24),
                Margin = new Padding(0, 0, 0, 12),
                Width = Math.Max(220, TextRenderer.MeasureText("Green (#00FF00)", Font).Width + 56)
            };
            colorComboBox.DropDownWidth = colorComboBox.Width;
            colorComboBox.DrawItem += ColorComboBox_DrawItem;
            colorComboBox.Items.Add(new ColorOption("Red", "#FF0000", Color.FromArgb(255, 0, 0)));
            colorComboBox.Items.Add(new ColorOption("Blue", "#0000FF", Color.FromArgb(0, 0, 255)));
            colorComboBox.Items.Add(new ColorOption("Green", "#00FF00", Color.FromArgb(0, 255, 0)));
            colorComboBox.Items.Add(new ColorOption("Black", "#000000", Color.FromArgb(0, 0, 0)));
            colorComboBox.SelectedIndex = FindColorIndex(options.ColorHex ?? "#FF0000");

            Label fontLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Text = "Citation font size (pt)"
            };

            fontSizeUpDown = new NumericUpDown
            {
                AutoSize = true,
                DecimalPlaces = 1,
                Minimum = (decimal)ZoteroLinkerOptions.MinFontSize,
                Maximum = (decimal)ZoteroLinkerOptions.MaxFontSize,
                Margin = new Padding(0, 0, 0, 14),
                Increment = 0.5M,
                Value = (decimal)ZoteroLinkerOptions.ClampFontSize(options.FontSize),
                Width = Math.Max(100, TextRenderer.MeasureText("00.0", Font).Width + 44)
            };

            saveButton = new Button
            {
                AutoSize = true,
                MinimumSize = new Size(84, 0),
                Text = "Save",
                DialogResult = DialogResult.OK
            };

            cancelButton = new Button
            {
                AutoSize = true,
                MinimumSize = new Size(84, 0),
                Text = "Cancel",
                DialogResult = DialogResult.Cancel
            };

            FlowLayoutPanel buttonsPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Right,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0),
                WrapContents = false
            };
            buttonsPanel.Controls.Add(saveButton);
            buttonsPanel.Controls.Add(cancelButton);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                RowCount = 5
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(colorLabel, 0, 0);
            layout.Controls.Add(colorComboBox, 0, 1);
            layout.Controls.Add(fontLabel, 0, 2);
            layout.Controls.Add(fontSizeUpDown, 0, 3);
            layout.Controls.Add(buttonsPanel, 0, 4);
            Controls.Add(layout);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
            ResumeLayout(false);
            PerformLayout();
        }

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

        private int FindColorIndex(string colorHex)
        {
            string normalized = ZoteroLinkerOptions.NormalizeColorHex(colorHex);
            for (int index = 0; index < colorComboBox.Items.Count; index += 1)
            {
                ColorOption option = (ColorOption)colorComboBox.Items[index];
                if (string.Equals(option.Hex, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }

            return 0;
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

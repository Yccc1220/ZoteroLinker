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

            Text = "Zotero Linker Options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(320, 180);
            Font = SystemFonts.MessageBoxFont;

            Label colorLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 18),
                Text = "Citation color"
            };

            colorComboBox = new ComboBox
            {
                DrawMode = DrawMode.OwnerDrawFixed,
                DropDownStyle = ComboBoxStyle.DropDownList,
                ItemHeight = 22,
                Location = new Point(16, 40),
                Width = 170
            };
            colorComboBox.DrawItem += ColorComboBox_DrawItem;
            colorComboBox.Items.Add(new ColorOption("Red", "#FF0000", Color.FromArgb(255, 0, 0)));
            colorComboBox.Items.Add(new ColorOption("Blue", "#0000FF", Color.FromArgb(0, 0, 255)));
            colorComboBox.Items.Add(new ColorOption("Green", "#00FF00", Color.FromArgb(0, 255, 0)));
            colorComboBox.Items.Add(new ColorOption("Black", "#000000", Color.FromArgb(0, 0, 0)));
            colorComboBox.SelectedIndex = FindColorIndex(options.ColorHex ?? "#FF0000");

            Label fontLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 80),
                Text = "Citation font size (pt)"
            };

            fontSizeUpDown = new NumericUpDown
            {
                Location = new Point(16, 102),
                Width = 100,
                DecimalPlaces = 1,
                Minimum = (decimal)ZoteroLinkerOptions.MinFontSize,
                Maximum = (decimal)ZoteroLinkerOptions.MaxFontSize,
                Increment = 0.5M,
                Value = (decimal)ZoteroLinkerOptions.ClampFontSize(options.FontSize)
            };

            saveButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Location = new Point(136, 136),
                Width = 80
            };

            cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(224, 136),
                Width = 80
            };

            Controls.Add(colorLabel);
            Controls.Add(colorComboBox);
            Controls.Add(fontLabel);
            Controls.Add(fontSizeUpDown);
            Controls.Add(saveButton);
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
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
            Rectangle swatch = new Rectangle(e.Bounds.Left + 6, e.Bounds.Top + 5, 18, 12);
            using (SolidBrush brush = new SolidBrush(option.Color))
            {
                e.Graphics.FillRectangle(brush, swatch);
            }

            e.Graphics.DrawRectangle(SystemPens.ControlDark, swatch);
            using (SolidBrush textBrush = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(
                    string.Format(CultureInfo.InvariantCulture, "{0} ({1})", option.Name, option.Hex),
                    e.Font,
                    textBrush,
                    e.Bounds.Left + 32,
                    e.Bounds.Top + 3);
            }

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

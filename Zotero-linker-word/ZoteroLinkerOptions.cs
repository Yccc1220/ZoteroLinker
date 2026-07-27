using Microsoft.Win32;
using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Zotero_linker
{
    internal sealed class ZoteroLinkerOptions
    {
        private const string RegistryPath = @"Software\ZoteroLinker";
        private const string ColorHexName = "CitationColorHex";
        private const string FontSizeName = "CitationFontSize";
        private const string DefaultColorHexValue = "#FF0000";
        private const float DefaultFontSizeValue = 10f;
        private const float MinFontSizeValue = 6f;
        private const float MaxFontSizeValue = 24f;

        internal ZoteroLinkerOptions()
        {
            ColorHex = DefaultColorHexValue;
            FontSize = DefaultFontSizeValue;
        }

        internal string ColorHex { get; set; }

        internal float FontSize { get; set; }

        internal Color CitationColor
        {
            get { return ParseColorHex(ColorHex); }
        }

        internal static float MinFontSize
        {
            get { return MinFontSizeValue; }
        }

        internal static float MaxFontSize
        {
            get { return MaxFontSizeValue; }
        }

        internal static ZoteroLinkerOptions Load()
        {
            ZoteroLinkerOptions options = new ZoteroLinkerOptions();
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
            {
                if (key == null)
                {
                    return options;
                }

                string colorHex = key.GetValue(ColorHexName) as string;
                if (IsValidColorHex(colorHex))
                {
                    options.ColorHex = NormalizeColorHex(colorHex);
                }

                object fontSizeValue = key.GetValue(FontSizeName);
                float fontSize;
                if (fontSizeValue != null &&
                    float.TryParse(Convert.ToString(fontSizeValue, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out fontSize))
                {
                    options.FontSize = ClampFontSize(fontSize);
                }
            }

            return options;
        }

        internal void Save()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
            {
                key.SetValue(ColorHexName, NormalizeColorHex(ColorHex), RegistryValueKind.String);
                key.SetValue(FontSizeName, ClampFontSize(FontSize).ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
            }
        }

        internal static bool IsValidColorHex(string value)
        {
            return Regex.IsMatch(value ?? string.Empty, "^#[0-9a-fA-F]{6}$");
        }

        internal static string NormalizeColorHex(string value)
        {
            return IsValidColorHex(value) ? value.Trim().ToUpperInvariant() : DefaultColorHexValue;
        }

        internal static float ClampFontSize(float value)
        {
            if (value < MinFontSizeValue)
            {
                return MinFontSizeValue;
            }

            if (value > MaxFontSizeValue)
            {
                return MaxFontSizeValue;
            }

            return value;
        }

        private static Color ParseColorHex(string value)
        {
            string normalized = NormalizeColorHex(value).TrimStart('#');
            int rgb = int.Parse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return Color.FromArgb((rgb >> 16) & 0xff, (rgb >> 8) & 0xff, rgb & 0xff);
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Zotero_linker
{
    public partial class FunctionArea
    {
        private void FunctionArea_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void ConfigureRuntimeLayout()
        {
            SetLargeControlSize(buttonLinkCitations);
            SetLargeControlSize(buttonRemoveLinks);
            SetLargeControlSize(buttonRestoreFormatting);
            SetLargeControlSize(buttonOptions);
        }

        private void buttonLinkCitations_Click(object sender, RibbonControlEventArgs e)
        {
            LinkCitations();
        }

        private void buttonRemoveLinks_Click(object sender, RibbonControlEventArgs e)
        {
            RemoveLinks();
        }

        private void buttonRestoreFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            RestoreFormatting();
        }

        private void buttonOptions_Click(object sender, RibbonControlEventArgs e)
        {
            ShowOptions();
        }

        internal void LinkCitations()
        {
            RunWithActiveDocument(document =>
            {
                ZoteroLinkerOptions options = GetOptions();
                LinkResult result = GetLinkerService().LinkCitations(
                    document,
                    options.CitationColor,
                    options.FontSize);

                ShowStatus(
                    "Link citations",
                    string.Format(
                        "Links {0}; backlinks {1}; hidden {2}",
                        result.Linked,
                        result.LinkedBacklinks,
                        result.SkippedCompressedItems),
                    string.Format(
                        "Unparsed {0}; missing bib {1}; bib fail {2}; range fail {3}",
                        result.SkippedMultiItem,
                        result.SkippedMissingBibliography,
                        result.FailedBibliographyMatch,
                        result.FailedCitationRange),
                    result.SkippedMultiItem > 0 ||
                        result.SkippedMissingBibliography > 0 ||
                        result.FailedBibliographyMatch > 0 ||
                        result.FailedCitationRange > 0);
            });
        }

        internal void RemoveLinks()
        {
            RunWithActiveDocument(document =>
            {
                RemoveResult result = GetLinkerService().RemoveCitationLinks(document);
                ZoteroLinkerOptions options = GetOptions();
                int changed = GetLinkerService().RestoreCitationFormatting(
                    document,
                    options.CitationColor,
                    options.FontSize);
                ShowStatus(
                    "Remove links",
                    string.Format(
                        "Removed links {0}; bookmarks {1}",
                        result.LinksRemoved,
                        result.BookmarksRemoved),
                    string.Format(
                        "Reset fields {0}; repaired {1}; done",
                        result.Recolored,
                        changed),
                    false);
            });
        }

        internal void RestoreFormatting()
        {
            RunWithActiveDocument(document =>
            {
                ZoteroLinkerOptions options = GetOptions();
                int changed = GetLinkerService().RestoreCitationFormatting(
                    document,
                    options.CitationColor,
                    options.FontSize);
                ShowStatus(
                    "Repair formatting",
                    string.Format("Repaired citation fields {0}", changed),
                    "Done.",
                    false);
            });
        }

        internal void ShowOptions()
        {
            try
            {
                ZoteroLinkerOptions options = GetOptions();
                using (ZoteroLinkerOptionsForm form = new ZoteroLinkerOptionsForm(options))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    options.ColorHex = form.ColorHex;
                    options.FontSize = form.FontSize;
                    options.Save();
                    Globals.ThisAddIn.RefreshOptions();

                    int changed = 0;
                    Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
                    if (document != null)
                    {
                        changed = GetLinkerService().RestoreCitationFormatting(
                            document,
                            options.CitationColor,
                            options.FontSize);
                    }

                    ShowStatus(
                        "Options saved",
                        string.Format("Updated citation fields {0}", changed),
                        string.Format("Color {0}; font size {1:0.#} pt; done", options.ColorHex, options.FontSize),
                        false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Zotero Linker",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void RunWithActiveDocument(Action<Word.Document> action)
        {
            try
            {
                Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
                if (document == null)
                {
                    ShowStatus("No document", "No active Word document.", " ", true);
                    return;
                }

                action(document);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Zotero Linker",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ShowStatus(string title, string line1, string line2, bool hasIssue)
        {
            labelStatusTitle.Label = TruncateStatusText(title);
            labelStatusLine1.Label = TruncateStatusText(line1);
            labelStatusLine2.Label = TruncateStatusText(line2);
            labelStatusLine2.Visible = !hasIssue;
            buttonStatusIssue.Label = TruncateStatusText(line2);
            buttonStatusIssue.Visible = hasIssue;
        }

        private static string TruncateStatusText(string value)
        {
            string text = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
            if (string.IsNullOrEmpty(text))
            {
                return " ";
            }

            const int maxLength = 58;
            return text.Length <= maxLength ? text : text.Substring(0, maxLength - 3) + "...";
        }

        private static ZoteroLinkerService GetLinkerService()
        {
            return Globals.ThisAddIn.LinkerService ?? new ZoteroLinkerService();
        }

        private static ZoteroLinkerOptions GetOptions()
        {
            return Globals.ThisAddIn.CurrentOptions ?? ZoteroLinkerOptions.Load();
        }

        private static void SetLargeControlSize(object control)
        {
            if (control == null)
            {
                return;
            }

            System.Reflection.PropertyInfo property = control.GetType().GetProperty("ControlSize");
            if (property == null || !property.PropertyType.IsEnum)
            {
                return;
            }

            object largeSize = Enum.Parse(property.PropertyType, "RibbonControlSizeLarge");
            property.SetValue(control, largeSize, null);
        }
    }
}

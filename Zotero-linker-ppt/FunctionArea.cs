using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;

namespace Zotero_linker_ppt
{
    public partial class FunctionArea
    {
        private void FunctionArea_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonInsertCitation_Click(object sender, RibbonControlEventArgs e)
        {
            RunZoteroCommand("addEditCitation", "Insert citation");
        }

        private void buttonAddEditBibliography_Click(object sender, RibbonControlEventArgs e)
        {
            RunZoteroCommand("addEditBibliography", "Add bibliography");
        }

        private void buttonDocumentPreferences_Click(object sender, RibbonControlEventArgs e)
        {
            RunZoteroCommand("setDocPrefs", "Document preferences");
        }

        private void buttonRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            RunZoteroCommand("refresh", "Refresh");
        }

        internal void RunZoteroCommand(string command, string title)
        {
            try
            {
                PowerPointZoteroDocument document = new PowerPointZoteroDocument(Globals.ThisAddIn.Application);
                GetIntegrationClient().Execute(command, document);
                ShowStatus(
                    title,
                    "Zotero command completed.",
                    "Done.",
                    false);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        private void ShowError(string message)
        {
            ShowStatus("Zotero Linker PPT", message, " ", true);
            MessageBox.Show(
                message,
                "Zotero Linker PPT",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
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

        private static ZoteroHttpIntegrationClient GetIntegrationClient()
        {
            return Globals.ThisAddIn.IntegrationClient ?? new ZoteroHttpIntegrationClient();
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
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
            SetLargeControlSize(buttonLinkDoi);
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

        private void buttonLinkDoi_Click(object sender, RibbonControlEventArgs e)
        {
            LinkDois();
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
            RunWithActiveDocument("链接引文", document =>
            {
                LinkResult result = GetLinkerService().LinkCitations(document);

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
            RunWithActiveDocument("移除链接", document =>
            {
                RemoveResult result = GetLinkerService().RemoveCitationLinks(document);
                ShowStatus(
                    "Remove links",
                    string.Format(
                        "Removed links {0}; bookmarks {1}",
                        result.LinksRemoved,
                        result.BookmarksRemoved),
                    "Original colors, underlines and sizes preserved",
                    false);
            });
        }

        internal void LinkDois()
        {
            RunWithActiveDocument("设置 DOI 超链接", document =>
            {
                DoiLinkResult result = GetLinkerService().LinkBibliographyDois(document);
                ShowStatus(
                    "Link DOI",
                    result.BibliographyFound
                        ? string.Format("DOI found {0}; linked {1}", result.Found, result.Linked)
                        : "No Zotero bibliography found",
                    string.Format("Existing links skipped {0}", result.SkippedExisting),
                    !result.BibliographyFound);
            });
        }

        internal void RestoreFormatting()
        {
            RunWithActiveDocument("修复格式", document =>
            {
                ZoteroLinkerOptions options = GetOptions();
                int changed = GetLinkerService().RestoreCitationFormatting(
                    document,
                    options.CitationColor);
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
                    form.ApplyColorRequested += delegate
                    {
                        ApplyOptionColor(form, options);
                    };
                    form.ApplyFontSizeRequested += delegate
                    {
                        ApplyOptionFontSize(form, options);
                    };
                    form.ShowDialog();
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

        private void ApplyOptionColor(ZoteroLinkerOptionsForm form, ZoteroLinkerOptions options)
        {
            try
            {
                options.ColorHex = form.ColorHex;
                options.Save();
                Globals.ThisAddIn.RefreshOptions();

                Word.Document document = GetEditableActiveDocument();
                int changed = document == null
                    ? 0
                    : GetLinkerService().RestoreCitationFormatting(document, options.CitationColor);
                if (document != null)
                {
                    RefreshDocumentScreen(document);
                }
                ShowStatus(
                    "Color applied",
                    document == null
                        ? "Color option updated"
                        : string.Format("Updated citation ranges {0}", changed),
                    "Existing citation sizes unchanged",
                    false);
            }
            catch (Exception ex)
            {
                ShowOperationError("应用引文颜色", ex);
            }
        }

        private void ApplyOptionFontSize(ZoteroLinkerOptionsForm form, ZoteroLinkerOptions options)
        {
            try
            {
                options.FontSize = form.FontSize;
                options.Save();
                Globals.ThisAddIn.RefreshOptions();

                Word.Document document = GetEditableActiveDocument();
                if (document == null)
                {
                    ShowStatus("Font size saved", "No active Word document", "Existing citation sizes unchanged", false);
                    return;
                }

                DialogResult applyResult = MessageBox.Show(
                    string.Format(
                        "Apply {0:0.#} pt to all Zotero citations in the active document?",
                        options.FontSize),
                    "Zotero Linker",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (applyResult != DialogResult.Yes)
                {
                    ShowStatus("Font size saved", "Existing citation sizes were not changed", "Cancelled.", false);
                    return;
                }

                int changed = GetLinkerService().ApplyCitationFontSize(document, options.FontSize);
                RefreshDocumentScreen(document);
                ShowStatus(
                    "Font size applied",
                    string.Format("Updated citation ranges {0}", changed),
                    "Done.",
                    false);
            }
            catch (Exception ex)
            {
                ShowOperationError("应用引文字号", ex);
            }
        }

        private static void ShowOperationError(string operationName, Exception ex)
        {
            MessageBox.Show(
                BuildOperationErrorMessage(operationName, ex),
                "Zotero Linker",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static string BuildOperationErrorMessage(string operationName, Exception ex)
        {
            string operation = string.IsNullOrWhiteSpace(operationName) ? "当前操作" : operationName;
            COMException comException = ex as COMException;
            if (comException != null)
            {
                return string.Format(
                    "{0}失败。Word 拒绝了对当前文档的修改（HRESULT 0x{1:X8}）。\r\n\r\n请先点击 Word 顶部的“启用编辑”，并关闭只读或文档保护后重试。",
                    operation,
                    unchecked((uint)comException.HResult));
            }

            return string.Format("{0}失败：{1}", operation, ex.Message);
        }

        private static Word.Document GetEditableActiveDocument()
        {
            Word.Application application = Globals.ThisAddIn.Application;
            if (application == null)
            {
                return null;
            }

            // 受保护视图中的文件没有可写的 ActiveDocument，直接访问时 Word 通常只返回无上下文的 E_FAIL。
            if (application.ActiveProtectedViewWindow != null)
            {
                throw new InvalidOperationException("当前文档处于受保护视图，请先点击“启用编辑”。");
            }

            if (application.Documents.Count == 0)
            {
                return null;
            }

            Word.Document document = application.ActiveDocument;
            if (document == null)
            {
                return null;
            }

            if (document.ReadOnly)
            {
                throw new InvalidOperationException("当前文档为只读状态，请启用编辑后重试。");
            }

            if (document.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                throw new InvalidOperationException("当前文档已启用编辑保护，请先停止保护后重试。");
            }

            return document;
        }

        private static void RefreshDocumentScreen(Word.Document document)
        {
            try
            {
                document.Application.ScreenRefresh();
            }
            catch
            {
            }
        }

        private void RunWithActiveDocument(string operationName, Action<Word.Document> action)
        {
            try
            {
                Word.Document document = GetEditableActiveDocument();
                if (document == null)
                {
                    ShowStatus("No document", "No active Word document.", " ", true);
                    return;
                }

                action(document);
            }
            catch (Exception ex)
            {
                ShowOperationError(operationName, ex);
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

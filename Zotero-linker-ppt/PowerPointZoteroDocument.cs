using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Zotero_linker_ppt
{
    internal sealed class PowerPointZoteroDocument
    {
        private const string MetadataShapeName = "ZoteroLinkerMetadata";
        private const string FieldIdTagName = "ZoteroFieldID";
        private const string CitationShapePrefix = "ZoteroCitation_";
        private const string BibliographyShapeName = "ZoteroBibliography";
        private const string ReferenceSlideTitle = "References";
        private const string DefaultPlaceholderText = "{Citation}";

        private readonly PowerPoint.Application application;
        private readonly JavaScriptSerializer serializer;
        private PowerPointZoteroState state;

        internal PowerPointZoteroDocument(PowerPoint.Application application)
        {
            if (application == null || application.ActivePresentation == null)
            {
                throw new InvalidOperationException("No active PowerPoint presentation.");
            }

            this.application = application;
            serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };
            state = LoadState();
            EnsureStateDefaults();
            if (string.IsNullOrWhiteSpace(state.DocumentId))
            {
                state.DocumentId = Guid.NewGuid().ToString("N");
                SaveState();
            }
        }

        internal string CurrentIntegrationCommand { get; set; }

        internal string DocumentId
        {
            get { return state.DocumentId; }
        }

        internal object Execute(string commandName, object[] rawArguments)
        {
            string method = LastCommandSegment(commandName);
            object[] args = StripDocumentIdArgument(rawArguments ?? new object[0]);

            switch (commandName)
            {
                case "Application.getActiveDocument":
                case "Application.getDocument":
                    return GetActiveDocument();
                case "Document.displayAlert":
                    return DisplayAlert(args);
                case "Document.activate":
                    return null;
                case "Document.canInsertField":
                    return true;
                case "Document.setDocumentData":
                    SetDocumentData(ToStringArg(args, 0));
                    return null;
                case "Document.getDocumentData":
                    return state.DocumentData ?? string.Empty;
                case "Document.cursorInField":
                    return CursorInField();
                case "Document.insertField":
                    return InsertField();
                case "Document.insertText":
                    InsertText(ToStringArg(args, 0));
                    return null;
                case "Document.getFields":
                    return GetFields();
                case "Document.convert":
                case "Document.convertPlaceholdersToFields":
                case "Document.setBibliographyStyle":
                case "Document.cleanup":
                case "Document.complete":
                    return null;
                case "Field.delete":
                    DeleteField(ToStringArg(args, 0));
                    return null;
                case "Field.select":
                    SelectField(ToStringArg(args, 0));
                    return null;
                case "Field.removeCode":
                    RemoveFieldCode(ToStringArg(args, 0));
                    return null;
                case "Field.setText":
                    SetFieldText(ToStringArg(args, 0), ToStringArg(args, 1));
                    return null;
                case "Field.getText":
                    return GetFieldText(ToStringArg(args, 0));
                case "Field.setCode":
                    SetFieldCode(ToStringArg(args, 0), ToStringArg(args, 1));
                    return null;
                case "Field.getCode":
                    return GetFieldCode(ToStringArg(args, 0));
                default:
                    throw new InvalidOperationException("Unsupported Zotero command: " + commandName + " (" + method + ")");
            }
        }

        private Dictionary<string, object> GetActiveDocument()
        {
            return new Dictionary<string, object>
            {
                { "documentID", state.DocumentId },
                { "outputFormat", "html" },
                { "supportedNotes", new object[0] },
                { "supportsImportExport", false }
            };
        }

        private int DisplayAlert(object[] args)
        {
            string text = ToStringArg(args, 0);
            int buttons = ToIntArg(args, 2);
            MessageBoxButtons messageBoxButtons = MessageBoxButtons.OK;
            if (buttons == 1)
            {
                messageBoxButtons = MessageBoxButtons.OKCancel;
            }
            else if (buttons == 2)
            {
                messageBoxButtons = MessageBoxButtons.YesNo;
            }
            else if (buttons == 3)
            {
                messageBoxButtons = MessageBoxButtons.YesNoCancel;
            }

            DialogResult result = MessageBox.Show(text, "Zotero", messageBoxButtons, MessageBoxIcon.Information);
            if (buttons == 2)
            {
                return result == DialogResult.Yes ? 1 : 0;
            }
            if (buttons == 3)
            {
                if (result == DialogResult.Yes)
                {
                    return 2;
                }
                return result == DialogResult.No ? 1 : 0;
            }

            return result == DialogResult.OK ? 1 : 0;
        }

        private void SetDocumentData(string documentData)
        {
            state.DocumentData = documentData ?? string.Empty;
            SaveState();
        }

        private object CursorInField()
        {
            PowerPoint.Shape shape = GetSelectedFieldShape();
            if (shape == null)
            {
                return null;
            }

            return BuildFieldObject(GetFieldId(shape), shape);
        }

        private object InsertField()
        {
            string fieldId = NewFieldId();
            PowerPoint.Shape shape = IsBibliographyCommand()
                ? EnsureBibliographyShape(fieldId)
                : CreateCitationShape(fieldId);
            SetFieldId(shape, fieldId);
            state.FieldCodes[fieldId] = string.Empty;
            SaveState();
            return BuildFieldObject(fieldId, shape);
        }

        private void InsertText(string html)
        {
            PowerPoint.Slide slide = GetCurrentSlide();
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                48,
                48,
                520,
                120);
            shape.TextFrame.TextRange.Text = HtmlToText(html);
            shape.TextFrame.WordWrap = MsoTriState.msoTrue;
        }

        private List<Dictionary<string, object>> GetFields()
        {
            List<Dictionary<string, object>> fields = new List<Dictionary<string, object>>();
            foreach (PowerPoint.Slide slide in application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string fieldId = GetFieldId(shape);
                    if (!string.IsNullOrWhiteSpace(fieldId))
                    {
                        fields.Add(BuildFieldObject(fieldId, shape));
                    }
                }
            }

            return fields;
        }

        private void DeleteField(string fieldId)
        {
            PowerPoint.Shape shape = FindFieldShape(fieldId);
            if (shape != null)
            {
                shape.Delete();
            }
            state.FieldCodes.Remove(fieldId);
            SaveState();
        }

        private void SelectField(string fieldId)
        {
            PowerPoint.Shape shape = FindFieldShape(fieldId);
            if (shape != null)
            {
                shape.Select(MsoTriState.msoTrue);
            }
        }

        private void RemoveFieldCode(string fieldId)
        {
            PowerPoint.Shape shape = FindFieldShape(fieldId);
            if (shape != null)
            {
                DeleteTag(shape, FieldIdTagName);
            }
            state.FieldCodes.Remove(fieldId);
            SaveState();
        }

        private void SetFieldText(string fieldId, string htmlText)
        {
            PowerPoint.Shape shape = FindFieldShape(fieldId);
            if (shape == null)
            {
                return;
            }

            string text = HtmlToText(htmlText);
            shape.TextFrame.TextRange.Text = string.IsNullOrWhiteSpace(text) ? DefaultPlaceholderText : text;
            shape.TextFrame.WordWrap = MsoTriState.msoTrue;
        }

        private string GetFieldText(string fieldId)
        {
            PowerPoint.Shape shape = FindFieldShape(fieldId);
            if (shape == null || shape.HasTextFrame != MsoTriState.msoTrue)
            {
                return string.Empty;
            }

            return shape.TextFrame.TextRange.Text ?? string.Empty;
        }

        private void SetFieldCode(string fieldId, string code)
        {
            state.FieldCodes[fieldId] = code ?? string.Empty;
            SaveState();
        }

        private string GetFieldCode(string fieldId)
        {
            string code;
            return state.FieldCodes.TryGetValue(fieldId, out code) ? code : string.Empty;
        }

        private Dictionary<string, object> BuildFieldObject(string fieldId, PowerPoint.Shape shape)
        {
            return new Dictionary<string, object>
            {
                { "id", fieldId },
                { "code", GetFieldCode(fieldId) },
                { "text", GetShapeText(shape) },
                { "noteIndex", 0 }
            };
        }

        private PowerPoint.Shape CreateCitationShape(string fieldId)
        {
            PowerPoint.Slide slide = GetCurrentSlide();
            float left = 48;
            float top = 48;
            PowerPoint.Shape selected = GetSelectedShape();
            if (selected != null)
            {
                left = selected.Left + selected.Width + 8;
                top = selected.Top;
            }

            PowerPoint.Shape shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                96,
                28);
            shape.Name = CitationShapePrefix + fieldId;
            shape.TextFrame.TextRange.Text = DefaultPlaceholderText;
            shape.TextFrame.TextRange.Font.Size = 12;
            shape.TextFrame.WordWrap = MsoTriState.msoTrue;
            SetFieldId(shape, fieldId);
            shape.Select(MsoTriState.msoTrue);
            return shape;
        }

        private PowerPoint.Shape EnsureBibliographyShape(string fieldId)
        {
            foreach (PowerPoint.Slide slide in application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (string.Equals(shape.Name, BibliographyShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        SetFieldId(shape, fieldId);
                        return shape;
                    }
                }
            }

            PowerPoint.Slide referenceSlide = EnsureReferenceSlide();
            PowerPoint.Shape body = referenceSlide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                40,
                72,
                640,
                420);
            body.Name = BibliographyShapeName;
            body.TextFrame.TextRange.Text = "{Bibliography}";
            body.TextFrame.TextRange.Font.Size = 11;
            body.TextFrame.WordWrap = MsoTriState.msoTrue;
            SetFieldId(body, fieldId);
            return body;
        }

        private PowerPoint.Slide EnsureReferenceSlide()
        {
            foreach (PowerPoint.Slide slide in application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (string.Equals(shape.Name, BibliographyShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return slide;
                    }
                }
            }

            PowerPoint.Presentation presentation = application.ActivePresentation;
            PowerPoint.Slide newSlide = presentation.Slides.Add(
                presentation.Slides.Count + 1,
                PowerPoint.PpSlideLayout.ppLayoutBlank);
            PowerPoint.Shape title = newSlide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                40,
                24,
                640,
                36);
            title.TextFrame.TextRange.Text = ReferenceSlideTitle;
            title.TextFrame.TextRange.Font.Size = 24;
            title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            return newSlide;
        }

        private PowerPoint.Shape GetSelectedFieldShape()
        {
            PowerPoint.Shape selected = GetSelectedShape();
            if (selected != null && !string.IsNullOrWhiteSpace(GetFieldId(selected)))
            {
                return selected;
            }

            return null;
        }

        private PowerPoint.Shape GetSelectedShape()
        {
            try
            {
                PowerPoint.Selection selection = application.ActiveWindow.Selection;
                if (selection != null &&
                    selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count > 0)
                {
                    return selection.ShapeRange[1];
                }
            }
            catch
            {
            }

            return null;
        }

        private PowerPoint.Slide GetCurrentSlide()
        {
            if (application.ActiveWindow != null &&
                application.ActiveWindow.View != null &&
                application.ActiveWindow.View.Slide != null)
            {
                return application.ActiveWindow.View.Slide as PowerPoint.Slide;
            }

            if (application.ActivePresentation.Slides.Count == 0)
            {
                return application.ActivePresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            }

            return application.ActivePresentation.Slides[1];
        }

        private PowerPoint.Shape FindFieldShape(string fieldId)
        {
            foreach (PowerPoint.Slide slide in application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (string.Equals(GetFieldId(shape), fieldId, StringComparison.OrdinalIgnoreCase))
                    {
                        return shape;
                    }
                }
            }

            return null;
        }

        private PowerPointZoteroState LoadState()
        {
            PowerPoint.Shape metadataShape = FindMetadataShape();
            if (metadataShape == null || metadataShape.HasTextFrame != MsoTriState.msoTrue)
            {
                return new PowerPointZoteroState();
            }

            string text = metadataShape.TextFrame.TextRange.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return new PowerPointZoteroState();
            }

            try
            {
                PowerPointZoteroState loaded = serializer.Deserialize<PowerPointZoteroState>(text);
                return loaded ?? new PowerPointZoteroState();
            }
            catch
            {
                return new PowerPointZoteroState();
            }
        }

        private void SaveState()
        {
            EnsureStateDefaults();
            PowerPoint.Shape metadataShape = EnsureMetadataShape();
            metadataShape.TextFrame.TextRange.Text = serializer.Serialize(state);
        }

        private void EnsureStateDefaults()
        {
            if (state == null)
            {
                state = new PowerPointZoteroState();
            }
            if (state.FieldCodes == null)
            {
                state.FieldCodes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private PowerPoint.Shape EnsureMetadataShape()
        {
            PowerPoint.Shape shape = FindMetadataShape();
            if (shape != null)
            {
                return shape;
            }

            PowerPoint.Slide slide = application.ActivePresentation.Slides.Count > 0
                ? application.ActivePresentation.Slides[1]
                : application.ActivePresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                -5000,
                -5000,
                10,
                10);
            shape.Name = MetadataShapeName;
            shape.Visible = MsoTriState.msoFalse;
            return shape;
        }

        private PowerPoint.Shape FindMetadataShape()
        {
            foreach (PowerPoint.Slide slide in application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (string.Equals(shape.Name, MetadataShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return shape;
                    }
                }
            }

            return null;
        }

        private bool IsBibliographyCommand()
        {
            return !string.IsNullOrWhiteSpace(CurrentIntegrationCommand) &&
                CurrentIntegrationCommand.IndexOf("Bibliography", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private object[] StripDocumentIdArgument(object[] args)
        {
            if (args.Length == 0)
            {
                return args;
            }

            string first = Convert.ToString(args[0], System.Globalization.CultureInfo.InvariantCulture);
            if (string.Equals(first, state.DocumentId, StringComparison.OrdinalIgnoreCase))
            {
                object[] result = new object[args.Length - 1];
                Array.Copy(args, 1, result, 0, result.Length);
                return result;
            }

            return args;
        }

        private static string NewFieldId()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 12);
        }

        private static string LastCommandSegment(string commandName)
        {
            int dot = (commandName ?? string.Empty).LastIndexOf('.');
            return dot >= 0 ? commandName.Substring(dot + 1) : commandName;
        }

        private static string GetFieldId(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return string.Empty;
            }

            try
            {
                return shape.Tags[FieldIdTagName] ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void SetFieldId(PowerPoint.Shape shape, string fieldId)
        {
            DeleteTag(shape, FieldIdTagName);
            shape.Tags.Add(FieldIdTagName, fieldId);
        }

        private static void DeleteTag(PowerPoint.Shape shape, string tagName)
        {
            try
            {
                shape.Tags.Delete(tagName);
            }
            catch
            {
            }
        }

        private static string GetShapeText(PowerPoint.Shape shape)
        {
            if (shape == null || shape.HasTextFrame != MsoTriState.msoTrue)
            {
                return string.Empty;
            }

            return shape.TextFrame.TextRange.Text ?? string.Empty;
        }

        private static string HtmlToText(string html)
        {
            string cslText = CslBibliographyHtmlToText(html);
            if (!string.IsNullOrWhiteSpace(cslText))
            {
                return cslText;
            }

            string text = html ?? string.Empty;
            text = Regex.Replace(text, "</(p|div|li)>", "\r", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<br\\s*/?>", "\r", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<[^>]+>", string.Empty);
            text = WebUtility.HtmlDecode(text);
            text = Regex.Replace(text, "[ \\t]+", " ");
            text = Regex.Replace(text, "\\s*\\r\\s*", "\r");
            return text.Trim();
        }

        private static string CslBibliographyHtmlToText(string html)
        {
            string source = html ?? string.Empty;
            if (source.IndexOf("csl-entry", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return string.Empty;
            }

            List<string> entries = new List<string>();
            MatchCollection splitEntryMatches = Regex.Matches(
                source,
                "<div[^>]*class\\s*=\\s*[\"'][^\"']*csl-entry[^\"']*[\"'][^>]*>\\s*" +
                "<div[^>]*class\\s*=\\s*[\"'][^\"']*csl-left-margin[^\"']*[\"'][^>]*>(.*?)</div>\\s*" +
                "<div[^>]*class\\s*=\\s*[\"'][^\"']*csl-right-inline[^\"']*[\"'][^>]*>(.*?)</div>\\s*</div>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match match in splitEntryMatches)
            {
                string number = HtmlFragmentToPlainText(match.Groups[1].Value).Trim();
                string entry = HtmlFragmentToPlainText(match.Groups[2].Value).Trim();
                entries.Add(NormalizeLine(number + " " + entry));
            }

            if (entries.Count > 0)
            {
                return string.Join("\r", entries.ToArray());
            }

            MatchCollection simpleEntryMatches = Regex.Matches(
                source,
                "<div[^>]*class\\s*=\\s*[\"'][^\"']*csl-entry[^\"']*[\"'][^>]*>(.*?)</div>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match match in simpleEntryMatches)
            {
                string entry = NormalizeLine(HtmlFragmentToPlainText(match.Groups[1].Value));
                if (!string.IsNullOrWhiteSpace(entry))
                {
                    entries.Add(entry);
                }
            }

            return entries.Count > 0 ? string.Join("\r", entries.ToArray()) : string.Empty;
        }

        private static string HtmlFragmentToPlainText(string html)
        {
            string text = html ?? string.Empty;
            text = Regex.Replace(text, "<br\\s*/?>", " ", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<[^>]+>", string.Empty);
            return WebUtility.HtmlDecode(text);
        }

        private static string NormalizeLine(string value)
        {
            return Regex.Replace(value ?? string.Empty, "\\s+", " ").Trim();
        }

        private static string ToStringArg(object[] args, int index)
        {
            if (args == null || index < 0 || index >= args.Length || args[index] == null)
            {
                return string.Empty;
            }

            return Convert.ToString(args[index], System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static int ToIntArg(object[] args, int index)
        {
            if (args == null || index < 0 || index >= args.Length || args[index] == null)
            {
                return 0;
            }

            int value;
            return int.TryParse(Convert.ToString(args[index], System.Globalization.CultureInfo.InvariantCulture), out value) ? value : 0;
        }
    }

    public sealed class PowerPointZoteroState
    {
        public PowerPointZoteroState()
        {
            FieldCodes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public string DocumentId { get; set; }
        public string DocumentData { get; set; }
        public Dictionary<string, string> FieldCodes { get; set; }
    }
}

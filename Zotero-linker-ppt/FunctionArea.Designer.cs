namespace Zotero_linker_ppt
{
    public partial class FunctionArea : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public FunctionArea()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupCitation = this.Factory.CreateRibbonGroup();
            this.buttonInsertCitation = this.Factory.CreateRibbonButton();
            this.buttonAddEditBibliography = this.Factory.CreateRibbonButton();
            this.buttonDocumentPreferences = this.Factory.CreateRibbonButton();
            this.buttonRefresh = this.Factory.CreateRibbonButton();
            this.groupStatus = this.Factory.CreateRibbonGroup();
            this.boxStatus = this.Factory.CreateRibbonBox();
            this.labelStatusTitle = this.Factory.CreateRibbonLabel();
            this.labelStatusLine1 = this.Factory.CreateRibbonLabel();
            this.labelStatusLine2 = this.Factory.CreateRibbonLabel();
            this.buttonStatusIssue = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupCitation.SuspendLayout();
            this.groupStatus.SuspendLayout();
            this.boxStatus.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.Groups.Add(this.groupCitation);
            this.tab1.Groups.Add(this.groupStatus);
            this.tab1.Label = "Zotero Linker";
            this.tab1.Name = "tab1";
            //
            // groupCitation
            //
            this.groupCitation.Items.Add(this.buttonInsertCitation);
            this.groupCitation.Items.Add(this.buttonAddEditBibliography);
            this.groupCitation.Items.Add(this.buttonDocumentPreferences);
            this.groupCitation.Items.Add(this.buttonRefresh);
            this.groupCitation.Label = "Zotero";
            this.groupCitation.Name = "groupCitation";
            //
            // buttonInsertCitation
            //
            this.buttonInsertCitation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonInsertCitation.Label = "Insert Citation";
            this.buttonInsertCitation.Name = "buttonInsertCitation";
            this.buttonInsertCitation.OfficeImageId = "HyperlinkInsert";
            this.buttonInsertCitation.ShowImage = true;
            this.buttonInsertCitation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonInsertCitation_Click);
            //
            // buttonAddEditBibliography
            //
            this.buttonAddEditBibliography.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddEditBibliography.Label = "Add Bibliography";
            this.buttonAddEditBibliography.Name = "buttonAddEditBibliography";
            this.buttonAddEditBibliography.OfficeImageId = "BibliographyInsert";
            this.buttonAddEditBibliography.ShowImage = true;
            this.buttonAddEditBibliography.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddEditBibliography_Click);
            //
            // buttonDocumentPreferences
            //
            this.buttonDocumentPreferences.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDocumentPreferences.Label = "Document Preferences";
            this.buttonDocumentPreferences.Name = "buttonDocumentPreferences";
            this.buttonDocumentPreferences.OfficeImageId = "PropertySheet";
            this.buttonDocumentPreferences.ShowImage = true;
            this.buttonDocumentPreferences.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDocumentPreferences_Click);
            //
            // buttonRefresh
            //
            this.buttonRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRefresh.Label = "Refresh";
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.OfficeImageId = "Refresh";
            this.buttonRefresh.ShowImage = true;
            this.buttonRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRefresh_Click);
            //
            // groupStatus
            //
            this.groupStatus.Items.Add(this.boxStatus);
            this.groupStatus.Label = "Status";
            this.groupStatus.Name = "groupStatus";
            //
            // boxStatus
            //
            this.boxStatus.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.boxStatus.Items.Add(this.labelStatusTitle);
            this.boxStatus.Items.Add(this.labelStatusLine1);
            this.boxStatus.Items.Add(this.labelStatusLine2);
            this.boxStatus.Items.Add(this.buttonStatusIssue);
            this.boxStatus.Name = "boxStatus";
            //
            // labelStatusTitle
            //
            this.labelStatusTitle.Label = "Ready";
            this.labelStatusTitle.Name = "labelStatusTitle";
            //
            // labelStatusLine1
            //
            this.labelStatusLine1.Label = "Insert a numbered citation into the active slide.";
            this.labelStatusLine1.Name = "labelStatusLine1";
            //
            // labelStatusLine2
            //
            this.labelStatusLine2.Label = " ";
            this.labelStatusLine2.Name = "labelStatusLine2";
            //
            // buttonStatusIssue
            //
            this.buttonStatusIssue.Label = " ";
            this.buttonStatusIssue.Name = "buttonStatusIssue";
            this.buttonStatusIssue.OfficeImageId = "ReviewRejectChange";
            this.buttonStatusIssue.ShowImage = true;
            this.buttonStatusIssue.Visible = false;
            //
            // FunctionArea
            //
            this.Name = "FunctionArea";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FunctionArea_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupCitation.ResumeLayout(false);
            this.groupCitation.PerformLayout();
            this.groupStatus.ResumeLayout(false);
            this.groupStatus.PerformLayout();
            this.boxStatus.ResumeLayout(false);
            this.boxStatus.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCitation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonInsertCitation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddEditBibliography;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDocumentPreferences;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusLine1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusLine2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStatusIssue;
    }
}

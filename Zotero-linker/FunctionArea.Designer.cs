namespace Zotero_linker
{
    public partial class FunctionArea : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FunctionArea()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            ConfigureRuntimeLayout();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupCitationLinks = this.Factory.CreateRibbonGroup();
            this.buttonLinkCitations = this.Factory.CreateRibbonButton();
            this.buttonRemoveLinks = this.Factory.CreateRibbonButton();
            this.buttonRestoreFormatting = this.Factory.CreateRibbonButton();
            this.buttonOptions = this.Factory.CreateRibbonButton();
            this.groupStatus = this.Factory.CreateRibbonGroup();
            this.boxStatus = this.Factory.CreateRibbonBox();
            this.labelStatusTitle = this.Factory.CreateRibbonLabel();
            this.labelStatusLine1 = this.Factory.CreateRibbonLabel();
            this.labelStatusLine2 = this.Factory.CreateRibbonLabel();
            this.buttonStatusIssue = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupCitationLinks.SuspendLayout();
            this.groupStatus.SuspendLayout();
            this.boxStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupCitationLinks);
            this.tab1.Groups.Add(this.groupStatus);
            this.tab1.Label = "Zotero Linker";
            this.tab1.Name = "tab1";
            // 
            // groupCitationLinks
            // 
            this.groupCitationLinks.Items.Add(this.buttonLinkCitations);
            this.groupCitationLinks.Items.Add(this.buttonRemoveLinks);
            this.groupCitationLinks.Items.Add(this.buttonRestoreFormatting);
            this.groupCitationLinks.Items.Add(this.buttonOptions);
            this.groupCitationLinks.Label = "Citation Links";
            this.groupCitationLinks.Name = "groupCitationLinks";
            // 
            // buttonLinkCitations
            // 
            this.buttonLinkCitations.Image = global::Zotero_linker.Properties.Resources.PixPin_2026_06_16_15_04_15;
            this.buttonLinkCitations.Label = "Link Citations";
            this.buttonLinkCitations.Name = "buttonLinkCitations";
            this.buttonLinkCitations.OfficeImageId = "HyperlinkInsert";
            this.buttonLinkCitations.ShowImage = true;
            this.buttonLinkCitations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLinkCitations_Click);
            // 
            // buttonRemoveLinks
            // 
            this.buttonRemoveLinks.Image = global::Zotero_linker.Properties.Resources.PixPin_2026_06_16_15_04_53;
            this.buttonRemoveLinks.Label = "Remove Links";
            this.buttonRemoveLinks.Name = "buttonRemoveLinks";
            this.buttonRemoveLinks.OfficeImageId = "HyperlinkRemove";
            this.buttonRemoveLinks.ShowImage = true;
            this.buttonRemoveLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRemoveLinks_Click);
            // 
            // buttonRestoreFormatting
            // 
            this.buttonRestoreFormatting.Image = global::Zotero_linker.Properties.Resources.PixPin_2026_06_16_15_05_05;
            this.buttonRestoreFormatting.Label = "Repair Formatting";
            this.buttonRestoreFormatting.Name = "buttonRestoreFormatting";
            this.buttonRestoreFormatting.OfficeImageId = "FontColorPicker";
            this.buttonRestoreFormatting.ShowImage = true;
            this.buttonRestoreFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRestoreFormatting_Click);
            // 
            // buttonOptions
            // 
            this.buttonOptions.Image = global::Zotero_linker.Properties.Resources.PixPin_2026_06_16_15_05_24;
            this.buttonOptions.Label = "Options";
            this.buttonOptions.Name = "buttonOptions";
            this.buttonOptions.OfficeImageId = "PropertySheet";
            this.buttonOptions.ShowImage = true;
            this.buttonOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonOptions_Click);
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
            this.labelStatusLine1.Label = "Run a command to show results here.";
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
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FunctionArea_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupCitationLinks.ResumeLayout(false);
            this.groupCitationLinks.PerformLayout();
            this.groupStatus.ResumeLayout(false);
            this.groupStatus.PerformLayout();
            this.boxStatus.ResumeLayout(false);
            this.boxStatus.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCitationLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLinkCitations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRemoveLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRestoreFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusLine1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStatusLine2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStatusIssue;
    }
}

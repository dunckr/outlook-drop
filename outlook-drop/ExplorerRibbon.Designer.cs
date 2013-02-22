namespace outlook_drop
{
    partial class ExplorerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExplorerRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.uploadButton = this.Factory.CreateRibbonButton();
            this.shareButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.uploadButton);
            this.group1.Items.Add(this.shareButton);
            this.group1.Label = "Drop";
            this.group1.Name = "group1";
            // 
            // uploadButton
            // 
            this.uploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.uploadButton.Image = global::outlook_drop.Properties.Resources.upload;
            this.uploadButton.Label = "Upload";
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.ShowImage = true;
            // 
            // shareButton
            // 
            this.shareButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.shareButton.Image = global::outlook_drop.Properties.Resources.share;
            this.shareButton.Label = "Share";
            this.shareButton.Name = "shareButton";
            this.shareButton.ShowImage = true;
            // 
            // ExplorerRibbon
            // 
            this.Name = "ExplorerRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExplorerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton uploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton shareButton;
    }

    partial class ThisRibbonCollection
    {
        internal ExplorerRibbon ExplorerRibbon
        {
            get { return this.GetRibbon<ExplorerRibbon>(); }
        }
    }
}

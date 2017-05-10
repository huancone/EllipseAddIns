namespace EllipseStatisticsProfileExcelAddIn
{
    partial class RibbonEllipse : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonEllipse()
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
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpStatisticsProfile = this.Factory.CreateRibbonGroup();
            this.btnFormatProfile = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.btnExecuteProfile = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpStatisticsProfile.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpStatisticsProfile);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpStatisticsProfile
            // 
            this.grpStatisticsProfile.Items.Add(this.box1);
            this.grpStatisticsProfile.Items.Add(this.drpEnviroment);
            this.grpStatisticsProfile.Items.Add(this.btnExecuteProfile);
            this.grpStatisticsProfile.Label = "Statistics Profile";
            this.grpStatisticsProfile.Name = "grpStatisticsProfile";
            // 
            // btnFormatProfile
            // 
            this.btnFormatProfile.Label = "Format Profile Sheet";
            this.btnFormatProfile.Name = "btnFormatProfile";
            this.btnFormatProfile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatProfile_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env. ";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // btnExecuteProfile
            // 
            this.btnExecuteProfile.Label = "Execute Profile Sheet";
            this.btnExecuteProfile.Name = "btnExecuteProfile";
            this.btnExecuteProfile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteProfile_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormatProfile);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Tag = "ELLIPSE 8";
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpStatisticsProfile.ResumeLayout(false);
            this.grpStatisticsProfile.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStatisticsProfile;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecuteProfile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatProfile;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

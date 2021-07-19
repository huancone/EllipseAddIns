namespace EllipseMSE140DeleteExcelAddIn
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpEllipse = this.Factory.CreateRibbonGroup();
            this.btnEllipseFormat = this.Factory.CreateRibbonButton();
            this.drpEllipseEnv = this.Factory.CreateRibbonDropDown();
            this.mnuOptions = this.Factory.CreateRibbonMenu();
            this.btnLoadData = this.Factory.CreateRibbonButton();
            this.btnEllipseExecute = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEllipse);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.btnEllipseFormat);
            this.grpEllipse.Items.Add(this.drpEllipseEnv);
            this.grpEllipse.Items.Add(this.mnuOptions);
            this.grpEllipse.Label = "MSE140 Delete v 1.0.1";
            this.grpEllipse.Name = "grpEllipse";
            // 
            // btnEllipseFormat
            // 
            this.btnEllipseFormat.Label = "Format New Sheet";
            this.btnEllipseFormat.Name = "btnEllipseFormat";
            this.btnEllipseFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEllipseFormat_Click);
            // 
            // drpEllipseEnv
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Contingencia";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            ribbonDropDownItemImpl4.Label = "Test";
            this.drpEllipseEnv.Items.Add(ribbonDropDownItemImpl1);
            this.drpEllipseEnv.Items.Add(ribbonDropDownItemImpl2);
            this.drpEllipseEnv.Items.Add(ribbonDropDownItemImpl3);
            this.drpEllipseEnv.Items.Add(ribbonDropDownItemImpl4);
            this.drpEllipseEnv.Label = "Env.";
            this.drpEllipseEnv.Name = "drpEllipseEnv";
            // 
            // mnuOptions
            // 
            this.mnuOptions.Items.Add(this.btnLoadData);
            this.mnuOptions.Items.Add(this.btnEllipseExecute);
            this.mnuOptions.Label = "Execute Options";
            this.mnuOptions.Name = "mnuOptions";
            // 
            // btnLoadData
            // 
            this.btnLoadData.Label = "Retrieve Data";
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.ShowImage = true;
            this.btnLoadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadData_Click);
            // 
            // btnEllipseExecute
            // 
            this.btnEllipseExecute.Label = "Execute Loader";
            this.btnEllipseExecute.Name = "btnEllipseExecute";
            this.btnEllipseExecute.ShowImage = true;
            this.btnEllipseExecute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEllipseExecute_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEllipseFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEllipseEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEllipseExecute;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadData;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

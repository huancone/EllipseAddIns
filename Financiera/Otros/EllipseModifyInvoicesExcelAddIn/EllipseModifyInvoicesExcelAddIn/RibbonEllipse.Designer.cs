namespace EllipseModifyInvoicesExcelAddIn
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
            this.grpModifyInvoices = this.Factory.CreateRibbonGroup();
            this.btnModifyInvoicesFormat = this.Factory.CreateRibbonButton();
            this.drpModifyInvoicesEnv = this.Factory.CreateRibbonDropDown();
            this.mnuOptions = this.Factory.CreateRibbonMenu();
            this.btnModifyInvoicesLoad = this.Factory.CreateRibbonButton();
            this.btnModifyInvoicesExecute = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpModifyInvoices.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpModifyInvoices);
            this.tabEllipse.Label = "ELLIPSE 8 ";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpModifyInvoices
            // 
            this.grpModifyInvoices.Items.Add(this.btnModifyInvoicesFormat);
            this.grpModifyInvoices.Items.Add(this.drpModifyInvoicesEnv);
            this.grpModifyInvoices.Items.Add(this.mnuOptions);
            this.grpModifyInvoices.Label = "MSO261 Opc 1 v 1.0.1";
            this.grpModifyInvoices.Name = "grpModifyInvoices";
            // 
            // btnModifyInvoicesFormat
            // 
            this.btnModifyInvoicesFormat.Label = "Format New Sheet";
            this.btnModifyInvoicesFormat.Name = "btnModifyInvoicesFormat";
            this.btnModifyInvoicesFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyInvoicesFormat_Click);
            // 
            // drpModifyInvoicesEnv
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Contingencia";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            ribbonDropDownItemImpl4.Label = "Test";
            this.drpModifyInvoicesEnv.Items.Add(ribbonDropDownItemImpl1);
            this.drpModifyInvoicesEnv.Items.Add(ribbonDropDownItemImpl2);
            this.drpModifyInvoicesEnv.Items.Add(ribbonDropDownItemImpl3);
            this.drpModifyInvoicesEnv.Items.Add(ribbonDropDownItemImpl4);
            this.drpModifyInvoicesEnv.Label = "Env.";
            this.drpModifyInvoicesEnv.Name = "drpModifyInvoicesEnv";
            // 
            // mnuOptions
            // 
            this.mnuOptions.Items.Add(this.btnModifyInvoicesLoad);
            this.mnuOptions.Items.Add(this.btnModifyInvoicesExecute);
            this.mnuOptions.Label = "Execute Options";
            this.mnuOptions.Name = "mnuOptions";
            // 
            // btnModifyInvoicesLoad
            // 
            this.btnModifyInvoicesLoad.Label = "Load States Invoices";
            this.btnModifyInvoicesLoad.Name = "btnModifyInvoicesLoad";
            this.btnModifyInvoicesLoad.ShowImage = true;
            this.btnModifyInvoicesLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyInvoicesLoad_Click);
            // 
            // btnModifyInvoicesExecute
            // 
            this.btnModifyInvoicesExecute.Label = "Modify Invoices";
            this.btnModifyInvoicesExecute.Name = "btnModifyInvoicesExecute";
            this.btnModifyInvoicesExecute.ShowImage = true;
            this.btnModifyInvoicesExecute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyInvoicesExecute_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpModifyInvoices.ResumeLayout(false);
            this.grpModifyInvoices.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpModifyInvoices;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyInvoicesExecute;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyInvoicesFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpModifyInvoicesEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyInvoicesLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuOptions;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

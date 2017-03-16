using EllipseMSO685Opc3ModifyExcelAddIn;

namespace EllipseMSO685Opc3ModifyExcelAddIn
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
            var ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            var ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            var ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            var ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpSubAssetDepreciation = this.Factory.CreateRibbonGroup();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.drpEnv = this.Factory.CreateRibbonDropDown();
            this.btnExecute = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpSubAssetDepreciation.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpSubAssetDepreciation);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpSubAssetDepreciation
            // 
            this.grpSubAssetDepreciation.Items.Add(this.btnFormat);
            this.grpSubAssetDepreciation.Items.Add(this.drpEnv);
            this.grpSubAssetDepreciation.Items.Add(this.btnExecute);
            this.grpSubAssetDepreciation.Label = "MSO685 Opc 3 Modify v 1.0.1";
            this.grpSubAssetDepreciation.Name = "grpSubAssetDepreciation";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Format New Sheet";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // drpEnv
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Contingencia";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            ribbonDropDownItemImpl4.Label = "Test";
            this.drpEnv.Items.Add(ribbonDropDownItemImpl1);
            this.drpEnv.Items.Add(ribbonDropDownItemImpl2);
            this.drpEnv.Items.Add(ribbonDropDownItemImpl3);
            this.drpEnv.Items.Add(ribbonDropDownItemImpl4);
            this.drpEnv.Label = "Env.";
            this.drpEnv.Name = "drpEnv";
            // 
            // btnExecute
            // 
            this.btnExecute.Label = "Modify Sub-Asset Depreciation";
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecute_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpSubAssetDepreciation.ResumeLayout(false);
            this.grpSubAssetDepreciation.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSubAssetDepreciation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecute;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

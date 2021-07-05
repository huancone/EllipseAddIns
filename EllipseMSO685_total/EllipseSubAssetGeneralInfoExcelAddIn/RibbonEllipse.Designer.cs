namespace EllipseSubAssetGeneralInfoExcelAddIn
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
            this.grpSubAssetGeneralInfo = this.Factory.CreateRibbonGroup();
            this.btnSubAssetGeneralInfoFormat = this.Factory.CreateRibbonButton();
            this.drpSubAssetGeneralInfoEnv = this.Factory.CreateRibbonDropDown();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.btnMSO685op1VC = this.Factory.CreateRibbonButton();
            this.btn_MSO685OP1VL = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_op3vl = this.Factory.CreateRibbonButton();
            this.btnOp4vc = this.Factory.CreateRibbonButton();
            this.btnmso685op4vl = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpSubAssetGeneralInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpSubAssetGeneralInfo);
            this.tabEllipse.Label = "ELLIPSE 9";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpSubAssetGeneralInfo
            // 
            this.grpSubAssetGeneralInfo.Items.Add(this.btnSubAssetGeneralInfoFormat);
            this.grpSubAssetGeneralInfo.Items.Add(this.drpSubAssetGeneralInfoEnv);
            this.grpSubAssetGeneralInfo.Items.Add(this.menu1);
            this.grpSubAssetGeneralInfo.Label = "MSO685  E9";
            this.grpSubAssetGeneralInfo.Name = "grpSubAssetGeneralInfo";
            // 
            // btnSubAssetGeneralInfoFormat
            // 
            this.btnSubAssetGeneralInfoFormat.Label = "Format New Sheet";
            this.btnSubAssetGeneralInfoFormat.Name = "btnSubAssetGeneralInfoFormat";
            this.btnSubAssetGeneralInfoFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSubAssetGeneralInfoFormat_Click);
            // 
            // drpSubAssetGeneralInfoEnv
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Contingencia";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            ribbonDropDownItemImpl4.Label = "Test";
            this.drpSubAssetGeneralInfoEnv.Items.Add(ribbonDropDownItemImpl1);
            this.drpSubAssetGeneralInfoEnv.Items.Add(ribbonDropDownItemImpl2);
            this.drpSubAssetGeneralInfoEnv.Items.Add(ribbonDropDownItemImpl3);
            this.drpSubAssetGeneralInfoEnv.Items.Add(ribbonDropDownItemImpl4);
            this.drpSubAssetGeneralInfoEnv.Label = "Env.";
            this.drpSubAssetGeneralInfoEnv.Name = "drpSubAssetGeneralInfoEnv";
            // 
            // menu1
            // 
            this.menu1.Items.Add(this.btnMSO685op1VC);
            this.menu1.Items.Add(this.btn_MSO685OP1VL);
            this.menu1.Items.Add(this.button1);
            this.menu1.Items.Add(this.btn_op3vl);
            this.menu1.Items.Add(this.btnOp4vc);
            this.menu1.Items.Add(this.btnmso685op4vl);
            this.menu1.Label = "Options";
            this.menu1.Name = "menu1";
            // 
            // btnMSO685op1VC
            // 
            this.btnMSO685op1VC.Label = "1. Maintain General Information VC";
            this.btnMSO685op1VC.Name = "btnMSO685op1VC";
            this.btnMSO685op1VC.ShowImage = true;
            this.btnMSO685op1VC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // btn_MSO685OP1VL
            // 
            this.btn_MSO685OP1VL.Label = "1. Maintain General Information VL";
            this.btn_MSO685OP1VL.Name = "btn_MSO685OP1VL";
            this.btn_MSO685OP1VL.ShowImage = true;
            this.btn_MSO685OP1VL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MSO685OP1VL_Click);
            // 
            // button1
            // 
            this.button1.Label = "3. Maintain Sub-Asset Depreciation Details VC";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // btn_op3vl
            // 
            this.btn_op3vl.Label = "3. Maintain Sub-Asset Depreciation Details VL";
            this.btn_op3vl.Name = "btn_op3vl";
            this.btn_op3vl.ShowImage = true;
            this.btn_op3vl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_op3vl_Click);
            // 
            // btnOp4vc
            // 
            this.btnOp4vc.Label = "4. Maintain Sub-Asset Valuation Details VC";
            this.btnOp4vc.Name = "btnOp4vc";
            this.btnOp4vc.ShowImage = true;
            this.btnOp4vc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOp4vc_Click);
            // 
            // btnmso685op4vl
            // 
            this.btnmso685op4vl.Label = "4. Maintain Sub-Asset Valuation Details VL";
            this.btnmso685op4vl.Name = "btnmso685op4vl";
            this.btnmso685op4vl.ShowImage = true;
            this.btnmso685op4vl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnmso685op4vl_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpSubAssetGeneralInfo.ResumeLayout(false);
            this.grpSubAssetGeneralInfo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSubAssetGeneralInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubAssetGeneralInfoFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpSubAssetGeneralInfoEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMSO685op1VC;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_MSO685OP1VL;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_op3vl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOp4vc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnmso685op4vl;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

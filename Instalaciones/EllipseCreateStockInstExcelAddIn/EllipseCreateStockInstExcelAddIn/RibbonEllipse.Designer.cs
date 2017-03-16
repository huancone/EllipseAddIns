namespace EllipseCreateStockInstExcelAddIn
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
            this.grpCreateStockInst = this.Factory.CreateRibbonGroup();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.btnCreateStock = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpCreateStockInst.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpCreateStockInst);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpCreateStockInst
            // 
            this.grpCreateStockInst.Items.Add(this.btnFormatSheet);
            this.grpCreateStockInst.Items.Add(this.drpEnviroment);
            this.grpCreateStockInst.Items.Add(this.btnCreateStock);
            this.grpCreateStockInst.Label = "Create Stock INST v1.0.0";
            this.grpCreateStockInst.Name = "grpCreateStockInst";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Format Create Stock Sheet";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            this.drpEnviroment.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drpEnviroment_SelectionChanged);
            // 
            // btnCreateStock
            // 
            this.btnCreateStock.Label = "Execute Create Stock Inst";
            this.btnCreateStock.Name = "btnCreateStock";
            this.btnCreateStock.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateStock_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpCreateStockInst.ResumeLayout(false);
            this.grpCreateStockInst.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCreateStockInst;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateStock;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

namespace EllipseStockCodesExcelAddIn
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
            this.grpStockCodeTrans = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewInventory = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnReviewPurchaseOrders = this.Factory.CreateRibbonButton();
            this.menuRequisitionActions = this.Factory.CreateRibbonMenu();
            this.cbValidOnly = this.Factory.CreateRibbonCheckBox();
            this.cbPreferedOnly = this.Factory.CreateRibbonCheckBox();
            this.tabEllipse.SuspendLayout();
            this.grpStockCodeTrans.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpStockCodeTrans);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpStockCodeTrans
            // 
            this.grpStockCodeTrans.Items.Add(this.box1);
            this.grpStockCodeTrans.Label = "StockCode";
            this.grpStockCodeTrans.Name = "grpStockCodeTrans";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.drpEnviroment);
            this.box1.Items.Add(this.menuActions);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.btnFormatSheet);
            this.box2.Items.Add(this.btnAbout);
            this.box2.Name = "box2";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatRequisitions_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReviewInventory);
            this.menuActions.Items.Add(this.cbValidOnly);
            this.menuActions.Items.Add(this.cbPreferedOnly);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewInventory
            // 
            this.btnReviewInventory.Label = "Consultar Inventario";
            this.btnReviewInventory.Name = "btnReviewInventory";
            this.btnReviewInventory.ShowImage = true;
            this.btnReviewInventory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewStockCodesRequisitions_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnReviewPurchaseOrders
            // 
            this.btnReviewPurchaseOrders.Label = "";
            this.btnReviewPurchaseOrders.Name = "btnReviewPurchaseOrders";
            // 
            // menuRequisitionActions
            // 
            this.menuRequisitionActions.Label = "";
            this.menuRequisitionActions.Name = "menuRequisitionActions";
            // 
            // cbValidOnly
            // 
            this.cbValidOnly.Label = "Sólo &Válidos";
            this.cbValidOnly.Name = "cbValidOnly";
            // 
            // cbPreferedOnly
            // 
            this.cbPreferedOnly.Label = "Sólo Preferidos";
            this.cbPreferedOnly.Name = "cbPreferedOnly";
            this.cbPreferedOnly.Visible = false;
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpStockCodeTrans.ResumeLayout(false);
            this.grpStockCodeTrans.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStockCodeTrans;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewInventory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewPurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRequisitionActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbValidOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbPreferedOnly;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

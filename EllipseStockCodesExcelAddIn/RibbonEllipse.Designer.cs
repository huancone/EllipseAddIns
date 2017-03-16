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
            this.menuFormatSheet = this.Factory.CreateRibbonMenu();
            this.btnFormatRequisitions = this.Factory.CreateRibbonButton();
            this.btnFormatPurchaseOrdersExtended = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.box3 = this.Factory.CreateRibbonBox();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.box4 = this.Factory.CreateRibbonBox();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewPurchaseOrders = this.Factory.CreateRibbonButton();
            this.menuRequisitionActions = this.Factory.CreateRibbonMenu();
            this.btnReviewStockCodesRequisitions = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpStockCodeTrans.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.box4.SuspendLayout();
            this.SuspendLayout();
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
            this.grpStockCodeTrans.Label = "StockCode v0.0.1";
            this.grpStockCodeTrans.Name = "grpStockCodeTrans";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.box3);
            this.box1.Items.Add(this.box4);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.menuFormatSheet);
            this.box2.Items.Add(this.btnAbout);
            this.box2.Name = "box2";
            // 
            // menuFormatSheet
            // 
            this.menuFormatSheet.Items.Add(this.btnFormatRequisitions);
            this.menuFormatSheet.Items.Add(this.btnFormatPurchaseOrdersExtended);
            this.menuFormatSheet.Label = "Formatear Hoja";
            this.menuFormatSheet.Name = "menuFormatSheet";
            // 
            // btnFormatRequisitions
            // 
            this.btnFormatRequisitions.Label = "Formato Requisiciones";
            this.btnFormatRequisitions.Name = "btnFormatRequisitions";
            this.btnFormatRequisitions.ShowImage = true;
            this.btnFormatRequisitions.Visible = false;
            this.btnFormatRequisitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatRequisitions_Click);
            // 
            // btnFormatPurchaseOrdersExtended
            // 
            this.btnFormatPurchaseOrdersExtended.Label = "Formatear Hoja";
            this.btnFormatPurchaseOrdersExtended.Name = "btnFormatPurchaseOrdersExtended";
            this.btnFormatPurchaseOrdersExtended.ShowImage = true;
            this.btnFormatPurchaseOrdersExtended.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatPurchaseOrdersExtended_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // box3
            // 
            this.box3.Items.Add(this.drpEnviroment);
            this.box3.Name = "box3";
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // box4
            // 
            this.box4.Items.Add(this.menuActions);
            this.box4.Name = "box4";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReviewPurchaseOrders);
            this.menuActions.Items.Add(this.menuRequisitionActions);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewPurchaseOrders
            // 
            this.btnReviewPurchaseOrders.Label = "&Consultar";
            this.btnReviewPurchaseOrders.Name = "btnReviewPurchaseOrders";
            this.btnReviewPurchaseOrders.ShowImage = true;
            this.btnReviewPurchaseOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewPurchaseOrders_Click);
            // 
            // menuRequisitionActions
            // 
            this.menuRequisitionActions.Items.Add(this.btnReviewStockCodesRequisitions);
            this.menuRequisitionActions.Label = "&Requisiciones";
            this.menuRequisitionActions.Name = "menuRequisitionActions";
            this.menuRequisitionActions.ShowImage = true;
            this.menuRequisitionActions.Visible = false;
            // 
            // btnReviewStockCodesRequisitions
            // 
            this.btnReviewStockCodesRequisitions.Label = "Consultar Vales por SCs";
            this.btnReviewStockCodesRequisitions.Name = "btnReviewStockCodesRequisitions";
            this.btnReviewStockCodesRequisitions.ShowImage = true;
            this.btnReviewStockCodesRequisitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewStockCodesRequisitions_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
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
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStockCodeTrans;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatRequisitions;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewStockCodesRequisitions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewPurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRequisitionActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatPurchaseOrdersExtended;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
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

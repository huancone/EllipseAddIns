namespace EllipseTransaccionesStockCodesExcelAddIn
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
            this.btnFormatPurchaseOrders = this.Factory.CreateRibbonButton();
            this.btnFormatPurchaseOrdersExtended = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.box3 = this.Factory.CreateRibbonBox();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.box4 = this.Factory.CreateRibbonBox();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuRequisitionActions = this.Factory.CreateRibbonMenu();
            this.btnReviewStockCodesRequisitions = this.Factory.CreateRibbonButton();
            this.menuPurchaseOrderActions = this.Factory.CreateRibbonMenu();
            this.btnReviewPurchaseOrders = this.Factory.CreateRibbonButton();
            this.btnModifyPurchaseOrders = this.Factory.CreateRibbonButton();
            this.btnDeletePurchaseOrders = this.Factory.CreateRibbonButton();
            this.btnDeletePurchaseOrderItem = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpStockCodeTrans.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.box4.SuspendLayout();
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
            this.grpStockCodeTrans.Label = "StockCode Trans. v1.0.4";
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
            this.menuFormatSheet.Items.Add(this.btnFormatPurchaseOrders);
            this.menuFormatSheet.Items.Add(this.btnFormatPurchaseOrdersExtended);
            this.menuFormatSheet.Label = "Formatear Hoja";
            this.menuFormatSheet.Name = "menuFormatSheet";
            // 
            // btnFormatRequisitions
            // 
            this.btnFormatRequisitions.Label = "Formato Requisiciones";
            this.btnFormatRequisitions.Name = "btnFormatRequisitions";
            this.btnFormatRequisitions.ShowImage = true;
            this.btnFormatRequisitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatRequisitions_Click);
            // 
            // btnFormatPurchaseOrders
            // 
            this.btnFormatPurchaseOrders.Label = "Formato Órdenes de Compra";
            this.btnFormatPurchaseOrders.Name = "btnFormatPurchaseOrders";
            this.btnFormatPurchaseOrders.ShowImage = true;
            this.btnFormatPurchaseOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatPurchaseOrders_Click);
            // 
            // btnFormatPurchaseOrdersExtended
            // 
            this.btnFormatPurchaseOrdersExtended.Label = "Formato Órdenes de Compra - Extendido";
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
            this.menuActions.Items.Add(this.menuRequisitionActions);
            this.menuActions.Items.Add(this.menuPurchaseOrderActions);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuRequisitionActions
            // 
            this.menuRequisitionActions.Items.Add(this.btnReviewStockCodesRequisitions);
            this.menuRequisitionActions.Label = "&Requisiciones";
            this.menuRequisitionActions.Name = "menuRequisitionActions";
            this.menuRequisitionActions.ShowImage = true;
            // 
            // btnReviewStockCodesRequisitions
            // 
            this.btnReviewStockCodesRequisitions.Label = "Consultar Vales por SCs";
            this.btnReviewStockCodesRequisitions.Name = "btnReviewStockCodesRequisitions";
            this.btnReviewStockCodesRequisitions.ShowImage = true;
            this.btnReviewStockCodesRequisitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewStockCodesRequisitions_Click);
            // 
            // menuPurchaseOrderActions
            // 
            this.menuPurchaseOrderActions.Items.Add(this.btnReviewPurchaseOrders);
            this.menuPurchaseOrderActions.Items.Add(this.btnModifyPurchaseOrders);
            this.menuPurchaseOrderActions.Items.Add(this.btnDeletePurchaseOrders);
            this.menuPurchaseOrderActions.Items.Add(this.btnDeletePurchaseOrderItem);
            this.menuPurchaseOrderActions.Label = "Órdenes de Com&pra";
            this.menuPurchaseOrderActions.Name = "menuPurchaseOrderActions";
            this.menuPurchaseOrderActions.ShowImage = true;
            // 
            // btnReviewPurchaseOrders
            // 
            this.btnReviewPurchaseOrders.Label = "&Consultar Órdenes de Compra";
            this.btnReviewPurchaseOrders.Name = "btnReviewPurchaseOrders";
            this.btnReviewPurchaseOrders.ShowImage = true;
            this.btnReviewPurchaseOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewPurchaseOrders_Click);
            // 
            // btnModifyPurchaseOrders
            // 
            this.btnModifyPurchaseOrders.Label = "&Modificar Órdenes";
            this.btnModifyPurchaseOrders.Name = "btnModifyPurchaseOrders";
            this.btnModifyPurchaseOrders.ShowImage = true;
            this.btnModifyPurchaseOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyPurchaseOrders_Click);
            // 
            // btnDeletePurchaseOrders
            // 
            this.btnDeletePurchaseOrders.Label = "&Eliminar Órdenes";
            this.btnDeletePurchaseOrders.Name = "btnDeletePurchaseOrders";
            this.btnDeletePurchaseOrders.ShowImage = true;
            this.btnDeletePurchaseOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeletePurchaseOrders_Click);
            // 
            // btnDeletePurchaseOrderItem
            // 
            this.btnDeletePurchaseOrderItem.Label = "Eliminar &Item";
            this.btnDeletePurchaseOrderItem.Name = "btnDeletePurchaseOrderItem";
            this.btnDeletePurchaseOrderItem.ShowImage = true;
            this.btnDeletePurchaseOrderItem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeletePurchaseOrderItem_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatPurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRequisitionActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuPurchaseOrderActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyPurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeletePurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeletePurchaseOrderItem;
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

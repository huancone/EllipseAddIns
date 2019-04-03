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
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.cbValidOnly = this.Factory.CreateRibbonCheckBox();
            this.cbPreferedOnly = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnReviewPurchaseOrders = this.Factory.CreateRibbonButton();
            this.menuRequisitionActions = this.Factory.CreateRibbonMenu();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tabEllipse.SuspendLayout();
            this.grpStockCodeTrans.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
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
            this.grpStockCodeTrans.Label = "StockCode";
            this.grpStockCodeTrans.Name = "grpStockCodeTrans";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.drpEnvironment);
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
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReview);
            this.menuActions.Items.Add(this.cbValidOnly);
            this.menuActions.Items.Add(this.cbPreferedOnly);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReview
            // 
            this.btnReview.Label = "Consultar";
            this.btnReview.Name = "btnReview";
            this.btnReview.ShowImage = true;
            this.btnReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // cbValidOnly
            // 
            this.cbValidOnly.Label = "Sólo PNs &Válidos";
            this.cbValidOnly.Name = "cbValidOnly";
            // 
            // cbPreferedOnly
            // 
            this.cbPreferedOnly.Checked = true;
            this.cbPreferedOnly.Label = "Sólo PN Preferido";
            this.cbPreferedOnly.Name = "cbPreferedOnly";
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
            // menu1
            // 
            this.menu1.Label = "menu1";
            this.menu1.Name = "menu1";
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStockCodeTrans;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewPurchaseOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRequisitionActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbValidOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbPreferedOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

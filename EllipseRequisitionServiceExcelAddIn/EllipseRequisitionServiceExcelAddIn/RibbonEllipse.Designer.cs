namespace EllipseRequisitionServiceExcelAddIn
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
            this.grpRequisitionService = this.Factory.CreateRibbonGroup();
            this.btnFormatNewSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuAction = this.Factory.CreateRibbonMenu();
            this.btnExcecuteRequisitionService = this.Factory.CreateRibbonButton();
            this.btnCreateReqIgError = this.Factory.CreateRibbonButton();
            this.btnCreateReqDirectOrderItems = this.Factory.CreateRibbonButton();
            this.cbMaxItems = this.Factory.CreateRibbonCheckBox();
            this.btnCleanSheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnManualCreditRequisitionMSE1VR = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpRequisitionService.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpRequisitionService);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpRequisitionService
            // 
            this.grpRequisitionService.Items.Add(this.btnFormatNewSheet);
            this.grpRequisitionService.Items.Add(this.drpEnviroment);
            this.grpRequisitionService.Items.Add(this.menuAction);
            this.grpRequisitionService.Label = "Requisition Service v1.2.1";
            this.grpRequisitionService.Name = "grpRequisitionService";
            // 
            // btnFormatNewSheet
            // 
            this.btnFormatNewSheet.Label = "&Formatear Hoja";
            this.btnFormatNewSheet.Name = "btnFormatNewSheet";
            this.btnFormatNewSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatNewSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuAction
            // 
            this.menuAction.Items.Add(this.btnExcecuteRequisitionService);
            this.menuAction.Items.Add(this.btnCreateReqIgError);
            this.menuAction.Items.Add(this.btnCreateReqDirectOrderItems);
            this.menuAction.Items.Add(this.btnManualCreditRequisitionMSE1VR);
            this.menuAction.Items.Add(this.cbMaxItems);
            this.menuAction.Items.Add(this.btnCleanSheet);
            this.menuAction.Items.Add(this.btnStopThread);
            this.menuAction.Label = "&Acciones";
            this.menuAction.Name = "menuAction";
            // 
            // btnExcecuteRequisitionService
            // 
            this.btnExcecuteRequisitionService.Label = "&Crear Requisición";
            this.btnExcecuteRequisitionService.Name = "btnExcecuteRequisitionService";
            this.btnExcecuteRequisitionService.ShowImage = true;
            this.btnExcecuteRequisitionService.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExcecuteRequisitionService_Click);
            // 
            // btnCreateReqIgError
            // 
            this.btnCreateReqIgError.Label = "Crear Req. - Ignorar Errores";
            this.btnCreateReqIgError.Name = "btnCreateReqIgError";
            this.btnCreateReqIgError.ShowImage = true;
            this.btnCreateReqIgError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateReqIgError_Click);
            // 
            // btnCreateReqDirectOrderItems
            // 
            this.btnCreateReqDirectOrderItems.Label = "Crear Req. - Items Orden Directa";
            this.btnCreateReqDirectOrderItems.Name = "btnCreateReqDirectOrderItems";
            this.btnCreateReqDirectOrderItems.ShowImage = true;
            this.btnCreateReqDirectOrderItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateReqDirectOrderItems_Click);
            // 
            // btnCleanSheet
            // 
            this.btnCleanSheet.Label = "&Limpiar Tabla";
            this.btnCleanSheet.Name = "btnCleanSheet";
            this.btnCleanSheet.ShowImage = true;
            this.btnCleanSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanSheet_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnManualCreditRequisitionMSE1VR
            // 
            this.btnManualCreditRequisitionMSE1VR.Label = "Devolucion Manual MSE1VR";
            this.btnManualCreditRequisitionMSE1VR.Name = "btnManualCreditRequisitionMSE1VR";
            this.btnManualCreditRequisitionMSE1VR.ShowImage = true;
            this.btnManualCreditRequisitionMSE1VR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnManualCreditRequisitionMSE1VR_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpRequisitionService.ResumeLayout(false);
            this.grpRequisitionService.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpRequisitionService;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatNewSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExcecuteRequisitionService;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateReqIgError;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateReqDirectOrderItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbMaxItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnManualCreditRequisitionMSE1VR;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

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
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatNewSheet = this.Factory.CreateRibbonButton();
            this.btnFormatMnttoAssurance = this.Factory.CreateRibbonButton();
            this.btnFormatExtended = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuAction = this.Factory.CreateRibbonMenu();
            this.menuQueries = this.Factory.CreateRibbonMenu();
            this.btnQueryRequisitions = this.Factory.CreateRibbonButton();
            this.btnReviewRequisitionControl = this.Factory.CreateRibbonButton();
            this.btnReReviewRequisitionControl = this.Factory.CreateRibbonButton();
            this.menuOptions = this.Factory.CreateRibbonMenu();
            this.cbMaxItems = this.Factory.CreateRibbonCheckBox();
            this.cbSortItems = this.Factory.CreateRibbonCheckBox();
            this.cbAssuranceProcess = this.Factory.CreateRibbonCheckBox();
            this.btnConfigAssuranceSettings = this.Factory.CreateRibbonButton();
            this.btnExcecuteRequisitionService = this.Factory.CreateRibbonButton();
            this.btnCreateReqIgError = this.Factory.CreateRibbonButton();
            this.btnCreateReqDirectOrderItems = this.Factory.CreateRibbonButton();
            this.btnManualCreditRequisitionMSE1VR = this.Factory.CreateRibbonButton();
            this.btnCleanSheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpRequisitionService.SuspendLayout();
            this.box1.SuspendLayout();
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
            this.grpRequisitionService.Items.Add(this.box1);
            this.grpRequisitionService.Items.Add(this.drpEnvironment);
            this.grpRequisitionService.Items.Add(this.menuAction);
            this.grpRequisitionService.Label = "Requisition Service";
            this.grpRequisitionService.Name = "grpRequisitionService";
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatNewSheet);
            this.menuFormat.Items.Add(this.btnFormatMnttoAssurance);
            this.menuFormat.Items.Add(this.btnFormatExtended);
            this.menuFormat.Label = "&Formatear";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatNewSheet
            // 
            this.btnFormatNewSheet.Label = "&Formato Base";
            this.btnFormatNewSheet.Name = "btnFormatNewSheet";
            this.btnFormatNewSheet.ShowImage = true;
            this.btnFormatNewSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatNewSheet_Click);
            // 
            // btnFormatMnttoAssurance
            // 
            this.btnFormatMnttoAssurance.Label = "Formato - Validación de Garantías";
            this.btnFormatMnttoAssurance.Name = "btnFormatMnttoAssurance";
            this.btnFormatMnttoAssurance.ShowImage = true;
            this.btnFormatMnttoAssurance.Visible = false;
            this.btnFormatMnttoAssurance.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatMnttoAssurance_Click);
            // 
            // btnFormatExtended
            // 
            this.btnFormatExtended.Label = "Formato E&xtendido";
            this.btnFormatExtended.Name = "btnFormatExtended";
            this.btnFormatExtended.ShowImage = true;
            this.btnFormatExtended.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatExtended_Click);
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
            // menuAction
            // 
            this.menuAction.Items.Add(this.menuQueries);
            this.menuAction.Items.Add(this.menuOptions);
            this.menuAction.Items.Add(this.btnExcecuteRequisitionService);
            this.menuAction.Items.Add(this.btnCreateReqIgError);
            this.menuAction.Items.Add(this.btnCreateReqDirectOrderItems);
            this.menuAction.Items.Add(this.btnManualCreditRequisitionMSE1VR);
            this.menuAction.Items.Add(this.btnCleanSheet);
            this.menuAction.Items.Add(this.btnStopThread);
            this.menuAction.Label = "&Acciones";
            this.menuAction.Name = "menuAction";
            // 
            // menuQueries
            // 
            this.menuQueries.Items.Add(this.btnQueryRequisitions);
            this.menuQueries.Items.Add(this.btnReviewRequisitionControl);
            this.menuQueries.Items.Add(this.btnReReviewRequisitionControl);
            this.menuQueries.Label = "Consultas";
            this.menuQueries.Name = "menuQueries";
            this.menuQueries.ShowImage = true;
            // 
            // btnQueryRequisitions
            // 
            this.btnQueryRequisitions.Label = "Consultar Requisiciones";
            this.btnQueryRequisitions.Name = "btnQueryRequisitions";
            this.btnQueryRequisitions.ShowImage = true;
            this.btnQueryRequisitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQueryRequisitions_Click);
            // 
            // btnReviewRequisitionControl
            // 
            this.btnReviewRequisitionControl.Label = "Consultar Detalle - Control";
            this.btnReviewRequisitionControl.Name = "btnReviewRequisitionControl";
            this.btnReviewRequisitionControl.ShowImage = true;
            this.btnReviewRequisitionControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRequisitionControl_Click);
            // 
            // btnReReviewRequisitionControl
            // 
            this.btnReReviewRequisitionControl.Label = "Reconsultar Detalle - Control";
            this.btnReReviewRequisitionControl.Name = "btnReReviewRequisitionControl";
            this.btnReReviewRequisitionControl.ShowImage = true;
            this.btnReReviewRequisitionControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewRequisitionControl_Click);
            // 
            // menuOptions
            // 
            this.menuOptions.Items.Add(this.cbMaxItems);
            this.menuOptions.Items.Add(this.cbSortItems);
            this.menuOptions.Items.Add(this.cbAssuranceProcess);
            this.menuOptions.Items.Add(this.btnConfigAssuranceSettings);
            this.menuOptions.Label = "Opciones";
            this.menuOptions.Name = "menuOptions";
            this.menuOptions.ShowImage = true;
            // 
            // cbMaxItems
            // 
            this.cbMaxItems.Label = "Máximo 100 Items por Pedido";
            this.cbMaxItems.Name = "cbMaxItems";
            this.cbMaxItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbMaxItems_Click);
            // 
            // cbSortItems
            // 
            this.cbSortItems.Label = "Auto&Ordenar Elementos";
            this.cbSortItems.Name = "cbSortItems";
            this.cbSortItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbSortItems_Click);
            // 
            // cbAssuranceProcess
            // 
            this.cbAssuranceProcess.Label = "Seguimiento a Garantías";
            this.cbAssuranceProcess.Name = "cbAssuranceProcess";
            this.cbAssuranceProcess.Visible = false;
            this.cbAssuranceProcess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAssuranceProcess_Click);
            // 
            // btnConfigAssuranceSettings
            // 
            this.btnConfigAssuranceSettings.Label = "Configurar Opciones de Garantías";
            this.btnConfigAssuranceSettings.Name = "btnConfigAssuranceSettings";
            this.btnConfigAssuranceSettings.ShowImage = true;
            this.btnConfigAssuranceSettings.Visible = false;
            this.btnConfigAssuranceSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfigAssuranceSettings_Click);
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
            // btnManualCreditRequisitionMSE1VR
            // 
            this.btnManualCreditRequisitionMSE1VR.Label = "Devolucion Manual MSE1VR";
            this.btnManualCreditRequisitionMSE1VR.Name = "btnManualCreditRequisitionMSE1VR";
            this.btnManualCreditRequisitionMSE1VR.ShowImage = true;
            this.btnManualCreditRequisitionMSE1VR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnManualCreditRequisitionMSE1VR_Click);
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
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpRequisitionService;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatNewSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExcecuteRequisitionService;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateReqIgError;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateReqDirectOrderItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnManualCreditRequisitionMSE1VR;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbMaxItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatExtended;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbSortItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuQueries;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQueryRequisitions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRequisitionControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewRequisitionControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAssuranceProcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfigAssuranceSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatMnttoAssurance;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

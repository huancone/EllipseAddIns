namespace EllipseWorkRequestExcelAddIn
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
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpWorkRequest = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnFormatMantto = this.Factory.CreateRibbonButton();
            this.btnFormatFcVagones = this.Factory.CreateRibbonButton();
            this.btnPlanFc = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuWorkRequest = this.Factory.CreateRibbonMenu();
            this.btnReviewWorkRequest = this.Factory.CreateRibbonButton();
            this.btnReReviewWorkRequest = this.Factory.CreateRibbonButton();
            this.btnCreateWorkRequest = this.Factory.CreateRibbonButton();
            this.btnModifyWorkRequest = this.Factory.CreateRibbonButton();
            this.btnDeleteWorkRequest = this.Factory.CreateRibbonButton();
            this.menuSla = this.Factory.CreateRibbonMenu();
            this.btnSetSla = this.Factory.CreateRibbonButton();
            this.btnResetSla = this.Factory.CreateRibbonButton();
            this.menuCloseWorkRequest = this.Factory.CreateRibbonMenu();
            this.btnReOpenWorkRequest = this.Factory.CreateRibbonButton();
            this.btnCloseWorkRequest = this.Factory.CreateRibbonButton();
            this.menuReferenceCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewRefCodes = this.Factory.CreateRibbonButton();
            this.btnReReviewRefCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateRefCodes = this.Factory.CreateRibbonButton();
            this.btnCleanSheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkRequest.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menu1
            // 
            this.menu1.Label = "menu1";
            this.menu1.Name = "menu1";
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpWorkRequest);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpWorkRequest
            // 
            this.grpWorkRequest.Items.Add(this.box1);
            this.grpWorkRequest.Items.Add(this.drpEnvironment);
            this.grpWorkRequest.Items.Add(this.menuActions);
            this.grpWorkRequest.Label = "WorkRequest";
            this.grpWorkRequest.Name = "grpWorkRequest";
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatSheet);
            this.menuFormat.Items.Add(this.btnFormatMantto);
            this.menuFormat.Items.Add(this.btnFormatFcVagones);
            this.menuFormat.Items.Add(this.btnPlanFc);
            this.menuFormat.Label = "&Formatear";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formato General";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // btnFormatMantto
            // 
            this.btnFormatMantto.Label = "Formatear &MTTO";
            this.btnFormatMantto.Name = "btnFormatMantto";
            this.btnFormatMantto.ShowImage = true;
            this.btnFormatMantto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatMantto_Click);
            // 
            // btnFormatFcVagones
            // 
            this.btnFormatFcVagones.Label = "Registro Fc Vagones";
            this.btnFormatFcVagones.Name = "btnFormatFcVagones";
            this.btnFormatFcVagones.ShowImage = true;
            this.btnFormatFcVagones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatFcVagones_Click);
            // 
            // btnPlanFc
            // 
            this.btnPlanFc.Label = "Registro Fc Solicitudes";
            this.btnPlanFc.Name = "btnPlanFc";
            this.btnPlanFc.ShowImage = true;
            this.btnPlanFc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlanFc_Click);
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
            this.menuActions.Items.Add(this.menuWorkRequest);
            this.menuActions.Items.Add(this.menuSla);
            this.menuActions.Items.Add(this.menuCloseWorkRequest);
            this.menuActions.Items.Add(this.menuReferenceCodes);
            this.menuActions.Items.Add(this.btnCleanSheet);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuWorkRequest
            // 
            this.menuWorkRequest.Items.Add(this.btnReviewWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnReReviewWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnCreateWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnModifyWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnDeleteWorkRequest);
            this.menuWorkRequest.Label = "&Work Request";
            this.menuWorkRequest.Name = "menuWorkRequest";
            this.menuWorkRequest.ShowImage = true;
            // 
            // btnReviewWorkRequest
            // 
            this.btnReviewWorkRequest.Label = "Consultar &Información";
            this.btnReviewWorkRequest.Name = "btnReviewWorkRequest";
            this.btnReviewWorkRequest.ShowImage = true;
            this.btnReviewWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewWorkRequest_Click);
            // 
            // btnReReviewWorkRequest
            // 
            this.btnReReviewWorkRequest.Label = "&Reconsultar Información";
            this.btnReReviewWorkRequest.Name = "btnReReviewWorkRequest";
            this.btnReReviewWorkRequest.ShowImage = true;
            this.btnReReviewWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewWorkRequest_Click);
            // 
            // btnCreateWorkRequest
            // 
            this.btnCreateWorkRequest.Label = "&Crear WorkRequest";
            this.btnCreateWorkRequest.Name = "btnCreateWorkRequest";
            this.btnCreateWorkRequest.ShowImage = true;
            this.btnCreateWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateWorkRequest_Click);
            // 
            // btnModifyWorkRequest
            // 
            this.btnModifyWorkRequest.Label = "&Actualizar WorkRequest";
            this.btnModifyWorkRequest.Name = "btnModifyWorkRequest";
            this.btnModifyWorkRequest.ShowImage = true;
            this.btnModifyWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyWorkRequest_Click);
            // 
            // btnDeleteWorkRequest
            // 
            this.btnDeleteWorkRequest.Label = "&Eliminar WorkRequests";
            this.btnDeleteWorkRequest.Name = "btnDeleteWorkRequest";
            this.btnDeleteWorkRequest.ShowImage = true;
            this.btnDeleteWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteWorkRequest_Click);
            // 
            // menuSla
            // 
            this.menuSla.Items.Add(this.btnSetSla);
            this.menuSla.Items.Add(this.btnResetSla);
            this.menuSla.Label = "Service &Level Agreement";
            this.menuSla.Name = "menuSla";
            this.menuSla.ShowImage = true;
            // 
            // btnSetSla
            // 
            this.btnSetSla.Label = "&Establecer SLA";
            this.btnSetSla.Name = "btnSetSla";
            this.btnSetSla.ShowImage = true;
            this.btnSetSla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetSla_Click);
            // 
            // btnResetSla
            // 
            this.btnResetSla.Label = "&Reset SLA";
            this.btnResetSla.Name = "btnResetSla";
            this.btnResetSla.ShowImage = true;
            this.btnResetSla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetSla_Click);
            // 
            // menuCloseWorkRequest
            // 
            this.menuCloseWorkRequest.Items.Add(this.btnReOpenWorkRequest);
            this.menuCloseWorkRequest.Items.Add(this.btnCloseWorkRequest);
            this.menuCloseWorkRequest.Label = "Cierre de WorkRequest";
            this.menuCloseWorkRequest.Name = "menuCloseWorkRequest";
            this.menuCloseWorkRequest.ShowImage = true;
            // 
            // btnReOpenWorkRequest
            // 
            this.btnReOpenWorkRequest.Label = "Re&Abrir WorkRequest";
            this.btnReOpenWorkRequest.Name = "btnReOpenWorkRequest";
            this.btnReOpenWorkRequest.ShowImage = true;
            this.btnReOpenWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReOpenWorkRequest_Click);
            // 
            // btnCloseWorkRequest
            // 
            this.btnCloseWorkRequest.Label = "&Cerrar WorkRequest";
            this.btnCloseWorkRequest.Name = "btnCloseWorkRequest";
            this.btnCloseWorkRequest.ShowImage = true;
            this.btnCloseWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCloseWorkRequest_Click);
            // 
            // menuReferenceCodes
            // 
            this.menuReferenceCodes.Items.Add(this.btnReviewRefCodes);
            this.menuReferenceCodes.Items.Add(this.btnReReviewRefCodes);
            this.menuReferenceCodes.Items.Add(this.btnUpdateRefCodes);
            this.menuReferenceCodes.Label = "&Reference Codes";
            this.menuReferenceCodes.Name = "menuReferenceCodes";
            this.menuReferenceCodes.ShowImage = true;
            // 
            // btnReviewRefCodes
            // 
            this.btnReviewRefCodes.Label = "&Consultar WR RefCodes";
            this.btnReviewRefCodes.Name = "btnReviewRefCodes";
            this.btnReviewRefCodes.ShowImage = true;
            this.btnReviewRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRefCodes_Click);
            // 
            // btnReReviewRefCodes
            // 
            this.btnReReviewRefCodes.Label = "&Reconsultar WR RefCodes";
            this.btnReReviewRefCodes.Name = "btnReReviewRefCodes";
            this.btnReReviewRefCodes.ShowImage = true;
            this.btnReReviewRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewRefCodes_Click);
            // 
            // btnUpdateRefCodes
            // 
            this.btnUpdateRefCodes.Label = "&Actualizar WR RefCodes";
            this.btnUpdateRefCodes.Name = "btnUpdateRefCodes";
            this.btnUpdateRefCodes.ShowImage = true;
            this.btnUpdateRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRefCodes_Click);
            // 
            // btnCleanSheet
            // 
            this.btnCleanSheet.Label = "&Limpiar Hoja";
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
            this.grpWorkRequest.ResumeLayout(false);
            this.grpWorkRequest.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloseWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCloseWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReOpenWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuReferenceCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatMantto;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatFcVagones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanFc;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

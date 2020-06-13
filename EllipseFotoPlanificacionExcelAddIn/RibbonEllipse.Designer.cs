namespace EllipseFotoPlanificacionExcelAddIn
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
            this.grpPhotoPlanner = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewEllipse = this.Factory.CreateRibbonButton();
            this.btnReviewSigman = this.Factory.CreateRibbonButton();
            this.btnUpdateSigman = this.Factory.CreateRibbonButton();
            this.menuUpdateExisting = this.Factory.CreateRibbonMenu();
            this.cbIgnoreUpdateError = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.cbDeactivateExisting = this.Factory.CreateRibbonCheckBox();
            this.cbDeleteExisting = this.Factory.CreateRibbonCheckBox();
            this.cbIgnoreExisting = this.Factory.CreateRibbonCheckBox();
            this.cbIgnoreNextTask = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpPhotoPlanner.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpPhotoPlanner);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpPhotoPlanner
            // 
            this.grpPhotoPlanner.Items.Add(this.box1);
            this.grpPhotoPlanner.Items.Add(this.drpEnvironment);
            this.grpPhotoPlanner.Items.Add(this.menuActions);
            this.grpPhotoPlanner.Label = "Foto Planificación";
            this.grpPhotoPlanner.Name = "grpPhotoPlanner";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "&Formatear Hoja";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
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
            this.menuActions.Items.Add(this.btnReviewEllipse);
            this.menuActions.Items.Add(this.btnReviewSigman);
            this.menuActions.Items.Add(this.btnUpdateSigman);
            this.menuActions.Items.Add(this.menuUpdateExisting);
            this.menuActions.Items.Add(this.cbIgnoreNextTask);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewEllipse
            // 
            this.btnReviewEllipse.Label = "Consultar de &Ellipse";
            this.btnReviewEllipse.Name = "btnReviewEllipse";
            this.btnReviewEllipse.ShowImage = true;
            this.btnReviewEllipse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewEllipse_Click);
            // 
            // btnReviewSigman
            // 
            this.btnReviewSigman.Label = "Consultar de &Sigman";
            this.btnReviewSigman.Name = "btnReviewSigman";
            this.btnReviewSigman.ShowImage = true;
            this.btnReviewSigman.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewSigman_Click);
            // 
            // btnUpdateSigman
            // 
            this.btnUpdateSigman.Label = "&Actualizar en Sigman";
            this.btnUpdateSigman.Name = "btnUpdateSigman";
            this.btnUpdateSigman.ShowImage = true;
            this.btnUpdateSigman.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateSigman_Click);
            // 
            // menuUpdateExisting
            // 
            this.menuUpdateExisting.Items.Add(this.cbIgnoreUpdateError);
            this.menuUpdateExisting.Items.Add(this.separator1);
            this.menuUpdateExisting.Items.Add(this.cbDeactivateExisting);
            this.menuUpdateExisting.Items.Add(this.cbDeleteExisting);
            this.menuUpdateExisting.Items.Add(this.cbIgnoreExisting);
            this.menuUpdateExisting.Label = "Actualizar Existentes";
            this.menuUpdateExisting.Name = "menuUpdateExisting";
            this.menuUpdateExisting.ShowImage = true;
            // 
            // cbIgnoreUpdateError
            // 
            this.cbIgnoreUpdateError.Label = "Ignorar Errores al Actualizar";
            this.cbIgnoreUpdateError.Name = "cbIgnoreUpdateError";
            this.cbIgnoreUpdateError.ScreenTip = "Si está desactivado cancela la actualización y hace rollback";
            this.cbIgnoreUpdateError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbIgnoreUpdateError_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // cbDeactivateExisting
            // 
            this.cbDeactivateExisting.Label = "Desactivar";
            this.cbDeactivateExisting.Name = "cbDeactivateExisting";
            this.cbDeactivateExisting.ScreenTip = "Se desactivarán los registros existentes que coincidan con los periodos y grupos " +
    "seleccionados";
            this.cbDeactivateExisting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbDeactivateExisting_Click);
            // 
            // cbDeleteExisting
            // 
            this.cbDeleteExisting.Label = "Eliminar";
            this.cbDeleteExisting.Name = "cbDeleteExisting";
            this.cbDeleteExisting.ScreenTip = "Se eliminarán los registros existentes que coincidan con los periodos y grupos se" +
    "leccionados";
            this.cbDeleteExisting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbDeleteExisting_Click);
            // 
            // cbIgnoreExisting
            // 
            this.cbIgnoreExisting.Label = "No Cambiar";
            this.cbIgnoreExisting.Name = "cbIgnoreExisting";
            this.cbIgnoreExisting.ScreenTip = "No hay cambios en los registros existentes. Solo se crearán nuevos registros";
            this.cbIgnoreExisting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbIgnoreExisting_Click);
            // 
            // cbIgnoreNextTask
            // 
            this.cbIgnoreNextTask.Label = "Ignorar Siguiente Fecha";
            this.cbIgnoreNextTask.Name = "cbIgnoreNextTask";
            this.cbIgnoreNextTask.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbIgnoreNextTask_Click);
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
            this.grpPhotoPlanner.ResumeLayout(false);
            this.grpPhotoPlanner.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPhotoPlanner;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewSigman;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateSigman;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIgnoreNextTask;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuUpdateExisting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDeactivateExisting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDeleteExisting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIgnoreExisting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIgnoreUpdateError;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

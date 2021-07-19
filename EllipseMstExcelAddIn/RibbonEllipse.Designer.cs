namespace EllipseMstExcelAddIn
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
            this.grpMaintenanceScheduleTask = this.Factory.CreateRibbonGroup();
            this.boxFormat = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewMsts = this.Factory.CreateRibbonButton();
            this.btnReReviewMst = this.Factory.CreateRibbonButton();
            this.btnCreateMst = this.Factory.CreateRibbonButton();
            this.btnUpdateMst = this.Factory.CreateRibbonButton();
            this.btnModifyNextSchedule = this.Factory.CreateRibbonButton();
            this.btnDeleteTask = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMaintenanceScheduleTask.SuspendLayout();
            this.boxFormat.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMaintenanceScheduleTask);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMaintenanceScheduleTask
            // 
            this.grpMaintenanceScheduleTask.Items.Add(this.boxFormat);
            this.grpMaintenanceScheduleTask.Items.Add(this.drpEnvironment);
            this.grpMaintenanceScheduleTask.Items.Add(this.menuActions);
            this.grpMaintenanceScheduleTask.Label = "Maint.Sched.Task";
            this.grpMaintenanceScheduleTask.Name = "grpMaintenanceScheduleTask";
            // 
            // boxFormat
            // 
            this.boxFormat.Items.Add(this.btnFormatSheet);
            this.boxFormat.Items.Add(this.btnAbout);
            this.boxFormat.Name = "boxFormat";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
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
            this.menuActions.Items.Add(this.btnReviewMsts);
            this.menuActions.Items.Add(this.btnReReviewMst);
            this.menuActions.Items.Add(this.btnCreateMst);
            this.menuActions.Items.Add(this.btnUpdateMst);
            this.menuActions.Items.Add(this.btnModifyNextSchedule);
            this.menuActions.Items.Add(this.btnDeleteTask);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewMsts
            // 
            this.btnReviewMsts.Label = "Consulta&r Tareas";
            this.btnReviewMsts.Name = "btnReviewMsts";
            this.btnReviewMsts.ShowImage = true;
            this.btnReviewMsts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewMsts_Click);
            // 
            // btnReReviewMst
            // 
            this.btnReReviewMst.Label = "&ReConsultar Tareas";
            this.btnReReviewMst.Name = "btnReReviewMst";
            this.btnReReviewMst.ShowImage = true;
            this.btnReReviewMst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewMst_Click);
            // 
            // btnCreateMst
            // 
            this.btnCreateMst.Label = "&Crear Tareas";
            this.btnCreateMst.Name = "btnCreateMst";
            this.btnCreateMst.ShowImage = true;
            this.btnCreateMst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateMsts_Click);
            // 
            // btnUpdateMst
            // 
            this.btnUpdateMst.Label = "Actualizar Tareas";
            this.btnUpdateMst.Name = "btnUpdateMst";
            this.btnUpdateMst.ShowImage = true;
            this.btnUpdateMst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateMst_Click);
            // 
            // btnModifyNextSchedule
            // 
            this.btnModifyNextSchedule.Label = "Ajustar &Programación";
            this.btnModifyNextSchedule.Name = "btnModifyNextSchedule";
            this.btnModifyNextSchedule.ShowImage = true;
            this.btnModifyNextSchedule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyNextSchedule_Click);
            // 
            // btnDeleteTask
            // 
            this.btnDeleteTask.Label = "&Eliminar Tareas";
            this.btnDeleteTask.Name = "btnDeleteTask";
            this.btnDeleteTask.ShowImage = true;
            this.btnDeleteTask.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteMsts_Click);
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
            this.grpMaintenanceScheduleTask.ResumeLayout(false);
            this.grpMaintenanceScheduleTask.PerformLayout();
            this.boxFormat.ResumeLayout(false);
            this.boxFormat.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMaintenanceScheduleTask;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateMst;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteTask;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewMsts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateMst;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyNextSchedule;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewMst;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxFormat;
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

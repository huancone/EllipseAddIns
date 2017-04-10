namespace EllipseStandardJobsExcelAddIn
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
            this.grpStandardJobs = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuStandardJobs = this.Factory.CreateRibbonMenu();
            this.btnStandardReview = this.Factory.CreateRibbonButton();
            this.btnQuickStandardReview = this.Factory.CreateRibbonButton();
            this.btnReReviewStandard = this.Factory.CreateRibbonButton();
            this.btnActivateStandard = this.Factory.CreateRibbonButton();
            this.btnDeactivateStandard = this.Factory.CreateRibbonButton();
            this.btnCreateStandard = this.Factory.CreateRibbonButton();
            this.btnModifyStandard = this.Factory.CreateRibbonButton();
            this.btnCleanStandardTable = this.Factory.CreateRibbonButton();
            this.menuTasks = this.Factory.CreateRibbonMenu();
            this.btnReviewTasks = this.Factory.CreateRibbonButton();
            this.btnExecuteTaskActions = this.Factory.CreateRibbonButton();
            this.btnCleanTasksTable = this.Factory.CreateRibbonButton();
            this.menuRequirements = this.Factory.CreateRibbonMenu();
            this.btnReviewRequirements = this.Factory.CreateRibbonButton();
            this.btnExecuteRequirements = this.Factory.CreateRibbonButton();
            this.btnGetAplRequirements = this.Factory.CreateRibbonButton();
            this.btnCleanRequirementTable = this.Factory.CreateRibbonButton();
            this.menuReferenceCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewStandardReferenceCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateStandardReferenceCodes = this.Factory.CreateRibbonButton();
            this.menuQualityReview = this.Factory.CreateRibbonMenu();
            this.btnReviewQualityStdJobs = this.Factory.CreateRibbonButton();
            this.btnUpdateQualityStdJobs = this.Factory.CreateRibbonButton();
            this.btnCleanQualityStdJobsTable = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpStandardJobs.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpStandardJobs);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpStandardJobs
            // 
            this.grpStandardJobs.Items.Add(this.box1);
            this.grpStandardJobs.Label = "StandardJobs";
            this.grpStandardJobs.Name = "grpStandardJobs";
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
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "&Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Dynamic = true;
            this.menuActions.Items.Add(this.menuStandardJobs);
            this.menuActions.Items.Add(this.menuTasks);
            this.menuActions.Items.Add(this.menuRequirements);
            this.menuActions.Items.Add(this.menuReferenceCodes);
            this.menuActions.Items.Add(this.menuQualityReview);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuStandardJobs
            // 
            this.menuStandardJobs.Dynamic = true;
            this.menuStandardJobs.Items.Add(this.btnStandardReview);
            this.menuStandardJobs.Items.Add(this.btnQuickStandardReview);
            this.menuStandardJobs.Items.Add(this.btnReReviewStandard);
            this.menuStandardJobs.Items.Add(this.btnActivateStandard);
            this.menuStandardJobs.Items.Add(this.btnDeactivateStandard);
            this.menuStandardJobs.Items.Add(this.btnCreateStandard);
            this.menuStandardJobs.Items.Add(this.btnModifyStandard);
            this.menuStandardJobs.Items.Add(this.btnCleanStandardTable);
            this.menuStandardJobs.Label = "&Standard Jobs";
            this.menuStandardJobs.Name = "menuStandardJobs";
            this.menuStandardJobs.ShowImage = true;
            // 
            // btnStandardReview
            // 
            this.btnStandardReview.Label = "Consultar &Estándares";
            this.btnStandardReview.Name = "btnStandardReview";
            this.btnStandardReview.ShowImage = true;
            this.btnStandardReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStandardReview_Click);
            // 
            // btnQuickStandardReview
            // 
            this.btnQuickStandardReview.Label = "Consulta &Rápida";
            this.btnQuickStandardReview.Name = "btnQuickStandardReview";
            this.btnQuickStandardReview.ShowImage = true;
            this.btnQuickStandardReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuickStandardReview_Click);
            // 
            // btnReReviewStandard
            // 
            this.btnReReviewStandard.Label = "ReConsultar &Tabla Estándares";
            this.btnReReviewStandard.Name = "btnReReviewStandard";
            this.btnReReviewStandard.ShowImage = true;
            this.btnReReviewStandard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewStandard_Click);
            // 
            // btnActivateStandard
            // 
            this.btnActivateStandard.Label = "Acti&var Estándares";
            this.btnActivateStandard.Name = "btnActivateStandard";
            this.btnActivateStandard.ShowImage = true;
            this.btnActivateStandard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActivateStandard_Click);
            // 
            // btnDeactivateStandard
            // 
            this.btnDeactivateStandard.Label = "&Desactivar Estándares";
            this.btnDeactivateStandard.Name = "btnDeactivateStandard";
            this.btnDeactivateStandard.ShowImage = true;
            this.btnDeactivateStandard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeactivateStandard_Click);
            // 
            // btnCreateStandard
            // 
            this.btnCreateStandard.Label = "&Crear Estándares";
            this.btnCreateStandard.Name = "btnCreateStandard";
            this.btnCreateStandard.ShowImage = true;
            this.btnCreateStandard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateStandard_Click);
            // 
            // btnModifyStandard
            // 
            this.btnModifyStandard.Label = "&Actualizar Estándares";
            this.btnModifyStandard.Name = "btnModifyStandard";
            this.btnModifyStandard.ShowImage = true;
            this.btnModifyStandard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyStandard_Click);
            // 
            // btnCleanStandardTable
            // 
            this.btnCleanStandardTable.Label = "&Limpiar Tabla de Estándares";
            this.btnCleanStandardTable.Name = "btnCleanStandardTable";
            this.btnCleanStandardTable.ShowImage = true;
            this.btnCleanStandardTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanStandardTable_Click);
            // 
            // menuTasks
            // 
            this.menuTasks.Dynamic = true;
            this.menuTasks.Items.Add(this.btnReviewTasks);
            this.menuTasks.Items.Add(this.btnExecuteTaskActions);
            this.menuTasks.Items.Add(this.btnCleanTasksTable);
            this.menuTasks.Label = "&Tareas";
            this.menuTasks.Name = "menuTasks";
            this.menuTasks.ShowImage = true;
            // 
            // btnReviewTasks
            // 
            this.btnReviewTasks.Label = "&Consultar Tareas";
            this.btnReviewTasks.Name = "btnReviewTasks";
            this.btnReviewTasks.ShowImage = true;
            this.btnReviewTasks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTasks_Click);
            // 
            // btnExecuteTaskActions
            // 
            this.btnExecuteTaskActions.Label = "&Ejecutar Acciones de Tareas";
            this.btnExecuteTaskActions.Name = "btnExecuteTaskActions";
            this.btnExecuteTaskActions.ShowImage = true;
            this.btnExecuteTaskActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteTaskActions_Click);
            // 
            // btnCleanTasksTable
            // 
            this.btnCleanTasksTable.Label = "&Limpiar Tabla de Tareas";
            this.btnCleanTasksTable.Name = "btnCleanTasksTable";
            this.btnCleanTasksTable.ShowImage = true;
            this.btnCleanTasksTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanTasksTable_Click);
            // 
            // menuRequirements
            // 
            this.menuRequirements.Dynamic = true;
            this.menuRequirements.Items.Add(this.btnReviewRequirements);
            this.menuRequirements.Items.Add(this.btnExecuteRequirements);
            this.menuRequirements.Items.Add(this.btnGetAplRequirements);
            this.menuRequirements.Items.Add(this.btnCleanRequirementTable);
            this.menuRequirements.Label = "&Requerimientos";
            this.menuRequirements.Name = "menuRequirements";
            this.menuRequirements.ShowImage = true;
            // 
            // btnReviewRequirements
            // 
            this.btnReviewRequirements.Label = "Consultar Requerimientos";
            this.btnReviewRequirements.Name = "btnReviewRequirements";
            this.btnReviewRequirements.ShowImage = true;
            this.btnReviewRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRequirements_Click);
            // 
            // btnExecuteRequirements
            // 
            this.btnExecuteRequirements.Label = "Ejecutar Acciones";
            this.btnExecuteRequirements.Name = "btnExecuteRequirements";
            this.btnExecuteRequirements.ShowImage = true;
            this.btnExecuteRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteRequirements_Click);
            // 
            // btnGetAplRequirements
            // 
            this.btnGetAplRequirements.Label = "Traer Recursos de APLs";
            this.btnGetAplRequirements.Name = "btnGetAplRequirements";
            this.btnGetAplRequirements.ShowImage = true;
            this.btnGetAplRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetAplRequirements_Click);
            // 
            // btnCleanRequirementTable
            // 
            this.btnCleanRequirementTable.Label = "Limpiar Tabla Requerimientos";
            this.btnCleanRequirementTable.Name = "btnCleanRequirementTable";
            this.btnCleanRequirementTable.ShowImage = true;
            this.btnCleanRequirementTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanRequirementTable_Click);
            // 
            // menuReferenceCodes
            // 
            this.menuReferenceCodes.Items.Add(this.btnReviewStandardReferenceCodes);
            this.menuReferenceCodes.Items.Add(this.btnUpdateStandardReferenceCodes);
            this.menuReferenceCodes.Label = "Reference Codes";
            this.menuReferenceCodes.Name = "menuReferenceCodes";
            this.menuReferenceCodes.ShowImage = true;
            // 
            // btnReviewStandardReferenceCodes
            // 
            this.btnReviewStandardReferenceCodes.Label = "Consultar Reference Codes";
            this.btnReviewStandardReferenceCodes.Name = "btnReviewStandardReferenceCodes";
            this.btnReviewStandardReferenceCodes.ShowImage = true;
            this.btnReviewStandardReferenceCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewStandardReferenceCodes_Click);
            // 
            // btnUpdateStandardReferenceCodes
            // 
            this.btnUpdateStandardReferenceCodes.Label = "Actualizar Reference Codes";
            this.btnUpdateStandardReferenceCodes.Name = "btnUpdateStandardReferenceCodes";
            this.btnUpdateStandardReferenceCodes.ShowImage = true;
            this.btnUpdateStandardReferenceCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateStandardReferenceCodes_Click);
            // 
            // menuQualityReview
            // 
            this.menuQualityReview.Items.Add(this.btnReviewQualityStdJobs);
            this.menuQualityReview.Items.Add(this.btnUpdateQualityStdJobs);
            this.menuQualityReview.Items.Add(this.btnCleanQualityStdJobsTable);
            this.menuQualityReview.Label = "Calidad de StdJobs";
            this.menuQualityReview.Name = "menuQualityReview";
            this.menuQualityReview.ShowImage = true;
            // 
            // btnReviewQualityStdJobs
            // 
            this.btnReviewQualityStdJobs.Label = "Consultar Estándares";
            this.btnReviewQualityStdJobs.Name = "btnReviewQualityStdJobs";
            this.btnReviewQualityStdJobs.ShowImage = true;
            this.btnReviewQualityStdJobs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewQualityStdJobs_Click);
            // 
            // btnUpdateQualityStdJobs
            // 
            this.btnUpdateQualityStdJobs.Label = "Actualizar Estándares";
            this.btnUpdateQualityStdJobs.Name = "btnUpdateQualityStdJobs";
            this.btnUpdateQualityStdJobs.ShowImage = true;
            this.btnUpdateQualityStdJobs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateQualityStdJobs_Click);
            // 
            // btnCleanQualityStdJobsTable
            // 
            this.btnCleanQualityStdJobsTable.Label = "Limpiar Tabla";
            this.btnCleanQualityStdJobsTable.Name = "btnCleanQualityStdJobsTable";
            this.btnCleanQualityStdJobsTable.ShowImage = true;
            this.btnCleanQualityStdJobsTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanQualityStdJobsTable_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Procesos";
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
            this.grpStandardJobs.ResumeLayout(false);
            this.grpStandardJobs.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStandardJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuStandardJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuTasks;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStandardReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickStandardReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateStandard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyStandard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActivateStandard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeactivateStandard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewStandard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanStandardTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewTasks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecuteTaskActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanTasksTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuQualityReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewQualityStdJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateQualityStdJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanQualityStdJobsTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecuteRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanRequirementTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetAplRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuReferenceCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewStandardReferenceCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateStandardReferenceCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
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

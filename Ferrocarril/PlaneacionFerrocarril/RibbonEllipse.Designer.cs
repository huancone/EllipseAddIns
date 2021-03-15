
namespace PlaneacionFerrocarril
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpEllipse = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuWeeklyPlanning = this.Factory.CreateRibbonMenu();
            this.btnFormatWeekPlanning = this.Factory.CreateRibbonButton();
            this.btnReviewWeekPlanningAndResources = this.Factory.CreateRibbonButton();
            this.btnReviewWeekPlanning = this.Factory.CreateRibbonButton();
            this.btnUpdateAvaResourceTable = this.Factory.CreateRibbonButton();
            this.btnUpdateReqResourceTable = this.Factory.CreateRibbonButton();
            this.menuForecast = this.Factory.CreateRibbonMenu();
            this.btnForecastFormat = this.Factory.CreateRibbonButton();
            this.btnReviewForcastTask = this.Factory.CreateRibbonButton();
            this.btnForecastReviewRequirements = this.Factory.CreateRibbonButton();
            this.menuWagonsTemperature = this.Factory.CreateRibbonMenu();
            this.btnLoadTempLogPlain = this.Factory.CreateRibbonButton();
            this.btnLoadTempLogMse345 = this.Factory.CreateRibbonButton();
            this.cbTempWagIgnoreLocomotives = this.Factory.CreateRibbonCheckBox();
            this.btnStop = this.Factory.CreateRibbonButton();
            this.menuPlanHistory = this.Factory.CreateRibbonMenu();
            this.btnPlanHistoryFormat = this.Factory.CreateRibbonButton();
            this.btnPlanHistoryLoad = this.Factory.CreateRibbonButton();
            this.btnPlanHistoryReview = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpEllipse);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.box1);
            this.grpEllipse.Items.Add(this.drpEnvironment);
            this.grpEllipse.Items.Add(this.menuActions);
            this.grpEllipse.Label = "Planeacion FFCC";
            this.grpEllipse.Name = "grpEllipse";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "&Formatear";
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
            this.menuActions.Items.Add(this.menuWeeklyPlanning);
            this.menuActions.Items.Add(this.menuForecast);
            this.menuActions.Items.Add(this.menuPlanHistory);
            this.menuActions.Items.Add(this.menuWagonsTemperature);
            this.menuActions.Items.Add(this.btnStop);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuWeeklyPlanning
            // 
            this.menuWeeklyPlanning.Items.Add(this.btnFormatWeekPlanning);
            this.menuWeeklyPlanning.Items.Add(this.btnReviewWeekPlanningAndResources);
            this.menuWeeklyPlanning.Items.Add(this.btnReviewWeekPlanning);
            this.menuWeeklyPlanning.Items.Add(this.btnUpdateAvaResourceTable);
            this.menuWeeklyPlanning.Items.Add(this.btnUpdateReqResourceTable);
            this.menuWeeklyPlanning.Label = "Programación Semanal";
            this.menuWeeklyPlanning.Name = "menuWeeklyPlanning";
            this.menuWeeklyPlanning.ShowImage = true;
            // 
            // btnFormatWeekPlanning
            // 
            this.btnFormatWeekPlanning.Label = "Formatear Programación";
            this.btnFormatWeekPlanning.Name = "btnFormatWeekPlanning";
            this.btnFormatWeekPlanning.ShowImage = true;
            this.btnFormatWeekPlanning.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatWeekPlanning_Click);
            // 
            // btnReviewWeekPlanningAndResources
            // 
            this.btnReviewWeekPlanningAndResources.Label = "Consultar Periodo de Programación y Recursos";
            this.btnReviewWeekPlanningAndResources.Name = "btnReviewWeekPlanningAndResources";
            this.btnReviewWeekPlanningAndResources.ShowImage = true;
            this.btnReviewWeekPlanningAndResources.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewWeekPlanning_Click);
            // 
            // btnReviewWeekPlanning
            // 
            this.btnReviewWeekPlanning.Label = "Consultar Periodo de Programación";
            this.btnReviewWeekPlanning.Name = "btnReviewWeekPlanning";
            this.btnReviewWeekPlanning.ShowImage = true;
            this.btnReviewWeekPlanning.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewWeekPlanning_Click_1);
            // 
            // btnUpdateAvaResourceTable
            // 
            this.btnUpdateAvaResourceTable.Label = "Actualizar Recursos Disponibles";
            this.btnUpdateAvaResourceTable.Name = "btnUpdateAvaResourceTable";
            this.btnUpdateAvaResourceTable.ShowImage = true;
            this.btnUpdateAvaResourceTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateAvaResourceTable_Click);
            // 
            // btnUpdateReqResourceTable
            // 
            this.btnUpdateReqResourceTable.Label = "Actualizar Recursos Programados";
            this.btnUpdateReqResourceTable.Name = "btnUpdateReqResourceTable";
            this.btnUpdateReqResourceTable.ShowImage = true;
            this.btnUpdateReqResourceTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateReqResourceTable_Click);
            // 
            // menuForecast
            // 
            this.menuForecast.Items.Add(this.btnForecastFormat);
            this.menuForecast.Items.Add(this.btnReviewForcastTask);
            this.menuForecast.Items.Add(this.btnForecastReviewRequirements);
            this.menuForecast.Label = "Proyección";
            this.menuForecast.Name = "menuForecast";
            this.menuForecast.ShowImage = true;
            // 
            // btnForecastFormat
            // 
            this.btnForecastFormat.Label = "Formatear Proyeccion";
            this.btnForecastFormat.Name = "btnForecastFormat";
            this.btnForecastFormat.ShowImage = true;
            this.btnForecastFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnForecastFormat_Click);
            // 
            // btnReviewForcastTask
            // 
            this.btnReviewForcastTask.Label = "Consultar Tareas";
            this.btnReviewForcastTask.Name = "btnReviewForcastTask";
            this.btnReviewForcastTask.ShowImage = true;
            this.btnReviewForcastTask.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewForecastTask_Click);
            // 
            // btnForecastReviewRequirements
            // 
            this.btnForecastReviewRequirements.Label = "Consultar Materiales";
            this.btnForecastReviewRequirements.Name = "btnForecastReviewRequirements";
            this.btnForecastReviewRequirements.ShowImage = true;
            this.btnForecastReviewRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnForecastReviewRequirements_Click);
            // 
            // menuWagonsTemperature
            // 
            this.menuWagonsTemperature.Items.Add(this.btnLoadTempLogPlain);
            this.menuWagonsTemperature.Items.Add(this.btnLoadTempLogMse345);
            this.menuWagonsTemperature.Items.Add(this.cbTempWagIgnoreLocomotives);
            this.menuWagonsTemperature.Label = "Temperatura Vagones";
            this.menuWagonsTemperature.Name = "menuWagonsTemperature";
            this.menuWagonsTemperature.ShowImage = true;
            // 
            // btnLoadTempLogPlain
            // 
            this.btnLoadTempLogPlain.Label = "Cargar Log a Formato Plano";
            this.btnLoadTempLogPlain.Name = "btnLoadTempLogPlain";
            this.btnLoadTempLogPlain.ShowImage = true;
            this.btnLoadTempLogPlain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTempLogPlain_Click);
            // 
            // btnLoadTempLogMse345
            // 
            this.btnLoadTempLogMse345.Label = "Transformar Plano a MSE345";
            this.btnLoadTempLogMse345.Name = "btnLoadTempLogMse345";
            this.btnLoadTempLogMse345.ShowImage = true;
            this.btnLoadTempLogMse345.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTempLogMse345_Click);
            // 
            // cbTempWagIgnoreLocomotives
            // 
            this.cbTempWagIgnoreLocomotives.Checked = true;
            this.cbTempWagIgnoreLocomotives.Label = "Ignorar Locomotoras al Transformar";
            this.cbTempWagIgnoreLocomotives.Name = "cbTempWagIgnoreLocomotives";
            this.cbTempWagIgnoreLocomotives.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbTempWagIgnoreLocomotives_Click);
            // 
            // btnStop
            // 
            this.btnStop.Label = "&Detener Procesos";
            this.btnStop.Name = "btnStop";
            this.btnStop.ShowImage = true;
            this.btnStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStop_Click);
            // 
            // menuPlanHistory
            // 
            this.menuPlanHistory.Items.Add(this.btnPlanHistoryFormat);
            this.menuPlanHistory.Items.Add(this.btnPlanHistoryReview);
            this.menuPlanHistory.Items.Add(this.btnPlanHistoryLoad);
            this.menuPlanHistory.Label = "Historia de Programación";
            this.menuPlanHistory.Name = "menuPlanHistory";
            this.menuPlanHistory.ShowImage = true;
            // 
            // btnPlanHistoryFormat
            // 
            this.btnPlanHistoryFormat.Label = "Formatear";
            this.btnPlanHistoryFormat.Name = "btnPlanHistoryFormat";
            this.btnPlanHistoryFormat.ShowImage = true;
            this.btnPlanHistoryFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlanHistoryFormat_Click);
            // 
            // btnPlanHistoryLoad
            // 
            this.btnPlanHistoryLoad.Label = "Cargar";
            this.btnPlanHistoryLoad.Name = "btnPlanHistoryLoad";
            this.btnPlanHistoryLoad.ShowImage = true;
            // 
            // btnPlanHistoryReview
            // 
            this.btnPlanHistoryReview.Label = "Consultar Historia";
            this.btnPlanHistoryReview.Name = "btnPlanHistoryReview";
            this.btnPlanHistoryReview.ShowImage = true;
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuWagonsTemperature;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadTempLogMse345;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadTempLogPlain;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuWeeklyPlanning;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewWeekPlanningAndResources;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateReqResourceTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateAvaResourceTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbTempWagIgnoreLocomotives;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatWeekPlanning;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewWeekPlanning;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuForecast;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnForecastFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnForecastReviewRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewForcastTask;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuPlanHistory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanHistoryFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanHistoryReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanHistoryLoad;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

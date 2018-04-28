namespace EllipsePlannerExcelAddIn
{
    partial class RibbonEllipse : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonEllipse()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpProyecto = this.Factory.CreateRibbonGroup();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewJobs = this.Factory.CreateRibbonButton();
            this.btnLoadData = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpProyecto.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpProyecto);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpProyecto
            // 
            this.grpProyecto.Items.Add(this.menuFormat);
            this.grpProyecto.Items.Add(this.drpEnviroment);
            this.grpProyecto.Items.Add(this.menuActions);
            this.grpProyecto.Items.Add(this.btnAbout);
            this.grpProyecto.Label = "Job Planner";
            this.grpProyecto.Name = "grpProyecto";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatSheet);
            this.menuFormat.Label = "&Formatear";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formato";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReviewJobs);
            this.menuActions.Items.Add(this.btnLoadData);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewJobs
            // 
            this.btnReviewJobs.Label = "Consultar &Información";
            this.btnReviewJobs.Name = "btnReviewJobs";
            this.btnReviewJobs.ShowImage = true;
            this.btnReviewJobs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewJobs_Click);
            // 
            // btnLoadData
            // 
            this.btnLoadData.Label = "Cargar Información";
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.ShowImage = true;
            this.btnLoadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadData_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpProyecto.ResumeLayout(false);
            this.grpProyecto.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProyecto;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadData;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

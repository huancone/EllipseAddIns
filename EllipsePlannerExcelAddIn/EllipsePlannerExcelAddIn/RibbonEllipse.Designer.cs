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
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewJobs = this.Factory.CreateRibbonButton();
            this.btnLoadData = this.Factory.CreateRibbonButton();
            this.btnUpdateEllipse = this.Factory.CreateRibbonButton();
            this.btnUpdateOrder = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.cbDeviationStats = this.Factory.CreateRibbonCheckBox();
            this.cbSplitTaskByResource = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.cbOverlappingDateSearch = this.Factory.CreateRibbonCheckBox();
            this.cbIncludeMsts = this.Factory.CreateRibbonCheckBox();
            this.tabEllipse.SuspendLayout();
            this.grpProyecto.SuspendLayout();
            this.box1.SuspendLayout();
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
            this.grpProyecto.Items.Add(this.box1);
            this.grpProyecto.Items.Add(this.drpEnvironment);
            this.grpProyecto.Items.Add(this.menuActions);
            this.grpProyecto.Label = "Job Planner";
            this.grpProyecto.Name = "grpProyecto";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormatSheet);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formato";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
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
            this.menuActions.Items.Add(this.btnReviewJobs);
            this.menuActions.Items.Add(this.btnLoadData);
            this.menuActions.Items.Add(this.btnUpdateEllipse);
            this.menuActions.Items.Add(this.btnUpdateOrder);
            this.menuActions.Items.Add(this.menu1);
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
            this.btnLoadData.Label = "Cargar Planes";
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.ShowImage = true;
            this.btnLoadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadData_Click);
            // 
            // btnUpdateEllipse
            // 
            this.btnUpdateEllipse.Label = "Actualizar Disponibilidad";
            this.btnUpdateEllipse.Name = "btnUpdateEllipse";
            this.btnUpdateEllipse.ShowImage = true;
            this.btnUpdateEllipse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateEllipse_Click);
            // 
            // btnUpdateOrder
            // 
            this.btnUpdateOrder.Label = "Actualizar Tareas";
            this.btnUpdateOrder.Name = "btnUpdateOrder";
            this.btnUpdateOrder.ShowImage = true;
            this.btnUpdateOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateOrder_Click);
            // 
            // menu1
            // 
            this.menu1.Items.Add(this.cbDeviationStats);
            this.menu1.Items.Add(this.cbSplitTaskByResource);
            this.menu1.Items.Add(this.cbIncludeMsts);
            this.menu1.Items.Add(this.cbOverlappingDateSearch);
            this.menu1.Label = "Opciones";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // cbDeviationStats
            // 
            this.cbDeviationStats.Checked = true;
            this.cbDeviationStats.Label = "Estadísticas de Desviación";
            this.cbDeviationStats.Name = "cbDeviationStats";
            this.cbDeviationStats.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbDeviationStats_Click);
            // 
            // cbSplitTaskByResource
            // 
            this.cbSplitTaskByResource.Checked = true;
            this.cbSplitTaskByResource.Label = "Separar Tareas Por Recurso";
            this.cbSplitTaskByResource.Name = "cbSplitTaskByResource";
            this.cbSplitTaskByResource.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbSplitTaskByResource_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // cbOverlappingDateSearch
            // 
            this.cbOverlappingDateSearch.Label = "Búsqueda Traslapada de Fechas";
            this.cbOverlappingDateSearch.Name = "cbOverlappingDateSearch";
            this.cbOverlappingDateSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbOverlappingDateSearch_Click);
            // 
            // cbIncludeMsts
            // 
            this.cbIncludeMsts.Label = "Incluír Msts";
            this.cbIncludeMsts.Name = "cbIncludeMsts";
            this.cbIncludeMsts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbIncludeMsts_Click);
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
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProyecto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewJobs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDeviationStats;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbSplitTaskByResource;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbOverlappingDateSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIncludeMsts;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

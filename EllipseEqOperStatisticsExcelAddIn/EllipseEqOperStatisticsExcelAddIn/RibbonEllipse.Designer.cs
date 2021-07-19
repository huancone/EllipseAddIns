namespace EllipseEqOperStatisticsExcelAddIn
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
            this.grpEqOperStatistics = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnLoadStatistics = this.Factory.CreateRibbonButton();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnDelete = this.Factory.CreateRibbonButton();
            this.btnRestoreEvents = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEqOperStatistics.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEqOperStatistics);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEqOperStatistics
            // 
            this.grpEqOperStatistics.Items.Add(this.box1);
            this.grpEqOperStatistics.Items.Add(this.drpEnvironment);
            this.grpEqOperStatistics.Items.Add(this.menuActions);
            this.grpEqOperStatistics.Label = "Oper. Statistics";
            this.grpEqOperStatistics.Name = "grpEqOperStatistics";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormatSheet);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear Hoja";
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
            this.menuActions.Items.Add(this.btnLoadStatistics);
            this.menuActions.Items.Add(this.btnReview);
            this.menuActions.Items.Add(this.btnDelete);
            this.menuActions.Items.Add(this.btnRestoreEvents);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnLoadStatistics
            // 
            this.btnLoadStatistics.Label = "Cargar Estadísticas";
            this.btnLoadStatistics.Name = "btnLoadStatistics";
            this.btnLoadStatistics.ShowImage = true;
            this.btnLoadStatistics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadStatistics_Click);
            // 
            // btnReview
            // 
            this.btnReview.Label = "Consultar Estadísticas";
            this.btnReview.Name = "btnReview";
            this.btnReview.ShowImage = true;
            this.btnReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Label = "Eliminar Estadisticas";
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.ShowImage = true;
            this.btnDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelete_Click);
            // 
            // btnRestoreEvents
            // 
            this.btnRestoreEvents.Label = "Restaurar AutoConsultas";
            this.btnRestoreEvents.Name = "btnRestoreEvents";
            this.btnRestoreEvents.ShowImage = true;
            this.btnRestoreEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRestoreEvents_Click);
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
            this.grpEqOperStatistics.ResumeLayout(false);
            this.grpEqOperStatistics.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEqOperStatistics;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadStatistics;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRestoreEvents;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

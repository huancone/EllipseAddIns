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
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.btnLoadStatistics = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEqOperStatistics.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEqOperStatistics);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEqOperStatistics
            // 
            this.grpEqOperStatistics.Items.Add(this.btnFormatSheet);
            this.grpEqOperStatistics.Items.Add(this.drpEnviroment);
            this.grpEqOperStatistics.Items.Add(this.btnLoadStatistics);
            this.grpEqOperStatistics.Label = "Oper. Statistics v1.1.0";
            this.grpEqOperStatistics.Name = "grpEqOperStatistics";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // btnLoadStatistics
            // 
            this.btnLoadStatistics.Label = "Cargar Estadísticas";
            this.btnLoadStatistics.Name = "btnLoadStatistics";
            this.btnLoadStatistics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadStatistics_Click);
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEqOperStatistics;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadStatistics;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

namespace VariacionesExcelAddIn
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
            this.grpVariaciones = this.Factory.CreateRibbonGroup();
            this.btnImportFile = this.Factory.CreateRibbonButton();
            this.drpYear = this.Factory.CreateRibbonDropDown();
            this.drpPeriodo = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.grpVariaciones.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpVariaciones);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpVariaciones
            // 
            this.grpVariaciones.Items.Add(this.btnImportFile);
            this.grpVariaciones.Items.Add(this.drpYear);
            this.grpVariaciones.Items.Add(this.drpPeriodo);
            this.grpVariaciones.Label = "Variaciones";
            this.grpVariaciones.Name = "grpVariaciones";
            // 
            // btnImportFile
            // 
            this.btnImportFile.Label = "Importar Archivo de Cognos";
            this.btnImportFile.Name = "btnImportFile";
            this.btnImportFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportFile_Click);
            // 
            // drpYear
            // 
            this.drpYear.Label = "Año";
            this.drpYear.Name = "drpYear";
            // 
            // drpPeriodo
            // 
            this.drpPeriodo.Label = "Mes";
            this.drpPeriodo.Name = "drpPeriodo";
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpVariaciones.ResumeLayout(false);
            this.grpVariaciones.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpVariaciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpYear;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpPeriodo;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

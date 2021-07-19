namespace EllipseConsolidacionesExcelAddIn
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
            this.gpConsolidaciones = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnConsolidations = this.Factory.CreateRibbonButton();
            this.btnServiceCategory = this.Factory.CreateRibbonButton();
            this.btnClean = this.Factory.CreateRibbonButton();
            this.btnStop = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.gpConsolidaciones.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.gpConsolidaciones);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // gpConsolidaciones
            // 
            this.gpConsolidaciones.Items.Add(this.box1);
            this.gpConsolidaciones.Items.Add(this.drpEnvironment);
            this.gpConsolidaciones.Items.Add(this.menuActions);
            this.gpConsolidaciones.Label = "Consolidaciones";
            this.gpConsolidaciones.Name = "gpConsolidaciones";
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
            this.menuActions.Items.Add(this.btnConsolidations);
            this.menuActions.Items.Add(this.btnServiceCategory);
            this.menuActions.Items.Add(this.btnClean);
            this.menuActions.Items.Add(this.btnStop);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnConsolidations
            // 
            this.btnConsolidations.Label = "Procesar Consolidación";
            this.btnConsolidations.Name = "btnConsolidations";
            this.btnConsolidations.ShowImage = true;
            this.btnConsolidations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsolidations_Click);
            // 
            // btnServiceCategory
            // 
            this.btnServiceCategory.Label = "Procesar Service Category";
            this.btnServiceCategory.Name = "btnServiceCategory";
            this.btnServiceCategory.ShowImage = true;
            this.btnServiceCategory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnServiceCategory_Click);
            // 
            // btnClean
            // 
            this.btnClean.Label = "Limpiar Hoja";
            this.btnClean.Name = "btnClean";
            this.btnClean.ShowImage = true;
            this.btnClean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClean_Click);
            // 
            // btnStop
            // 
            this.btnStop.Label = "Detener Proceso";
            this.btnStop.Name = "btnStop";
            this.btnStop.ShowImage = true;
            this.btnStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStop_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.gpConsolidaciones.ResumeLayout(false);
            this.gpConsolidaciones.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpConsolidaciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsolidations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnServiceCategory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStop;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

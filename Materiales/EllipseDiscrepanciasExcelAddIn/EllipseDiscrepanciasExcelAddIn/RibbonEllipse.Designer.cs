namespace EllipseDiscrepanciasExcelAddIn
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
            this.grpDiscrepancias = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.mAcciones = this.Factory.CreateRibbonMenu();
            this.bProcesar = this.Factory.CreateRibbonButton();
            this.bProcesarMSE1TD = this.Factory.CreateRibbonButton();
            this.bLimpiar = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.tabEllipse.SuspendLayout();
            this.grpDiscrepancias.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpDiscrepancias);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpDiscrepancias
            // 
            this.grpDiscrepancias.Items.Add(this.box1);
            this.grpDiscrepancias.Items.Add(this.drpEnvironment);
            this.grpDiscrepancias.Items.Add(this.mAcciones);
            this.grpDiscrepancias.Label = "Discrepancias";
            this.grpDiscrepancias.Name = "grpDiscrepancias";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // mAcciones
            // 
            this.mAcciones.Items.Add(this.bProcesar);
            this.mAcciones.Items.Add(this.bProcesarMSE1TD);
            this.mAcciones.Items.Add(this.bLimpiar);
            this.mAcciones.Label = "Acciones";
            this.mAcciones.Name = "mAcciones";
            // 
            // bProcesar
            // 
            this.bProcesar.Label = "Procesar MSE1SF-1SX";
            this.bProcesar.Name = "bProcesar";
            this.bProcesar.ShowImage = true;
            this.bProcesar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bProcesar_Click);
            // 
            // bProcesarMSE1TD
            // 
            this.bProcesarMSE1TD.Label = "Procesar MSE1TD";
            this.bProcesarMSE1TD.Name = "bProcesarMSE1TD";
            this.bProcesarMSE1TD.ShowImage = true;
            this.bProcesarMSE1TD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bProcesarMSE1TD_Click);
            // 
            // bLimpiar
            // 
            this.bLimpiar.Label = "Limpiar";
            this.bLimpiar.Name = "bLimpiar";
            this.bLimpiar.ShowImage = true;
            this.bLimpiar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bLimpiar_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Formatear";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpDiscrepancias.ResumeLayout(false);
            this.grpDiscrepancias.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDiscrepancias;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bProcesar;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bProcesarMSE1TD;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
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

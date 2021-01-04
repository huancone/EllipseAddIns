namespace EllipseAddInInfoPm
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
            this.grpProyecto = this.Factory.CreateRibbonGroup();
            this.btnFormatear = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuAcciones = this.Factory.CreateRibbonMenu();
            this.btnConsultar = this.Factory.CreateRibbonButton();
            this.btnLimpiar = this.Factory.CreateRibbonButton();
            this.Stop = this.Factory.CreateRibbonButton();
            this.btnProy = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnPrueba = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpProyecto.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpProyecto);
            this.tab1.Label = "Ellipse 9";
            this.tab1.Name = "tab1";
            // 
            // grpProyecto
            // 
            this.grpProyecto.Items.Add(this.btnFormatear);
            this.grpProyecto.Items.Add(this.drpEnviroment);
            this.grpProyecto.Items.Add(this.menuAcciones);
            this.grpProyecto.Items.Add(this.btnAbout);
            this.grpProyecto.Label = "Info PM";
            this.grpProyecto.Name = "grpProyecto";
            // 
            // btnFormatear
            // 
            this.btnFormatear.Label = "Formatear";
            this.btnFormatear.Name = "btnFormatear";
            this.btnFormatear.ShowImage = true;
            this.btnFormatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatear_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuAcciones
            // 
            this.menuAcciones.Items.Add(this.btnConsultar);
            this.menuAcciones.Items.Add(this.btnLimpiar);
            this.menuAcciones.Items.Add(this.Stop);
            this.menuAcciones.Items.Add(this.btnProy);
            this.menuAcciones.Label = "Acciones";
            this.menuAcciones.Name = "menuAcciones";
            // 
            // btnConsultar
            // 
            this.btnConsultar.Label = "Consultar PM";
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.ShowImage = true;
            this.btnConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultar_Click);
            // 
            // btnLimpiar
            // 
            this.btnLimpiar.Label = "Limpiar";
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.ShowImage = true;
            this.btnLimpiar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLimpiar_Click);
            // 
            // Stop
            // 
            this.Stop.Label = "Detener";
            this.Stop.Name = "Stop";
            this.Stop.ShowImage = true;
            this.Stop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Stop_Click);
            // 
            // btnProy
            // 
            this.btnProy.Label = "Generar Proy";
            this.btnProy.Name = "btnProy";
            this.btnProy.ShowImage = true;
            this.btnProy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProy_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // btnPrueba
            // 
            this.btnPrueba.Label = "Prueba";
            this.btnPrueba.Name = "btnPrueba";
            this.btnPrueba.ShowImage = true;
            this.btnPrueba.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Prueba_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpProyecto.ResumeLayout(false);
            this.grpProyecto.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProyecto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Stop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrueba;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

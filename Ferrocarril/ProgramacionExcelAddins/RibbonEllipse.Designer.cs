namespace ProgramacionExcelAddins
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
            this.grpProyecto = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuAcciones = this.Factory.CreateRibbonMenu();
            this.menuHistoria = this.Factory.CreateRibbonMenu();
            this.btnConsultar = this.Factory.CreateRibbonButton();
            this.btnCargar = this.Factory.CreateRibbonButton();
            this.btnEliminar = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpProyecto.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpProyecto);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpProyecto
            // 
            this.grpProyecto.Items.Add(this.box1);
            this.grpProyecto.Items.Add(this.drpEnvironment);
            this.grpProyecto.Items.Add(this.menuAcciones);
            this.grpProyecto.Label = "Programacion";
            this.grpProyecto.Name = "grpProyecto";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Formatear";
            this.btnFormat.Name = "btnFormat";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuAcciones
            // 
            this.menuAcciones.Items.Add(this.menuHistoria);
            this.menuAcciones.Label = "Acciones";
            this.menuAcciones.Name = "menuAcciones";
            // 
            // menuHistoria
            // 
            this.menuHistoria.Items.Add(this.btnConsultar);
            this.menuHistoria.Items.Add(this.btnCargar);
            this.menuHistoria.Items.Add(this.btnEliminar);
            this.menuHistoria.Label = "Historia";
            this.menuHistoria.Name = "menuHistoria";
            this.menuHistoria.ShowImage = true;
            // 
            // btnConsultar
            // 
            this.btnConsultar.Label = "Consultar";
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.ShowImage = true;
            this.btnConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultar_Click);
            // 
            // btnCargar
            // 
            this.btnCargar.Label = "Cargar";
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.ShowImage = true;
            this.btnCargar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCargar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Label = "Eliminar";
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.ShowImage = true;
            this.btnEliminar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEliminar_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuHistoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCargar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEliminar;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

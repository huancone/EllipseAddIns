namespace EllipseAddinGanttEQ
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.Pruebas = this.Factory.CreateRibbonButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpProyecto = this.Factory.CreateRibbonGroup();
            this.btnFormatear = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuAcciones = this.Factory.CreateRibbonMenu();
            this.btnConsultar = this.Factory.CreateRibbonButton();
            this.btnConsultarLab = this.Factory.CreateRibbonButton();
            this.btnConsultDur = this.Factory.CreateRibbonButton();
            this.btnConsultVale = this.Factory.CreateRibbonButton();
            this.btnActualizarGantt = this.Factory.CreateRibbonButton();
            this.btnActualizaDatos = this.Factory.CreateRibbonButton();
            this.btnActualizarDurLab = this.Factory.CreateRibbonButton();
            this.IdCumplimiento = this.Factory.CreateRibbonButton();
            this.btnLimpiar = this.Factory.CreateRibbonButton();
            this.Stop = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpProyecto.SuspendLayout();
            this.SuspendLayout();
            // 
            // Pruebas
            // 
            this.Pruebas.Label = "Pruebas";
            this.Pruebas.Name = "Pruebas";
            this.Pruebas.ShowImage = true;
            this.Pruebas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Pruebas_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpProyecto);
            this.tab1.Label = "Ellipse 8";
            this.tab1.Name = "tab1";
            // 
            // grpProyecto
            // 
            this.grpProyecto.DialogLauncher = ribbonDialogLauncherImpl1;
            this.grpProyecto.Items.Add(this.btnFormatear);
            this.grpProyecto.Items.Add(this.drpEnvironment);
            this.grpProyecto.Items.Add(this.menuAcciones);
            this.grpProyecto.Items.Add(this.btnAbout);
            this.grpProyecto.Label = "Gantt de Parada";
            this.grpProyecto.Name = "grpProyecto";
            // 
            // btnFormatear
            // 
            this.btnFormatear.Label = "Formatear";
            this.btnFormatear.Name = "btnFormatear";
            this.btnFormatear.ShowImage = true;
            this.btnFormatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuAcciones
            // 
            this.menuAcciones.Items.Add(this.btnConsultar);
            this.menuAcciones.Items.Add(this.btnConsultarLab);
            this.menuAcciones.Items.Add(this.btnConsultDur);
            this.menuAcciones.Items.Add(this.btnConsultVale);
            this.menuAcciones.Items.Add(this.btnActualizarGantt);
            this.menuAcciones.Items.Add(this.btnActualizaDatos);
            this.menuAcciones.Items.Add(this.btnActualizarDurLab);
            this.menuAcciones.Items.Add(this.IdCumplimiento);
            this.menuAcciones.Items.Add(this.btnLimpiar);
            this.menuAcciones.Items.Add(this.Stop);
            this.menuAcciones.Label = "Acciones";
            this.menuAcciones.Name = "menuAcciones";
            // 
            // btnConsultar
            // 
            this.btnConsultar.Label = "Consultar OT";
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.ShowImage = true;
            this.btnConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecution_Click);
            // 
            // btnConsultarLab
            // 
            this.btnConsultarLab.Label = "Consultar Labor";
            this.btnConsultarLab.Name = "btnConsultarLab";
            this.btnConsultarLab.ShowImage = true;
            this.btnConsultarLab.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultarLab_Click);
            // 
            // btnConsultDur
            // 
            this.btnConsultDur.Label = "Consultar Duracion";
            this.btnConsultDur.Name = "btnConsultDur";
            this.btnConsultDur.ShowImage = true;
            this.btnConsultDur.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultDur_Click);
            // 
            // btnConsultVale
            // 
            this.btnConsultVale.Label = "Consultar Vales";
            this.btnConsultVale.Name = "btnConsultVale";
            this.btnConsultVale.ShowImage = true;
            this.btnConsultVale.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultVale_Click);
            // 
            // btnActualizarGantt
            // 
            this.btnActualizarGantt.Label = "Actualizar Por Gantt";
            this.btnActualizarGantt.Name = "btnActualizarGantt";
            this.btnActualizarGantt.ShowImage = true;
            this.btnActualizarGantt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActualizarGantt_Click);
            // 
            // btnActualizaDatos
            // 
            this.btnActualizaDatos.Label = "Actualizar Por Datos";
            this.btnActualizaDatos.Name = "btnActualizaDatos";
            this.btnActualizaDatos.ShowImage = true;
            this.btnActualizaDatos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActualizarDatos_Click);
            // 
            // btnActualizarDurLab
            // 
            this.btnActualizarDurLab.Label = "Acciones Labor";
            this.btnActualizarDurLab.Name = "btnActualizarDurLab";
            this.btnActualizarDurLab.ShowImage = true;
            this.btnActualizarDurLab.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActualizarDurLab_Click);
            // 
            // IdCumplimiento
            // 
            this.IdCumplimiento.Label = "Seguimiento";
            this.IdCumplimiento.Name = "IdCumplimiento";
            this.IdCumplimiento.ShowImage = true;
            this.IdCumplimiento.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IdCumplimiento_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActualizarDurLab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActualizarGantt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActualizaDatos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultarLab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton IdCumplimiento;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultDur;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Stop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Pruebas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultVale;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

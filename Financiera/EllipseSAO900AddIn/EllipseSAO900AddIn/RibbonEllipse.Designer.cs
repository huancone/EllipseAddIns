namespace EllipseSAO900AddIn
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
        /// <param name="disposing">true si los recursos administrados se deben desechar; en caso contrario, false.</param>
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
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpSAO900 = this.Factory.CreateRibbonGroup();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatoReclasificaciones = this.Factory.CreateRibbonButton();
            this.btnFormatoModificaciones = this.Factory.CreateRibbonButton();
            this.btnFormatoCausaciones = this.Factory.CreateRibbonButton();
            this.btnFormatoDistribuciones = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuAction = this.Factory.CreateRibbonMenu();
            this.btnValidar = this.Factory.CreateRibbonButton();
            this.btnExportar = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpSAO900.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpSAO900);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpSAO900
            // 
            this.grpSAO900.Items.Add(this.box1);
            this.grpSAO900.Items.Add(this.drpEnvironment);
            this.grpSAO900.Items.Add(this.menuAction);
            this.grpSAO900.Label = "SAO900";
            this.grpSAO900.Name = "grpSAO900";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatoReclasificaciones);
            this.menuFormat.Items.Add(this.btnFormatoModificaciones);
            this.menuFormat.Items.Add(this.btnFormatoCausaciones);
            this.menuFormat.Items.Add(this.btnFormatoDistribuciones);
            this.menuFormat.Label = "&Formatos";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatoReclasificaciones
            // 
            this.btnFormatoReclasificaciones.Label = "&Reclasificaciones";
            this.btnFormatoReclasificaciones.Name = "btnFormatoReclasificaciones";
            this.btnFormatoReclasificaciones.ShowImage = true;
            this.btnFormatoReclasificaciones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatoReclasificaciones_Click);
            // 
            // btnFormatoModificaciones
            // 
            this.btnFormatoModificaciones.Label = "&Modificaciones";
            this.btnFormatoModificaciones.Name = "btnFormatoModificaciones";
            this.btnFormatoModificaciones.ShowImage = true;
            this.btnFormatoModificaciones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatoModificaciones_Click);
            // 
            // btnFormatoCausaciones
            // 
            this.btnFormatoCausaciones.Label = "&Causaciones";
            this.btnFormatoCausaciones.Name = "btnFormatoCausaciones";
            this.btnFormatoCausaciones.ShowImage = true;
            this.btnFormatoCausaciones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatoCausaciones_Click);
            // 
            // btnFormatoDistribuciones
            // 
            this.btnFormatoDistribuciones.Label = "&Distribuciones";
            this.btnFormatoDistribuciones.Name = "btnFormatoDistribuciones";
            this.btnFormatoDistribuciones.ShowImage = true;
            this.btnFormatoDistribuciones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatoDistribuciones_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuAction
            // 
            this.menuAction.Items.Add(this.btnValidar);
            this.menuAction.Items.Add(this.btnExportar);
            this.menuAction.Label = "&Acciones";
            this.menuAction.Name = "menuAction";
            // 
            // btnValidar
            // 
            this.btnValidar.Label = "&Validar";
            this.btnValidar.Name = "btnValidar";
            this.btnValidar.ShowImage = true;
            this.btnValidar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidar_Click);
            // 
            // btnExportar
            // 
            this.btnExportar.Label = "&Exportar";
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.ShowImage = true;
            this.btnExportar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportar_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
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
            this.grpSAO900.ResumeLayout(false);
            this.grpSAO900.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSAO900;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatoReclasificaciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatoModificaciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatoCausaciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatoDistribuciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportar;
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

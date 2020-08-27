namespace EllipseMonitoreoExcelAddIn
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
            this.components = new System.ComponentModel.Container();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.imageList2 = new System.Windows.Forms.ImageList(this.components);
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.Monitoreo = this.Factory.CreateRibbonGroup();
            this.menuHoja = this.Factory.CreateRibbonMenu();
            this.formatear = this.Factory.CreateRibbonButton();
            this.buttonLimpiar = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuAcciones = this.Factory.CreateRibbonMenu();
            this.cargar = this.Factory.CreateRibbonButton();
            this.borrar = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.Monitoreo.SuspendLayout();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // imageList2
            // 
            this.imageList2.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList2.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.Monitoreo);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // Monitoreo
            // 
            this.Monitoreo.Items.Add(this.menuHoja);
            this.Monitoreo.Items.Add(this.drpEnviroment);
            this.Monitoreo.Items.Add(this.menuAcciones);
            this.Monitoreo.Items.Add(this.btnAbout);
            this.Monitoreo.Label = "Monitoreo V3.0";
            this.Monitoreo.Name = "Monitoreo";
            // 
            // menuHoja
            // 
            this.menuHoja.Items.Add(this.formatear);
            this.menuHoja.Items.Add(this.buttonLimpiar);
            this.menuHoja.Label = "Hoja";
            this.menuHoja.Name = "menuHoja";
            // 
            // formatear
            // 
            this.formatear.Label = "Formatear";
            this.formatear.Name = "formatear";
            this.formatear.ShowImage = true;
            this.formatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatear_Click);
            // 
            // buttonLimpiar
            // 
            this.buttonLimpiar.Label = "Limpiar";
            this.buttonLimpiar.Name = "buttonLimpiar";
            this.buttonLimpiar.ShowImage = true;
            this.buttonLimpiar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLimpiar_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            this.drpEnviroment.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drpEnviroment_SelectionChanged);
            // 
            // menuAcciones
            // 
            this.menuAcciones.Items.Add(this.cargar);
            this.menuAcciones.Items.Add(this.borrar);
            this.menuAcciones.Label = "Acciones";
            this.menuAcciones.Name = "menuAcciones";
            // 
            // cargar
            // 
            this.cargar.Label = "Cargar";
            this.cargar.Name = "cargar";
            this.cargar.ShowImage = true;
            this.cargar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cargar_Click);
            // 
            // borrar
            // 
            this.borrar.Label = "Borrar";
            this.borrar.Name = "borrar";
            this.borrar.ShowImage = true;
            this.borrar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.borrar_Click);
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
            this.Monitoreo.ResumeLayout(false);
            this.Monitoreo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Monitoreo;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ImageList imageList2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuHoja;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cargar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton borrar;
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

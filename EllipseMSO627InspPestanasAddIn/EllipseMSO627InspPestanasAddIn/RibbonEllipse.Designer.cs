namespace EllipseMSO627InspPestanasAddIn
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
            this.grpInspPestanas = this.Factory.CreateRibbonGroup();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnLimpiezaCarbon = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpInspPestanas.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpInspPestanas);
            this.tabEllipse.Label = "ELLIPSE 8 ";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpInspPestanas
            // 
            this.grpInspPestanas.Items.Add(this.box1);
            this.grpInspPestanas.Items.Add(this.drpEnviroment);
            this.grpInspPestanas.Items.Add(this.menuActions);
            this.grpInspPestanas.Label = "MSO627 FFCC";
            this.grpInspPestanas.Name = "grpInspPestanas";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormat);
            this.menuFormat.Items.Add(this.btnLimpiezaCarbon);
            this.menuFormat.Label = "Formato";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Inspeccion Pestañas";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.ShowImage = true;
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // btnLimpiezaCarbon
            // 
            this.btnLimpiezaCarbon.Label = "Limpieza de Carbon";
            this.btnLimpiezaCarbon.Name = "btnLimpiezaCarbon";
            this.btnLimpiezaCarbon.ShowImage = true;
            this.btnLimpiezaCarbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLimpiezaCarbon_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnLoad);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Cargar hoja";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.ShowImage = true;
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener Procesos";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
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
            this.grpInspPestanas.ResumeLayout(false);
            this.grpInspPestanas.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInspPestanas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLimpiezaCarbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
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

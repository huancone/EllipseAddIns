namespace EllipseAddinManejoInfoMCL
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnFormatear = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuAcciones = this.Factory.CreateRibbonMenu();
            this.btnConsultar = this.Factory.CreateRibbonButton();
            this.BtnAcciones = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.bLimpiar = this.Factory.CreateRibbonButton();
            this.btnRestoreEvents = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "AddInMC&L";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnFormatear);
            this.group1.Items.Add(this.drpEnvironment);
            this.group1.Items.Add(this.menuAcciones);
            this.group1.Items.Add(this.btnAbout);
            this.group1.Label = "Menu";
            this.group1.Name = "group1";
            // 
            // btnFormatear
            // 
            this.btnFormatear.Label = "Formatear";
            this.btnFormatear.Name = "btnFormatear";
            this.btnFormatear.ShowImage = true;
            this.btnFormatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatear_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuAcciones
            // 
            this.menuAcciones.Items.Add(this.btnConsultar);
            this.menuAcciones.Items.Add(this.BtnAcciones);
            this.menuAcciones.Items.Add(this.btnStopThread);
            this.menuAcciones.Items.Add(this.bLimpiar);
            this.menuAcciones.Items.Add(this.btnRestoreEvents);
            this.menuAcciones.Label = "Acciones";
            this.menuAcciones.Name = "menuAcciones";
            // 
            // btnConsultar
            // 
            this.btnConsultar.Label = "Consultar";
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.ShowImage = true;
            this.btnConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultar_Click);
            // 
            // BtnAcciones
            // 
            this.BtnAcciones.Label = "Acciones";
            this.BtnAcciones.Name = "BtnAcciones";
            this.BtnAcciones.ShowImage = true;
            this.BtnAcciones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAcciones_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            // 
            // bLimpiar
            // 
            this.bLimpiar.Label = "Limpiar";
            this.bLimpiar.Name = "bLimpiar";
            this.bLimpiar.ShowImage = true;
            this.bLimpiar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bLimpiar_Click);
            // 
            // btnRestoreEvents
            // 
            this.btnRestoreEvents.Label = "Restaurar AutoConsultas";
            this.btnRestoreEvents.Name = "btnRestoreEvents";
            this.btnRestoreEvents.ShowImage = true;
            this.btnRestoreEvents.Visible = false;
            this.btnRestoreEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRestoreEvents_Click);
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRestoreEvents;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

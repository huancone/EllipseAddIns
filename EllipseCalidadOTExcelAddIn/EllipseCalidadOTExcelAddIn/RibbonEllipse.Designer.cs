namespace EllipseCalidadOTExcelAddIn
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
            this.grpCalidadOt = this.Factory.CreateRibbonGroup();
            this.bFormatear = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.bConsultar = this.Factory.CreateRibbonButton();
            this.btnConsulta2 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.bCalificar = this.Factory.CreateRibbonButton();
            this.bLimpiar = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpCalidadOt.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpCalidadOt);
            this.tabEllipse.Label = "ELLIPSE8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpCalidadOt
            // 
            this.grpCalidadOt.Items.Add(this.bFormatear);
            this.grpCalidadOt.Items.Add(this.drpEnvironment);
            this.grpCalidadOt.Items.Add(this.menuActions);
            this.grpCalidadOt.Items.Add(this.btnAbout);
            this.grpCalidadOt.Label = "Calidad OT";
            this.grpCalidadOt.Name = "grpCalidadOt";
            // 
            // bFormatear
            // 
            this.bFormatear.Label = "Formatear";
            this.bFormatear.Name = "bFormatear";
            this.bFormatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bFormatear_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            this.drpEnvironment.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drpEnviroment_SelectionChanged);
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.bConsultar);
            this.menuActions.Items.Add(this.btnConsulta2);
            this.menuActions.Items.Add(this.button1);
            this.menuActions.Items.Add(this.bCalificar);
            this.menuActions.Items.Add(this.bLimpiar);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // bConsultar
            // 
            this.bConsultar.Label = "Consultar";
            this.bConsultar.Name = "bConsultar";
            this.bConsultar.ShowImage = true;
            this.bConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bConsultar_Click);
            // 
            // btnConsulta2
            // 
            this.btnConsulta2.Label = "Consultar servicios contratados";
            this.btnConsulta2.Name = "btnConsulta2";
            this.btnConsulta2.ShowImage = true;
            this.btnConsulta2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsulta2_Click);
            // 
            // button1
            // 
            this.button1.Label = "Re consultar";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // bCalificar
            // 
            this.bCalificar.Label = "Calificar";
            this.bCalificar.Name = "bCalificar";
            this.bCalificar.ShowImage = true;
            this.bCalificar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bCalificar_Click);
            // 
            // bLimpiar
            // 
            this.bLimpiar.Label = "Limpiar Formato";
            this.bLimpiar.Name = "bLimpiar";
            this.bLimpiar.ShowImage = true;
            this.bLimpiar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bLimpiar_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener &Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click_1);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpCalidadOt.ResumeLayout(false);
            this.grpCalidadOt.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCalidadOt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bFormatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bCalificar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bLimpiar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsulta2;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

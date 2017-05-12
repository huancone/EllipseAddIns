namespace MSO627BombasExcelAddIn
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
            this.grpEllipse = this.Factory.CreateRibbonGroup();
            this.btnFormato = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnCargar = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEllipse);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.box1);
            this.grpEllipse.Items.Add(this.drpEnviroment);
            this.grpEllipse.Items.Add(this.menuActions);
            this.grpEllipse.Label = "MSO627 Bombas";
            this.grpEllipse.Name = "grpEllipse";
            // 
            // btnFormato
            // 
            this.btnFormato.Label = "&Formato";
            this.btnFormato.Name = "btnFormato";
            this.btnFormato.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormato_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "&Env. ";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnCargar);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnCargar
            // 
            this.btnCargar.Label = "&Cargar Datos";
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.ShowImage = true;
            this.btnCargar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCargar_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormato);
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
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormato;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCargar;
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

namespace EllipseMSO210ExcelAddin
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
            this.grpMSO210 = this.Factory.CreateRibbonGroup();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO210.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO210);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO210
            // 
            this.grpMSO210.Items.Add(this.box1);
            this.grpMSO210.Items.Add(this.drpEnviroment);
            this.grpMSO210.Items.Add(this.btnLoad);
            this.grpMSO210.Label = "MSO210";
            this.grpMSO210.Name = "grpMSO210";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Format Sheet";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env. ";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Load Sheet";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
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
            this.grpMSO210.ResumeLayout(false);
            this.grpMSO210.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO210;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
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

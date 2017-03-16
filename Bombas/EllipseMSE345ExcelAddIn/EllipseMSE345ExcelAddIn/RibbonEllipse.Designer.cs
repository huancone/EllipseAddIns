namespace EllipseMSE345ExcelAddIn
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpCondMonit = this.Factory.CreateRibbonGroup();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.Crear = this.Factory.CreateRibbonButton();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnStopProcess = this.Factory.CreateRibbonButton();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpCondMonit.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpCondMonit);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpCondMonit
            // 
            this.grpCondMonit.Items.Add(this.btnFormat);
            this.grpCondMonit.Items.Add(this.drpEnviroment);
            this.grpCondMonit.Items.Add(this.menuActions);
            this.grpCondMonit.Label = "MSE345 v 1.0.1";
            this.grpCondMonit.Name = "grpCondMonit";
            // 
            // drpEnviroment
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Test";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl1);
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl2);
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl3);
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // Crear
            // 
            this.Crear.Label = "Cargar Info";
            this.Crear.Name = "Crear";
            this.Crear.ShowImage = true;
            this.Crear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Crear_Click);
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.Crear);
            this.menuActions.Items.Add(this.btnStopProcess);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnStopProcess
            // 
            this.btnStopProcess.Label = "&Detener Proceso";
            this.btnStopProcess.Name = "btnStopProcess";
            this.btnStopProcess.ShowImage = true;
            this.btnStopProcess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopProcess_Click);
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Formato";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpCondMonit.ResumeLayout(false);
            this.grpCondMonit.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCondMonit;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Crear;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopProcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

namespace EllipseMSO905ExcelAddIn
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
            this.grpMSO905 = this.Factory.CreateRibbonGroup();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.mnActions = this.Factory.CreateRibbonMenu();
            this.btnValidate = this.Factory.CreateRibbonButton();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO905.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO905);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO905
            // 
            this.grpMSO905.Items.Add(this.btnFormatSheet);
            this.grpMSO905.Items.Add(this.drpEnviroment);
            this.grpMSO905.Items.Add(this.mnActions);
            this.grpMSO905.Label = "MSO905 v1.0.0";
            this.grpMSO905.Name = "grpMSO905";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Format Sheet";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // mnActions
            // 
            this.mnActions.Items.Add(this.btnValidate);
            this.mnActions.Items.Add(this.btnLoad);
            this.mnActions.Label = "Actions";
            this.mnActions.Name = "mnActions";
            // 
            // btnValidate
            // 
            this.btnValidate.Label = "Validate";
            this.btnValidate.Name = "btnValidate";
            this.btnValidate.ShowImage = true;
            this.btnValidate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidate_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Load Sheet";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.ShowImage = true;
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpMSO905.ResumeLayout(false);
            this.grpMSO905.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO905;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

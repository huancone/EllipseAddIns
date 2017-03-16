namespace EllipseMSO685ExcelAddIn
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
            this.grpMSO685 = this.Factory.CreateRibbonGroup();
            this.menuFormats = this.Factory.CreateRibbonMenu();
            this.btnFormatSubAssetsDep = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnAccion3 = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO685.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO685);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO685
            // 
            this.grpMSO685.Items.Add(this.menuFormats);
            this.grpMSO685.Items.Add(this.drpEnviroment);
            this.grpMSO685.Items.Add(this.menuActions);
            this.grpMSO685.Label = "group1";
            this.grpMSO685.Name = "grpMSO685";
            // 
            // menuFormats
            // 
            this.menuFormats.Items.Add(this.btnFormatSubAssetsDep);
            this.menuFormats.Label = "Formatos";
            this.menuFormats.Name = "menuFormats";
            // 
            // btnFormatSubAssetsDep
            // 
            this.btnFormatSubAssetsDep.Label = "Maintain Sub-Asset Depreciation";
            this.btnFormatSubAssetsDep.Name = "btnFormatSubAssetsDep";
            this.btnFormatSubAssetsDep.ShowImage = true;
            this.btnFormatSubAssetsDep.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSubAssetsDep_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnAccion3);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnAccion3
            // 
            this.btnAccion3.Label = "Maintain Sub-Asset Depreciation";
            this.btnAccion3.Name = "btnAccion3";
            this.btnAccion3.ShowImage = true;
            this.btnAccion3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAccion3_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpMSO685.ResumeLayout(false);
            this.grpMSO685.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO685;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormats;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSubAssetsDep;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAccion3;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

namespace EllipseMSO200ExcelAddIn
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
            this.grpMSO200 = this.Factory.CreateRibbonGroup();
            this.menuFormats = this.Factory.CreateRibbonMenu();
            this.btnChangeAccounts = this.Factory.CreateRibbonButton();
            this.btnFormatInactiveBusiness = this.Factory.CreateRibbonButton();
            this.btnInactivateSupplier = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.ValidateAccounts = this.Factory.CreateRibbonButton();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO200.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO200);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO200
            // 
            this.grpMSO200.Items.Add(this.box1);
            this.grpMSO200.Items.Add(this.menuFormats);
            this.grpMSO200.Items.Add(this.drpEnviroment);
            this.grpMSO200.Label = "MSO200";
            this.grpMSO200.Name = "grpMSO200";
            // 
            // menuFormats
            // 
            this.menuFormats.Items.Add(this.btnChangeAccounts);
            this.menuFormats.Items.Add(this.btnFormatInactiveBusiness);
            this.menuFormats.Items.Add(this.btnInactivateSupplier);
            this.menuFormats.Label = "Formatos";
            this.menuFormats.Name = "menuFormats";
            // 
            // btnChangeAccounts
            // 
            this.btnChangeAccounts.Label = "Formato Cambio Cuentas";
            this.btnChangeAccounts.Name = "btnChangeAccounts";
            this.btnChangeAccounts.ShowImage = true;
            this.btnChangeAccounts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeAccounts_Click);
            // 
            // btnFormatInactiveBusiness
            // 
            this.btnFormatInactiveBusiness.Label = "Formato Inactivar Supplier Business";
            this.btnFormatInactiveBusiness.Name = "btnFormatInactiveBusiness";
            this.btnFormatInactiveBusiness.ShowImage = true;
            this.btnFormatInactiveBusiness.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatInactiveBusiness_Click);
            // 
            // btnInactivateSupplier
            // 
            this.btnInactivateSupplier.Label = "Formato Inactivar Supplier";
            this.btnInactivateSupplier.Name = "btnInactivateSupplier";
            this.btnInactivateSupplier.ShowImage = true;
            this.btnInactivateSupplier.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInactivateSupplier_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.ValidateAccounts);
            this.menuActions.Items.Add(this.btnLoad);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // ValidateAccounts
            // 
            this.ValidateAccounts.Label = "Validar Cuentas";
            this.ValidateAccounts.Name = "ValidateAccounts";
            this.ValidateAccounts.ShowImage = true;
            this.ValidateAccounts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ValidateAccounts_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Cargar a Ellipse";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.ShowImage = true;
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuActions);
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
            this.grpMSO200.ResumeLayout(false);
            this.grpMSO200.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO200;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormats;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeAccounts;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ValidateAccounts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatInactivateSupplierBusiness;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatInactiveBusiness;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInactivateSupplier;
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

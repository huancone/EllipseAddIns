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
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormats = this.Factory.CreateRibbonMenu();
            this.btnChangeAccounts = this.Factory.CreateRibbonButton();
            this.btnInactivateSupplier = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnValidateAccounts = this.Factory.CreateRibbonButton();
            this.btnLoadAccounts = this.Factory.CreateRibbonButton();
            this.btnInactivateBussiness = this.Factory.CreateRibbonButton();
            this.btnInactivareAddress = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnSuspender = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO200.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO200);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO200
            // 
            this.grpMSO200.Items.Add(this.box1);
            this.grpMSO200.Items.Add(this.drpEnvironment);
            this.grpMSO200.Items.Add(this.menuActions);
            this.grpMSO200.Label = "MSO200 ";
            this.grpMSO200.Name = "grpMSO200";
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormats);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // menuFormats
            // 
            this.menuFormats.Items.Add(this.btnChangeAccounts);
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
            // btnInactivateSupplier
            // 
            this.btnInactivateSupplier.Label = "Formato Inactivar Supplier";
            this.btnInactivateSupplier.Name = "btnInactivateSupplier";
            this.btnInactivateSupplier.ShowImage = true;
            this.btnInactivateSupplier.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInactivateSupplier_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnValidateAccounts);
            this.menuActions.Items.Add(this.btnLoadAccounts);
            this.menuActions.Items.Add(this.btnInactivateBussiness);
            this.menuActions.Items.Add(this.btnInactivareAddress);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Items.Add(this.btnSuspender);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnValidateAccounts
            // 
            this.btnValidateAccounts.Label = "Validar Cuentas";
            this.btnValidateAccounts.Name = "btnValidateAccounts";
            this.btnValidateAccounts.ShowImage = true;
            this.btnValidateAccounts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateAccounts_Click);
            // 
            // btnLoadAccounts
            // 
            this.btnLoadAccounts.Label = "Cambiar Cuentas";
            this.btnLoadAccounts.Name = "btnLoadAccounts";
            this.btnLoadAccounts.ShowImage = true;
            this.btnLoadAccounts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // btnInactivateBussiness
            // 
            this.btnInactivateBussiness.Label = "Inactivar Supplier Bussiness";
            this.btnInactivateBussiness.Name = "btnInactivateBussiness";
            this.btnInactivateBussiness.ShowImage = true;
            this.btnInactivateBussiness.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInactivateBussiness_Click);
            // 
            // btnInactivareAddress
            // 
            this.btnInactivareAddress.Label = "Inactivar Supplier Address";
            this.btnInactivareAddress.Name = "btnInactivareAddress";
            this.btnInactivareAddress.ShowImage = true;
            this.btnInactivareAddress.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInactivareAddress_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnSuspender
            // 
            this.btnSuspender.Label = "Suspender Supplier";
            this.btnSuspender.Name = "btnSuspender";
            this.btnSuspender.ShowImage = true;
            this.btnSuspender.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSuspender_Click);
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO200;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormats;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeAccounts;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateAccounts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadAccounts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatInactivateSupplierBusiness;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInactivateSupplier;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInactivateBussiness;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInactivareAddress;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSuspender;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

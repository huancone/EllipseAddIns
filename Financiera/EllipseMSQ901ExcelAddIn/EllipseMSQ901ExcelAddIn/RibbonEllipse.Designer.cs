namespace EllipseMSQ901ExcelAddIn
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
            this.grpMSQ901 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatoSupplierInvoice = this.Factory.CreateRibbonButton();
            this.btnJournal = this.Factory.CreateRibbonButton();
            this.btnCustomerInvoice = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnConsultar = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSQ901.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSQ901);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSQ901
            // 
            this.grpMSQ901.Items.Add(this.box1);
            this.grpMSQ901.Items.Add(this.drpEnvironment);
            this.grpMSQ901.Items.Add(this.menuActions);
            this.grpMSQ901.Label = "MSQ901";
            this.grpMSQ901.Name = "grpMSQ901";
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatoSupplierInvoice);
            this.menuFormat.Items.Add(this.btnJournal);
            this.menuFormat.Items.Add(this.btnCustomerInvoice);
            this.menuFormat.Label = "&Formatos";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatoSupplierInvoice
            // 
            this.btnFormatoSupplierInvoice.Label = "&Supplier/Invoice";
            this.btnFormatoSupplierInvoice.Name = "btnFormatoSupplierInvoice";
            this.btnFormatoSupplierInvoice.ShowImage = true;
            this.btnFormatoSupplierInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatoSupplierInvoice_Click);
            // 
            // btnJournal
            // 
            this.btnJournal.Label = "&Journal";
            this.btnJournal.Name = "btnJournal";
            this.btnJournal.ShowImage = true;
            this.btnJournal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnJournal_Click);
            // 
            // btnCustomerInvoice
            // 
            this.btnCustomerInvoice.Label = "&Customer/Invoice";
            this.btnCustomerInvoice.Name = "btnCustomerInvoice";
            this.btnCustomerInvoice.ShowImage = true;
            this.btnCustomerInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCustomerInvoice_Click);
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
            this.menuActions.Items.Add(this.btnConsultar);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnConsultar
            // 
            this.btnConsultar.Label = "C&onsultar";
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.ShowImage = true;
            this.btnConsultar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultar_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener &Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpMSQ901.ResumeLayout(false);
            this.grpMSQ901.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSQ901;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatoSupplierInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnJournal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCustomerInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

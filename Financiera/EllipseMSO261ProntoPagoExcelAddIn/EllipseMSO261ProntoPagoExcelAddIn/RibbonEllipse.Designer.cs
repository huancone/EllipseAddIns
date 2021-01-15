namespace EllipseMSO261ProntoPagoExcelAddIn
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
            this.grpProntoPago = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnGetInvoice = this.Factory.CreateRibbonButton();
            this.btnReloadParameters = this.Factory.CreateRibbonButton();
            this.btnCalculateDiscount = this.Factory.CreateRibbonButton();
            this.btnModifyInvoice = this.Factory.CreateRibbonButton();
            this.btnVerifyInvoice = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpProntoPago.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpProntoPago);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpProntoPago
            // 
            this.grpProntoPago.Items.Add(this.box1);
            this.grpProntoPago.Items.Add(this.drpEnvironment);
            this.grpProntoPago.Items.Add(this.menuActions);
            this.grpProntoPago.Label = "Pronto Pago";
            this.grpProntoPago.Name = "grpProntoPago";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Formato";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env. ";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnGetInvoice);
            this.menuActions.Items.Add(this.btnReloadParameters);
            this.menuActions.Items.Add(this.btnCalculateDiscount);
            this.menuActions.Items.Add(this.btnModifyInvoice);
            this.menuActions.Items.Add(this.btnVerifyInvoice);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnGetInvoice
            // 
            this.btnGetInvoice.Label = "1. Consultar Facturas";
            this.btnGetInvoice.Name = "btnGetInvoice";
            this.btnGetInvoice.ShowImage = true;
            this.btnGetInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetInvoice_Click);
            // 
            // btnReloadParameters
            // 
            this.btnReloadParameters.Label = "2. Recargar Parametros";
            this.btnReloadParameters.Name = "btnReloadParameters";
            this.btnReloadParameters.ShowImage = true;
            this.btnReloadParameters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReloadParameters_Click);
            // 
            // btnCalculateDiscount
            // 
            this.btnCalculateDiscount.Label = "3. Calcular Descuentos";
            this.btnCalculateDiscount.Name = "btnCalculateDiscount";
            this.btnCalculateDiscount.ShowImage = true;
            this.btnCalculateDiscount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalculateDiscount_Click);
            // 
            // btnModifyInvoice
            // 
            this.btnModifyInvoice.Label = "4. Modificar Facturas";
            this.btnModifyInvoice.Name = "btnModifyInvoice";
            this.btnModifyInvoice.ShowImage = true;
            this.btnModifyInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyInvoice_Click);
            // 
            // btnVerifyInvoice
            // 
            this.btnVerifyInvoice.Label = "5. Verificar Facturas";
            this.btnVerifyInvoice.Name = "btnVerifyInvoice";
            this.btnVerifyInvoice.ShowImage = true;
            this.btnVerifyInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVerifyInvoice_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpProntoPago.ResumeLayout(false);
            this.grpProntoPago.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProntoPago;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReloadParameters;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalculateDiscount;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVerifyInvoice;
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

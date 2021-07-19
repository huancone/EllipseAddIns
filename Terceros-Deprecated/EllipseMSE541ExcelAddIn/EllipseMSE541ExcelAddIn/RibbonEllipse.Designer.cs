namespace EllipseMSE541ExcelAddIn
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
            this.grpWorkRequest = this.Factory.CreateRibbonGroup();
            this.Formatear = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.actions = this.Factory.CreateRibbonMenu();
            this.Ejecutar = this.Factory.CreateRibbonButton();
            this.Consulta = this.Factory.CreateRibbonButton();
            this.modificar = this.Factory.CreateRibbonButton();
            this.cerrar = this.Factory.CreateRibbonButton();
            this.clean = this.Factory.CreateRibbonButton();
            this.procesos = this.Factory.CreateRibbonMenu();
            this.wrCerrar = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkRequest.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpWorkRequest);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpWorkRequest
            // 
            this.grpWorkRequest.Items.Add(this.Formatear);
            this.grpWorkRequest.Items.Add(this.drpEnvironment);
            this.grpWorkRequest.Items.Add(this.actions);
            this.grpWorkRequest.Label = "Work Request v1.0";
            this.grpWorkRequest.Name = "grpWorkRequest";
            // 
            // Formatear
            // 
            this.Formatear.Label = "Format Single Sheet";
            this.Formatear.Name = "Formatear";
            this.Formatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Formatear_Click);
            // 
            // drpEnvironment
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Test";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            this.drpEnvironment.Items.Add(ribbonDropDownItemImpl1);
            this.drpEnvironment.Items.Add(ribbonDropDownItemImpl2);
            this.drpEnvironment.Items.Add(ribbonDropDownItemImpl3);
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // actions
            // 
            this.actions.Items.Add(this.Ejecutar);
            this.actions.Items.Add(this.Consulta);
            this.actions.Items.Add(this.modificar);
            this.actions.Items.Add(this.cerrar);
            this.actions.Items.Add(this.clean);
            this.actions.Items.Add(this.procesos);
            this.actions.Label = "Actions";
            this.actions.Name = "actions";
            // 
            // Ejecutar
            // 
            this.Ejecutar.Label = "Cargar";
            this.Ejecutar.Name = "Ejecutar";
            this.Ejecutar.ShowImage = true;
            this.Ejecutar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Ejecutar_Click);
            // 
            // Consulta
            // 
            this.Consulta.Label = "Consultar";
            this.Consulta.Name = "Consulta";
            this.Consulta.ShowImage = true;
            this.Consulta.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Consulta_Click);
            // 
            // modificar
            // 
            this.modificar.Label = "Modificar";
            this.modificar.Name = "modificar";
            this.modificar.ShowImage = true;
            this.modificar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.modificar_Click_1);
            // 
            // cerrar
            // 
            this.cerrar.Label = "Cerrar";
            this.cerrar.Name = "cerrar";
            this.cerrar.ShowImage = true;
            this.cerrar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cerrar_Click);
            // 
            // clean
            // 
            this.clean.Label = "Limpiar Hoja";
            this.clean.Name = "clean";
            this.clean.ShowImage = true;
            this.clean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clean_Click);
            // 
            // procesos
            // 
            this.procesos.Items.Add(this.wrCerrar);
            this.procesos.Label = "Procesos";
            this.procesos.Name = "procesos";
            this.procesos.ShowImage = true;
            // 
            // wrCerrar
            // 
            this.wrCerrar.Label = "Cierre WR Completados";
            this.wrCerrar.Name = "wrCerrar";
            this.wrCerrar.ShowImage = true;
            this.wrCerrar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wrCerrar_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpWorkRequest.ResumeLayout(false);
            this.grpWorkRequest.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Formatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Ejecutar;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu actions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Consulta;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clean;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton modificar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cerrar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wrCerrar;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu procesos;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

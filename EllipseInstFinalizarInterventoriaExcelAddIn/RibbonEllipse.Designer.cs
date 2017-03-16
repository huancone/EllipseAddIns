namespace EllipseInstFinalizarInterventoriaExcelAddIn
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
            this.grpInstFinalizarInterventoria = this.Factory.CreateRibbonGroup();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnClearSheet = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpInstFinalizarInterventoria.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpInstFinalizarInterventoria);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpInstFinalizarInterventoria
            // 
            this.grpInstFinalizarInterventoria.Items.Add(this.btnFormatSheet);
            this.grpInstFinalizarInterventoria.Items.Add(this.drpEnviroment);
            this.grpInstFinalizarInterventoria.Items.Add(this.menuActions);
            this.grpInstFinalizarInterventoria.Label = "Finalizar Inter. v1.0.1";
            this.grpInstFinalizarInterventoria.Name = "grpInstFinalizarInterventoria";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnUpdate);
            this.menuActions.Items.Add(this.btnClearSheet);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnUpdate
            // 
            this.btnUpdate.Label = "&Actualización";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // btnClearSheet
            // 
            this.btnClearSheet.Label = "&Limpiar Hoja";
            this.btnClearSheet.Name = "btnClearSheet";
            this.btnClearSheet.ShowImage = true;
            this.btnClearSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearSheet_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpInstFinalizarInterventoria.ResumeLayout(false);
            this.grpInstFinalizarInterventoria.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInstFinalizarInterventoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

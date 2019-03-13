namespace EllipseNonOrderInvoiceEntryExcelAddIn
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
            this.grpMSO265 = this.Factory.CreateRibbonGroup();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.mnActions = this.Factory.CreateRibbonMenu();
            this.btnLoadFile = this.Factory.CreateRibbonButton();
            this.btnValidateSheet = this.Factory.CreateRibbonButton();
            this.btnLoadSheet = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO265.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO265);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO265
            // 
            this.grpMSO265.Items.Add(this.btnFormat);
            this.grpMSO265.Items.Add(this.drpEnvironment);
            this.grpMSO265.Items.Add(this.mnActions);
            this.grpMSO265.Label = "MSO265 v1.0.0.0";
            this.grpMSO265.Name = "grpMSO265";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Format Sheet";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env,";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // mnActions
            // 
            this.mnActions.Items.Add(this.btnLoadFile);
            this.mnActions.Items.Add(this.btnValidateSheet);
            this.mnActions.Items.Add(this.btnLoadSheet);
            this.mnActions.Label = "Actions";
            this.mnActions.Name = "mnActions";
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Label = "Load File";
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.ShowImage = true;
            this.btnLoadFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadFile_Click);
            // 
            // btnValidateSheet
            // 
            this.btnValidateSheet.Label = "Validate Sheet";
            this.btnValidateSheet.Name = "btnValidateSheet";
            this.btnValidateSheet.ShowImage = true;
            // 
            // btnLoadSheet
            // 
            this.btnLoadSheet.Label = "Load Sheet";
            this.btnLoadSheet.Name = "btnLoadSheet";
            this.btnLoadSheet.ShowImage = true;
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpMSO265.ResumeLayout(false);
            this.grpMSO265.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMSO265;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadSheet;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

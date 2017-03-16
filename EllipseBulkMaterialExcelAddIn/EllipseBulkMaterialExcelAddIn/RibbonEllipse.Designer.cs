namespace EllipseBulkMaterialExcelAddIn
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
            this.grpBulkMaterial = this.Factory.CreateRibbonGroup();
            this.btnBulkMaterialFormatMultiple = this.Factory.CreateRibbonButton();
            this.drpBulkMaterialEnv = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnValidateStats = this.Factory.CreateRibbonButton();
            this.btnImport = this.Factory.CreateRibbonButton();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.btnUnApplyDelete = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpBulkMaterial.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpBulkMaterial);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpBulkMaterial
            // 
            this.grpBulkMaterial.Items.Add(this.btnBulkMaterialFormatMultiple);
            this.grpBulkMaterial.Items.Add(this.drpBulkMaterialEnv);
            this.grpBulkMaterial.Items.Add(this.menuActions);
            this.grpBulkMaterial.Label = "Bulk Material v2.0.0";
            this.grpBulkMaterial.Name = "grpBulkMaterial";
            // 
            // btnBulkMaterialFormatMultiple
            // 
            this.btnBulkMaterialFormatMultiple.Label = "Format Bulk Material Sheet";
            this.btnBulkMaterialFormatMultiple.Name = "btnBulkMaterialFormatMultiple";
            this.btnBulkMaterialFormatMultiple.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBulkMaterialFormatMultiple_Click);
            // 
            // drpBulkMaterialEnv
            // 
            this.drpBulkMaterialEnv.Label = "Env.";
            this.drpBulkMaterialEnv.Name = "drpBulkMaterialEnv";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnValidateStats);
            this.menuActions.Items.Add(this.btnImport);
            this.menuActions.Items.Add(this.btnLoad);
            this.menuActions.Items.Add(this.btnUnApplyDelete);
            this.menuActions.Label = "Actions";
            this.menuActions.Name = "menuActions";
            // 
            // btnValidateStats
            // 
            this.btnValidateStats.Label = "Validate Stats";
            this.btnValidateStats.Name = "btnValidateStats";
            this.btnValidateStats.ShowImage = true;
            this.btnValidateStats.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateStats_Click);
            // 
            // btnImport
            // 
            this.btnImport.Label = "Import CSV File";
            this.btnImport.Name = "btnImport";
            this.btnImport.ShowImage = true;
            this.btnImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImport_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Load Data";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.ShowImage = true;
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // btnUnApplyDelete
            // 
            this.btnUnApplyDelete.Label = "Unapply - Delete";
            this.btnUnApplyDelete.Name = "btnUnApplyDelete";
            this.btnUnApplyDelete.ShowImage = true;
            this.btnUnApplyDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnApplyDelete_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpBulkMaterial.ResumeLayout(false);
            this.grpBulkMaterial.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpBulkMaterial;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpBulkMaterialEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBulkMaterialFormatMultiple;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnApplyDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateStats;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

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
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnBulkMaterialFormatMultiple = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.btnLoadSecond = this.Factory.CreateRibbonButton();
            this.btnValidateStats = this.Factory.CreateRibbonButton();
            this.btnImport = this.Factory.CreateRibbonButton();
            this.btnUnApplyDelete = this.Factory.CreateRibbonButton();
            this.menuListActions = this.Factory.CreateRibbonMenu();
            this.btnReviewEquipList = this.Factory.CreateRibbonButton();
            this.btnReviewFromBulkSheet = this.Factory.CreateRibbonButton();
            this.btnAddToList = this.Factory.CreateRibbonButton();
            this.btnRemoveFromList = this.Factory.CreateRibbonButton();
            this.cbAutoSortItems = this.Factory.CreateRibbonCheckBox();
            this.cbIgnoreItemError = this.Factory.CreateRibbonCheckBox();
            this.cbAccountElementOverride = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpBulkMaterial.SuspendLayout();
            this.box1.SuspendLayout();
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
            this.grpBulkMaterial.Items.Add(this.box1);
            this.grpBulkMaterial.Items.Add(this.drpEnvironment);
            this.grpBulkMaterial.Items.Add(this.menuActions);
            this.grpBulkMaterial.Label = "Bulk Material";
            this.grpBulkMaterial.Name = "grpBulkMaterial";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnBulkMaterialFormatMultiple);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnBulkMaterialFormatMultiple
            // 
            this.btnBulkMaterialFormatMultiple.Label = "&Formatear";
            this.btnBulkMaterialFormatMultiple.Name = "btnBulkMaterialFormatMultiple";
            this.btnBulkMaterialFormatMultiple.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBulkMaterialFormatMultiple_Click);
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
            this.menuActions.Items.Add(this.btnLoad);
            this.menuActions.Items.Add(this.btnLoadSecond);
            this.menuActions.Items.Add(this.btnValidateStats);
            this.menuActions.Items.Add(this.btnImport);
            this.menuActions.Items.Add(this.btnUnApplyDelete);
            this.menuActions.Items.Add(this.menuListActions);
            this.menuActions.Items.Add(this.cbAutoSortItems);
            this.menuActions.Items.Add(this.cbIgnoreItemError);
            this.menuActions.Items.Add(this.cbAccountElementOverride);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Actions";
            this.menuActions.Name = "menuActions";
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Load Data";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.ShowImage = true;
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // btnLoadSecond
            // 
            this.btnLoadSecond.Label = "Load Data (No Post)";
            this.btnLoadSecond.Name = "btnLoadSecond";
            this.btnLoadSecond.ShowImage = true;
            this.btnLoadSecond.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadSecond_Click);
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
            // btnUnApplyDelete
            // 
            this.btnUnApplyDelete.Label = "Unapply - Delete";
            this.btnUnApplyDelete.Name = "btnUnApplyDelete";
            this.btnUnApplyDelete.ShowImage = true;
            this.btnUnApplyDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnApplyDelete_Click);
            // 
            // menuListActions
            // 
            this.menuListActions.Items.Add(this.btnReviewEquipList);
            this.menuListActions.Items.Add(this.btnReviewFromBulkSheet);
            this.menuListActions.Items.Add(this.btnAddToList);
            this.menuListActions.Items.Add(this.btnRemoveFromList);
            this.menuListActions.Label = "Lista";
            this.menuListActions.Name = "menuListActions";
            this.menuListActions.ShowImage = true;
            // 
            // btnReviewEquipList
            // 
            this.btnReviewEquipList.Label = "&Consultar Listas";
            this.btnReviewEquipList.Name = "btnReviewEquipList";
            this.btnReviewEquipList.ShowImage = true;
            this.btnReviewEquipList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewEquipList_Click);
            // 
            // btnReviewFromBulkSheet
            // 
            this.btnReviewFromBulkSheet.Label = "C&onsultar desde Hoja de Combustible";
            this.btnReviewFromBulkSheet.Name = "btnReviewFromBulkSheet";
            this.btnReviewFromBulkSheet.ShowImage = true;
            this.btnReviewFromBulkSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewFromBulkSheet_Click);
            // 
            // btnAddToList
            // 
            this.btnAddToList.Label = "&Agregar a la Lista";
            this.btnAddToList.Name = "btnAddToList";
            this.btnAddToList.ShowImage = true;
            this.btnAddToList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddToList_Click);
            // 
            // btnRemoveFromList
            // 
            this.btnRemoveFromList.Label = "&Quitar de la Lista";
            this.btnRemoveFromList.Name = "btnRemoveFromList";
            this.btnRemoveFromList.ShowImage = true;
            this.btnRemoveFromList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveFromList_Click);
            // 
            // cbAutoSortItems
            // 
            this.cbAutoSortItems.Checked = true;
            this.cbAutoSortItems.Label = "Ordenar Automáticamente";
            this.cbAutoSortItems.Name = "cbAutoSortItems";
            // 
            // cbIgnoreItemError
            // 
            this.cbIgnoreItemError.Label = "Ignorar Errores en Ítems";
            this.cbIgnoreItemError.Name = "cbIgnoreItemError";
            // 
            // cbAccountElementOverride
            // 
            this.cbAccountElementOverride.Label = "Autoasignar Centro de Costo";
            this.cbAccountElementOverride.Name = "cbAccountElementOverride";
            this.cbAccountElementOverride.ScreenTip = "Asignará el Centro de Costo ignorando el escrito y utilizará el relacionado con e" +
    "l equipo y el tipo de material";
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
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
            this.grpBulkMaterial.ResumeLayout(false);
            this.grpBulkMaterial.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpBulkMaterial;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBulkMaterialFormatMultiple;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnApplyDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateStats;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuListActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewEquipList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddToList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveFromList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewFromBulkSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIgnoreItemError;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAutoSortItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadSecond;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAccountElementOverride;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

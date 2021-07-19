﻿namespace EllipseBulkMaterialExcelAddIn
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
            this.btnUnApplyDelete = this.Factory.CreateRibbonButton();
            this.btnValidateStats = this.Factory.CreateRibbonButton();
            this.btnImport = this.Factory.CreateRibbonButton();
            this.menuListActions = this.Factory.CreateRibbonMenu();
            this.btnReviewEquipList = this.Factory.CreateRibbonButton();
            this.btnReviewFromBulkSheet = this.Factory.CreateRibbonButton();
            this.btnAddToList = this.Factory.CreateRibbonButton();
            this.btnRemoveFromList = this.Factory.CreateRibbonButton();
            this.menuOptions = this.Factory.CreateRibbonMenu();
            this.menuAutoasignAccountCode = this.Factory.CreateRibbonMenu();
            this.cbAccountElementOverrideDisable = this.Factory.CreateRibbonCheckBox();
            this.cbAccountElementOverrideDefault = this.Factory.CreateRibbonCheckBox();
            this.cbAccountElementOverrideAlways = this.Factory.CreateRibbonCheckBox();
            this.cbAccountElementOverrideMntto = this.Factory.CreateRibbonCheckBox();
            this.cbAutoSortItems = this.Factory.CreateRibbonCheckBox();
            this.cbIgnoreItemError = this.Factory.CreateRibbonCheckBox();
            this.cbAllowBackgroundWork = this.Factory.CreateRibbonCheckBox();
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
            this.tabEllipse.Label = "ELLIPSE";
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
            this.menuActions.Items.Add(this.btnUnApplyDelete);
            this.menuActions.Items.Add(this.btnValidateStats);
            this.menuActions.Items.Add(this.btnImport);
            this.menuActions.Items.Add(this.menuListActions);
            this.menuActions.Items.Add(this.menuOptions);
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
            // btnUnApplyDelete
            // 
            this.btnUnApplyDelete.Label = "Unapply - Delete";
            this.btnUnApplyDelete.Name = "btnUnApplyDelete";
            this.btnUnApplyDelete.ShowImage = true;
            this.btnUnApplyDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnApplyDelete_Click);
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
            // menuListActions
            // 
            this.menuListActions.Items.Add(this.btnReviewEquipList);
            this.menuListActions.Items.Add(this.btnReviewFromBulkSheet);
            this.menuListActions.Items.Add(this.btnAddToList);
            this.menuListActions.Items.Add(this.btnRemoveFromList);
            this.menuListActions.Label = "Lista de Equipos";
            this.menuListActions.Name = "menuListActions";
            this.menuListActions.ShowImage = true;
            // 
            // btnReviewEquipList
            // 
            this.btnReviewEquipList.Label = "&Consultar Listas de Equipos";
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
            // menuOptions
            // 
            this.menuOptions.Items.Add(this.menuAutoasignAccountCode);
            this.menuOptions.Items.Add(this.cbAutoSortItems);
            this.menuOptions.Items.Add(this.cbIgnoreItemError);
            this.menuOptions.Items.Add(this.cbAllowBackgroundWork);
            this.menuOptions.Label = "Opciones";
            this.menuOptions.Name = "menuOptions";
            this.menuOptions.ShowImage = true;
            // 
            // menuAutoasignAccountCode
            // 
            this.menuAutoasignAccountCode.Items.Add(this.cbAccountElementOverrideDisable);
            this.menuAutoasignAccountCode.Items.Add(this.cbAccountElementOverrideDefault);
            this.menuAutoasignAccountCode.Items.Add(this.cbAccountElementOverrideAlways);
            this.menuAutoasignAccountCode.Items.Add(this.cbAccountElementOverrideMntto);
            this.menuAutoasignAccountCode.Label = "Centro de Costo";
            this.menuAutoasignAccountCode.Name = "menuAutoasignAccountCode";
            this.menuAutoasignAccountCode.ShowImage = true;
            // 
            // cbAccountElementOverrideDisable
            // 
            this.cbAccountElementOverrideDisable.Label = "Autosignar Desactivado";
            this.cbAccountElementOverrideDisable.Name = "cbAccountElementOverrideDisable";
            this.cbAccountElementOverrideDisable.ScreenTip = "Asignará el Centro de Costo ignorando el escrito y utilizará el relacionado con e" +
    "l equipo y el tipo de material";
            this.cbAccountElementOverrideDisable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAccountElementOverrideDisable_Click);
            // 
            // cbAccountElementOverrideDefault
            // 
            this.cbAccountElementOverrideDefault.Label = "Autosignar Predeterminado";
            this.cbAccountElementOverrideDefault.Name = "cbAccountElementOverrideDefault";
            this.cbAccountElementOverrideDefault.ScreenTip = "Solo si no existe para el encabezado";
            this.cbAccountElementOverrideDefault.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAccountElementOverrideDefault_Click);
            // 
            // cbAccountElementOverrideAlways
            // 
            this.cbAccountElementOverrideAlways.Label = "Autosignar Siempre";
            this.cbAccountElementOverrideAlways.Name = "cbAccountElementOverrideAlways";
            this.cbAccountElementOverrideAlways.ScreenTip = "Ignora el Centro de Costo escrito y busca siempre el del item";
            this.cbAccountElementOverrideAlways.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAccountElementOverrideAlways_Click);
            // 
            // cbAccountElementOverrideMntto
            // 
            this.cbAccountElementOverrideMntto.Label = "Autosignar Mantenimiento";
            this.cbAccountElementOverrideMntto.Name = "cbAccountElementOverrideMntto";
            this.cbAccountElementOverrideMntto.ScreenTip = "Ignora el centro de costo escrito y busca el del ítem si pertenece a Mantenimient" +
    "o. Predeterminado para los demás casos";
            this.cbAccountElementOverrideMntto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAccountElementOverrideMntto_Click);
            // 
            // cbAutoSortItems
            // 
            this.cbAutoSortItems.Checked = true;
            this.cbAutoSortItems.Label = "Ordenar Automáticamente";
            this.cbAutoSortItems.Name = "cbAutoSortItems";
            this.cbAutoSortItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAutoSortItems_Click);
            // 
            // cbIgnoreItemError
            // 
            this.cbIgnoreItemError.Label = "Ignorar Errores en Ítems";
            this.cbIgnoreItemError.Name = "cbIgnoreItemError";
            this.cbIgnoreItemError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbIgnoreItemError_Click);
            // 
            // cbAllowBackgroundWork
            // 
            this.cbAllowBackgroundWork.Label = "Trabajo en Segundo Plano";
            this.cbAllowBackgroundWork.Name = "cbAllowBackgroundWork";
            this.cbAllowBackgroundWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAllowBackgroundWork_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAccountElementOverrideDisable;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAccountElementOverrideMntto;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAutoasignAccountCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAccountElementOverrideDefault;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAccountElementOverrideAlways;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAllowBackgroundWork;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

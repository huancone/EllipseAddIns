namespace LogsheetDatamodelAdmin
{
    partial class RibbonLsdm : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonLsdm()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonLsdm));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpEllipse = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuDatasheet = this.Factory.CreateRibbonMenu();
            this.btnDatasheetSearch = this.Factory.CreateRibbonButton();
            this.btnDatasheetUpdate = this.Factory.CreateRibbonButton();
            this.btnDatasheetDelete = this.Factory.CreateRibbonButton();
            this.menuAttributes = this.Factory.CreateRibbonMenu();
            this.btnAttributeSearch = this.Factory.CreateRibbonButton();
            this.btnAttributeSearchEach = this.Factory.CreateRibbonButton();
            this.btnAttributeUpdate = this.Factory.CreateRibbonButton();
            this.btnAttributeDelete = this.Factory.CreateRibbonButton();
            this.menuModel = this.Factory.CreateRibbonMenu();
            this.btnModelSearch = this.Factory.CreateRibbonButton();
            this.btnModelSearchEach = this.Factory.CreateRibbonButton();
            this.btnModelUpdate = this.Factory.CreateRibbonButton();
            this.btnModelDelete = this.Factory.CreateRibbonButton();
            this.menuMeasure = this.Factory.CreateRibbonMenu();
            this.btnMeasureSearch = this.Factory.CreateRibbonButton();
            this.btnMeasureSearchEach = this.Factory.CreateRibbonButton();
            this.btnMeasureUpdate = this.Factory.CreateRibbonButton();
            this.btnMeasureDelete = this.Factory.CreateRibbonButton();
            this.menuMeasureType = this.Factory.CreateRibbonMenu();
            this.btnMeasureTypeSearch = this.Factory.CreateRibbonButton();
            this.btnMeasureTypeSearchEach = this.Factory.CreateRibbonButton();
            this.btnMeasureTypeUpdate = this.Factory.CreateRibbonButton();
            this.btnMeasureTypeDelete = this.Factory.CreateRibbonButton();
            this.menuValidationItems = this.Factory.CreateRibbonMenu();
            this.btnValidItemsSearch = this.Factory.CreateRibbonButton();
            this.btnValidItemsSearchEach = this.Factory.CreateRibbonButton();
            this.btnValidItemsUpdate = this.Factory.CreateRibbonButton();
            this.btnValidItemsDelete = this.Factory.CreateRibbonButton();
            this.menuValidationSources = this.Factory.CreateRibbonMenu();
            this.btnValidSourcesSearch = this.Factory.CreateRibbonButton();
            this.btnValidSourcesSearchEach = this.Factory.CreateRibbonButton();
            this.btnValidSourcesUpdate = this.Factory.CreateRibbonButton();
            this.btnValidSourcesDelete = this.Factory.CreateRibbonButton();
            this.btnStop = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpEllipse);
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.box1);
            this.grpEllipse.Items.Add(this.menuActions);
            resources.ApplyResources(this.grpEllipse, "grpEllipse");
            this.grpEllipse.Name = "grpEllipse";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            resources.ApplyResources(this.btnFormat, "btnFormat");
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // btnAbout
            // 
            resources.ApplyResources(this.btnAbout, "btnAbout");
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.menuDatasheet);
            this.menuActions.Items.Add(this.menuAttributes);
            this.menuActions.Items.Add(this.menuModel);
            this.menuActions.Items.Add(this.menuMeasure);
            this.menuActions.Items.Add(this.menuMeasureType);
            this.menuActions.Items.Add(this.menuValidationItems);
            this.menuActions.Items.Add(this.menuValidationSources);
            this.menuActions.Items.Add(this.btnStop);
            resources.ApplyResources(this.menuActions, "menuActions");
            this.menuActions.Name = "menuActions";
            // 
            // menuDatasheet
            // 
            this.menuDatasheet.Items.Add(this.btnDatasheetSearch);
            this.menuDatasheet.Items.Add(this.btnDatasheetUpdate);
            this.menuDatasheet.Items.Add(this.btnDatasheetDelete);
            resources.ApplyResources(this.menuDatasheet, "menuDatasheet");
            this.menuDatasheet.Name = "menuDatasheet";
            this.menuDatasheet.ShowImage = true;
            // 
            // btnDatasheetSearch
            // 
            resources.ApplyResources(this.btnDatasheetSearch, "btnDatasheetSearch");
            this.btnDatasheetSearch.Name = "btnDatasheetSearch";
            this.btnDatasheetSearch.ShowImage = true;
            this.btnDatasheetSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDatasheetSearch_Click);
            // 
            // btnDatasheetUpdate
            // 
            resources.ApplyResources(this.btnDatasheetUpdate, "btnDatasheetUpdate");
            this.btnDatasheetUpdate.Name = "btnDatasheetUpdate";
            this.btnDatasheetUpdate.ShowImage = true;
            this.btnDatasheetUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDatasheetUpdate_Click);
            // 
            // btnDatasheetDelete
            // 
            resources.ApplyResources(this.btnDatasheetDelete, "btnDatasheetDelete");
            this.btnDatasheetDelete.Name = "btnDatasheetDelete";
            this.btnDatasheetDelete.ShowImage = true;
            this.btnDatasheetDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDatasheetDelete_Click);
            // 
            // menuAttributes
            // 
            this.menuAttributes.Items.Add(this.btnAttributeSearch);
            this.menuAttributes.Items.Add(this.btnAttributeSearchEach);
            this.menuAttributes.Items.Add(this.btnAttributeUpdate);
            this.menuAttributes.Items.Add(this.btnAttributeDelete);
            resources.ApplyResources(this.menuAttributes, "menuAttributes");
            this.menuAttributes.Name = "menuAttributes";
            this.menuAttributes.ShowImage = true;
            // 
            // btnAttributeSearch
            // 
            resources.ApplyResources(this.btnAttributeSearch, "btnAttributeSearch");
            this.btnAttributeSearch.Name = "btnAttributeSearch";
            this.btnAttributeSearch.ShowImage = true;
            this.btnAttributeSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAttributeSearch_Click);
            // 
            // btnAttributeSearchEach
            // 
            resources.ApplyResources(this.btnAttributeSearchEach, "btnAttributeSearchEach");
            this.btnAttributeSearchEach.Name = "btnAttributeSearchEach";
            this.btnAttributeSearchEach.ShowImage = true;
            this.btnAttributeSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAttributeSearchEach_Click);
            // 
            // btnAttributeUpdate
            // 
            resources.ApplyResources(this.btnAttributeUpdate, "btnAttributeUpdate");
            this.btnAttributeUpdate.Name = "btnAttributeUpdate";
            this.btnAttributeUpdate.ShowImage = true;
            this.btnAttributeUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAttributeUpdate_Click);
            // 
            // btnAttributeDelete
            // 
            resources.ApplyResources(this.btnAttributeDelete, "btnAttributeDelete");
            this.btnAttributeDelete.Name = "btnAttributeDelete";
            this.btnAttributeDelete.ShowImage = true;
            this.btnAttributeDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAttributeDelete_Click);
            // 
            // menuModel
            // 
            this.menuModel.Items.Add(this.btnModelSearch);
            this.menuModel.Items.Add(this.btnModelSearchEach);
            this.menuModel.Items.Add(this.btnModelUpdate);
            this.menuModel.Items.Add(this.btnModelDelete);
            resources.ApplyResources(this.menuModel, "menuModel");
            this.menuModel.Name = "menuModel";
            this.menuModel.ShowImage = true;
            // 
            // btnModelSearch
            // 
            resources.ApplyResources(this.btnModelSearch, "btnModelSearch");
            this.btnModelSearch.Name = "btnModelSearch";
            this.btnModelSearch.ShowImage = true;
            this.btnModelSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModelSearch_Click);
            // 
            // btnModelSearchEach
            // 
            resources.ApplyResources(this.btnModelSearchEach, "btnModelSearchEach");
            this.btnModelSearchEach.Name = "btnModelSearchEach";
            this.btnModelSearchEach.ShowImage = true;
            this.btnModelSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModelSearchEach_Click);
            // 
            // btnModelUpdate
            // 
            resources.ApplyResources(this.btnModelUpdate, "btnModelUpdate");
            this.btnModelUpdate.Name = "btnModelUpdate";
            this.btnModelUpdate.ShowImage = true;
            this.btnModelUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModelUpdate_Click);
            // 
            // btnModelDelete
            // 
            resources.ApplyResources(this.btnModelDelete, "btnModelDelete");
            this.btnModelDelete.Name = "btnModelDelete";
            this.btnModelDelete.ShowImage = true;
            this.btnModelDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModelDelete_Click);
            // 
            // menuMeasure
            // 
            this.menuMeasure.Items.Add(this.btnMeasureSearch);
            this.menuMeasure.Items.Add(this.btnMeasureSearchEach);
            this.menuMeasure.Items.Add(this.btnMeasureUpdate);
            this.menuMeasure.Items.Add(this.btnMeasureDelete);
            resources.ApplyResources(this.menuMeasure, "menuMeasure");
            this.menuMeasure.Name = "menuMeasure";
            this.menuMeasure.ShowImage = true;
            // 
            // btnMeasureSearch
            // 
            resources.ApplyResources(this.btnMeasureSearch, "btnMeasureSearch");
            this.btnMeasureSearch.Name = "btnMeasureSearch";
            this.btnMeasureSearch.ShowImage = true;
            this.btnMeasureSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureSearch_Click);
            // 
            // btnMeasureSearchEach
            // 
            resources.ApplyResources(this.btnMeasureSearchEach, "btnMeasureSearchEach");
            this.btnMeasureSearchEach.Name = "btnMeasureSearchEach";
            this.btnMeasureSearchEach.ShowImage = true;
            this.btnMeasureSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureSearchEach_Click);
            // 
            // btnMeasureUpdate
            // 
            resources.ApplyResources(this.btnMeasureUpdate, "btnMeasureUpdate");
            this.btnMeasureUpdate.Name = "btnMeasureUpdate";
            this.btnMeasureUpdate.ShowImage = true;
            this.btnMeasureUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureUpdate_Click);
            // 
            // btnMeasureDelete
            // 
            resources.ApplyResources(this.btnMeasureDelete, "btnMeasureDelete");
            this.btnMeasureDelete.Name = "btnMeasureDelete";
            this.btnMeasureDelete.ShowImage = true;
            this.btnMeasureDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureDelete_Click);
            // 
            // menuMeasureType
            // 
            this.menuMeasureType.Items.Add(this.btnMeasureTypeSearch);
            this.menuMeasureType.Items.Add(this.btnMeasureTypeSearchEach);
            this.menuMeasureType.Items.Add(this.btnMeasureTypeUpdate);
            this.menuMeasureType.Items.Add(this.btnMeasureTypeDelete);
            resources.ApplyResources(this.menuMeasureType, "menuMeasureType");
            this.menuMeasureType.Name = "menuMeasureType";
            this.menuMeasureType.ShowImage = true;
            // 
            // btnMeasureTypeSearch
            // 
            resources.ApplyResources(this.btnMeasureTypeSearch, "btnMeasureTypeSearch");
            this.btnMeasureTypeSearch.Name = "btnMeasureTypeSearch";
            this.btnMeasureTypeSearch.ShowImage = true;
            this.btnMeasureTypeSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureTypeSearch_Click);
            // 
            // btnMeasureTypeSearchEach
            // 
            resources.ApplyResources(this.btnMeasureTypeSearchEach, "btnMeasureTypeSearchEach");
            this.btnMeasureTypeSearchEach.Name = "btnMeasureTypeSearchEach";
            this.btnMeasureTypeSearchEach.ShowImage = true;
            this.btnMeasureTypeSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureTypeSearchEach_Click);
            // 
            // btnMeasureTypeUpdate
            // 
            resources.ApplyResources(this.btnMeasureTypeUpdate, "btnMeasureTypeUpdate");
            this.btnMeasureTypeUpdate.Name = "btnMeasureTypeUpdate";
            this.btnMeasureTypeUpdate.ShowImage = true;
            this.btnMeasureTypeUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureTypeUpdate_Click);
            // 
            // btnMeasureTypeDelete
            // 
            resources.ApplyResources(this.btnMeasureTypeDelete, "btnMeasureTypeDelete");
            this.btnMeasureTypeDelete.Name = "btnMeasureTypeDelete";
            this.btnMeasureTypeDelete.ShowImage = true;
            this.btnMeasureTypeDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMeasureTypeDelete_Click);
            // 
            // menuValidationItems
            // 
            this.menuValidationItems.Items.Add(this.btnValidItemsSearch);
            this.menuValidationItems.Items.Add(this.btnValidItemsSearchEach);
            this.menuValidationItems.Items.Add(this.btnValidItemsUpdate);
            this.menuValidationItems.Items.Add(this.btnValidItemsDelete);
            resources.ApplyResources(this.menuValidationItems, "menuValidationItems");
            this.menuValidationItems.Name = "menuValidationItems";
            this.menuValidationItems.ShowImage = true;
            // 
            // btnValidItemsSearch
            // 
            resources.ApplyResources(this.btnValidItemsSearch, "btnValidItemsSearch");
            this.btnValidItemsSearch.Name = "btnValidItemsSearch";
            this.btnValidItemsSearch.ShowImage = true;
            this.btnValidItemsSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidItemsSearch_Click);
            // 
            // btnValidItemsSearchEach
            // 
            resources.ApplyResources(this.btnValidItemsSearchEach, "btnValidItemsSearchEach");
            this.btnValidItemsSearchEach.Name = "btnValidItemsSearchEach";
            this.btnValidItemsSearchEach.ShowImage = true;
            this.btnValidItemsSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidItemsSearchEach_Click);
            // 
            // btnValidItemsUpdate
            // 
            resources.ApplyResources(this.btnValidItemsUpdate, "btnValidItemsUpdate");
            this.btnValidItemsUpdate.Name = "btnValidItemsUpdate";
            this.btnValidItemsUpdate.ShowImage = true;
            this.btnValidItemsUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidItemsUpdate_Click);
            // 
            // btnValidItemsDelete
            // 
            resources.ApplyResources(this.btnValidItemsDelete, "btnValidItemsDelete");
            this.btnValidItemsDelete.Name = "btnValidItemsDelete";
            this.btnValidItemsDelete.ShowImage = true;
            this.btnValidItemsDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidItemsDelete_Click);
            // 
            // menuValidationSources
            // 
            this.menuValidationSources.Items.Add(this.btnValidSourcesSearch);
            this.menuValidationSources.Items.Add(this.btnValidSourcesSearchEach);
            this.menuValidationSources.Items.Add(this.btnValidSourcesUpdate);
            this.menuValidationSources.Items.Add(this.btnValidSourcesDelete);
            resources.ApplyResources(this.menuValidationSources, "menuValidationSources");
            this.menuValidationSources.Name = "menuValidationSources";
            this.menuValidationSources.ShowImage = true;
            // 
            // btnValidSourcesSearch
            // 
            resources.ApplyResources(this.btnValidSourcesSearch, "btnValidSourcesSearch");
            this.btnValidSourcesSearch.Name = "btnValidSourcesSearch";
            this.btnValidSourcesSearch.ShowImage = true;
            this.btnValidSourcesSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidSourcesSearch_Click);
            // 
            // btnValidSourcesSearchEach
            // 
            resources.ApplyResources(this.btnValidSourcesSearchEach, "btnValidSourcesSearchEach");
            this.btnValidSourcesSearchEach.Name = "btnValidSourcesSearchEach";
            this.btnValidSourcesSearchEach.ShowImage = true;
            this.btnValidSourcesSearchEach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidSourcesSearchEach_Click);
            // 
            // btnValidSourcesUpdate
            // 
            resources.ApplyResources(this.btnValidSourcesUpdate, "btnValidSourcesUpdate");
            this.btnValidSourcesUpdate.Name = "btnValidSourcesUpdate";
            this.btnValidSourcesUpdate.ShowImage = true;
            this.btnValidSourcesUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidSourcesUpdate_Click);
            // 
            // btnValidSourcesDelete
            // 
            resources.ApplyResources(this.btnValidSourcesDelete, "btnValidSourcesDelete");
            this.btnValidSourcesDelete.Name = "btnValidSourcesDelete";
            this.btnValidSourcesDelete.ShowImage = true;
            this.btnValidSourcesDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidSourcesDelete_Click);
            // 
            // btnStop
            // 
            resources.ApplyResources(this.btnStop, "btnStop");
            this.btnStop.Name = "btnStop";
            this.btnStop.ShowImage = true;
            this.btnStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStop_Click);
            // 
            // RibbonLsdm
            // 
            this.Name = "RibbonLsdm";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuMeasureType;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureTypeSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureTypeUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureTypeDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureTypeSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuMeasure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMeasureDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAttributes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAttributeSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAttributeSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAttributeUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAttributeDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuValidationItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidItemsSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidItemsSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidItemsUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidItemsDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuValidationSources;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidSourcesSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidSourcesSearchEach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidSourcesUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidSourcesDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuDatasheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDatasheetSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDatasheetUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDatasheetDelete;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonLsdm RibbonEllipse
        {
            get { return this.GetRibbon<RibbonLsdm>(); }
        }
    }
}

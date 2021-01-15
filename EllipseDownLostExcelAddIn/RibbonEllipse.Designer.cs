namespace EllipseDownLostExcelAddIn
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonEllipse));
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpDownLost = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnFormatDownPbv = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewDL = this.Factory.CreateRibbonButton();
            this.btnReviewDLPbv = this.Factory.CreateRibbonButton();
            this.btnCreateDL = this.Factory.CreateRibbonButton();
            this.btnCreatIgnoreDuplicate = this.Factory.CreateRibbonButton();
            this.btnGenerateCollection = this.Factory.CreateRibbonButton();
            this.btnDeleteDL = this.Factory.CreateRibbonButton();
            this.btnClearTable = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpDownLost.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpDownLost);
            resources.ApplyResources(this.tabEllipse, "tabEllipse");
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpDownLost
            // 
            this.grpDownLost.Items.Add(this.box1);
            this.grpDownLost.Items.Add(this.drpEnvironment);
            this.grpDownLost.Items.Add(this.menuActions);
            resources.ApplyResources(this.grpDownLost, "grpDownLost");
            this.grpDownLost.Name = "grpDownLost";
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatSheet);
            this.menuFormat.Items.Add(this.btnFormatDownPbv);
            resources.ApplyResources(this.menuFormat, "menuFormat");
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatSheet
            // 
            resources.ApplyResources(this.btnFormatSheet, "btnFormatSheet");
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // btnFormatDownPbv
            // 
            resources.ApplyResources(this.btnFormatDownPbv, "btnFormatDownPbv");
            this.btnFormatDownPbv.Name = "btnFormatDownPbv";
            this.btnFormatDownPbv.ShowImage = true;
            this.btnFormatDownPbv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatDownPbv_Click);
            // 
            // btnAbout
            // 
            resources.ApplyResources(this.btnAbout, "btnAbout");
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // drpEnvironment
            // 
            resources.ApplyResources(this.drpEnvironment, "drpEnvironment");
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReviewDL);
            this.menuActions.Items.Add(this.btnReviewDLPbv);
            this.menuActions.Items.Add(this.btnCreateDL);
            this.menuActions.Items.Add(this.btnCreatIgnoreDuplicate);
            this.menuActions.Items.Add(this.btnGenerateCollection);
            this.menuActions.Items.Add(this.btnDeleteDL);
            this.menuActions.Items.Add(this.btnClearTable);
            this.menuActions.Items.Add(this.btnStopThread);
            resources.ApplyResources(this.menuActions, "menuActions");
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewDL
            // 
            resources.ApplyResources(this.btnReviewDL, "btnReviewDL");
            this.btnReviewDL.Name = "btnReviewDL";
            this.btnReviewDL.ShowImage = true;
            this.btnReviewDL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // btnReviewDLPbv
            // 
            resources.ApplyResources(this.btnReviewDLPbv, "btnReviewDLPbv");
            this.btnReviewDLPbv.Name = "btnReviewDLPbv";
            this.btnReviewDLPbv.ShowImage = true;
            this.btnReviewDLPbv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewDLPbv_Click);
            // 
            // btnCreateDL
            // 
            resources.ApplyResources(this.btnCreateDL, "btnCreateDL");
            this.btnCreateDL.Name = "btnCreateDL";
            this.btnCreateDL.ShowImage = true;
            this.btnCreateDL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateDL_Click);
            // 
            // btnCreatIgnoreDuplicate
            // 
            resources.ApplyResources(this.btnCreatIgnoreDuplicate, "btnCreatIgnoreDuplicate");
            this.btnCreatIgnoreDuplicate.Name = "btnCreatIgnoreDuplicate";
            this.btnCreatIgnoreDuplicate.ShowImage = true;
            this.btnCreatIgnoreDuplicate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreatIgnoreDuplicate_Click);
            // 
            // btnGenerateCollection
            // 
            resources.ApplyResources(this.btnGenerateCollection, "btnGenerateCollection");
            this.btnGenerateCollection.Name = "btnGenerateCollection";
            this.btnGenerateCollection.ShowImage = true;
            this.btnGenerateCollection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerateCollection_Click);
            // 
            // btnDeleteDL
            // 
            resources.ApplyResources(this.btnDeleteDL, "btnDeleteDL");
            this.btnDeleteDL.Name = "btnDeleteDL";
            this.btnDeleteDL.ShowImage = true;
            this.btnDeleteDL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteDL_Click);
            // 
            // btnClearTable
            // 
            resources.ApplyResources(this.btnClearTable, "btnClearTable");
            this.btnClearTable.Name = "btnClearTable";
            this.btnClearTable.ShowImage = true;
            this.btnClearTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearTable_Click);
            // 
            // btnStopThread
            // 
            resources.ApplyResources(this.btnStopThread, "btnStopThread");
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
            this.grpDownLost.ResumeLayout(false);
            this.grpDownLost.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDownLost;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewDL;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteDL;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateDL;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreatIgnoreDuplicate;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatDownPbv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewDLPbv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerateCollection;
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

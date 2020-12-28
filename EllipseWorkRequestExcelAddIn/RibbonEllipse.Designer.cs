namespace EllipseWorkRequestExcelAddIn
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
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpWorkRequest = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnFormatMantto = this.Factory.CreateRibbonButton();
            this.btnFormatFcVagones = this.Factory.CreateRibbonButton();
            this.btnPlanFc = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuWorkRequest = this.Factory.CreateRibbonMenu();
            this.btnReviewWorkRequest = this.Factory.CreateRibbonButton();
            this.btnReReviewWorkRequest = this.Factory.CreateRibbonButton();
            this.btnCreateWorkRequest = this.Factory.CreateRibbonButton();
            this.btnModifyWorkRequest = this.Factory.CreateRibbonButton();
            this.btnDeleteWorkRequest = this.Factory.CreateRibbonButton();
            this.menuSla = this.Factory.CreateRibbonMenu();
            this.btnSetSla = this.Factory.CreateRibbonButton();
            this.btnResetSla = this.Factory.CreateRibbonButton();
            this.menuCloseWorkRequest = this.Factory.CreateRibbonMenu();
            this.btnReOpenWorkRequest = this.Factory.CreateRibbonButton();
            this.btnCloseWorkRequest = this.Factory.CreateRibbonButton();
            this.menuReferenceCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewRefCodes = this.Factory.CreateRibbonButton();
            this.btnReReviewRefCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateRefCodes = this.Factory.CreateRibbonButton();
            this.btnCleanSheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkRequest.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menu1
            // 
            resources.ApplyResources(this.menu1, "menu1");
            this.menu1.Name = "menu1";
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpWorkRequest);
            resources.ApplyResources(this.tabEllipse, "tabEllipse");
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpWorkRequest
            // 
            this.grpWorkRequest.Items.Add(this.box1);
            this.grpWorkRequest.Items.Add(this.drpEnvironment);
            this.grpWorkRequest.Items.Add(this.menuActions);
            resources.ApplyResources(this.grpWorkRequest, "grpWorkRequest");
            this.grpWorkRequest.Name = "grpWorkRequest";
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
            this.menuFormat.Items.Add(this.btnFormatMantto);
            this.menuFormat.Items.Add(this.btnFormatFcVagones);
            this.menuFormat.Items.Add(this.btnPlanFc);
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
            // btnFormatMantto
            // 
            resources.ApplyResources(this.btnFormatMantto, "btnFormatMantto");
            this.btnFormatMantto.Name = "btnFormatMantto";
            this.btnFormatMantto.ShowImage = true;
            this.btnFormatMantto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatMantto_Click);
            // 
            // btnFormatFcVagones
            // 
            resources.ApplyResources(this.btnFormatFcVagones, "btnFormatFcVagones");
            this.btnFormatFcVagones.Name = "btnFormatFcVagones";
            this.btnFormatFcVagones.ShowImage = true;
            this.btnFormatFcVagones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatFcVagones_Click);
            // 
            // btnPlanFc
            // 
            resources.ApplyResources(this.btnPlanFc, "btnPlanFc");
            this.btnPlanFc.Name = "btnPlanFc";
            this.btnPlanFc.ShowImage = true;
            this.btnPlanFc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlanFc_Click);
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
            this.menuActions.Items.Add(this.menuWorkRequest);
            this.menuActions.Items.Add(this.menuSla);
            this.menuActions.Items.Add(this.menuCloseWorkRequest);
            this.menuActions.Items.Add(this.menuReferenceCodes);
            this.menuActions.Items.Add(this.btnCleanSheet);
            this.menuActions.Items.Add(this.btnStopThread);
            resources.ApplyResources(this.menuActions, "menuActions");
            this.menuActions.Name = "menuActions";
            // 
            // menuWorkRequest
            // 
            this.menuWorkRequest.Items.Add(this.btnReviewWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnReReviewWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnCreateWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnModifyWorkRequest);
            this.menuWorkRequest.Items.Add(this.btnDeleteWorkRequest);
            resources.ApplyResources(this.menuWorkRequest, "menuWorkRequest");
            this.menuWorkRequest.Name = "menuWorkRequest";
            this.menuWorkRequest.ShowImage = true;
            // 
            // btnReviewWorkRequest
            // 
            resources.ApplyResources(this.btnReviewWorkRequest, "btnReviewWorkRequest");
            this.btnReviewWorkRequest.Name = "btnReviewWorkRequest";
            this.btnReviewWorkRequest.ShowImage = true;
            this.btnReviewWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewWorkRequest_Click);
            // 
            // btnReReviewWorkRequest
            // 
            resources.ApplyResources(this.btnReReviewWorkRequest, "btnReReviewWorkRequest");
            this.btnReReviewWorkRequest.Name = "btnReReviewWorkRequest";
            this.btnReReviewWorkRequest.ShowImage = true;
            this.btnReReviewWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewWorkRequest_Click);
            // 
            // btnCreateWorkRequest
            // 
            resources.ApplyResources(this.btnCreateWorkRequest, "btnCreateWorkRequest");
            this.btnCreateWorkRequest.Name = "btnCreateWorkRequest";
            this.btnCreateWorkRequest.ShowImage = true;
            this.btnCreateWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateWorkRequest_Click);
            // 
            // btnModifyWorkRequest
            // 
            resources.ApplyResources(this.btnModifyWorkRequest, "btnModifyWorkRequest");
            this.btnModifyWorkRequest.Name = "btnModifyWorkRequest";
            this.btnModifyWorkRequest.ShowImage = true;
            this.btnModifyWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifyWorkRequest_Click);
            // 
            // btnDeleteWorkRequest
            // 
            resources.ApplyResources(this.btnDeleteWorkRequest, "btnDeleteWorkRequest");
            this.btnDeleteWorkRequest.Name = "btnDeleteWorkRequest";
            this.btnDeleteWorkRequest.ShowImage = true;
            this.btnDeleteWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteWorkRequest_Click);
            // 
            // menuSla
            // 
            this.menuSla.Items.Add(this.btnSetSla);
            this.menuSla.Items.Add(this.btnResetSla);
            resources.ApplyResources(this.menuSla, "menuSla");
            this.menuSla.Name = "menuSla";
            this.menuSla.ShowImage = true;
            // 
            // btnSetSla
            // 
            resources.ApplyResources(this.btnSetSla, "btnSetSla");
            this.btnSetSla.Name = "btnSetSla";
            this.btnSetSla.ShowImage = true;
            this.btnSetSla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetSla_Click);
            // 
            // btnResetSla
            // 
            resources.ApplyResources(this.btnResetSla, "btnResetSla");
            this.btnResetSla.Name = "btnResetSla";
            this.btnResetSla.ShowImage = true;
            this.btnResetSla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetSla_Click);
            // 
            // menuCloseWorkRequest
            // 
            this.menuCloseWorkRequest.Items.Add(this.btnReOpenWorkRequest);
            this.menuCloseWorkRequest.Items.Add(this.btnCloseWorkRequest);
            resources.ApplyResources(this.menuCloseWorkRequest, "menuCloseWorkRequest");
            this.menuCloseWorkRequest.Name = "menuCloseWorkRequest";
            this.menuCloseWorkRequest.ShowImage = true;
            // 
            // btnReOpenWorkRequest
            // 
            resources.ApplyResources(this.btnReOpenWorkRequest, "btnReOpenWorkRequest");
            this.btnReOpenWorkRequest.Name = "btnReOpenWorkRequest";
            this.btnReOpenWorkRequest.ShowImage = true;
            this.btnReOpenWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReOpenWorkRequest_Click);
            // 
            // btnCloseWorkRequest
            // 
            resources.ApplyResources(this.btnCloseWorkRequest, "btnCloseWorkRequest");
            this.btnCloseWorkRequest.Name = "btnCloseWorkRequest";
            this.btnCloseWorkRequest.ShowImage = true;
            this.btnCloseWorkRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCloseWorkRequest_Click);
            // 
            // menuReferenceCodes
            // 
            this.menuReferenceCodes.Items.Add(this.btnReviewRefCodes);
            this.menuReferenceCodes.Items.Add(this.btnReReviewRefCodes);
            this.menuReferenceCodes.Items.Add(this.btnUpdateRefCodes);
            resources.ApplyResources(this.menuReferenceCodes, "menuReferenceCodes");
            this.menuReferenceCodes.Name = "menuReferenceCodes";
            this.menuReferenceCodes.ShowImage = true;
            // 
            // btnReviewRefCodes
            // 
            resources.ApplyResources(this.btnReviewRefCodes, "btnReviewRefCodes");
            this.btnReviewRefCodes.Name = "btnReviewRefCodes";
            this.btnReviewRefCodes.ShowImage = true;
            this.btnReviewRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRefCodes_Click);
            // 
            // btnReReviewRefCodes
            // 
            resources.ApplyResources(this.btnReReviewRefCodes, "btnReReviewRefCodes");
            this.btnReReviewRefCodes.Name = "btnReReviewRefCodes";
            this.btnReReviewRefCodes.ShowImage = true;
            this.btnReReviewRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewRefCodes_Click);
            // 
            // btnUpdateRefCodes
            // 
            resources.ApplyResources(this.btnUpdateRefCodes, "btnUpdateRefCodes");
            this.btnUpdateRefCodes.Name = "btnUpdateRefCodes";
            this.btnUpdateRefCodes.ShowImage = true;
            this.btnUpdateRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRefCodes_Click);
            // 
            // btnCleanSheet
            // 
            resources.ApplyResources(this.btnCleanSheet, "btnCleanSheet");
            this.btnCleanSheet.Name = "btnCleanSheet";
            this.btnCleanSheet.ShowImage = true;
            this.btnCleanSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanSheet_Click);
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
            this.grpWorkRequest.ResumeLayout(false);
            this.grpWorkRequest.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetSla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloseWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCloseWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReOpenWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuWorkRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuReferenceCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatMantto;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatFcVagones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanFc;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

namespace EllipseStdTextExcelAddIn
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpStdText = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuStdText = this.Factory.CreateRibbonMenu();
            this.btnGetHeaderAndText = this.Factory.CreateRibbonButton();
            this.btnUpdateHeaderAndText = this.Factory.CreateRibbonButton();
            this.btnGetHeaderOnly = this.Factory.CreateRibbonButton();
            this.btnSetHeaderOnly = this.Factory.CreateRibbonButton();
            this.btnGetTextOnly = this.Factory.CreateRibbonButton();
            this.btnSetTextOnly = this.Factory.CreateRibbonButton();
            this.menuRefCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewRefCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateRefCodes = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnCleanTable = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpStdText.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpStdText);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpStdText
            // 
            this.grpStdText.Items.Add(this.box1);
            this.grpStdText.Label = "StdText";
            this.grpStdText.Name = "grpStdText";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.drpEnvironment);
            this.box1.Items.Add(this.menuActions);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.btnFormatSheet);
            this.box2.Items.Add(this.btnAbout);
            this.box2.Name = "box2";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
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
            this.menuActions.Items.Add(this.menuStdText);
            this.menuActions.Items.Add(this.menuRefCodes);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Items.Add(this.btnCleanTable);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuStdText
            // 
            this.menuStdText.Items.Add(this.btnGetHeaderAndText);
            this.menuStdText.Items.Add(this.btnUpdateHeaderAndText);
            this.menuStdText.Items.Add(this.btnGetHeaderOnly);
            this.menuStdText.Items.Add(this.btnSetHeaderOnly);
            this.menuStdText.Items.Add(this.btnGetTextOnly);
            this.menuStdText.Items.Add(this.btnSetTextOnly);
            this.menuStdText.Label = "&Std Text";
            this.menuStdText.Name = "menuStdText";
            this.menuStdText.ShowImage = true;
            // 
            // btnGetHeaderAndText
            // 
            this.btnGetHeaderAndText.Label = "Consultar Encabezado y Texto";
            this.btnGetHeaderAndText.Name = "btnGetHeaderAndText";
            this.btnGetHeaderAndText.ShowImage = true;
            this.btnGetHeaderAndText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetHeaderAndText_Click);
            // 
            // btnUpdateHeaderAndText
            // 
            this.btnUpdateHeaderAndText.Label = "Actualizar Encabezado y Texto";
            this.btnUpdateHeaderAndText.Name = "btnUpdateHeaderAndText";
            this.btnUpdateHeaderAndText.ShowImage = true;
            this.btnUpdateHeaderAndText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateHeaderAndText_Click);
            // 
            // btnGetHeaderOnly
            // 
            this.btnGetHeaderOnly.Label = "Consultar Encabezado";
            this.btnGetHeaderOnly.Name = "btnGetHeaderOnly";
            this.btnGetHeaderOnly.ShowImage = true;
            this.btnGetHeaderOnly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetHeaderOnly_Click);
            // 
            // btnSetHeaderOnly
            // 
            this.btnSetHeaderOnly.Label = "Actualizar Encabezado";
            this.btnSetHeaderOnly.Name = "btnSetHeaderOnly";
            this.btnSetHeaderOnly.ShowImage = true;
            this.btnSetHeaderOnly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetHeaderOnly_Click);
            // 
            // btnGetTextOnly
            // 
            this.btnGetTextOnly.Label = "Consultar Texto";
            this.btnGetTextOnly.Name = "btnGetTextOnly";
            this.btnGetTextOnly.ShowImage = true;
            this.btnGetTextOnly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetTextOnly_Click);
            // 
            // btnSetTextOnly
            // 
            this.btnSetTextOnly.Label = "Actualizar Texto";
            this.btnSetTextOnly.Name = "btnSetTextOnly";
            this.btnSetTextOnly.ShowImage = true;
            this.btnSetTextOnly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetTextOnly_Click);
            // 
            // menuRefCodes
            // 
            this.menuRefCodes.Items.Add(this.btnReviewRefCodes);
            this.menuRefCodes.Items.Add(this.btnUpdateRefCodes);
            this.menuRefCodes.Label = "&Reference Codes";
            this.menuRefCodes.Name = "menuRefCodes";
            this.menuRefCodes.ShowImage = true;
            // 
            // btnReviewRefCodes
            // 
            this.btnReviewRefCodes.Label = "&Consultar RefCodes";
            this.btnReviewRefCodes.Name = "btnReviewRefCodes";
            this.btnReviewRefCodes.ShowImage = true;
            this.btnReviewRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRefCodes_Click);
            // 
            // btnUpdateRefCodes
            // 
            this.btnUpdateRefCodes.Label = "Actualizar RefCodes";
            this.btnUpdateRefCodes.Name = "btnUpdateRefCodes";
            this.btnUpdateRefCodes.ShowImage = true;
            this.btnUpdateRefCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRefCodes_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnCleanTable
            // 
            this.btnCleanTable.Label = "&Limpiar Tabla";
            this.btnCleanTable.Name = "btnCleanTable";
            this.btnCleanTable.ShowImage = true;
            this.btnCleanTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanTable_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpStdText.ResumeLayout(false);
            this.grpStdText.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStdText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetHeaderAndText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateHeaderAndText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetHeaderOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetHeaderOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetTextOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetTextOnly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuStdText;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

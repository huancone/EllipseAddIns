namespace EllipseCreateJournalExcelAddIn
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpCreateJournal = this.Factory.CreateRibbonGroup();
            this.btnCreateJournalFormat = this.Factory.CreateRibbonButton();
            this.drpCreateJournalEnv = this.Factory.CreateRibbonDropDown();
            this.btnCreateJournalExecute = this.Factory.CreateRibbonButton();
            this.btnValidateAccountCode = this.Factory.CreateRibbonButton();
            this.btnValidateNit = this.Factory.CreateRibbonButton();
            this.btnPesos = this.Factory.CreateRibbonButton();
            this.btnDolares = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpCreateJournal.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpCreateJournal);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpCreateJournal
            // 
            this.grpCreateJournal.Items.Add(this.btnCreateJournalFormat);
            this.grpCreateJournal.Items.Add(this.drpCreateJournalEnv);
            this.grpCreateJournal.Items.Add(this.btnCreateJournalExecute);
            this.grpCreateJournal.Items.Add(this.btnValidateAccountCode);
            this.grpCreateJournal.Items.Add(this.btnValidateNit);
            this.grpCreateJournal.Items.Add(this.btnPesos);
            this.grpCreateJournal.Items.Add(this.btnDolares);
            this.grpCreateJournal.Label = "MSO905 Opc 3 v 1.0.1";
            this.grpCreateJournal.Name = "grpCreateJournal";
            // 
            // btnCreateJournalFormat
            // 
            this.btnCreateJournalFormat.Label = "Format New Sheet";
            this.btnCreateJournalFormat.Name = "btnCreateJournalFormat";
            this.btnCreateJournalFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateJournalFormat_Click);
            // 
            // drpCreateJournalEnv
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Contingencia";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            ribbonDropDownItemImpl4.Label = "Test";
            this.drpCreateJournalEnv.Items.Add(ribbonDropDownItemImpl1);
            this.drpCreateJournalEnv.Items.Add(ribbonDropDownItemImpl2);
            this.drpCreateJournalEnv.Items.Add(ribbonDropDownItemImpl3);
            this.drpCreateJournalEnv.Items.Add(ribbonDropDownItemImpl4);
            this.drpCreateJournalEnv.Label = "Env.";
            this.drpCreateJournalEnv.Name = "drpCreateJournalEnv";
            // 
            // btnCreateJournalExecute
            // 
            this.btnCreateJournalExecute.Label = "Execute CreateJournal";
            this.btnCreateJournalExecute.Name = "btnCreateJournalExecute";
            this.btnCreateJournalExecute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateJournalExecute_Click);
            // 
            // btnValidateAccountCode
            // 
            this.btnValidateAccountCode.Label = "Validate Account Code";
            this.btnValidateAccountCode.Name = "btnValidateAccountCode";
            this.btnValidateAccountCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateAccountCode_Click);
            // 
            // btnValidateNit
            // 
            this.btnValidateNit.Label = "Validate Nit";
            this.btnValidateNit.Name = "btnValidateNit";
            this.btnValidateNit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateNit_Click);
            // 
            // btnPesos
            // 
            this.btnPesos.Label = "Pesos";
            this.btnPesos.Name = "btnPesos";
            this.btnPesos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPesos_Click);
            // 
            // btnDolares
            // 
            this.btnDolares.Label = "Dollars";
            this.btnDolares.Name = "btnDolares";
            this.btnDolares.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDolares_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpCreateJournal.ResumeLayout(false);
            this.grpCreateJournal.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCreateJournal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateJournalFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateJournalExecute;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpCreateJournalEnv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDolares;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPesos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateNit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateAccountCode;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

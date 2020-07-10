namespace ContratosGeneradorDocumentos
{
    partial class RibbonContracts : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonContracts()
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
            this.grpDocumentGenerator = this.Factory.CreateRibbonGroup();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAction = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpDocumentGenerator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpDocumentGenerator);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpDocumentGenerator
            // 
            this.grpDocumentGenerator.Items.Add(this.btnFormat);
            this.grpDocumentGenerator.Items.Add(this.btnAction);
            this.grpDocumentGenerator.Label = "Document Generator";
            this.grpDocumentGenerator.Name = "grpDocumentGenerator";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "Formatear";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // btnAction
            // 
            this.btnAction.Label = "Acción";
            this.btnAction.Name = "btnAction";
            this.btnAction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAction_Click);
            // 
            // RibbonContracts
            // 
            this.Name = "RibbonContracts";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonContracts_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpDocumentGenerator.ResumeLayout(false);
            this.grpDocumentGenerator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDocumentGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonContracts RibbonContracts
        {
            get { return this.GetRibbon<RibbonContracts>(); }
        }
    }
}

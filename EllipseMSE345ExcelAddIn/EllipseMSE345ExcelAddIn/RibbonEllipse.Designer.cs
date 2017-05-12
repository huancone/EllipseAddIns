namespace EllipseMSE345ExcelAddIn
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
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpCondMonit = this.Factory.CreateRibbonGroup();
            this.mFormatear = this.Factory.CreateRibbonMenu();
            this.Formatear = this.Factory.CreateRibbonButton();
            this.fMantto = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.Crear = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpCondMonit.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpCondMonit);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpCondMonit
            // 
            this.grpCondMonit.Items.Add(this.box1);
            this.grpCondMonit.Items.Add(this.drpEnviroment);
            this.grpCondMonit.Items.Add(this.Crear);
            this.grpCondMonit.Label = "MSE345";
            this.grpCondMonit.Name = "grpCondMonit";
            // 
            // mFormatear
            // 
            this.mFormatear.Items.Add(this.Formatear);
            this.mFormatear.Items.Add(this.fMantto);
            this.mFormatear.Label = "Formatear";
            this.mFormatear.Name = "mFormatear";
            // 
            // Formatear
            // 
            this.Formatear.Label = "Formato Estandar";
            this.Formatear.Name = "Formatear";
            this.Formatear.ShowImage = true;
            this.Formatear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Formatear_Click);
            // 
            // fMantto
            // 
            this.fMantto.Label = "Formato Mantto";
            this.fMantto.Name = "fMantto";
            this.fMantto.ShowImage = true;
            this.fMantto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fMantto_Click);
            // 
            // drpEnviroment
            // 
            ribbonDropDownItemImpl1.Label = "Productivo";
            ribbonDropDownItemImpl2.Label = "Test";
            ribbonDropDownItemImpl3.Label = "Desarrollo";
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl1);
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl2);
            this.drpEnviroment.Items.Add(ribbonDropDownItemImpl3);
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // Crear
            // 
            this.Crear.Label = "Cargar Info";
            this.Crear.Name = "Crear";
            this.Crear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Crear_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.mFormatear);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpCondMonit.ResumeLayout(false);
            this.grpCondMonit.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCondMonit;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Crear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Formatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mFormatear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fMantto;
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

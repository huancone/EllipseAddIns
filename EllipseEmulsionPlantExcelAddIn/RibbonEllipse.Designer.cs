namespace EllipseEmulsionPlantExcelAddIn
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
            this.grpEmulsionPlant = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnLoadEmulsion = this.Factory.CreateRibbonButton();
            this.btnLoadSolox = this.Factory.CreateRibbonButton();
            this.btnGetModuleEmulsion = this.Factory.CreateRibbonButton();
            this.btnGetModuleSolox = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEmulsionPlant.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEmulsionPlant);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEmulsionPlant
            // 
            this.grpEmulsionPlant.Items.Add(this.box1);
            this.grpEmulsionPlant.Items.Add(this.drpEnvironment);
            this.grpEmulsionPlant.Items.Add(this.menuActions);
            this.grpEmulsionPlant.Label = "Planta de Emulsión";
            this.grpEmulsionPlant.Name = "grpEmulsionPlant";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "&Formatear";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
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
            this.menuActions.Items.Add(this.btnLoadEmulsion);
            this.menuActions.Items.Add(this.btnLoadSolox);
            this.menuActions.Items.Add(this.btnGetModuleEmulsion);
            this.menuActions.Items.Add(this.btnGetModuleSolox);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnLoadEmulsion
            // 
            this.btnLoadEmulsion.Label = "Cargar &Emulsión";
            this.btnLoadEmulsion.Name = "btnLoadEmulsion";
            this.btnLoadEmulsion.ShowImage = true;
            this.btnLoadEmulsion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadEmulsion_Click);
            // 
            // btnLoadSolox
            // 
            this.btnLoadSolox.Label = "Cargar &Solución Oxidante";
            this.btnLoadSolox.Name = "btnLoadSolox";
            this.btnLoadSolox.ShowImage = true;
            this.btnLoadSolox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadSolox_Click);
            // 
            // btnGetModuleEmulsion
            // 
            this.btnGetModuleEmulsion.Label = "Obtener Datos de Emulsión";
            this.btnGetModuleEmulsion.Name = "btnGetModuleEmulsion";
            this.btnGetModuleEmulsion.ShowImage = true;
            this.btnGetModuleEmulsion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetModuleEmulsion_Click);
            // 
            // btnGetModuleSolox
            // 
            this.btnGetModuleSolox.Label = "Obtener Datos de Solución";
            this.btnGetModuleSolox.Name = "btnGetModuleSolox";
            this.btnGetModuleSolox.ShowImage = true;
            this.btnGetModuleSolox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetModuleSolox_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpEmulsionPlant.ResumeLayout(false);
            this.grpEmulsionPlant.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEmulsionPlant;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadEmulsion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadSolox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetModuleEmulsion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetModuleSolox;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

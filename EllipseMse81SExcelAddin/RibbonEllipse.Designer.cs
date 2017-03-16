namespace EllipseMse81SExcelAddin
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
            this.grpMse81s = this.Factory.CreateRibbonGroup();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnCreateEmployee = this.Factory.CreateRibbonButton();
            this.btnReReviewEmployees = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMse81s.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMse81s);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMse81s
            // 
            this.grpMse81s.Items.Add(this.btnFormatSheet);
            this.grpMse81s.Items.Add(this.drpEnviroment);
            this.grpMse81s.Items.Add(this.menuActions);
            this.grpMse81s.Label = "Mse81S v.1.0.0";
            this.grpMse81s.Name = "grpMse81s";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formatear Hoja";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnCreateEmployee);
            this.menuActions.Items.Add(this.btnReReviewEmployees);
            this.menuActions.Items.Add(this.btnUpdate);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnCreateEmployee
            // 
            this.btnCreateEmployee.Label = "&Crear Registros";
            this.btnCreateEmployee.Name = "btnCreateEmployee";
            this.btnCreateEmployee.ShowImage = true;
            this.btnCreateEmployee.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateEmployee_Click);
            // 
            // btnReReviewEmployees
            // 
            this.btnReReviewEmployees.Label = "&Reconsultar Empleados";
            this.btnReReviewEmployees.Name = "btnReReviewEmployees";
            this.btnReReviewEmployees.ShowImage = true;
            this.btnReReviewEmployees.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewEmployees_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Label = "&Actualizar Registros";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
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
            this.grpMse81s.ResumeLayout(false);
            this.grpMse81s.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMse81s;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateEmployee;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReviewEmployees;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

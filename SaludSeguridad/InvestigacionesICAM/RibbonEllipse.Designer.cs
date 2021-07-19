
namespace InvestigacionesICAM
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
            this.grpInvestigacionesIcam = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReviewAll = this.Factory.CreateRibbonButton();
            this.btnReviewAccidents = this.Factory.CreateRibbonButton();
            this.btnReviewRecomendations = this.Factory.CreateRibbonButton();
            this.btnReviewPlans = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpInvestigacionesIcam.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpInvestigacionesIcam);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpInvestigacionesIcam
            // 
            this.grpInvestigacionesIcam.Items.Add(this.box1);
            this.grpInvestigacionesIcam.Items.Add(this.drpEnvironment);
            this.grpInvestigacionesIcam.Items.Add(this.menuActions);
            this.grpInvestigacionesIcam.Label = "Investigaciones ICAM";
            this.grpInvestigacionesIcam.Name = "grpInvestigacionesIcam";
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
            this.drpEnvironment.Label = "&Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReviewAll);
            this.menuActions.Items.Add(this.btnReviewAccidents);
            this.menuActions.Items.Add(this.btnReviewRecomendations);
            this.menuActions.Items.Add(this.btnReviewPlans);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReviewAll
            // 
            this.btnReviewAll.Label = "&Consultar Todo";
            this.btnReviewAll.Name = "btnReviewAll";
            this.btnReviewAll.ShowImage = true;
            this.btnReviewAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecution_Click);
            // 
            // btnReviewAccidents
            // 
            this.btnReviewAccidents.Label = "Consultar &Accidentes";
            this.btnReviewAccidents.Name = "btnReviewAccidents";
            this.btnReviewAccidents.ShowImage = true;
            this.btnReviewAccidents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewAccidents_Click);
            // 
            // btnReviewRecomendations
            // 
            this.btnReviewRecomendations.Label = "Consultar &Recomendaciones";
            this.btnReviewRecomendations.Name = "btnReviewRecomendations";
            this.btnReviewRecomendations.ShowImage = true;
            this.btnReviewRecomendations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRecomendations_Click);
            // 
            // btnReviewPlans
            // 
            this.btnReviewPlans.Label = "Consultar &Planes de Acción";
            this.btnReviewPlans.Name = "btnReviewPlans";
            this.btnReviewPlans.ShowImage = true;
            this.btnReviewPlans.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewPlans_Click);
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
            this.grpInvestigacionesIcam.ResumeLayout(false);
            this.grpInvestigacionesIcam.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInvestigacionesIcam;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewAccidents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewRecomendations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewPlans;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

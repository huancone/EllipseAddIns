using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace EllipseFinalizeWorkOrderExcelAddIn
{
    partial class RibbonEllipse : RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.grpWorkOrder = this.Factory.CreateRibbonGroup();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnReReview = this.Factory.CreateRibbonButton();
            this.btnFinalize = this.Factory.CreateRibbonButton();
            this.btnCleanWorkOrderSheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.box4 = this.Factory.CreateRibbonBox();
            this.box5 = this.Factory.CreateRibbonBox();
            this.butAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkOrder.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.box4.SuspendLayout();
            this.box5.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpWorkOrder);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpWorkOrder
            // 
            this.grpWorkOrder.Items.Add(this.box1);
            this.grpWorkOrder.Items.Add(this.box2);
            this.grpWorkOrder.Items.Add(this.box3);
            this.grpWorkOrder.Label = "Finalize Orders";
            this.grpWorkOrder.Name = "grpWorkOrder";
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
            this.drpEnviroment.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drpEnviroment_SelectionChanged);
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReview);
            this.menuActions.Items.Add(this.btnReReview);
            this.menuActions.Items.Add(this.btnFinalize);
            this.menuActions.Items.Add(this.btnCleanWorkOrderSheet);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReview
            // 
            this.btnReview.Label = "Consultar OTs";
            this.btnReview.Name = "btnReview";
            this.btnReview.ShowImage = true;
            this.btnReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // btnReReview
            // 
            this.btnReReview.Label = "ReConsultar OTs Tabla";
            this.btnReReview.Name = "btnReReview";
            this.btnReReview.ShowImage = true;
            this.btnReReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReview_Click);
            // 
            // btnFinalize
            // 
            this.btnFinalize.Label = "Finali&zar OTs";
            this.btnFinalize.Name = "btnFinalize";
            this.btnFinalize.ShowImage = true;
            this.btnFinalize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFinalize_Click);
            // 
            // btnCleanWorkOrderSheet
            // 
            this.btnCleanWorkOrderSheet.Label = "&Limpiar Hoja";
            this.btnCleanWorkOrderSheet.Name = "btnCleanWorkOrderSheet";
            this.btnCleanWorkOrderSheet.ShowImage = true;
            this.btnCleanWorkOrderSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanWorkOrderSheet_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener &Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.box4);
            this.box1.Items.Add(this.box5);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.drpEnviroment);
            this.box2.Name = "box2";
            // 
            // box3
            // 
            this.box3.Items.Add(this.menuActions);
            this.box3.Name = "box3";
            // 
            // box4
            // 
            this.box4.Items.Add(this.btnFormatSheet);
            this.box4.Name = "box4";
            // 
            // box5
            // 
            this.box5.Items.Add(this.butAbout);
            this.box5.Name = "box5";
            // 
            // butAbout
            // 
            this.butAbout.Label = "?";
            this.butAbout.Name = "butAbout";
            this.butAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butAbout_Click);
            // 
            // RibbonEllipse
            // 
            this.Name = "RibbonEllipse";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEllipse);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEllipse_Load);
            this.tabEllipse.ResumeLayout(false);
            this.tabEllipse.PerformLayout();
            this.grpWorkOrder.ResumeLayout(false);
            this.grpWorkOrder.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();

        }

        #endregion

        internal RibbonTab tabEllipse;
        internal RibbonGroup grpWorkOrder;
        internal RibbonDropDown drpEnviroment;
        internal RibbonButton btnReview;
        internal RibbonButton btnFormatSheet;
        internal RibbonMenu menuActions;
        internal RibbonButton btnFinalize;
        internal RibbonButton btnReReview;
        internal RibbonButton btnCleanWorkOrderSheet;
        internal RibbonButton btnStopThread;
        internal RibbonBox box1;
        internal RibbonBox box4;
        internal RibbonBox box5;
        internal RibbonButton butAbout;
        internal RibbonBox box2;
        internal RibbonBox box3;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

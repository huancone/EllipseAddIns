using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace EllipseMSO010ExcelAddIn
{
    partial class RibbonEllipse : RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private IContainer components = null;

        public RibbonEllipse()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; en caso contrario, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabEllipse = this.Factory.CreateRibbonTab();
            this.grpMSO010 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.box5 = this.Factory.CreateRibbonBox();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.box6 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.box3 = this.Factory.CreateRibbonBox();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.box4 = this.Factory.CreateRibbonBox();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnCreate = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpMSO010.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box5.SuspendLayout();
            this.box6.SuspendLayout();
            this.box3.SuspendLayout();
            this.box4.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpMSO010);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpMSO010
            // 
            this.grpMSO010.Items.Add(this.box1);
            this.grpMSO010.Label = "MSO010 Codes";
            this.grpMSO010.Name = "grpMSO010";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.box3);
            this.box1.Items.Add(this.box4);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.box5);
            this.box2.Items.Add(this.box6);
            this.box2.Name = "box2";
            // 
            // box5
            // 
            this.box5.Items.Add(this.btnFormatSheet);
            this.box5.Name = "box5";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "Formatear";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatCesantias_Click);
            // 
            // box6
            // 
            this.box6.Items.Add(this.btnAbout);
            this.box6.Name = "box6";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // box3
            // 
            this.box3.Items.Add(this.drpEnviroment);
            this.box3.Name = "box3";
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // box4
            // 
            this.box4.Items.Add(this.menuActions);
            this.box4.Name = "box4";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnReview);
            this.menuActions.Items.Add(this.btnCreate);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnReview
            // 
            this.btnReview.Label = "Consulta&r";
            this.btnReview.Name = "btnReview";
            this.btnReview.ShowImage = true;
            this.btnReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // btnCreate
            // 
            this.btnCreate.Label = "&Crear Registro";
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.ShowImage = true;
            this.btnCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreate_Click);
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
            this.grpMSO010.ResumeLayout(false);
            this.grpMSO010.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();

        }

        #endregion

        internal RibbonTab tabEllipse;
        internal RibbonGroup grpMSO010;
        internal RibbonDropDown drpEnviroment;
        internal RibbonMenu menuActions;
        internal RibbonButton btnReview;
        internal RibbonButton btnFormatSheet;
        internal RibbonButton btnCreate;
        internal RibbonButton btnStopThread;
        internal RibbonBox box1;
        internal RibbonBox box2;
        internal RibbonBox box5;
        internal RibbonBox box6;
        internal RibbonButton btnAbout;
        internal RibbonBox box3;
        internal RibbonBox box4;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

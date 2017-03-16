namespace EllipseEquipmentExcelAddIn
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
            this.grpEllipse = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.menuFormatSheet = this.Factory.CreateRibbonMenu();
            this.btnFormatFull = this.Factory.CreateRibbonButton();
            this.box3 = this.Factory.CreateRibbonBox();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.box4 = this.Factory.CreateRibbonBox();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuEquipments = this.Factory.CreateRibbonMenu();
            this.btnCreateEquipment = this.Factory.CreateRibbonButton();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnReReview = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentData = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentStatus = this.Factory.CreateRibbonButton();
            this.btnDeleteEquipment = this.Factory.CreateRibbonButton();
            this.menuCompMovement = this.Factory.CreateRibbonMenu();
            this.btnTraceAction = this.Factory.CreateRibbonButton();
            this.btnReviewFitments = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.box5 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.box4.SuspendLayout();
            this.box5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEllipse);
            this.tabEllipse.Label = "ELLIPSE 8";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.box1);
            this.grpEllipse.Label = "Equipments v1.0.0";
            this.grpEllipse.Name = "grpEllipse";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box2);
            this.box1.Items.Add(this.box3);
            this.box1.Items.Add(this.box4);
            this.box1.Items.Add(this.box5);
            this.box1.Name = "box1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.menuFormatSheet);
            this.box2.Name = "box2";
            // 
            // menuFormatSheet
            // 
            this.menuFormatSheet.Items.Add(this.btnFormatFull);
            this.menuFormatSheet.Label = "&Formatear Hoja";
            this.menuFormatSheet.Name = "menuFormatSheet";
            // 
            // btnFormatFull
            // 
            this.btnFormatFull.Label = "Formato &Completo";
            this.btnFormatFull.Name = "btnFormatFull";
            this.btnFormatFull.ShowImage = true;
            this.btnFormatFull.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatFull_Click);
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
            this.menuActions.Items.Add(this.menuEquipments);
            this.menuActions.Items.Add(this.menuCompMovement);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuEquipments
            // 
            this.menuEquipments.Items.Add(this.btnCreateEquipment);
            this.menuEquipments.Items.Add(this.btnReview);
            this.menuEquipments.Items.Add(this.btnReReview);
            this.menuEquipments.Items.Add(this.btnUpdateEquipmentData);
            this.menuEquipments.Items.Add(this.btnUpdateEquipmentStatus);
            this.menuEquipments.Items.Add(this.btnDeleteEquipment);
            this.menuEquipments.Label = "&Equipos";
            this.menuEquipments.Name = "menuEquipments";
            this.menuEquipments.ShowImage = true;
            // 
            // btnCreateEquipment
            // 
            this.btnCreateEquipment.Label = "&Crear Equipo";
            this.btnCreateEquipment.Name = "btnCreateEquipment";
            this.btnCreateEquipment.ShowImage = true;
            this.btnCreateEquipment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateEquipment_Click);
            // 
            // btnReview
            // 
            this.btnReview.Label = "&Consultar Equipos";
            this.btnReview.Name = "btnReview";
            this.btnReview.ShowImage = true;
            this.btnReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReview_Click);
            // 
            // btnReReview
            // 
            this.btnReReview.Label = "&Reconsultar Equipos";
            this.btnReReview.Name = "btnReReview";
            this.btnReReview.ShowImage = true;
            this.btnReReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReview_Click);
            // 
            // btnUpdateEquipmentData
            // 
            this.btnUpdateEquipmentData.Label = "&Actualizar Datos Equipo";
            this.btnUpdateEquipmentData.Name = "btnUpdateEquipmentData";
            this.btnUpdateEquipmentData.ShowImage = true;
            this.btnUpdateEquipmentData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateEquipmentData_Click);
            // 
            // btnUpdateEquipmentStatus
            // 
            this.btnUpdateEquipmentStatus.Label = "&Actualizar Estado Equipo";
            this.btnUpdateEquipmentStatus.Name = "btnUpdateEquipmentStatus";
            this.btnUpdateEquipmentStatus.ShowImage = true;
            this.btnUpdateEquipmentStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateEquipmentStatus_Click);
            // 
            // btnDeleteEquipment
            // 
            this.btnDeleteEquipment.Label = "&Eliminar Equipo";
            this.btnDeleteEquipment.Name = "btnDeleteEquipment";
            this.btnDeleteEquipment.ShowImage = true;
            this.btnDeleteEquipment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteEquipment_Click);
            // 
            // menuCompMovement
            // 
            this.menuCompMovement.Items.Add(this.btnTraceAction);
            this.menuCompMovement.Items.Add(this.btnReviewFitments);
            this.menuCompMovement.Label = "&Mov. Componentes";
            this.menuCompMovement.Name = "menuCompMovement";
            this.menuCompMovement.ShowImage = true;
            // 
            // btnTraceAction
            // 
            this.btnTraceAction.Label = "&Realizar Acción";
            this.btnTraceAction.Name = "btnTraceAction";
            this.btnTraceAction.ShowImage = true;
            this.btnTraceAction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDefitment_Click);
            // 
            // btnReviewFitments
            // 
            this.btnReviewFitments.Label = "Consultar Historia";
            this.btnReviewFitments.Name = "btnReviewFitments";
            this.btnReviewFitments.ShowImage = true;
            this.btnReviewFitments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewFitments_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "&Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // box5
            // 
            this.box5.Items.Add(this.btnAbout);
            this.box5.Name = "box5";
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
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatFull;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnviroment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateEquipmentData;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuFormatSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateEquipment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteEquipment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateEquipmentStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCompMovement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTraceAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuEquipments;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewFitments;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

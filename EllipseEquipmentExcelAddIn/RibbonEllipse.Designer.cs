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
            this.box5 = this.Factory.CreateRibbonBox();
            this.menuFormatSheet = this.Factory.CreateRibbonMenu();
            this.btnFormatFull = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuEquipments = this.Factory.CreateRibbonMenu();
            this.btnCreateEquipment = this.Factory.CreateRibbonButton();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnReReview = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentData = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentStatus = this.Factory.CreateRibbonButton();
            this.btnDisposal = this.Factory.CreateRibbonButton();
            this.btnDeleteEquipment = this.Factory.CreateRibbonButton();
            this.menuListEquipment = this.Factory.CreateRibbonMenu();
            this.btnReviewListEquips = this.Factory.CreateRibbonButton();
            this.btnReviewFromEquipmentList = this.Factory.CreateRibbonButton();
            this.btnAddEquipToList = this.Factory.CreateRibbonButton();
            this.btnDeleteEquipFromList = this.Factory.CreateRibbonButton();
            this.menuCompMovement = this.Factory.CreateRibbonMenu();
            this.btnTraceAction = this.Factory.CreateRibbonButton();
            this.btnReviewCurrentFitment = this.Factory.CreateRibbonButton();
            this.btnDeleteAction = this.Factory.CreateRibbonButton();
            this.cbIgnoreRefCodes = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEllipse.SuspendLayout();
            this.box1.SuspendLayout();
            this.box5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpEllipse);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpEllipse
            // 
            this.grpEllipse.Items.Add(this.box1);
            this.grpEllipse.Label = "Equipments";
            this.grpEllipse.Name = "grpEllipse";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.box5);
            this.box1.Items.Add(this.drpEnvironment);
            this.box1.Items.Add(this.menuActions);
            this.box1.Name = "box1";
            // 
            // box5
            // 
            this.box5.Items.Add(this.menuFormatSheet);
            this.box5.Items.Add(this.btnAbout);
            this.box5.Name = "box5";
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
            this.menuActions.Items.Add(this.menuEquipments);
            this.menuActions.Items.Add(this.menuListEquipment);
            this.menuActions.Items.Add(this.menuCompMovement);
            this.menuActions.Items.Add(this.cbIgnoreRefCodes);
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
            this.menuEquipments.Items.Add(this.btnDisposal);
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
            this.btnReview.Label = "C&onsultar Equipos";
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
            this.btnUpdateEquipmentStatus.Label = "Actualizar E&stado Equipo";
            this.btnUpdateEquipmentStatus.Name = "btnUpdateEquipmentStatus";
            this.btnUpdateEquipmentStatus.ShowImage = true;
            this.btnUpdateEquipmentStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateEquipmentStatus_Click);
            // 
            // btnDisposal
            // 
            this.btnDisposal.Label = "&Disposal";
            this.btnDisposal.Name = "btnDisposal";
            this.btnDisposal.ShowImage = true;
            this.btnDisposal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisposal_Click);
            // 
            // btnDeleteEquipment
            // 
            this.btnDeleteEquipment.Label = "&Eliminar Equipo";
            this.btnDeleteEquipment.Name = "btnDeleteEquipment";
            this.btnDeleteEquipment.ShowImage = true;
            this.btnDeleteEquipment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteEquipment_Click);
            // 
            // menuListEquipment
            // 
            this.menuListEquipment.Items.Add(this.btnReviewListEquips);
            this.menuListEquipment.Items.Add(this.btnReviewFromEquipmentList);
            this.menuListEquipment.Items.Add(this.btnAddEquipToList);
            this.menuListEquipment.Items.Add(this.btnDeleteEquipFromList);
            this.menuListEquipment.Label = "&Listas de Equipos";
            this.menuListEquipment.Name = "menuListEquipment";
            this.menuListEquipment.ShowImage = true;
            // 
            // btnReviewListEquips
            // 
            this.btnReviewListEquips.Label = "&Consultar";
            this.btnReviewListEquips.Name = "btnReviewListEquips";
            this.btnReviewListEquips.ShowImage = true;
            this.btnReviewListEquips.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewListEquips_Click);
            // 
            // btnReviewFromEquipmentList
            // 
            this.btnReviewFromEquipmentList.Label = "Consultar de Hoja de &Equipos";
            this.btnReviewFromEquipmentList.Name = "btnReviewFromEquipmentList";
            this.btnReviewFromEquipmentList.ScreenTip = "Consulta los equipos de la hoja de equipos para saber en qué listado existen";
            this.btnReviewFromEquipmentList.ShowImage = true;
            this.btnReviewFromEquipmentList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewFromEquipmentList_Click);
            // 
            // btnAddEquipToList
            // 
            this.btnAddEquipToList.Label = "&Agregar a Lista";
            this.btnAddEquipToList.Name = "btnAddEquipToList";
            this.btnAddEquipToList.ShowImage = true;
            this.btnAddEquipToList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddEquipToList_Click);
            // 
            // btnDeleteEquipFromList
            // 
            this.btnDeleteEquipFromList.Label = "&Quitar de List";
            this.btnDeleteEquipFromList.Name = "btnDeleteEquipFromList";
            this.btnDeleteEquipFromList.ShowImage = true;
            this.btnDeleteEquipFromList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteEquipFromList_Click);
            // 
            // menuCompMovement
            // 
            this.menuCompMovement.Items.Add(this.btnTraceAction);
            this.menuCompMovement.Items.Add(this.btnReviewCurrentFitment);
            this.menuCompMovement.Items.Add(this.btnDeleteAction);
            this.menuCompMovement.Label = "&Mov. Componentes";
            this.menuCompMovement.Name = "menuCompMovement";
            this.menuCompMovement.ShowImage = true;
            // 
            // btnTraceAction
            // 
            this.btnTraceAction.Label = "&Realizar Acción";
            this.btnTraceAction.Name = "btnTraceAction";
            this.btnTraceAction.ShowImage = true;
            this.btnTraceAction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTraceAction_Click);
            // 
            // btnReviewCurrentFitment
            // 
            this.btnReviewCurrentFitment.Label = "Consultar &Ultimo Movimiento";
            this.btnReviewCurrentFitment.Name = "btnReviewCurrentFitment";
            this.btnReviewCurrentFitment.ShowImage = true;
            this.btnReviewCurrentFitment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewFitments_Click);
            // 
            // btnDeleteAction
            // 
            this.btnDeleteAction.Label = "&Eliminar Acción";
            this.btnDeleteAction.Name = "btnDeleteAction";
            this.btnDeleteAction.ShowImage = true;
            this.btnDeleteAction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteAction_Click);
            // 
            // cbIgnoreRefCodes
            // 
            this.cbIgnoreRefCodes.Label = "Ignorar Reference Codes";
            this.cbIgnoreRefCodes.Name = "cbIgnoreRefCodes";
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
            this.grpEllipse.ResumeLayout(false);
            this.grpEllipse.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEllipse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatFull;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewCurrentFitment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuListEquipment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewListEquips;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReviewFromEquipmentList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddEquipToList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteEquipFromList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisposal;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbIgnoreRefCodes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteAction;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

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
            this.menuFormatSheet = this.Factory.CreateRibbonMenu();
            this.btnFormatFull = this.Factory.CreateRibbonButton();
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.btnCreateEquipment = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentData = this.Factory.CreateRibbonButton();
            this.btnUpdateEquipmentStatus = this.Factory.CreateRibbonButton();
            this.btnDeleteEquipment = this.Factory.CreateRibbonButton();
            this.menuCompMovement = this.Factory.CreateRibbonMenu();
            this.btnTraceAction = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpEllipse.SuspendLayout();
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
            this.grpEllipse.Items.Add(this.menuFormatSheet);
            this.grpEllipse.Items.Add(this.drpEnviroment);
            this.grpEllipse.Items.Add(this.menuActions);
            this.grpEllipse.Label = "Equipments v0.1.0";
            this.grpEllipse.Name = "grpEllipse";
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
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.btnCreateEquipment);
            this.menuActions.Items.Add(this.btnUpdateEquipmentData);
            this.menuActions.Items.Add(this.btnUpdateEquipmentStatus);
            this.menuActions.Items.Add(this.btnDeleteEquipment);
            this.menuActions.Items.Add(this.menuCompMovement);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "&Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // btnCreateEquipment
            // 
            this.btnCreateEquipment.Label = "&Crear Equipo";
            this.btnCreateEquipment.Name = "btnCreateEquipment";
            this.btnCreateEquipment.ShowImage = true;
            this.btnCreateEquipment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateEquipment_Click);
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
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace EllipseWorkOrderExcelAddIn
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
            this.drpEnviroment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuGeneral = this.Factory.CreateRibbonMenu();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnReReview = this.Factory.CreateRibbonButton();
            this.btnCreate = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnCleanWorkOrderSheet = this.Factory.CreateRibbonButton();
            this.menuComplete = this.Factory.CreateRibbonMenu();
            this.btnClose = this.Factory.CreateRibbonButton();
            this.btnReOpen = this.Factory.CreateRibbonButton();
            this.btnReviewCloseText = this.Factory.CreateRibbonButton();
            this.btnUpdateCloseText = this.Factory.CreateRibbonButton();
            this.cbIgnoreClosedStatus = this.Factory.CreateRibbonCheckBox();
            this.btnCleanCloseSheets = this.Factory.CreateRibbonButton();
            this.menuDurations = this.Factory.CreateRibbonMenu();
            this.btnDurationsReview = this.Factory.CreateRibbonButton();
            this.btnDurationsAction = this.Factory.CreateRibbonButton();
            this.btnCleanDuration = this.Factory.CreateRibbonButton();
            this.menuReferenceCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewReferenceCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateReferenceCodes = this.Factory.CreateRibbonButton();
            this.menuQuality = this.Factory.CreateRibbonMenu();
            this.btnReviewQuality = this.Factory.CreateRibbonButton();
            this.btnReReviewQuality = this.Factory.CreateRibbonButton();
            this.btnCleanQualitySheet = this.Factory.CreateRibbonButton();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnFormatDetail = this.Factory.CreateRibbonButton();
            this.btnFormatQuality = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkOrder.SuspendLayout();
            this.box2.SuspendLayout();
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
            this.grpWorkOrder.Items.Add(this.box2);
            this.grpWorkOrder.Items.Add(this.drpEnviroment);
            this.grpWorkOrder.Items.Add(this.menuActions);
            this.grpWorkOrder.Label = "WorkOrders";
            this.grpWorkOrder.Name = "grpWorkOrder";
            // 
            // drpEnviroment
            // 
            this.drpEnviroment.Label = "Env.";
            this.drpEnviroment.Name = "drpEnviroment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.menuGeneral);
            this.menuActions.Items.Add(this.menuComplete);
            this.menuActions.Items.Add(this.menuDurations);
            this.menuActions.Items.Add(this.menuReferenceCodes);
            this.menuActions.Items.Add(this.menuQuality);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuGeneral
            // 
            this.menuGeneral.Items.Add(this.btnReview);
            this.menuGeneral.Items.Add(this.btnReReview);
            this.menuGeneral.Items.Add(this.btnCreate);
            this.menuGeneral.Items.Add(this.btnUpdate);
            this.menuGeneral.Items.Add(this.btnCleanWorkOrderSheet);
            this.menuGeneral.Label = "&WorkOrders";
            this.menuGeneral.Name = "menuGeneral";
            this.menuGeneral.ShowImage = true;
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
            // btnCreate
            // 
            this.btnCreate.Label = "Crear OTs";
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.ShowImage = true;
            this.btnCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreate_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Label = "Actualizar OTs";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // btnCleanWorkOrderSheet
            // 
            this.btnCleanWorkOrderSheet.Label = "&Limpiar Hoja";
            this.btnCleanWorkOrderSheet.Name = "btnCleanWorkOrderSheet";
            this.btnCleanWorkOrderSheet.ShowImage = true;
            this.btnCleanWorkOrderSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanWorkOrderSheet_Click);
            // 
            // menuComplete
            // 
            this.menuComplete.Items.Add(this.btnClose);
            this.menuComplete.Items.Add(this.btnReOpen);
            this.menuComplete.Items.Add(this.btnReviewCloseText);
            this.menuComplete.Items.Add(this.btnUpdateCloseText);
            this.menuComplete.Items.Add(this.cbIgnoreClosedStatus);
            this.menuComplete.Items.Add(this.btnCleanCloseSheets);
            this.menuComplete.Label = "&Cierre de OTs";
            this.menuComplete.Name = "menuComplete";
            this.menuComplete.ShowImage = true;
            // 
            // btnClose
            // 
            this.btnClose.Label = "Cerrar OTs";
            this.btnClose.Name = "btnClose";
            this.btnClose.ShowImage = true;
            this.btnClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClose_Click);
            // 
            // btnReOpen
            // 
            this.btnReOpen.Label = "ReAbrir OT";
            this.btnReOpen.Name = "btnReOpen";
            this.btnReOpen.ShowImage = true;
            this.btnReOpen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReOpen_Click);
            // 
            // btnReviewCloseText
            // 
            this.btnReviewCloseText.Label = "Consultar Comentarios";
            this.btnReviewCloseText.Name = "btnReviewCloseText";
            this.btnReviewCloseText.ShowImage = true;
            this.btnReviewCloseText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewCloseText_Click);
            // 
            // btnUpdateCloseText
            // 
            this.btnUpdateCloseText.Label = "Actualizar Comentarios";
            this.btnUpdateCloseText.Name = "btnUpdateCloseText";
            this.btnUpdateCloseText.ShowImage = true;
            this.btnUpdateCloseText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateCloseText_Click);
            // 
            // cbIgnoreClosedStatus
            // 
            this.cbIgnoreClosedStatus.Label = "&Ignorar Estado de Cierre";
            this.cbIgnoreClosedStatus.Name = "cbIgnoreClosedStatus";
            // 
            // btnCleanCloseSheets
            // 
            this.btnCleanCloseSheets.Label = "Limpiar Hoja";
            this.btnCleanCloseSheets.Name = "btnCleanCloseSheets";
            this.btnCleanCloseSheets.ShowImage = true;
            this.btnCleanCloseSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanCloseSheets_Click);
            // 
            // menuDurations
            // 
            this.menuDurations.Items.Add(this.btnDurationsReview);
            this.menuDurations.Items.Add(this.btnDurationsAction);
            this.menuDurations.Items.Add(this.btnCleanDuration);
            this.menuDurations.Label = "&Duration de OTs";
            this.menuDurations.Name = "menuDurations";
            this.menuDurations.ShowImage = true;
            // 
            // btnDurationsReview
            // 
            this.btnDurationsReview.Label = "Consultar Duraciones";
            this.btnDurationsReview.Name = "btnDurationsReview";
            this.btnDurationsReview.ShowImage = true;
            this.btnDurationsReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDurationsReview_Click);
            // 
            // btnDurationsAction
            // 
            this.btnDurationsAction.Label = "Ejecutar Acciones";
            this.btnDurationsAction.Name = "btnDurationsAction";
            this.btnDurationsAction.ShowImage = true;
            this.btnDurationsAction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDurationsAction_Click);
            // 
            // btnCleanDuration
            // 
            this.btnCleanDuration.Label = "Limpiar Hoja";
            this.btnCleanDuration.Name = "btnCleanDuration";
            this.btnCleanDuration.ShowImage = true;
            this.btnCleanDuration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanDuration_Click);
            // 
            // menuReferenceCodes
            // 
            this.menuReferenceCodes.Items.Add(this.btnReviewReferenceCodes);
            this.menuReferenceCodes.Items.Add(this.btnUpdateReferenceCodes);
            this.menuReferenceCodes.Label = "&Reference Codes";
            this.menuReferenceCodes.Name = "menuReferenceCodes";
            this.menuReferenceCodes.ShowImage = true;
            // 
            // btnReviewReferenceCodes
            // 
            this.btnReviewReferenceCodes.Label = "&Consultar Reference Codes";
            this.btnReviewReferenceCodes.Name = "btnReviewReferenceCodes";
            this.btnReviewReferenceCodes.ShowImage = true;
            this.btnReviewReferenceCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewReferenceCodes_Click);
            // 
            // btnUpdateReferenceCodes
            // 
            this.btnUpdateReferenceCodes.Label = "&Actualizar Reference Codes";
            this.btnUpdateReferenceCodes.Name = "btnUpdateReferenceCodes";
            this.btnUpdateReferenceCodes.ShowImage = true;
            this.btnUpdateReferenceCodes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateReferenceCodes_Click);
            // 
            // menuQuality
            // 
            this.menuQuality.Items.Add(this.btnReviewQuality);
            this.menuQuality.Items.Add(this.btnReReviewQuality);
            this.menuQuality.Items.Add(this.btnCleanQualitySheet);
            this.menuQuality.Label = "C&alidad de OTs";
            this.menuQuality.Name = "menuQuality";
            this.menuQuality.ShowImage = true;
            // 
            // btnReviewQuality
            // 
            this.btnReviewQuality.Label = "&Consultar OTs";
            this.btnReviewQuality.Name = "btnReviewQuality";
            this.btnReviewQuality.ShowImage = true;
            this.btnReviewQuality.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewQuality_Click);
            // 
            // btnReReviewQuality
            // 
            this.btnReReviewQuality.Label = "&ReConsultar OTs";
            this.btnReReviewQuality.Name = "btnReReviewQuality";
            this.btnReReviewQuality.ShowImage = true;
            this.btnReReviewQuality.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewQuality_Click);
            // 
            // btnCleanQualitySheet
            // 
            this.btnCleanQualitySheet.Label = "&Limpiar Hoja";
            this.btnCleanQualitySheet.Name = "btnCleanQualitySheet";
            this.btnCleanQualitySheet.ShowImage = true;
            this.btnCleanQualitySheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanQualitySheet_Click);
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener &Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // box2
            // 
            this.box2.Items.Add(this.menuFormat);
            this.box2.Items.Add(this.btnAbout);
            this.box2.Name = "box2";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatSheet);
            this.menuFormat.Items.Add(this.btnFormatDetail);
            this.menuFormat.Items.Add(this.btnFormatQuality);
            this.menuFormat.Label = "&Formatos";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatSheet
            // 
            this.btnFormatSheet.Label = "&Formato Base";
            this.btnFormatSheet.Name = "btnFormatSheet";
            this.btnFormatSheet.ShowImage = true;
            this.btnFormatSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheet_Click);
            // 
            // btnFormatDetail
            // 
            this.btnFormatDetail.Label = "Formato &Detallado";
            this.btnFormatDetail.Name = "btnFormatDetail";
            this.btnFormatDetail.ShowImage = true;
            this.btnFormatDetail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatDetail_Click);
            // 
            // btnFormatQuality
            // 
            this.btnFormatQuality.Label = "&Calidad de OTs";
            this.btnFormatQuality.Name = "btnFormatQuality";
            this.btnFormatQuality.ShowImage = true;
            this.btnFormatQuality.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatQuality_Click);
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
            this.grpWorkOrder.ResumeLayout(false);
            this.grpWorkOrder.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal RibbonTab tabEllipse;
        internal RibbonGroup grpWorkOrder;
        internal RibbonDropDown drpEnviroment;
        internal RibbonButton btnReview;
        internal RibbonMenu menuActions;
        internal RibbonButton btnCreate;
        internal RibbonButton btnUpdate;
        internal RibbonButton btnClose;
        internal RibbonButton btnReOpen;
        internal RibbonMenu menuGeneral;
        internal RibbonMenu menuComplete;
        internal RibbonMenu menuDurations;
        internal RibbonButton btnDurationsReview;
        internal RibbonButton btnDurationsAction;
        internal RibbonButton btnReviewCloseText;
        internal RibbonButton btnUpdateCloseText;
        internal RibbonButton btnReReview;
        internal RibbonButton btnCleanWorkOrderSheet;
        internal RibbonButton btnCleanCloseSheets;
        internal RibbonButton btnCleanDuration;
        internal RibbonButton btnStopThread;
        internal RibbonMenu menuQuality;
        internal RibbonButton btnReviewQuality;
        internal RibbonButton btnReReviewQuality;
        internal RibbonButton btnCleanQualitySheet;
        internal RibbonCheckBox cbIgnoreClosedStatus;
        internal RibbonMenu menuReferenceCodes;
        internal RibbonButton btnReviewReferenceCodes;
        internal RibbonButton btnUpdateReferenceCodes;
        internal RibbonBox box2;
        internal RibbonMenu menuFormat;
        internal RibbonButton btnFormatSheet;
        internal RibbonButton btnFormatDetail;
        internal RibbonButton btnFormatQuality;
        internal RibbonButton btnAbout;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

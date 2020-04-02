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
            this.box2 = this.Factory.CreateRibbonBox();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheet = this.Factory.CreateRibbonButton();
            this.btnFormatDetail = this.Factory.CreateRibbonButton();
            this.btnFormatQuality = this.Factory.CreateRibbonButton();
            this.btnFormatCriticalControls = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuGeneral = this.Factory.CreateRibbonMenu();
            this.btnReview = this.Factory.CreateRibbonButton();
            this.btnReReview = this.Factory.CreateRibbonButton();
            this.btnCreate = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.btnFlagEstDuration = this.Factory.CreateRibbonCheckBox();
            this.btnCleanWorkOrderSheet = this.Factory.CreateRibbonButton();
            this.menuTasks = this.Factory.CreateRibbonMenu();
            this.btnReviewTasks = this.Factory.CreateRibbonButton();
            this.btnExecuteTaskActions = this.Factory.CreateRibbonButton();
            this.btnValidateTaskPlanDates = this.Factory.CreateRibbonCheckBox();
            this.btnCleanTasksTable = this.Factory.CreateRibbonButton();
            this.menuRequirements = this.Factory.CreateRibbonMenu();
            this.btnReviewRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewLabRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewMatRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewEqpRequirements = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnReviewTaskRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewTaskLabRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewTaskMatRequirements = this.Factory.CreateRibbonButton();
            this.btnReviewTaskEqpRequirements = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnExecuteRequirements = this.Factory.CreateRibbonButton();
            this.btnGetAplRequirements = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.btnCleanRequirementTable = this.Factory.CreateRibbonButton();
            this.menuComplete = this.Factory.CreateRibbonMenu();
            this.btnClose = this.Factory.CreateRibbonButton();
            this.btnReOpen = this.Factory.CreateRibbonButton();
            this.btnReviewCloseText = this.Factory.CreateRibbonButton();
            this.btnUpdateCloseText = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.cbIgnoreClosedStatus = this.Factory.CreateRibbonCheckBox();
            this.btnCleanCloseSheets = this.Factory.CreateRibbonButton();
            this.menuDurations = this.Factory.CreateRibbonMenu();
            this.btnDurationsReview = this.Factory.CreateRibbonButton();
            this.btnDurationsAction = this.Factory.CreateRibbonButton();
            this.btnCleanDuration = this.Factory.CreateRibbonButton();
            this.menuWorkProgress = this.Factory.CreateRibbonMenu();
            this.btnReviewWorkProgress = this.Factory.CreateRibbonButton();
            this.btnUpdatePercentProgress = this.Factory.CreateRibbonButton();
            this.btnUpdateUnitsProgress = this.Factory.CreateRibbonButton();
            this.btnUpdateUnitsRequired = this.Factory.CreateRibbonButton();
            this.menuToDo = this.Factory.CreateRibbonMenu();
            this.btnToDoReviewWorkOrders = this.Factory.CreateRibbonButton();
            this.btnToDoReviewTasks = this.Factory.CreateRibbonButton();
            this.btnCreateToDo = this.Factory.CreateRibbonButton();
            this.btnUpdateToDo = this.Factory.CreateRibbonButton();
            this.btnDeleteToDo = this.Factory.CreateRibbonButton();
            this.btnCleanToDo = this.Factory.CreateRibbonButton();
            this.menuReferenceCodes = this.Factory.CreateRibbonMenu();
            this.btnReviewReferenceCodes = this.Factory.CreateRibbonButton();
            this.btnUpdateReferenceCodes = this.Factory.CreateRibbonButton();
            this.menuQuality = this.Factory.CreateRibbonMenu();
            this.btnReviewQuality = this.Factory.CreateRibbonButton();
            this.btnReReviewQuality = this.Factory.CreateRibbonButton();
            this.btnCleanQualitySheet = this.Factory.CreateRibbonButton();
            this.menuCriticalControls = this.Factory.CreateRibbonMenu();
            this.btnReviewCriticalControls = this.Factory.CreateRibbonButton();
            this.btnReReviewCritialControls = this.Factory.CreateRibbonButton();
            this.btnExportCriticalControls = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpWorkOrder.SuspendLayout();
            this.box2.SuspendLayout();
            this.SuspendLayout();
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
            this.grpWorkOrder.Items.Add(this.drpEnvironment);
            this.grpWorkOrder.Items.Add(this.menuActions);
            this.grpWorkOrder.Label = "WorkOrders";
            this.grpWorkOrder.Name = "grpWorkOrder";
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
            this.menuFormat.Items.Add(this.btnFormatCriticalControls);
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
            // btnFormatCriticalControls
            // 
            this.btnFormatCriticalControls.Label = "Cont&roles Críticos";
            this.btnFormatCriticalControls.Name = "btnFormatCriticalControls";
            this.btnFormatCriticalControls.ShowImage = true;
            this.btnFormatCriticalControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatCriticalControls_Click);
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
            this.menuActions.Items.Add(this.menuGeneral);
            this.menuActions.Items.Add(this.menuTasks);
            this.menuActions.Items.Add(this.menuRequirements);
            this.menuActions.Items.Add(this.menuComplete);
            this.menuActions.Items.Add(this.menuDurations);
            this.menuActions.Items.Add(this.menuWorkProgress);
            this.menuActions.Items.Add(this.menuToDo);
            this.menuActions.Items.Add(this.menuReferenceCodes);
            this.menuActions.Items.Add(this.menuQuality);
            this.menuActions.Items.Add(this.menuCriticalControls);
            this.menuActions.Items.Add(this.separator1);
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
            this.menuGeneral.Items.Add(this.separator6);
            this.menuGeneral.Items.Add(this.btnFlagEstDuration);
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
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // btnFlagEstDuration
            // 
            this.btnFlagEstDuration.Checked = true;
            this.btnFlagEstDuration.Label = "Estimados de Horas Calculados";
            this.btnFlagEstDuration.Name = "btnFlagEstDuration";
            // 
            // btnCleanWorkOrderSheet
            // 
            this.btnCleanWorkOrderSheet.Label = "&Limpiar Hoja";
            this.btnCleanWorkOrderSheet.Name = "btnCleanWorkOrderSheet";
            this.btnCleanWorkOrderSheet.ShowImage = true;
            this.btnCleanWorkOrderSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanWorkOrderSheet_Click);
            // 
            // menuTasks
            // 
            this.menuTasks.Items.Add(this.btnReviewTasks);
            this.menuTasks.Items.Add(this.btnExecuteTaskActions);
            this.menuTasks.Items.Add(this.btnValidateTaskPlanDates);
            this.menuTasks.Items.Add(this.btnCleanTasksTable);
            this.menuTasks.Label = "&Tareas";
            this.menuTasks.Name = "menuTasks";
            this.menuTasks.ShowImage = true;
            // 
            // btnReviewTasks
            // 
            this.btnReviewTasks.Label = "&Consultar Tareas";
            this.btnReviewTasks.Name = "btnReviewTasks";
            this.btnReviewTasks.ShowImage = true;
            this.btnReviewTasks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTasks_Click);
            // 
            // btnExecuteTaskActions
            // 
            this.btnExecuteTaskActions.Label = "&Ejecutar Acciones de Tareas";
            this.btnExecuteTaskActions.Name = "btnExecuteTaskActions";
            this.btnExecuteTaskActions.ShowImage = true;
            this.btnExecuteTaskActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteTaskActions_Click);
            // 
            // btnValidateTaskPlanDates
            // 
            this.btnValidateTaskPlanDates.Checked = true;
            this.btnValidateTaskPlanDates.Label = "Validar Fechas Plan en OT";
            this.btnValidateTaskPlanDates.Name = "btnValidateTaskPlanDates";
            // 
            // btnCleanTasksTable
            // 
            this.btnCleanTasksTable.Label = "&Limpiar Tabla de Tareas";
            this.btnCleanTasksTable.Name = "btnCleanTasksTable";
            this.btnCleanTasksTable.ShowImage = true;
            this.btnCleanTasksTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanTasksTable_Click);
            // 
            // menuRequirements
            // 
            this.menuRequirements.Items.Add(this.btnReviewRequirements);
            this.menuRequirements.Items.Add(this.btnReviewLabRequirements);
            this.menuRequirements.Items.Add(this.btnReviewMatRequirements);
            this.menuRequirements.Items.Add(this.btnReviewEqpRequirements);
            this.menuRequirements.Items.Add(this.separator3);
            this.menuRequirements.Items.Add(this.btnReviewTaskRequirements);
            this.menuRequirements.Items.Add(this.btnReviewTaskLabRequirements);
            this.menuRequirements.Items.Add(this.btnReviewTaskMatRequirements);
            this.menuRequirements.Items.Add(this.btnReviewTaskEqpRequirements);
            this.menuRequirements.Items.Add(this.separator2);
            this.menuRequirements.Items.Add(this.btnExecuteRequirements);
            this.menuRequirements.Items.Add(this.btnGetAplRequirements);
            this.menuRequirements.Items.Add(this.separator5);
            this.menuRequirements.Items.Add(this.btnCleanRequirementTable);
            this.menuRequirements.Label = "&Requerimientos";
            this.menuRequirements.Name = "menuRequirements";
            this.menuRequirements.ShowImage = true;
            // 
            // btnReviewRequirements
            // 
            this.btnReviewRequirements.Label = "Recursos por Orden";
            this.btnReviewRequirements.Name = "btnReviewRequirements";
            this.btnReviewRequirements.ShowImage = true;
            this.btnReviewRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewRequirements_Click);
            // 
            // btnReviewLabRequirements
            // 
            this.btnReviewLabRequirements.Label = "Labor por Orden";
            this.btnReviewLabRequirements.Name = "btnReviewLabRequirements";
            this.btnReviewLabRequirements.ShowImage = true;
            this.btnReviewLabRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewLabRequirements_Click);
            // 
            // btnReviewMatRequirements
            // 
            this.btnReviewMatRequirements.Label = "Materiales por Orden";
            this.btnReviewMatRequirements.Name = "btnReviewMatRequirements";
            this.btnReviewMatRequirements.ShowImage = true;
            this.btnReviewMatRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewMatRequirements_Click);
            // 
            // btnReviewEqpRequirements
            // 
            this.btnReviewEqpRequirements.Label = "Equipos por Orden";
            this.btnReviewEqpRequirements.Name = "btnReviewEqpRequirements";
            this.btnReviewEqpRequirements.ShowImage = true;
            this.btnReviewEqpRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewEqpRequirements_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btnReviewTaskRequirements
            // 
            this.btnReviewTaskRequirements.Label = "Recursos por Tarea";
            this.btnReviewTaskRequirements.Name = "btnReviewTaskRequirements";
            this.btnReviewTaskRequirements.ShowImage = true;
            this.btnReviewTaskRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTaskRequirements_Click);
            // 
            // btnReviewTaskLabRequirements
            // 
            this.btnReviewTaskLabRequirements.Label = "Labor por Tarea";
            this.btnReviewTaskLabRequirements.Name = "btnReviewTaskLabRequirements";
            this.btnReviewTaskLabRequirements.ShowImage = true;
            this.btnReviewTaskLabRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTaskLabRequirements_Click);
            // 
            // btnReviewTaskMatRequirements
            // 
            this.btnReviewTaskMatRequirements.Label = "Materiales por Tarea";
            this.btnReviewTaskMatRequirements.Name = "btnReviewTaskMatRequirements";
            this.btnReviewTaskMatRequirements.ShowImage = true;
            this.btnReviewTaskMatRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTaskMatRequirements_Click);
            // 
            // btnReviewTaskEqpRequirements
            // 
            this.btnReviewTaskEqpRequirements.Label = "Equipos por Tarea";
            this.btnReviewTaskEqpRequirements.Name = "btnReviewTaskEqpRequirements";
            this.btnReviewTaskEqpRequirements.ShowImage = true;
            this.btnReviewTaskEqpRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewTaskEqpRequirements_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnExecuteRequirements
            // 
            this.btnExecuteRequirements.Label = "Ejecutar Acciones";
            this.btnExecuteRequirements.Name = "btnExecuteRequirements";
            this.btnExecuteRequirements.ShowImage = true;
            this.btnExecuteRequirements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteRequirements_Click);
            // 
            // btnGetAplRequirements
            // 
            this.btnGetAplRequirements.Label = "Traer Recursos de APLs";
            this.btnGetAplRequirements.Name = "btnGetAplRequirements";
            this.btnGetAplRequirements.ShowImage = true;
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // btnCleanRequirementTable
            // 
            this.btnCleanRequirementTable.Label = "Limpiar Tabla Requerimientos";
            this.btnCleanRequirementTable.Name = "btnCleanRequirementTable";
            this.btnCleanRequirementTable.ShowImage = true;
            this.btnCleanRequirementTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanRequirementTable_Click);
            // 
            // menuComplete
            // 
            this.menuComplete.Items.Add(this.btnClose);
            this.menuComplete.Items.Add(this.btnReOpen);
            this.menuComplete.Items.Add(this.btnReviewCloseText);
            this.menuComplete.Items.Add(this.btnUpdateCloseText);
            this.menuComplete.Items.Add(this.separator4);
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
            // separator4
            // 
            this.separator4.Name = "separator4";
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
            // menuWorkProgress
            // 
            this.menuWorkProgress.Items.Add(this.btnReviewWorkProgress);
            this.menuWorkProgress.Items.Add(this.btnUpdatePercentProgress);
            this.menuWorkProgress.Items.Add(this.btnUpdateUnitsProgress);
            this.menuWorkProgress.Items.Add(this.btnUpdateUnitsRequired);
            this.menuWorkProgress.Label = "&Progreso de OTs";
            this.menuWorkProgress.Name = "menuWorkProgress";
            this.menuWorkProgress.ShowImage = true;
            // 
            // btnReviewWorkProgress
            // 
            this.btnReviewWorkProgress.Label = "&Consultar Progreso";
            this.btnReviewWorkProgress.Name = "btnReviewWorkProgress";
            this.btnReviewWorkProgress.ShowImage = true;
            this.btnReviewWorkProgress.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewWorkProgress_Click);
            // 
            // btnUpdatePercentProgress
            // 
            this.btnUpdatePercentProgress.Label = "Actualizar &Porcentaje";
            this.btnUpdatePercentProgress.Name = "btnUpdatePercentProgress";
            this.btnUpdatePercentProgress.ShowImage = true;
            this.btnUpdatePercentProgress.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdatePercentProgress_Click);
            // 
            // btnUpdateUnitsProgress
            // 
            this.btnUpdateUnitsProgress.Label = "&Actualizar Completadas";
            this.btnUpdateUnitsProgress.Name = "btnUpdateUnitsProgress";
            this.btnUpdateUnitsProgress.ShowImage = true;
            this.btnUpdateUnitsProgress.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateUnitsProgress_Click);
            // 
            // btnUpdateUnitsRequired
            // 
            this.btnUpdateUnitsRequired.Label = "Actualizar &Requeridas";
            this.btnUpdateUnitsRequired.Name = "btnUpdateUnitsRequired";
            this.btnUpdateUnitsRequired.ShowImage = true;
            this.btnUpdateUnitsRequired.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateUnitsRequired_Click);
            // 
            // menuToDo
            // 
            this.menuToDo.Items.Add(this.btnToDoReviewWorkOrders);
            this.menuToDo.Items.Add(this.btnToDoReviewTasks);
            this.menuToDo.Items.Add(this.btnCreateToDo);
            this.menuToDo.Items.Add(this.btnUpdateToDo);
            this.menuToDo.Items.Add(this.btnDeleteToDo);
            this.menuToDo.Items.Add(this.btnCleanToDo);
            this.menuToDo.Label = "To Do de OTs";
            this.menuToDo.Name = "menuToDo";
            this.menuToDo.ShowImage = true;
            // 
            // btnToDoReviewWorkOrders
            // 
            this.btnToDoReviewWorkOrders.Label = "Consultar de OTs";
            this.btnToDoReviewWorkOrders.Name = "btnToDoReviewWorkOrders";
            this.btnToDoReviewWorkOrders.ShowImage = true;
            this.btnToDoReviewWorkOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToDoReviewWorkOrders_Click);
            // 
            // btnToDoReviewTasks
            // 
            this.btnToDoReviewTasks.Label = "Consultar de Tareas";
            this.btnToDoReviewTasks.Name = "btnToDoReviewTasks";
            this.btnToDoReviewTasks.ShowImage = true;
            this.btnToDoReviewTasks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToDoReviewTasks_Click);
            // 
            // btnCreateToDo
            // 
            this.btnCreateToDo.Label = "Crear To Do";
            this.btnCreateToDo.Name = "btnCreateToDo";
            this.btnCreateToDo.ShowImage = true;
            this.btnCreateToDo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateToDo_Click);
            // 
            // btnUpdateToDo
            // 
            this.btnUpdateToDo.Label = "Actualizar To Do";
            this.btnUpdateToDo.Name = "btnUpdateToDo";
            this.btnUpdateToDo.ShowImage = true;
            this.btnUpdateToDo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateToDo_Click);
            // 
            // btnDeleteToDo
            // 
            this.btnDeleteToDo.Label = "Eliminar To Do";
            this.btnDeleteToDo.Name = "btnDeleteToDo";
            this.btnDeleteToDo.ShowImage = true;
            this.btnDeleteToDo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteToDo_Click);
            // 
            // btnCleanToDo
            // 
            this.btnCleanToDo.Label = "&Limpiar Hoja";
            this.btnCleanToDo.Name = "btnCleanToDo";
            this.btnCleanToDo.ShowImage = true;
            this.btnCleanToDo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanToDo_Click);
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
            // menuCriticalControls
            // 
            this.menuCriticalControls.Items.Add(this.btnReviewCriticalControls);
            this.menuCriticalControls.Items.Add(this.btnReReviewCritialControls);
            this.menuCriticalControls.Items.Add(this.btnExportCriticalControls);
            this.menuCriticalControls.Label = "Controles Críticos";
            this.menuCriticalControls.Name = "menuCriticalControls";
            this.menuCriticalControls.ShowImage = true;
            // 
            // btnReviewCriticalControls
            // 
            this.btnReviewCriticalControls.Label = "&Consultar OTs";
            this.btnReviewCriticalControls.Name = "btnReviewCriticalControls";
            this.btnReviewCriticalControls.ShowImage = true;
            this.btnReviewCriticalControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReviewCriticalControls_Click);
            // 
            // btnReReviewCritialControls
            // 
            this.btnReReviewCritialControls.Label = "&ReConsultar OTs";
            this.btnReReviewCritialControls.Name = "btnReReviewCritialControls";
            this.btnReReviewCritialControls.ShowImage = true;
            this.btnReReviewCritialControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReReviewCritialControls_Click);
            // 
            // btnExportCriticalControls
            // 
            this.btnExportCriticalControls.Label = "&Exportar a RTF";
            this.btnExportCriticalControls.Name = "btnExportCriticalControls";
            this.btnExportCriticalControls.ShowImage = true;
            this.btnExportCriticalControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportCriticalControls_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener &Proceso";
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
            this.grpWorkOrder.ResumeLayout(false);
            this.grpWorkOrder.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal RibbonTab tabEllipse;
        internal RibbonGroup grpWorkOrder;
        internal RibbonDropDown drpEnvironment;
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
        internal RibbonMenu menuWorkProgress;
        internal RibbonButton btnReviewWorkProgress;
        internal RibbonButton btnUpdatePercentProgress;
        internal RibbonButton btnUpdateUnitsProgress;
        internal RibbonButton btnUpdateUnitsRequired;
        internal RibbonButton btnFormatCriticalControls;
        internal RibbonMenu menuCriticalControls;
        internal RibbonButton btnReviewCriticalControls;
        internal RibbonButton btnReReviewCritialControls;
        internal RibbonButton btnExportCriticalControls;
        internal RibbonMenu menuTasks;
        internal RibbonButton btnReviewTasks;
        internal RibbonButton btnExecuteTaskActions;
        internal RibbonButton btnCleanTasksTable;
        internal RibbonMenu menuRequirements;
        internal RibbonButton btnReviewRequirements;
        internal RibbonButton btnExecuteRequirements;
        internal RibbonButton btnGetAplRequirements;
        internal RibbonButton btnCleanRequirementTable;
        internal RibbonButton btnReviewMatRequirements;
        internal RibbonMenu menuToDo;
        internal RibbonButton btnToDoReviewWorkOrders;
        internal RibbonButton btnToDoReviewTasks;
        internal RibbonButton btnCreateToDo;
        internal RibbonButton btnDeleteToDo;
        internal RibbonButton btnCleanToDo;
        internal RibbonButton btnUpdateToDo;
        internal RibbonCheckBox btnValidateTaskPlanDates;
        internal RibbonCheckBox btnFlagEstDuration;
        internal RibbonButton btnReviewLabRequirements;
        internal RibbonButton btnReviewEqpRequirements;
        internal RibbonSeparator separator1;
        internal RibbonSeparator separator2;
        internal RibbonSeparator separator3;
        internal RibbonButton btnReviewTaskRequirements;
        internal RibbonButton btnReviewTaskLabRequirements;
        internal RibbonButton btnReviewTaskMatRequirements;
        internal RibbonButton btnReviewTaskEqpRequirements;
        internal RibbonSeparator separator6;
        internal RibbonSeparator separator5;
        internal RibbonSeparator separator4;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEllipse RibbonEllipse
        {
            get { return this.GetRibbon<RibbonEllipse>(); }
        }
    }
}

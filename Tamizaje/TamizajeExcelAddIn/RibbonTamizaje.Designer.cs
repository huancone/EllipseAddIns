
namespace TamizajeExcelAddIn
{
    partial class RibbonTamizaje : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTamizaje()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpTamizaje = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnFormat = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuActions = this.Factory.CreateRibbonMenu();
            this.menuQuestionaries = this.Factory.CreateRibbonMenu();
            this.btnLoadQuestionary = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.cbUpdateExistingRecords = this.Factory.CreateRibbonCheckBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.cbAllowBackgroundWork = this.Factory.CreateRibbonCheckBox();
            this.btnStopThread = this.Factory.CreateRibbonButton();
            this.btnValidateUser = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpTamizaje.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpTamizaje);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpTamizaje
            // 
            this.grpTamizaje.Items.Add(this.box1);
            this.grpTamizaje.Items.Add(this.drpEnvironment);
            this.grpTamizaje.Items.Add(this.menuActions);
            this.grpTamizaje.Label = "Tamizaje";
            this.grpTamizaje.Name = "grpTamizaje";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btnFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
            // 
            // btnFormat
            // 
            this.btnFormat.Label = "&Format";
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "?";
            this.btnAbout.Name = "btnAbout";
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Serv.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuActions
            // 
            this.menuActions.Items.Add(this.menuQuestionaries);
            this.menuActions.Items.Add(this.btnValidateUser);
            this.menuActions.Items.Add(this.cbAllowBackgroundWork);
            this.menuActions.Items.Add(this.btnStopThread);
            this.menuActions.Label = "Acciones";
            this.menuActions.Name = "menuActions";
            // 
            // menuQuestionaries
            // 
            this.menuQuestionaries.Items.Add(this.btnLoadQuestionary);
            this.menuQuestionaries.Items.Add(this.separator1);
            this.menuQuestionaries.Items.Add(this.cbUpdateExistingRecords);
            this.menuQuestionaries.Items.Add(this.checkBox1);
            this.menuQuestionaries.Label = "Cuestionarios";
            this.menuQuestionaries.Name = "menuQuestionaries";
            this.menuQuestionaries.ShowImage = true;
            // 
            // btnLoadQuestionary
            // 
            this.btnLoadQuestionary.Label = "Cargar Cuestionario";
            this.btnLoadQuestionary.Name = "btnLoadQuestionary";
            this.btnLoadQuestionary.ShowImage = true;
            this.btnLoadQuestionary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadQuestionary_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // cbUpdateExistingRecords
            // 
            this.cbUpdateExistingRecords.Checked = true;
            this.cbUpdateExistingRecords.Label = "Actualizar Existentes";
            this.cbUpdateExistingRecords.Name = "cbUpdateExistingRecords";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "AutoInsertar Valores No Existentes";
            this.checkBox1.Name = "checkBox1";
            // 
            // cbAllowBackgroundWork
            // 
            this.cbAllowBackgroundWork.Label = "Permitir Trabajo en Segundo Plano";
            this.cbAllowBackgroundWork.Name = "cbAllowBackgroundWork";
            // 
            // btnStopThread
            // 
            this.btnStopThread.Label = "Detener Proceso";
            this.btnStopThread.Name = "btnStopThread";
            this.btnStopThread.ShowImage = true;
            this.btnStopThread.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopThread_Click);
            // 
            // btnValidateUser
            // 
            this.btnValidateUser.Label = "Validar Usuario";
            this.btnValidateUser.Name = "btnValidateUser";
            this.btnValidateUser.ShowImage = true;
            this.btnValidateUser.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateUser_Click);
            // 
            // RibbonTamizaje
            // 
            this.Name = "RibbonTamizaje";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTamizaje_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpTamizaje.ResumeLayout(false);
            this.grpTamizaje.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTamizaje;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpEnvironment;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadQuestionary;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAllowBackgroundWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuQuestionaries;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUpdateExistingRecords;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopThread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateUser;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTamizaje RibbonTamizaje
        {
            get { return this.GetRibbon<RibbonTamizaje>(); }
        }
    }
}

using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace EllipseInstDemorasEnOTsExcelAddIn
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
            this.grpDemorasOTs = this.Factory.CreateRibbonGroup();
            this.menuFormat = this.Factory.CreateRibbonMenu();
            this.btnFormatSheetImis = this.Factory.CreateRibbonButton();
            this.btnFormatSheetAires = this.Factory.CreateRibbonButton();
            this.drpEnvironment = this.Factory.CreateRibbonDropDown();
            this.menuOpciones = this.Factory.CreateRibbonMenu();
            this.btnUpdateDemoras = this.Factory.CreateRibbonButton();
            this.btnConsultarDemoras = this.Factory.CreateRibbonButton();
            this.btnEliminar = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEllipse.SuspendLayout();
            this.grpDemorasOTs.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tabEllipse
            // 
            this.tabEllipse.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEllipse.Groups.Add(this.grpDemorasOTs);
            this.tabEllipse.Label = "ELLIPSE";
            this.tabEllipse.Name = "tabEllipse";
            // 
            // grpDemorasOTs
            // 
            this.grpDemorasOTs.Items.Add(this.box1);
            this.grpDemorasOTs.Items.Add(this.drpEnvironment);
            this.grpDemorasOTs.Items.Add(this.menuOpciones);
            this.grpDemorasOTs.Label = "Demoras OTs";
            this.grpDemorasOTs.Name = "grpDemorasOTs";
            // 
            // menuFormat
            // 
            this.menuFormat.Items.Add(this.btnFormatSheetImis);
            this.menuFormat.Items.Add(this.btnFormatSheetAires);
            this.menuFormat.Label = "&Formatear";
            this.menuFormat.Name = "menuFormat";
            // 
            // btnFormatSheetImis
            // 
            this.btnFormatSheetImis.Label = "Formato IMIS";
            this.btnFormatSheetImis.Name = "btnFormatSheetImis";
            this.btnFormatSheetImis.ShowImage = true;
            this.btnFormatSheetImis.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheetImis_Click);
            // 
            // btnFormatSheetAires
            // 
            this.btnFormatSheetAires.Label = "Formato AIRES";
            this.btnFormatSheetAires.Name = "btnFormatSheetAires";
            this.btnFormatSheetAires.ShowImage = true;
            this.btnFormatSheetAires.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatSheetAires_Click);
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.Label = "Env.";
            this.drpEnvironment.Name = "drpEnvironment";
            // 
            // menuOpciones
            // 
            this.menuOpciones.Items.Add(this.btnUpdateDemoras);
            this.menuOpciones.Items.Add(this.btnConsultarDemoras);
            this.menuOpciones.Items.Add(this.btnEliminar);
            this.menuOpciones.Label = "Opciones";
            this.menuOpciones.Name = "menuOpciones";
            // 
            // btnUpdateDemoras
            // 
            this.btnUpdateDemoras.Label = "Crear/Actualizar Demoras";
            this.btnUpdateDemoras.Name = "btnUpdateDemoras";
            this.btnUpdateDemoras.ShowImage = true;
            this.btnUpdateDemoras.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateDemoras_Click);
            // 
            // btnConsultarDemoras
            // 
            this.btnConsultarDemoras.Label = "Consultar Demoras";
            this.btnConsultarDemoras.Name = "btnConsultarDemoras";
            this.btnConsultarDemoras.ShowImage = true;
            this.btnConsultarDemoras.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultarDemoras_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Label = "Eliminar Demoras";
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.ShowImage = true;
            this.btnEliminar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEliminar_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.menuFormat);
            this.box1.Items.Add(this.btnAbout);
            this.box1.Name = "box1";
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
            this.grpDemorasOTs.ResumeLayout(false);
            this.grpDemorasOTs.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal RibbonTab tabEllipse;
        internal RibbonGroup grpDemorasOTs;
        internal RibbonDropDown drpEnvironment;
        internal RibbonButton btnFormatSheetImis;
        internal RibbonButton btnConsultarDemoras;
        internal RibbonMenu menuOpciones;
        internal RibbonButton btnUpdateDemoras;
        internal RibbonButton btnEliminar;
        internal RibbonMenu menuFormat;
        internal RibbonButton btnFormatSheetAires;
        internal RibbonBox box1;
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

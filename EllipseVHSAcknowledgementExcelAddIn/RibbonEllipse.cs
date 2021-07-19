using System;
using System.Diagnostics;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse;
using EllipseVHSAcknowledgementExcelAddIn.IssueRequisitionItemStocklessService;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace EllipseVHSAcknowledgementExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        Application _excelApp;
        private const string SheetName01 = "VHSAcknowledgement";
        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 12;
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }
        private void btnAction_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
            {
                ReviewData();
            }
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para realizar la acción");
            
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "VHS ACKNOWLEDGEMENT - ELLIPSE";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A4").Value = "SUPPLIER";
                _cells.GetCell("A5").Value = "REQUISITION";
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                //TITLE
                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, TitleRow01).Value = "REQUISICIÓN";
                _cells.GetCell(2, TitleRow01).Value = "REQ. ITEM";
                _cells.GetCell(3, TitleRow01).Value = "STOCK CODE";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPCIÓN";
                _cells.GetCell(5, TitleRow01).Value = "CANT. REQUERIDA";
                _cells.GetCell(6, TitleRow01).Value = "CANT. PENDIENTE";
                _cells.GetCell(7, TitleRow01).Value = "CANT. ADMITIDA";
                _cells.GetCell(8, TitleRow01).Value = "UNIDAD";
                _cells.GetCell(9, TitleRow01).Value = "ACCIÓN";
                _cells.GetCell(9, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell(10, TitleRow01).Value = "PROVEEDOR";
                _cells.GetCell(11, TitleRow01).Value = "NOMBRE PROVEEDOR";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        
        }

        public void ReviewData()
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() == DialogResult.OK)
            {
                //Variables de gestión
                string district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                string supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B4").Value);
                string requisition = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B5").Value);
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

                //Variables de operación del servicio
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var service = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();
                var requestSheet = new IssueRequisitionItemStocklessSearchParam();

                var itemDto = new IssueRequisitionItemStocklessDTO();
                service.Url = urlService;

                //gestionamos y enviamos la solicitud
                requestSheet.districtCode = district;
                requestSheet.supplierNumber = supplier;
                requestSheet.issueRequisitionNumber = requisition;
                requestSheet.defaultQuantityAcknowledged = false;
                requestSheet.defaultQuantityAcknowledgedSpecified = true;

                service.search(opSheet, requestSheet, itemDto);
                //IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessAcknowledgeDTO ackDTO = new IssueRequisitionItemStocklessAcknowledgeDTO();
                //ackDTO.activityCounter = "000";
                //ackDTO.authorisedStatus = "A";
                //ackDTO.custodianId = "6799DD7F11AA429C9CD44E5E846DB58F";
                //ackDTO.customsValue = 0;
                //ackDTO.customsValueSecondaryCurrency = 0;
                //ackDTO.description = "MANTENIMIENTO INSTALACIONES PTO BOLIVAR; 0002ESPECIALISTA EN BOMBAS";

                //proxySheet.acknowledge(opSheet, ackDTO);
            }


            
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}

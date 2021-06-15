using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Ellipse.Constants;
using System.Web.Services.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseRequisitionServiceExcelAddIn.IssueRequisitionItemStocklessService;
using EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary;
using EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.Assurance;
using EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.SearchConstants;
using SharedClassLibrary.Ellipse.Connections;
using Screen = SharedClassLibrary.Ellipse.ScreenService;

namespace EllipseRequisitionServiceExcelAddIn
{
    public partial class RibbonEllipse
    {
        Excel.Application _excelApp;
        ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;

        private const int TitleRow01 = 5;
        private const int ResultColumn01 = 19;

        private const int TitleRow01Ext = 5;
        private const int ResultColumn01Ext = 21;

        private const int TitleRow02 = 8;
        private const int ResultColumn02 = 27;
        private const int TitleRow03 = 5;

        private const string SheetName01 = "RequisitionService";
        private const string SheetName02 = "Consultas";
        private const string SheetName03 = "DetalleConsultas";
        private const string TableName01 = "RequisitionServiceTable";
        private const string TableName02 = "RequisitionQueriesTable";
        private const string TableName03 = "RequisitionDetailedTable";
        private const string ValidationSheetName = "ValidationRequisition";



        private bool _ignoreItemError;

        private Thread _thread;
        public List<SpecialRestriction.SpecialRestrictionItem> RestrictionList;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            try
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

                settings.SetDefaultCustomSettingValue("maxItemValue", "true");
                settings.SetDefaultCustomSettingValue("autoSortElements", "true");
                settings.SetDefaultCustomSettingValue("assuranceProcessTrack", "false");

                //Setting of Configuration Options from Config File (or default)
                try
                {
                    settings.LoadCustomSettings();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                cbMaxItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("maxItemValue"));
                cbSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("autoSortElements"));
                cbAssuranceProcess.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("assuranceProcessTrack"));

                //
                settings.SaveCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void btnFormatNewSheet_Click(object sender, RibbonControlEventArgs e)
        {
            RequisitionServiceFormat();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(
                    @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnFormatMnttoAssurance_Click(object sender, RibbonControlEventArgs e)
        {
            RequisitionServiceAssuranceFormat();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(
                    @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnFormatExtended_Click(object sender, RibbonControlEventArgs e)
        {
            RequisitionServiceExtendedFormat();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(
                    @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnExcecuteRequisitionService_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                {
                    _ignoreItemError = false;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(() => CreateRequisitionService(true));
                    else
                        _thread = new Thread(() => CreateRequisitionService(false));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateRequisition(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCreateReqIgError_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                {
                    _ignoreItemError = true;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(() => CreateRequisitionService(true));
                    else
                        _thread = new Thread(() => CreateRequisitionService(false));

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateRequisition(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCreateReqDirectOrderItems_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                {
                    _ignoreItemError = false;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(CreateRequisitionScreenServiceExtended);
                    else
                        _thread = new Thread(CreateRequisitionScreenService);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateRequisition(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnManualCreditRequisitionMSE1VR_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                {
                    _ignoreItemError = false;
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(ManualCreditRequisitionExtended);
                    else
                        _thread = new Thread(ManualCreditRequisition);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateRequisition(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnQueryRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewRequisitionList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }



        private void btnReviewRequisitionControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewRequisitionControlList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionControlList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReReviewRequisitionControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewRequisitionControlList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReReviewRequisitionControlList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }

        /// <summary>
        /// Da Formato a la Hoja de Excel Creando los
        /// </summary>
        private void RequisitionServiceFormat()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheetName);

                #region Hoja 1

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01;

                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "REQUISITION SERVICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "G2");

                _cells.GetCell("H1").Value = "OBLIGATORIO";
                _cells.GetCell("H1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("H2").Value = "OPCIONAL";
                _cells.GetCell("H2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("H3").Value = "INFORMATIVO";
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                //
                _cells.GetCell(1, titleRow).Value = "Requested By";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "Requested By Position";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, titleRow).Value = "Indicador de Serie";
                _cells.GetCell(3, titleRow).AddComment("Indica una serie diferente para vales con encabezados comunes (Ej. Flota, número, secuencia A, etc)");
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(4, titleRow).Value = "Requisition Number";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(5, titleRow).Value = "Requisition Type";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionReqTypeList = new List<string>
                {
                    "NI - NORMAL REQUISITION",
                    "PR - PURCHASE REQUISITION",
                    "CR - CREDIT REQUISITION",
                    "LN - LOAN REQUISITION"
                };
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), optionReqTypeList, ValidationSheetName, 1, false);
                _cells.GetCell(6, titleRow).Value = "Transaction Type";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), optionTransTypeList, ValidationSheetName, 2, false);
                _cells.GetCell(7, titleRow).Value = "Required By Date";
                _cells.GetCell(7, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "Original Warehouse";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "Priority Code";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionPriorList = _eFunctions.GetItemCodesString("PI");

                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), optionPriorList, ValidationSheetName, 3, false);

                _cells.GetCell(10, titleRow).Value = "Reference Type";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionRefTypeList = new List<string> {"Work Order", "Equipment No.", "Project No.", "Account Code"};
                _cells.SetValidationList(_cells.GetCell(10, titleRow + 1), optionRefTypeList, ValidationSheetName, 4);

                _cells.GetCell(11, titleRow).Value = "Reference";
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(12, titleRow).Value = "Delivery Instructions A"; //120 caracteres (60/60)
                _cells.GetCell(12, titleRow).AddComment("60 caracteres");
                _cells.GetCell(12, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationLength(_cells.GetCell(12, titleRow + 1), 60);
                
                _cells.GetCell(13, titleRow).Value = "Delivery Instructions B"; //120 caracteres (60/60)
                _cells.GetCell(13, titleRow).AddComment("60 caracteres");
                _cells.GetCell(13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationLength(_cells.GetCell(13, titleRow + 1), 60);
                
                _cells.GetCell(14, titleRow).Value = "Return Cause";
                var returnCauseList = _eFunctions.GetItemCodesString("I2");
                _cells.SetValidationList(_cells.GetCell(14, titleRow + 1), returnCauseList, ValidationSheetName, 7, false);
                _cells.GetCell(14, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                
                _cells.GetCell(15, titleRow).Value = "Issue Question";
                _cells.GetCell(15, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var optionIssueList = new List<string> {"A - VENTAS", "B - RUBROS"};
                _cells.SetValidationList(_cells.GetCell(15, titleRow + 1), optionIssueList, ValidationSheetName, 5, false);

                _cells.GetCell(16, titleRow).Value = "Stock Code";
                _cells.GetCell(16, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(17, titleRow).Value = "Unit Of Issue";
                _cells.GetCell(17, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(18, titleRow).Value = "Quantity";
                _cells.GetCell(18, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Result";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName01);

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                #region Hoja 2

                titleRow = TitleRow02;
                resultColumn = ResultColumn02;

                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "CONSULTA GENERAL - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                var districtList = Districts.GetDistrictList();
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = RequisitionStatus.GetStatusList(true).Select(g => g.Value).ToList();
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 8);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.GetCell("A4").AddComment("--ÁREA GERENCIAL/SUPERINTENDENCIA--\n" +
                                                "INST: IMIS, MINA\n" +
                                                "" + ManagementArea.ManejoDeCarbon.Key + ": " + QuarterMasters.Ferrocarril.Key + ", " + QuarterMasters.PuertoBolivar.Key + ", " + QuarterMasters.PlantasDeCarbon.Key + "\n" +
                                                "" + ManagementArea.Mantenimiento.Key + ": MINA\n" +
                                                "" + ManagementArea.SoporteOperacion.Key + ": ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 9);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 10, false);
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetCell("A6").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B6"), statusList, ValidationSheetName, 11);
                _cells.GetRange("A3", "A6").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B6").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Creation.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 12);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(1, titleRow).Value = "District";
                _cells.GetCell(2, titleRow).Value = "Work Group";
                _cells.GetCell(3, titleRow).Value = "Equipment No.";
                _cells.GetCell(4, titleRow).Value = "Work Order";
                _cells.GetCell(5, titleRow).Value = "Work Order Desc.";
                _cells.GetCell(6, titleRow).Value = "Project No.";
                _cells.GetCell(7, titleRow).Value = "Account Code";
                _cells.GetCell(8, titleRow).Value = "Requisition Number";
                _cells.GetCell(9, titleRow).Value = "Number of Items";
                _cells.GetCell(10, titleRow).Value = "Wo Raised Date";
                _cells.GetCell(11, titleRow).Value = "Wo Plan Date";
                _cells.GetCell(12, titleRow).Value = "Creation Date";
                _cells.GetCell(13, titleRow).Value = "Required Date";
                _cells.GetCell(14, titleRow).Value = "Authorization Date";
                _cells.GetCell(15, titleRow).Value = "Created By";
                _cells.GetCell(16, titleRow).Value = "Requested By";
                _cells.GetCell(17, titleRow).Value = "Requested Pos.";
                _cells.GetCell(18, titleRow).Value = "Authorized By";
                _cells.GetCell(19, titleRow).Value = "Authorized Pos.";
                _cells.GetCell(20, titleRow).Value = "Requisition Status";
                _cells.GetCell(21, titleRow).Value = "Authorized Status";
                _cells.GetCell(22, titleRow).Value = "Requisition Type";
                _cells.GetCell(23, titleRow).Value = "Transaction Type";
                _cells.GetCell(24, titleRow).Value = "Original Warehouse";
                _cells.GetCell(25, titleRow).Value = "Priority Code";
                _cells.GetCell(26, titleRow).Value = "Egi";

                _cells.GetCell(resultColumn, titleRow).Value = "Result";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName02);

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                #region Hoja 3

                //titleRow = TitleRow03;
                //resultColumn = 10;

                _excelApp.ActiveWorkbook.Sheets[3].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "DETALLE CONSULTA - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }


        /// <summary>
        /// Da Formato a la Hoja de Excel Creando los
        /// </summary>
        private void RequisitionServiceAssuranceFormat()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheetName);

                #region Hoja 1

                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01;

                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "REQUISITION SERVICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "G2");

                _cells.GetCell("H1").Value = "OBLIGATORIO";
                _cells.GetCell("H1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("H2").Value = "OPCIONAL";
                _cells.GetCell("H2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("H3").Value = "INFORMATIVO";
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                //
                _cells.GetCell(1, titleRow).Value = "Requested By";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "Requested By Position";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, titleRow).Value = "Indicador de Serie";
                _cells.GetCell(3, titleRow).AddComment("Indica una serie diferente para vales con encabezados comunes (Ej. Flota, número, secuencia A, etc)");
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(4, titleRow).Value = "Requisition Number";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(5, titleRow).Value = "Requisition Type";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionReqTypeList = new List<string>
                {
                    "NI - NORMAL REQUISITION",
                    "PR - PURCHASE REQUISITION",
                    "CR - CREDIT REQUISITION",
                    "LN - LOAN REQUISITION"
                };
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), optionReqTypeList, ValidationSheetName, 1, false);
                _cells.GetCell(6, titleRow).Value = "Transaction Type";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), optionTransTypeList, ValidationSheetName, 2, false);
                _cells.GetCell(7, titleRow).Value = "Required By Date";
                _cells.GetCell(7, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "Original Warehouse";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "Priority Code";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionPriorList = _eFunctions.GetItemCodesString("PI");
                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), optionPriorList, ValidationSheetName, 3, false);

                _cells.GetCell(10, titleRow).Value = "Reference Type";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionRefTypeList = new List<string> { "Work Order", "Equipment No.", "Project No.", "Account Code" };
                _cells.SetValidationList(_cells.GetCell(10, titleRow + 1), optionRefTypeList, ValidationSheetName, 4);

                _cells.GetCell(11, titleRow).Value = "Reference";
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(12, titleRow).Value = "Delivery Instructions A"; //60 caracteres
                _cells.GetCell(12, titleRow).AddComment("60 caracteres");
                _cells.GetCell(12, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationLength(_cells.GetCell(12, titleRow + 1), 60);

                _cells.GetCell(13, titleRow).Value = "Delivery Instructions B"; //60 caracteres
                _cells.GetCell(13, titleRow).AddComment("60 caracteres");
                _cells.GetCell(13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationLength(_cells.GetCell(13, titleRow + 1), 60);

                _cells.GetCell(14, titleRow).Value = "Return Cause";
                _cells.GetCell(14, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var returnCauseList = _eFunctions.GetItemCodesString("I2");
                _cells.SetValidationList(_cells.GetCell(14, titleRow + 1), returnCauseList, ValidationSheetName, 7, false);

                _cells.GetCell(15, titleRow).Value = "Issue Question";
                _cells.GetCell(15, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var optionIssueList = new List<string> { "A - VENTAS", "B - RUBROS" };
                _cells.SetValidationList(_cells.GetCell(15, titleRow + 1), optionIssueList, ValidationSheetName, 5, false);

                _cells.GetCell(16, titleRow).Value = "Stock Code";
                _cells.GetCell(16, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(17, titleRow).Value = "Unit Of Issue";
                _cells.GetCell(17, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(18, titleRow).Value = "Quantity";
                _cells.GetCell(18, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Result";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName01);

                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                #region Hoja 2

                titleRow = TitleRow02;
                resultColumn = ResultColumn02;

                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "CONSULTA GENERAL - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                var districtList = Districts.GetDistrictList();
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = RequisitionStatus.GetStatusList(true).Select(g => g.Value).ToList();
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 8);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.GetCell("A4").AddComment("--ÁREA GERENCIAL/SUPERINTENDENCIA--\n" +
                                                "INST: IMIS, MINA\n" +
                                                "" + ManagementArea.ManejoDeCarbon.Key + ": " + QuarterMasters.Ferrocarril.Key + ", " + QuarterMasters.PuertoBolivar.Key + ", " + QuarterMasters.PlantasDeCarbon.Key + "\n" +
                                                "" + ManagementArea.Mantenimiento.Key + ": MINA\n" +
                                                "" + ManagementArea.SoporteOperacion.Key + ": ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 9);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 10, false);
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetCell("A6").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B6"), statusList, ValidationSheetName, 11);
                _cells.GetRange("A3", "A6").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B6").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Creation.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 12);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(1, titleRow).Value = "District";
                _cells.GetCell(2, titleRow).Value = "Work Group";
                _cells.GetCell(3, titleRow).Value = "Equipment No.";
                _cells.GetCell(4, titleRow).Value = "Work Order";
                _cells.GetCell(5, titleRow).Value = "Work Order Desc.";
                _cells.GetCell(6, titleRow).Value = "Project No.";
                _cells.GetCell(7, titleRow).Value = "Account Code";
                _cells.GetCell(8, titleRow).Value = "Requisition Number";
                _cells.GetCell(9, titleRow).Value = "Number of Items";
                _cells.GetCell(10, titleRow).Value = "Wo Raised Date";
                _cells.GetCell(11, titleRow).Value = "Wo Plan Date";
                _cells.GetCell(12, titleRow).Value = "Creation Date";
                _cells.GetCell(13, titleRow).Value = "Required Date";
                _cells.GetCell(14, titleRow).Value = "Authorization Date";
                _cells.GetCell(15, titleRow).Value = "Created By";
                _cells.GetCell(16, titleRow).Value = "Requested By";
                _cells.GetCell(17, titleRow).Value = "Requested Pos.";
                _cells.GetCell(18, titleRow).Value = "Authorized By";
                _cells.GetCell(19, titleRow).Value = "Authorized Pos.";
                _cells.GetCell(20, titleRow).Value = "Requisition Status";
                _cells.GetCell(21, titleRow).Value = "Authorized Status";
                _cells.GetCell(22, titleRow).Value = "Requisition Type";
                _cells.GetCell(23, titleRow).Value = "Transaction Type";
                _cells.GetCell(24, titleRow).Value = "Original Warehouse";
                _cells.GetCell(25, titleRow).Value = "Priority Code";
                _cells.GetCell(26, titleRow).Value = "Egi";

                _cells.GetCell(resultColumn, titleRow).Value = "Result";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName02);

                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                #region Hoja 3

                //titleRow = TitleRow03;
                //resultColumn = 10;

                _excelApp.ActiveWorkbook.Sheets[3].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "DETALLE CONSULTA - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }


        /// <summary>
        /// Da Formato a la Hoja de Excel de forma extendida
        /// </summary>
        private void RequisitionServiceExtendedFormat()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheetName);
                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01 + "Ext";

                #region FormatHeader

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "REQUISITION SERVICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "G2");

                _cells.GetCell("H1").Value = "OBLIGATORIO";
                _cells.GetCell("H1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("H2").Value = "OPCIONAL";
                _cells.GetCell("H2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("H3").Value = "INFORMATIVO";
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                #endregion

                var titleRow = TitleRow01Ext;
                var resultColumn = ResultColumn01Ext;

                //
                _cells.GetCell(1, titleRow).Value = "Requested By";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "Requested By Position";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, titleRow).Value = "Indicador de Serie";
                _cells.GetCell(3, titleRow).AddComment("Indica una serie diferente para vales con encabezados comunes (Ej. Flota, número, secuencia A, etc)");
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(4, titleRow).Value = "Requisition Number";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(5, titleRow).Value = "Requisition Type";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionReqTypeList = new List<string>
                {
                    "NI - NORMAL REQUISITION",
                    "PR - PURCHASE REQUISITION",
                    "CR - CREDIT REQUISITION",
                    "LN - LOAN REQUISITION"
                };
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), optionReqTypeList, ValidationSheetName, 1, false);
                _cells.GetCell(6, titleRow).Value = "Transaction Type";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), optionTransTypeList, ValidationSheetName, 2, false);
                _cells.GetCell(7, titleRow).Value = "Required By Date";
                _cells.GetCell(7, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "Original Warehouse";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "Priority Code";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionPriorList = _eFunctions.GetItemCodesString("PI");
                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), optionPriorList, ValidationSheetName, 3, false);

                _cells.GetCell(10, titleRow).Value = "Work Order";
                _cells.GetCell(11, titleRow).Value = "Equipment No.";
                _cells.GetCell(12, titleRow).Value = "Project No.";
                _cells.GetCell(13, titleRow).Value = "Account Code";
                _cells.GetRange(10, titleRow, 13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(14, titleRow).Value = "Delivery Instructions A"; //60 caracteres
                _cells.GetCell(14, titleRow).AddComment("60 caracteres");
                _cells.GetCell(14, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationLength(_cells.GetCell(14, titleRow + 1), 60);

                _cells.GetCell(15, titleRow).Value = "Delivery Instructions B"; //60 caracteres
                _cells.GetCell(15, titleRow).AddComment("60 caracteres");
                _cells.GetCell(15, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationLength(_cells.GetCell(15, titleRow + 1), 60);

                _cells.GetCell(16, titleRow).Value = "Return Cause";
                var returnCauseList = _eFunctions.GetItemCodesString("I2");
                _cells.SetValidationList(_cells.GetCell(16, titleRow + 1), returnCauseList, ValidationSheetName, 7, false);
                _cells.GetCell(16, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                
                _cells.GetCell(17, titleRow).Value = "Issue Question";
                _cells.GetCell(17, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var optionIssueList = new List<string> {"A - VENTAS", "B - RUBROS"};
                _cells.SetValidationList(_cells.GetCell(17, titleRow + 1), optionIssueList, ValidationSheetName, 5, false);

                _cells.GetCell(18, titleRow).Value = "Stock Code";
                _cells.GetCell(18, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(19, titleRow).Value = "Unit Of Issue";
                _cells.GetCell(19, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(20, titleRow).Value = "Quantity";
                _cells.GetCell(20, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Result";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName01);

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public RequisitionHeader PopulateRequisitionHeader(int currentRow, bool isExtended)
        {
            string allocPcA;
            string districtCode;
            string costDistrictAllocation;
            string requestedBy;
            string requiredByPosition;
            string seriesIndicator;
            string requisitionType;
            string issueTranType;
            string requiredByDate;
            string originalWarehouse;
            string priorityCode;
            bool partIssueIndicator = true;
            bool protectedIndicator = false;
            string deliveryInstructionsA;
            string deliveryInstructionsB;
            string answerB;
            string answerD;
            string answerN;
            string answerP;
            string workOrderAllocation;
            string workProjectIndicatorAllocation;
            string equipmentAllocation;
            string projectAllocation;
            string costCentreAllocation;
            string requisitionNumber;



            if (isExtended)
            {
                allocPcA = "100";
                districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct) ? "ICOR" : _frmAuth.EllipseDstrct;
                costDistrictAllocation = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct) ? "ICOR" : _frmAuth.EllipseDstrct;
                requestedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null ? _frmAuth.EllipseUser : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                requiredByPosition = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) == null ? _frmAuth.EllipsePost : _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                seriesIndicator = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                requisitionNumber = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                requisitionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value));
                issueTranType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                requiredByDate = _cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value);
                originalWarehouse = _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value);
                priorityCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));

                workOrderAllocation = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                equipmentAllocation = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                projectAllocation = _cells.GetNullOrTrimmedValue(_cells.GetCell(12, currentRow).Value);
                costCentreAllocation = _cells.GetNullOrTrimmedValue(_cells.GetCell(13, currentRow).Value);
                workProjectIndicatorAllocation = string.IsNullOrWhiteSpace(projectAllocation) ? "W" : "P";
                deliveryInstructionsA = _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value);
                deliveryInstructionsB = _cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value);
                if (requisitionType != null && requisitionType.Equals("CR"))
                {
                    answerB = null;
                    answerD = null;
                    answerN = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(16, currentRow).Value));
                    answerP = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(17, currentRow).Value));

                    //revalidación porque el campo no acepta caracteres vacíos y getCodeKey devuelve vacío si no hay código
                    answerN = string.IsNullOrWhiteSpace(answerN) ? null : answerN;
                    answerP = string.IsNullOrWhiteSpace(answerP) ? null : answerP;
                }
                else
                {
                    answerB = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(16, currentRow).Value));
                    answerD = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(17, currentRow).Value));
                    answerN = null;
                    answerP = null;

                    //revalidación porque el campo no acepta caracteres vacíos y getCodeKey devuelve vacío si no hay código
                    answerB = string.IsNullOrWhiteSpace(answerB) ? null : answerB;
                    answerD = string.IsNullOrWhiteSpace(answerD) ? null : answerD;
                }
            }
            else
            {
                allocPcA = "100";
                districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct) ? "ICOR" : _frmAuth.EllipseDstrct;
                costDistrictAllocation = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct) ? "ICOR" : _frmAuth.EllipseDstrct;
                requestedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null ? _frmAuth.EllipseUser : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                requiredByPosition = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) == null ? _frmAuth.EllipsePost : _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                seriesIndicator = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                requisitionNumber = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                requisitionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value));
                issueTranType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                requiredByDate = _cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value);
                originalWarehouse = _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value);
                priorityCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));

                workOrderAllocation = null;
                equipmentAllocation = null;
                projectAllocation = null;
                costCentreAllocation = null;
                workProjectIndicatorAllocation = null;

                deliveryInstructionsA = _cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value);
                deliveryInstructionsB = _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value);

                if (requisitionType != null && requisitionType.Equals("CR"))
                {
                    answerB = null;
                    answerD = null;
                    answerN = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value));
                    answerP = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value));

                    //revalidación porque el campo no acepta caracteres vacíos y getCodeKey devuelve vacío si no hay código
                    answerN = string.IsNullOrWhiteSpace(answerN) ? null : answerN;
                    answerP = string.IsNullOrWhiteSpace(answerP) ? null : answerP;
                }
                else
                {
                    answerB = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value));
                    answerD = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value));
                    answerN = null;
                    answerP = null;

                    //revalidación porque el campo no acepta caracteres vacíos y getCodeKey devuelve vacío si no hay código
                    answerB = string.IsNullOrWhiteSpace(answerB) ? null : answerB;
                    answerD = string.IsNullOrWhiteSpace(answerD) ? null : answerD;
                }

                string switchCase = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                var reference = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                switch (switchCase)
                {
                    case "Work Order":
                        workOrderAllocation = reference;
                        workProjectIndicatorAllocation = "W"; //Solo aplica para MSO140
                        break;
                    case "Equipment No.":
                        equipmentAllocation = reference;
                        break;
                    case "Project No.":
                        projectAllocation = reference;
                        workProjectIndicatorAllocation = "P"; //Solo aplica para MSO140
                        break;
                    case "Account Code":
                        costCentreAllocation = reference;
                        break;
                }
            }


            var requisitionHeader = new RequisitionHeader
            {
                AllocPcA = allocPcA,
                DistrictCode = districtCode,
                CostDistrictA = costDistrictAllocation,
                RequestedBy = requestedBy,
                RequiredByPos = requiredByPosition,
                IndSerie = seriesIndicator,
                IreqNo = requisitionNumber,
                IreqType = requisitionType,
                IssTranType = issueTranType,
                RequiredByDate = requiredByDate,
                OrigWhouseId = originalWarehouse,
                PriorityCode = priorityCode,
                PartIssue = partIssueIndicator,
                PartIssueSpecified = true,
                ProtectedInd = protectedIndicator,
                ProtectedIndSpecified = true,
                DelivInstrA = deliveryInstructionsA,
                DelivInstrB = deliveryInstructionsB,
                AnswerB = answerB,
                AnswerD = answerD,
                AnswerN = answerN,
                AnswerP = answerP,
                WorkOrderA = workOrderAllocation,
                WorkProjectIndA = workProjectIndicatorAllocation,
                EquipmentA = equipmentAllocation,
                ProjectA = projectAllocation,
                CostCentreA = costCentreAllocation
            };

            //La condición dice:
            //Si la posición existe en la tabla debe cumplir con la prioridad
            //Si la prioridad existe en la tabla debe cumplir con la posición
            //Si la posición existe en la tabla y cumple con la prioridad, debe cumplir con el flag de orden
            var isRestricted = false;

            //Valida si existe en la tabla de restricciones
            foreach (var resItem in RestrictionList)
            {
                if (resItem.Position.Trim().ToUpper().Equals(_frmAuth.EllipsePost.Trim().ToUpper()) || resItem.Code.Trim().ToUpper().Equals(requisitionHeader.PriorityCode.Trim().ToUpper()))
                {
                    isRestricted = true;
                    break;
                }
            }

            //Si hay restricción valida el tipo
            if (isRestricted)
            {
                var isPositionRestrictionValid = false;
                var isMandatoryOrder = false;
                foreach (var resItem in RestrictionList)
                {
                    if (resItem.Position.Trim().ToUpper().Equals(_frmAuth.EllipsePost.Trim().ToUpper()) && resItem.Code.Trim().ToUpper().Equals(requisitionHeader.PriorityCode.Trim().ToUpper()))
                    {
                        isPositionRestrictionValid = true;
                        isMandatoryOrder = resItem.MandatoryWorkOrder;
                        break;
                    }
                }

                if (isMandatoryOrder && string.IsNullOrWhiteSpace(requisitionHeader.WorkOrderA))
                    throw new Exception(@"MANDATORY WORK ORDER IN PRIORITY CODE " + requisitionHeader.PriorityCode.Trim().ToUpper() + " FOR LOGGED POSITION " + _frmAuth.EllipsePost.Trim().ToUpper());
                if (!isPositionRestrictionValid)
                    throw new Exception(@"UNAUTHORISED PRIORITY CODE " + requisitionHeader.PriorityCode.Trim().ToUpper() + " FOR LOGGED POSITION " + _frmAuth.EllipsePost.Trim().ToUpper());
            }

            return requisitionHeader;
        }

        public RequisitionItem PopulateRequisitionItem(int currentRow, int indexList, bool isExtended)
        {
            const bool partialAllowed = true; //forced
            string stockCode;
            string unitOfMeasure;
            decimal quantityRequired;


            if (isExtended)
            {
                stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(18, currentRow).Value);
                unitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(19, currentRow).Value);
                quantityRequired = Convert.ToDecimal(_cells.GetCell(20, currentRow).Value);
            }
            else
            {
                stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value);
                unitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(17, currentRow).Value);
                quantityRequired = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);
            }


            var item = new RequisitionItem
            {
                Index = indexList,
                ItemType = "S",
                PartIssueSpecified = true,
                PartIssue = partialAllowed,
                StockCode = stockCode,
                UnitOfMeasure = unitOfMeasure,
                QuantityRequired = quantityRequired
            };
            item.StockCode = (item.StockCode != null && item.StockCode.Length < 9) ? item.StockCode.PadLeft(9, '0') : item.StockCode;

            //si es item de orden directa o no
            var sqlQuery = Queries.GetItemDirectOrder(item.StockCode);
            var odr = _eFunctions.GetQueryResult(sqlQuery);
            if (odr.Read() && MyUtilities.IsTrue(odr["DIRECT_ORDER_IND"]))
                item.DirectOrderIndicator = true;

            //Obtengo la unidad del Stock Code que voy a registrar
            if (string.IsNullOrWhiteSpace(item.UnitOfMeasure))
            {
                sqlQuery = Queries.GetItemUnitOfIssue(item.StockCode);
                odr = _eFunctions.GetQueryResult(sqlQuery);

                //si se pudo obtener la Unidad
                if (odr.Read())
                    item.UnitOfMeasure = "" + odr["UNIT_OF_ISSUE"];
            }

            return item;
        }


        /// <summary>
        /// Recorre y Crea los vales de a tabla de Excel
        /// </summary>
        private void CreateRequisitionService(bool isExtended)
        {
            //instancia del Servicio
            var requisitionService = new RequisitionService.RequisitionService();

            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();


                int resultColumn, titleRow;
                if (isExtended)
                {
                    resultColumn = ResultColumn01Ext;
                    titleRow = TitleRow01Ext;
                }
                else
                {
                    resultColumn = ResultColumn01;
                    titleRow = TitleRow01;
                }


                #region SortFields

                if (cbSortItems.Checked)
                {
                    Excel.ListObject excelSheetItems = _cells.GetRange(TableName01).ListObject;
                    //Organiza las celdas de forma que se creen la menor cantidad de vales posibles
                    if (excelSheetItems.Sort.SortFields.Count > 0)
                        excelSheetItems.Sort.SortFields.Clear();

                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(1, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(5, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(7, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(8, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(12, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(13, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(14, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    //_excelSheetItems.Sort.SortFields.Add(_cells.GetCell(16, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    //    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.Apply();
                }

                #endregion
                
                
                //Header
                var opRequisition = new RequisitionService.OperationContext();
                
                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                requisitionService.Url = urlService + "/RequisitionService";

                opRequisition.district = _frmAuth.EllipseDstrct;
                opRequisition.maxInstances = 100;
                opRequisition.position = _frmAuth.EllipsePost;
                opRequisition.returnWarnings = false;
                
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);


                var currentRow = titleRow + 1;
                var currentRowHeader = currentRow;
                const int seriesIndicatorColumn = 3;
                var itemIndicatorColumn = isExtended ? 18 : 16;
                const int requisitionNoColumn = 4;
                _cells.ClearTableRangeColumn(TableName01, resultColumn);
                _cells.ClearTableRangeColumn(TableName01, requisitionNoColumn);
                RestrictionList = SpecialRestriction.GetPositionRestrictions(_eFunctions);

                var itemList = new List<RequisitionItem>();
                RequisitionService.RequisitionServiceCreateHeaderReplyDTO headerCreateReply = null;

                RequisitionHeader prevReqHeader = null;
                var abortRequisition = false;

                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndicatorColumn, currentRow).Value) != null ||
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(seriesIndicatorColumn, currentRow).Value) != null)
                {
                    try
                    {
                        //obtengo los datos para el encabezado
                        var curReqHeader = PopulateRequisitionHeader(currentRow, isExtended);

                        //si es el primer elemento creo un nuevo encabezado
                        if (prevReqHeader == null)
                        {
                            var headerCreateRequest = curReqHeader.GetCreateHeaderRequest();

                            headerCreateReply = requisitionService.createHeader(opRequisition, headerCreateRequest);
                            curReqHeader.IreqNo = headerCreateReply.ireqNo;
                            prevReqHeader = curReqHeader;
                            currentRowHeader = currentRow;
                            abortRequisition = false;
                        }
                        //comparo si el nuevo registro corresponde a un nuevo encabezado o si he alcanzado 99 items. Si es así, envío el encabezado anterior y creo un encabezado
                        else if (!prevReqHeader.Equals(curReqHeader) || (cbMaxItems.Checked && itemList.Count >= 99))
                        {
                            //agrego los items que tenga hasta el momento al encabezado
                            foreach (var item in itemList)
                            {
                                try
                                {
                                    #region itemDto

                                    var itemDto = new List<RequisitionService.RequisitionItemDTO>
                                    {
                                        item.GetRequisitionItemDto()
                                    };

                                    var itemRequest = new RequisitionService.RequisitionServiceCreateItemRequestDTO
                                    {
                                        districtCode = prevReqHeader.DistrictCode,
                                        ireqNo = prevReqHeader.IreqNo,
                                        ireqType = prevReqHeader.IreqType,
                                        requisitionItems = itemDto.ToArray(),

                                    };

                                    #endregion

                                    requisitionService.createItem(opRequisition, itemRequest);
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = itemRequest.ireqNo;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }
                                catch (Exception ex)
                                {
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 += "ERROR: " + ex.Message;
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = prevReqHeader.IreqNo;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    abortRequisition = true;
                                }

                                _cells.GetCell(resultColumn, currentRowHeader + item.Index).Select();
                            }

                            //aborto o finalizo según el resultado de los items
                            if (abortRequisition && !_ignoreItemError)
                            {
                                #region FailedItemProcess

                                var addMessage = "";
                                try
                                {
                                    DeleteHeader(requisitionService, headerCreateReply, opRequisition);
                                }
                                catch (Exception ex)
                                {
                                    addMessage = ". ERROR AL ELIMINAR. " + ex.Message;
                                }

                                foreach (var item in itemList)
                                {
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 += " - ELIMINADO";
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 += " - VALE ELIMINADO" + addMessage;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                }

                                prevReqHeader = null;
                                abortRequisition = false;

                                #endregion
                            }
                            else
                            {
                                var finaliseRequest = new RequisitionService.RequisitionServiceFinaliseRequestDTO
                                {
                                    ireqNo = prevReqHeader.IreqNo,
                                    ireqType = prevReqHeader.IreqType,
                                    districtCode = prevReqHeader.DistrictCode
                                };

                                //Se añade este bloque try/catch porque el tiempo excesivo de finalización afecta el siguiente item de la lista. Cuando esto ocurra no afectará el proceso
                                try
                                {
                                    requisitionService.finalise(opRequisition, finaliseRequest);
                                }
                                catch (TimeoutException ex)
                                {
                                    _cells.GetCell(resultColumn, currentRow - 1).Value2 = _cells.GetCell(resultColumn, currentRow - 1).Value2 + " " + ex.Message;
                                    _cells.GetCell(resultColumn, currentRow - 1).Style = StyleConstants.Warning;
                                    _cells.GetCell(requisitionNoColumn, currentRow - 1).Style = StyleConstants.Warning;
                                }
                            }

                            //creo el nuevo encabezado y reinicio variables
                            prevReqHeader = null; //no es una línea inservible. Es necesaria por si se produce una excepción al momento de creación de un nuevo encabezado
                            currentRowHeader = currentRow;
                            abortRequisition = false;
                            itemList = new List<RequisitionItem>();
                            var headerCreateRequest = curReqHeader.GetCreateHeaderRequest();
                            headerCreateReply = requisitionService.createHeader(opRequisition, headerCreateRequest);
                            curReqHeader.IreqNo = headerCreateReply.ireqNo;
                            prevReqHeader = curReqHeader;
                        }

                        //Obtengo los datos para el item
                        var curItem = PopulateRequisitionItem(currentRow, itemList.Count, isExtended);

                        if (string.IsNullOrWhiteSpace(curItem.UnitOfMeasure))
                        {
                            abortRequisition = true;
                            _cells.GetCell(resultColumn, currentRow).Value2 += curItem.StockCode + " NO EXISTE UNIDAD DE MEDIDA EN EL CATALOGO PARA ESTE STOCK CODE";
                            _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                            curItem.StockCode = ""; //Se vacía el campo para conservar la estructura del vale, pero para que indique el error
                        }

                        if (curItem.DirectOrderIndicator)
                        {
                            abortRequisition = true;

                            _cells.GetCell(resultColumn, currentRow).Value2 += curItem.StockCode + " ITEM DE ORDEN DIRECTA. DEBE CREAR EL VALE CON OTRO MÉTODO";
                            _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                            curItem.StockCode = ""; //Se vacía el campo para conservar la estructura del vale, pero para que indique el error
                        }

                        itemList.Add(curItem);
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRow).Value2 = ex.Message;
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(requisitionNoColumn, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }
                } //finaliza el while del proceso completo

                //para el último encabezado
                if (prevReqHeader == null) return;
                //agrego los items que tenga hasta el momento al encabezado
                foreach (var item in itemList)
                {
                    try
                    {
                        var itemListDto = new List<RequisitionService.RequisitionItemDTO> {item.GetRequisitionItemDto()};

                        var itemRequest = new RequisitionService.RequisitionServiceCreateItemRequestDTO
                        {
                            districtCode = prevReqHeader.DistrictCode,
                            ireqNo = prevReqHeader.IreqNo,
                            ireqType = prevReqHeader.IreqType,
                            requisitionItems = itemListDto.ToArray()
                        };
                        requisitionService.createItem(opRequisition, itemRequest);

                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = itemRequest.ireqNo;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 += "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = prevReqHeader.IreqNo;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        abortRequisition = true;

                    }

                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Select();
                }

                //aborto o finalizo según el resultado de los items
                if (abortRequisition && !_ignoreItemError)
                {
                    var addMessage = "";
                    try
                    {
                        DeleteHeader(requisitionService, headerCreateReply, opRequisition);
                    }
                    catch (Exception ex)
                    {
                        addMessage = ". ERROR AL ELIMINAR. " + ex.Message;
                    }

                    foreach (var item in itemList)
                    {
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 += " - ELIMINADO";
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 += " - VALE ELIMINADO" + addMessage;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                    }
                }
                else
                {
                    var finaliseRequest = new RequisitionService.RequisitionServiceFinaliseRequestDTO
                    {
                        ireqNo = prevReqHeader.IreqNo,
                        ireqType = prevReqHeader.IreqType,
                        districtCode = prevReqHeader.DistrictCode
                    };

                    try
                    {
                        requisitionService.finalise(opRequisition, finaliseRequest);
                    }
                    catch (TimeoutException ex)
                    {
                        _cells.GetCell(resultColumn, currentRow - 1).Value2 = _cells.GetCell(resultColumn, currentRow - 1).Value2 + " " + ex.Message;
                        _cells.GetCell(resultColumn, currentRow - 1).Style = StyleConstants.Warning;
                        _cells.GetCell(requisitionNoColumn, currentRow - 1).Style = StyleConstants.Warning;
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true); //Se restaura por 'Pooled Connection Request Timed Out'
                _eFunctions.CloseConnection();

                requisitionService?.Dispose();
            }
        }

        private void CreateRequisitionScreenServiceExtended()
        {
            CreateRequisitionScreenService(true);
        }

        private void CreateRequisitionScreenService()
        {
            CreateRequisitionScreenService(false);
        }

        /// <summary>
        /// Recorre y Crea los vales de a tabla de Excel para items catalogados como de Orden Directa
        /// </summary>
        private void CreateRequisitionScreenService(bool isExtended)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                int resultColumn, titleRow;
                if (isExtended)
                {
                    resultColumn = ResultColumn01Ext;
                    titleRow = TitleRow01Ext;
                }
                else
                {
                    resultColumn = ResultColumn01;
                    titleRow = TitleRow01;
                }

                _cells.ClearTableRangeColumn(TableName01, resultColumn);
                _cells.ClearTableRangeColumn(TableName01, 4);

                #region SortItems

                if (cbSortItems.Checked)
                {
                    Excel.ListObject excelSheetItems = _cells.GetRange(TableName01).ListObject;
                    //Organiza las celdas de forma que se creen la menor cantidad de vales posibles
                    if (excelSheetItems.Sort.SortFields.Count > 0)
                        excelSheetItems.Sort.SortFields.Clear();
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(1, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(5, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(7, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(8, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(12, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(13, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.SortFields.Add(_cells.GetCell(14, titleRow), Excel.XlSortOn.xlSortOnValues,
                        Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    //_excelSheetItems.Sort.SortFields.Add(_cells.GetCell(16, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    //    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                    excelSheetItems.Sort.Apply();
                }

                #endregion

                #region ScreenService

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'

                //ScreenService Opción en reemplazo de los servicios
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                var proxySheet = new Screen.ScreenService {Url = urlService + "/ScreenService"};
                ////ScreenService
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                #endregion

                var currentRow = titleRow + 1;
                var currentRowHeader = currentRow;

                var itemList = new List<RequisitionItem>();

                const int seriesIndicatorColumn = 3;
                var itemIndicatorColumn = isExtended ? 18 : 16;
                const int requisitionNoColumn = 4;

                RequisitionHeader prevReqHeader = null;
                RequisitionHeader curReqHeader;
                RestrictionList = SpecialRestriction.GetPositionRestrictions(_eFunctions);
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndicatorColumn, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(seriesIndicatorColumn, currentRow).Value) != null)
                {
                    try
                    {
                        //obtengo los datos para el encabezado
                        curReqHeader = PopulateRequisitionHeader(currentRow, false);

                        //si el nuevo registro corresponde a un encabezado nuevo diferente creo el vale anterior con sus items respectivos

                        #region CreateNewRequisition

                        if ((prevReqHeader != null && !prevReqHeader.Equals(curReqHeader) && itemList.Count > 0) || (cbMaxItems.Checked && itemList.Count >= 99))
                        {
                            try
                            {
                                //Crear el encabezado en el MSO140
                                _eFunctions.RevertOperation(opSheet, proxySheet);
                                //ejecutamos el programa
                                var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                //Validamos el ingreso
                                if (reply.mapName != "MSM140A")
                                    throw new Exception("ERROR: Se ha producido un error al intentar ingresar al programa. No se puede acceder al MSO140/MSM140A. " + reply.message);

                                //se adicionan los valores a los campos
                                var arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("REQ_NO1I", "" + prevReqHeader.IreqNo);
                                arrayFields.Add("TRAN_TYPE1I", "" + prevReqHeader.IssTranType);
                                arrayFields.Add("REQ_BY_DATE1I", "" + prevReqHeader.RequiredByDate);
                                arrayFields.Add("WHOUSE_ID1I", "" + prevReqHeader.OrigWhouseId);
                                arrayFields.Add("PART_ISSUE1I", "Y");
                                arrayFields.Add("PROT_IND1I", "Y");
                                arrayFields.Add("ALLOC_PCA1I", "" + prevReqHeader.AllocPcA);
                                arrayFields.Add("WORK_PROJ_INDA1I", prevReqHeader.WorkProjectIndA);
                                arrayFields.Add("WORK_PROJA1I", prevReqHeader.WorkOrderA ?? prevReqHeader.ProjectA);
                                arrayFields.Add("COST_CENTREA1I", "" + prevReqHeader.CostCentreA);
                                arrayFields.Add("EQUIP_REFA1I", "" + prevReqHeader.EquipmentA);
                                arrayFields.Add("DELIV_INSTRA1I", "" + prevReqHeader.DelivInstrA);
                                arrayFields.Add("DELIV_INSTRB1I", "" + prevReqHeader.DelivInstrB);
                                arrayFields.Add("PRIORITY_CODE1I", "" + prevReqHeader.PriorityCode);
                                arrayFields.Add("ANSWER_B1I", "" + prevReqHeader.AnswerB);
                                arrayFields.Add("ANSWER_D1I", "" + prevReqHeader.AnswerD);
                                arrayFields.Add("ANSWER_N1I", "" + prevReqHeader.AnswerN);
                                arrayFields.Add("ANSWER_P1I", "" + prevReqHeader.AnswerP);
                                arrayFields.Add("REQUESTED_BY1I", "" + prevReqHeader.RequestedBy);

                                //enviar el encabezado MSM140A
                                var request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };
                                reply = proxySheet.submit(opSheet, request);
                                //Confirmar el encabezado MSM140A
                                while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
                                {
                                    request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                }

                                //no hay errores ni advertencias
                                if (reply == null || _eFunctions.CheckReplyError(reply))
                                    throw new Exception("ERROR: " + reply.message);

                                //MSM14BA
                                if (reply.mapName != "MSM14BA")
                                    throw new Exception("ERROR: Se ha producido un error al crear el encabezado. No se puede acceder al MSO140/MSM14BA. " + reply.message);

                                var parItemIndex = 0; //controla el par de items por pantalla
                                //agrego los items que tenga hasta el momento al encabezado
                                foreach (var item in itemList)
                                {
                                    //asigno número de vale si se genera
                                    var screenValues = new ArrayScreenNameValue(reply.screenFields);
                                    if (!string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                        prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                    if (parItemIndex % 2 == 0)
                                        arrayFields = new ArrayScreenNameValue();
                                    arrayFields.Add("QTY_REQD1I" + (parItemIndex + 1), "" + item.QuantityRequired);
                                    arrayFields.Add("UOM1I" + (parItemIndex + 1), item.UnitOfMeasure);
                                    arrayFields.Add("TYPE1I" + (parItemIndex + 1), "S");
                                    arrayFields.Add("DESCR_A1I" + (parItemIndex + 1), item.StockCode);
                                    arrayFields.Add("PART_ISSUE1I" + (parItemIndex + 1), "Y");

                                    //envío si es el último item de la lista o si es el segundo de la pantalla
                                    if (item == itemList[itemList.Count - 1] || parItemIndex > 0)
                                    {
                                        request = new Screen.ScreenSubmitRequestDTO
                                        {
                                            screenFields = arrayFields.ToArray(),
                                            screenKey = "1"
                                        };
                                        reply = proxySheet.submit(opSheet, request);

                                        //evalúo si hay error y cancelo
                                        if (_eFunctions.CheckReplyError(reply))
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);
                                        //mientras confirmación o bodega
                                        while (reply != null && (reply.mapName == "MSM14BA" || reply.mapName == "MSM14CA"))
                                        {
                                            //evalúo error
                                            if (_eFunctions.CheckReplyError(reply))
                                                throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);

                                            //asigno número de vale si se genera
                                            screenValues = new ArrayScreenNameValue(reply.screenFields);
                                            if (screenValues.GetField("IREQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                                prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                            //si es una nueva pantalla de items
                                            if (reply.mapName == "MSM14BA" && item != itemList[itemList.Count - 1])
                                                if (screenValues.GetField("DESCR_A1I1") != null && string.IsNullOrWhiteSpace(screenValues.GetField("DESCR_A1I1").value))
                                                    break;
                                            ////MSM14CA  - Warehouse Holdings
                                            if (reply.mapName == "MSM14CA")
                                            {
                                                if (screenValues.GetField("TOTAL_REQD1I") == null)
                                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "No se ha encontrado el valor total requerido al asignar a bodega. MSM14CA");
                                                string selWarehouseIndex = "";
                                                //obtengo solo la lista de pares del objeto para actualizarla
                                                var screenArray = screenValues.ToArray();
                                                foreach (var parValue in screenArray)
                                                {
                                                    if (parValue.fieldName != null && parValue.fieldName.StartsWith("WHOUSE_ID_") && parValue.value == prevReqHeader.OrigWhouseId)
                                                        selWarehouseIndex = parValue.fieldName.Replace("WHOUSE_ID_", "");
                                                    if (parValue.fieldName != null && parValue.fieldName.StartsWith("QTY_REQD_"))
                                                        parValue.value = "";
                                                }

                                                //reingreso la lista al objeto del screen y actualizo la cantidad del w/h que quiero de acuerdo a lo realizado anteriormente
                                                if (string.IsNullOrWhiteSpace(selWarehouseIndex))
                                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "El item no está catalogado en la bodega seleccionada. MSM14CA");
                                                screenValues = new ArrayScreenNameValue(screenArray);
                                                screenValues.SetField("QTY_REQD_" + selWarehouseIndex, "" + screenValues.GetField("TOTAL_REQD1I").value);

                                                //envío el proceso
                                                request = new Screen.ScreenSubmitRequestDTO
                                                {
                                                    screenFields = screenValues.ToArray(),
                                                    screenKey = "1"
                                                };
                                                reply = proxySheet.submit(opSheet, request);
                                                continue; //continúa con el siguiente while
                                            }

                                            ////Confirm MSM14BA o cualquier otra confirmación que no requiera datos
                                            request = new Screen.ScreenSubmitRequestDTO
                                            {
                                                screenKey = "1"
                                            };
                                            reply = proxySheet.submit(opSheet, request);
                                        }
                                    }

                                    parItemIndex++;
                                    if (parItemIndex > 1)
                                        parItemIndex = 0;

                                    //OK del item
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }

                                //Confirmo la creación de todos los items. Si no llega a este punto es por algún problema presentado
                                foreach (var item in itemList)
                                {
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                                    _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }
                            }
                            catch (Exception ex)
                            {
                                var addMessage = "" + ex.Message;
                                try
                                {
                                    if (string.IsNullOrWhiteSpace(prevReqHeader.IreqNo))
                                    {
                                        addMessage += " .NO SE HA REALIZADO NINGUNA ACCIÓN";
                                    }
                                    else
                                    {
                                        //Eliminación del vale por el MSO140
                                        _eFunctions.RevertOperation(opSheet, proxySheet);
                                        //ejecutamos el programa
                                        var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                        var screenValues = new ArrayScreenNameValue(reply.screenFields);

                                        if (_eFunctions.CheckReplyError(reply))
                                            throw new Exception("" + reply.message);

                                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" &&
                                               (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")) &&
                                               (screenValues.GetField("REQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("REQ_NO1I").value)))
                                        {
                                            var arrayFields = new ArrayScreenNameValue();
                                            arrayFields.Add("COMP_DEL1I", "D");

                                            var request = new Screen.ScreenSubmitRequestDTO
                                            {
                                                screenFields = arrayFields.ToArray(),
                                                screenKey = "1"
                                            };
                                            reply = proxySheet.submit(opSheet, request);
                                            screenValues = new ArrayScreenNameValue(reply.screenFields);

                                            if (_eFunctions.CheckReplyError(reply))
                                                throw new Exception(". ERROR AL ELIMINAR " + prevReqHeader.IreqNo + ": " + reply.message);
                                        }

                                        addMessage += " - VALE ELIMINADO " + prevReqHeader.IreqNo;
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    addMessage += ex2;
                                }
                                finally
                                {
                                    addMessage = addMessage.Replace("X2:0011 - INPUT REQUIRED  \"C\" TO COMPLETE OR \"D\" TO DELETE", "X2:0011 - EXISTE UNA ORDEN INCOMPLETA EN PROCESO. INGRESE AL MSO 140 PARA COMPLETARLA/ELIMINARLA");
                                    foreach (var item in itemList)
                                    {
                                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = !string.IsNullOrWhiteSpace(prevReqHeader.IreqNo) ? prevReqHeader.IreqNo + " - ELIMINADO" : "";
                                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = addMessage;
                                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    }
                                }
                            }
                            finally
                            {
                                //creo el nuevo encabezado y reinicio variables
                                currentRowHeader = currentRow;
                                itemList = new List<RequisitionItem>();
                                prevReqHeader = curReqHeader;
                            }
                        }

                        #endregion

                        //Obtengo los datos para el item
                        var curItem = PopulateRequisitionItem(currentRow, itemList.Count, false);
                        itemList.Add(curItem);
                        prevReqHeader = curReqHeader;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRow).Value2 = ex.Message;
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(requisitionNoColumn, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }
                } //finaliza el while del proceso completo

                //para el último vale a crear

                #region CreateLastRequisition

                // ReSharper disable once InvertIf
                if (itemList.Count > 0)
                {
                    try
                    {
                        //Crear el encabezado en el MSO140
                        _eFunctions.RevertOperation(opSheet, proxySheet);
                        //ejecutamos el programa
                        var reply = proxySheet.executeScreen(opSheet, "MSO140");
                        //Validamos el ingreso
                        if (reply.mapName != "MSM140A")
                            throw new Exception("ERROR: Se ha producido un error al intentar ingresar al programa. No se puede acceder al MSO140/MSM140A. " + reply.message);

                        //se adicionan los valores a los campos
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("REQ_NO1I", "" + prevReqHeader.IreqNo);
                        arrayFields.Add("TRAN_TYPE1I", "" + prevReqHeader.IssTranType);
                        arrayFields.Add("REQ_BY_DATE1I", "" + prevReqHeader.RequiredByDate);
                        arrayFields.Add("WHOUSE_ID1I", "" + prevReqHeader.OrigWhouseId);
                        arrayFields.Add("PART_ISSUE1I", "Y");
                        arrayFields.Add("PROT_IND1I", "Y");
                        arrayFields.Add("ALLOC_PCA1I", "" + prevReqHeader.AllocPcA);
                        arrayFields.Add("WORK_PROJ_INDA1I", prevReqHeader.WorkProjectIndA);
                        arrayFields.Add("WORK_PROJA1I", prevReqHeader.WorkOrderA ?? prevReqHeader.ProjectA);
                        arrayFields.Add("COST_CENTREA1I", "" + prevReqHeader.CostCentreA);
                        arrayFields.Add("EQUIP_REFA1I", "" + prevReqHeader.EquipmentA);
                        arrayFields.Add("DELIV_INSTRA1I", "" + prevReqHeader.DelivInstrA);
                        arrayFields.Add("DELIV_INSTRB1I", "" + prevReqHeader.DelivInstrB);
                        arrayFields.Add("PRIORITY_CODE1I", "" + prevReqHeader.PriorityCode);
                        arrayFields.Add("ANSWER_B1I", "" + prevReqHeader.AnswerB);
                        arrayFields.Add("ANSWER_D1I", "" + prevReqHeader.AnswerD);
                        arrayFields.Add("ANSWER_B1I", "" + prevReqHeader.AnswerN);
                        arrayFields.Add("ANSWER_D1I", "" + prevReqHeader.AnswerP);
                        arrayFields.Add("REQUESTED_BY1I", "" + prevReqHeader.RequestedBy);

                        //enviar el encabezado MSM140A
                        var request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);
                        //Confirmar el encabezado MSM140A
                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
                        {
                            request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenKey = "1"
                            };
                            reply = proxySheet.submit(opSheet, request);
                        }

                        //no hay errores ni advertencias
                        if (reply == null || _eFunctions.CheckReplyError(reply))
                            throw new Exception("ERROR: " + reply.message);

                        //MSM14BA
                        if (reply.mapName != "MSM14BA")
                            throw new Exception("ERROR: Se ha producido un error al crear el encabezado. No se puede acceder al MSO140/MSM14BA. " + reply.message);

                        var parItemIndex = 0; //controla el par de items por pantalla
                        //agrego los items que tenga hasta el momento al encabezado
                        foreach (var item in itemList)
                        {
                            //asigno número de vale si se genera
                            var screenValues = new ArrayScreenNameValue(reply.screenFields);
                            if (!string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                            if (parItemIndex % 2 == 0)
                                arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("QTY_REQD1I" + (parItemIndex + 1), "" + item.QuantityRequired);
                            arrayFields.Add("UOM1I" + (parItemIndex + 1), item.UnitOfMeasure);
                            arrayFields.Add("TYPE1I" + (parItemIndex + 1), "S");
                            arrayFields.Add("DESCR_A1I" + (parItemIndex + 1), item.StockCode);
                            arrayFields.Add("PART_ISSUE1I" + (parItemIndex + 1), "Y");

                            //envío si es el último item de la lista o si es el segundo de la pantalla
                            if (item == itemList[itemList.Count - 1] || parItemIndex > 0)
                            {
                                request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };
                                reply = proxySheet.submit(opSheet, request);

                                //evalúo si hay error y cancelo
                                if (_eFunctions.CheckReplyError(reply))
                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);
                                //mientras confirmación o bodega
                                while (reply != null && (reply.mapName == "MSM14BA" || reply.mapName == "MSM14CA"))
                                {
                                    //evalúo error
                                    if (_eFunctions.CheckReplyError(reply))
                                        throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);

                                    //asigno número de vale si se genera
                                    screenValues = new ArrayScreenNameValue(reply.screenFields);
                                    if (screenValues.GetField("IREQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                        prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                    //si es una nueva pantalla de items
                                    if (reply.mapName == "MSM14BA" && item != itemList[itemList.Count - 1])
                                        if (screenValues.GetField("DESCR_A1I1") != null && string.IsNullOrWhiteSpace(screenValues.GetField("DESCR_A1I1").value))
                                            break;
                                    ////MSM14CA  - Warehouse Holdings
                                    if (reply.mapName == "MSM14CA")
                                    {
                                        if (screenValues.GetField("TOTAL_REQD1I") == null)
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "No se ha encontrado el valor total requerido al asignar a bodega. MSM14CA");
                                        string selWarehouseIndex = "";
                                        //obtengo solo la lista de pares del objeto para actualizarla
                                        var screenArray = screenValues.ToArray();
                                        foreach (var parValue in screenArray)
                                        {
                                            if (parValue.fieldName != null && parValue.fieldName.StartsWith("WHOUSE_ID_") && parValue.value == prevReqHeader.OrigWhouseId)
                                                selWarehouseIndex = parValue.fieldName.Replace("WHOUSE_ID_", "");
                                            if (parValue.fieldName != null && parValue.fieldName.StartsWith("QTY_REQD_"))
                                                parValue.value = "";
                                        }

                                        //reingreso la lista al objeto del screen y actualizo la cantidad del w/h que quiero de acuerdo a lo realizado anteriormente
                                        if (string.IsNullOrWhiteSpace(selWarehouseIndex))
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "El item no está catalogado en la bodega seleccionada. MSM14CA");
                                        screenValues = new ArrayScreenNameValue(screenArray);
                                        screenValues.SetField("QTY_REQD_" + selWarehouseIndex, "" + screenValues.GetField("TOTAL_REQD1I").value);

                                        //envío el proceso
                                        request = new Screen.ScreenSubmitRequestDTO
                                        {
                                            screenFields = screenValues.ToArray(),
                                            screenKey = "1"
                                        };
                                        reply = proxySheet.submit(opSheet, request);
                                        continue; //continúa con el siguiente while
                                    }

                                    ////Confirm MSM14BA o cualquier otra confirmación que no requiera datos
                                    request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                }
                            }

                            parItemIndex++;
                            if (parItemIndex > 1)
                                parItemIndex = 0;

                            //OK del item
                            _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                            _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                            _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                            _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        }

                        //Confirmo la creación de todos los items. Si no llega a este punto es por algún problema presentado
                        foreach (var item in itemList)
                        {
                            _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                            _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                            _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                            _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        }
                    }
                    catch (Exception ex)
                    {
                        var addMessage = "" + ex.Message;
                        try
                        {
                            if (string.IsNullOrWhiteSpace(prevReqHeader.IreqNo))
                            {
                                addMessage += " .NO SE HA REALIZADO NINGUNA ACCIÓN";
                            }
                            else
                            {
                                //Eliminación del vale por el MSO140
                                _eFunctions.RevertOperation(opSheet, proxySheet);
                                //ejecutamos el programa
                                var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                var screenValues = new ArrayScreenNameValue(reply.screenFields);

                                if (_eFunctions.CheckReplyError(reply))
                                    throw new Exception("" + reply.message);

                                while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" &&
                                       (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")) &&
                                       (screenValues.GetField("REQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("REQ_NO1I").value)))
                                {
                                    var arrayFields = new ArrayScreenNameValue();
                                    arrayFields.Add("COMP_DEL1I", "D");

                                    var request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenFields = arrayFields.ToArray(),
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                    screenValues = new ArrayScreenNameValue(reply.screenFields);

                                    if (_eFunctions.CheckReplyError(reply))
                                        throw new Exception(". ERROR AL ELIMINAR " + prevReqHeader.IreqNo + ": " + reply.message);
                                }

                                addMessage += " - VALE ELIMINADO " + prevReqHeader.IreqNo;
                            }
                        }
                        catch (Exception ex2)
                        {
                            addMessage += ex2;
                        }
                        finally
                        {
                            addMessage = addMessage.Replace("X2:0011 - INPUT REQUIRED  \"C\" TO COMPLETE OR \"D\" TO DELETE", "X2:0011 - EXISTE UNA ORDEN INCOMPLETA EN PROCESO. INGRESE AL MSO 140 PARA COMPLETARLA/ELIMINARLA");
                            foreach (var item in itemList)
                            {
                                _cells.GetCell(4, currentRowHeader + item.Index).Value2 = !string.IsNullOrWhiteSpace(prevReqHeader.IreqNo) ? prevReqHeader.IreqNo + " - ELIMINADO" : "";
                                _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = addMessage;
                                _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                            }
                        }
                    }
                }

                #endregion

            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true); //Se restaura por 'Pooled Connection Request Timed Out'
            }
        }

        /// <summary>
        /// Borra el header de un vale cuando este no se puede finalizar
        /// </summary>
        /// <param name="proxyRequisition"></param>
        /// <param name="createHeaderReply"></param>
        /// <param name="opRequisition"></param>
        private static void DeleteHeader(RequisitionService.RequisitionService proxyRequisition, RequisitionService.RequisitionServiceCreateHeaderReplyDTO createHeaderReply, RequisitionService.OperationContext opRequisition)
        {
            if (createHeaderReply == null)
                return;
            //new RequisitionService.RequisitionServiceDeleteHeaderReplyDTO();
            var deleteHeaderRequest = CreateDeleteRequestDto(createHeaderReply);

            proxyRequisition.deleteHeader(opRequisition, deleteHeaderRequest);
        }

        /// <summary>
        /// Esta funcion copia el encabezado de la creacion de la requisicion en el objeto del encabezado para el borrado
        /// </summary>
        /// <param name="createHeaderReply">Encabezado de la requisicion a borrar</param>
        /// <returns></returns>
        private static RequisitionService.RequisitionServiceDeleteHeaderRequestDTO CreateDeleteRequestDto(RequisitionService.RequisitionServiceCreateHeaderReplyDTO createHeaderReply)
        {
            var deleteHeaderRequest = new RequisitionService.RequisitionServiceDeleteHeaderRequestDTO
            {
                allocPcA = createHeaderReply.allocPcA,
                allocPcB = createHeaderReply.allocPcB,
                assignToTeam = createHeaderReply.assignToTeam,
                authorisedStatusDesc = createHeaderReply.authorisedStatusDesc,
                //authsdBy = createHeaderReply.authsdBy,//Removidos en E9
                //authsdByName = createHeaderReply.authsdByName,//Removidos en E9
                authsdDate = createHeaderReply.authsdDate,
                authsdItmAmt = createHeaderReply.authsdItmAmt,
                //authsdPosition = createHeaderReply.authsdPosition,//Removidos en E9
                //authsdPositionDesc = createHeaderReply.authsdPositionDesc,//Removidos en E9
                authsdStatus = createHeaderReply.authsdStatus,
                authsdStatusDesc = createHeaderReply.authsdStatusDesc,
                authsdTime = createHeaderReply.authsdTime,
                authsdTotAmt = createHeaderReply.authsdTotAmt,
                completedDate = createHeaderReply.completedDate,
                completeItems = createHeaderReply.completeItems,
                confirmDelete = createHeaderReply.confirmDelete,
                costCentreA = createHeaderReply.costCentreA,
                costCentreB = createHeaderReply.costCentreB,
                costDistrictA = createHeaderReply.costDistrictA,
                costDistrictB = createHeaderReply.costDistrictB,
                createdBy = createHeaderReply.createdBy,
                createdByName = createHeaderReply.createdByName,
                creationDate = createHeaderReply.creationDate,
                creationTime = createHeaderReply.creationTime,
                custNo = createHeaderReply.custNo,
                custNoDesc = createHeaderReply.custNoDesc,
                delivInstrA = createHeaderReply.delivInstrA,
                delivInstrB = createHeaderReply.delivInstrB,
                delivLocation = createHeaderReply.delivLocation,
                delivLocationDesc = createHeaderReply.delivLocationDesc,
                directPurchOrd = createHeaderReply.directPurchOrd,
                districtCode = createHeaderReply.districtCode,
                districtName = createHeaderReply.districtName,
                entitlementPeriod = createHeaderReply.entitlementPeriod,
                equipmentA = createHeaderReply.equipmentA,
                equipmentB = createHeaderReply.equipmentB,
                equipmentRefA = createHeaderReply.equipmentRefA,
                equipmentRefB = createHeaderReply.equipmentRefB,
                groupClass = createHeaderReply.groupClass,
                hdr140Status = createHeaderReply.hdr140Status,
                hdr140StatusDesc = createHeaderReply.hdr140StatusDesc,
                headerType = createHeaderReply.headerType,
                inabilityDate = createHeaderReply.inabilityDate,
                inabilityRsn = createHeaderReply.inabilityRsn,
                inspectCode = createHeaderReply.inspectCode,
                inventCat = createHeaderReply.inventCat,
                inventCatDesc = createHeaderReply.inventCatDesc,
                ireqNo = createHeaderReply.ireqNo,
                ireqType = createHeaderReply.ireqType,
                issTranType = createHeaderReply.issTranType,
                issTranTypeDesc = createHeaderReply.issTranTypeDesc,
                issueRequisitionTypeDesc = createHeaderReply.issueRequisitionTypeDesc,
                lastAcqDate = createHeaderReply.lastAcqDate,
                loanDuration = createHeaderReply.loanDuration,
                loanRequisitionNo = createHeaderReply.loanRequisitionNo,
                lstAmodDate = createHeaderReply.lstAmodDate,
                lstAmodTime = createHeaderReply.lstAmodTime,
                lstAmodUser = createHeaderReply.lstAmodUser,
                matGroupCode = createHeaderReply.matGroupCode,
                matGroupCodeDesc = createHeaderReply.matGroupCodeDesc,
                moreInstr = createHeaderReply.moreInstr,
                msg140Data = createHeaderReply.msg140Data,
                narrative = createHeaderReply.narrative,
                numOfItems = createHeaderReply.numOfItems,
                orderStatusDesc = createHeaderReply.orderStatusDesc,
                origWhouseId = createHeaderReply.origWhouseId,
                origWhouseIdDesc = createHeaderReply.origWhouseIdDesc,
                partIssue = createHeaderReply.partIssue,
                partIssueSpecified = createHeaderReply.partIssueSpecified,
                password = createHeaderReply.password,
                preqNo = createHeaderReply.preqNo,
                priorityCode = createHeaderReply.priorityCode,
                projectA = createHeaderReply.projectA,
                projectB = createHeaderReply.projectB,
                protectedInd = createHeaderReply.protectedInd,
                purchaseOrdNo = createHeaderReply.purchaseOrdNo,
                purchDelivInstr = createHeaderReply.purchDelivInstr,
                purchInstr = createHeaderReply.purchInstr,
                purchInstruction = createHeaderReply.purchInstruction,
                purchOfficer = createHeaderReply.purchOfficer,
                rcvngWhouse = createHeaderReply.rcvngWhouse,
                rcvngWhouseDesc = createHeaderReply.rcvngWhouseDesc,
                relatedWhReq = createHeaderReply.relatedWhReq,
                repairRequest = createHeaderReply.repairRequest,
                requestedBy = createHeaderReply.requestedBy,
                requestedByName = createHeaderReply.requestedByName,
                requiredByDate = createHeaderReply.requiredByDate,
                requiredByPos = createHeaderReply.requiredByPos,
                requiredByPosDesc = createHeaderReply.requiredByPosDesc,
                requisitionItemStatusDesc = createHeaderReply.requisitionItemStatusDesc,
                reversePeriodStart = createHeaderReply.reversePeriodStart,
                rotnRequisitionNo = createHeaderReply.rotnRequisitionNo,
                sentType = createHeaderReply.sentType,
                sentTypeDesc = createHeaderReply.sentTypeDesc,
                statsUpdatedInd = createHeaderReply.statsUpdatedInd,
                suggestedSupp = createHeaderReply.suggestedSupp,
                surveyNo = createHeaderReply.surveyNo,
                transType = createHeaderReply.transType,
                useByDate = createHeaderReply.useByDate,
                workOrderA = createHeaderReply.workOrderA,
                workOrderB = createHeaderReply.workOrderB,
                workProjA = createHeaderReply.workProjA,
                workProjB = createHeaderReply.workProjB,
                workProjIndA = createHeaderReply.workProjIndA,
                workProjIndB = createHeaderReply.workProjIndB
            };


            return deleteHeaderRequest;
        }

        private void btnCleanSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
        }

        private void ManualCreditRequisition()
        {

            //instancia del Servicio
            var issueRequisitionItemStocklessService = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();

            var currentRow = TitleRow01 + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01);



                //Header
                var opRequisition = new OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                issueRequisitionItemStocklessService.Url = urlService + "/IssueRequisitionItemStocklessService";

                opRequisition.district = _frmAuth.EllipseDstrct;
                opRequisition.maxInstances = 100;
                opRequisition.position = _frmAuth.EllipsePost;
                opRequisition.returnWarnings = false;

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var headerCreateReturnReply = new ImmediateReturnStocklessDTO();


                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value) != null ||
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
                {

                    try
                    {
                        string switchCase = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                        var reference = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                        switch (switchCase)
                        {
                            case "Work Order":
                                headerCreateReturnReply.workOrderx1 = reference;
                                break;
                            case "Equipment No.":
                                headerCreateReturnReply.equipmentReferencex1 = reference;
                                break;
                            case "Project No.":
                                headerCreateReturnReply.projectNumberx1 = reference;
                                break;
                            case "Account Code":
                                headerCreateReturnReply.costCodex1 = reference;
                                break;
                        }

                        headerCreateReturnReply.districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct)
                            ? "ICOR"
                            : _frmAuth.EllipseDstrct;
                        headerCreateReturnReply.processedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.percentageAllocatedx1 = 100;
                        headerCreateReturnReply.requestedByEmployee =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.requestedByPositionId = _frmAuth.EllipsePost;
                        headerCreateReturnReply.warehouseId =
                            _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value);
                        headerCreateReturnReply.authorisedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.transactionType =
                            MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                        headerCreateReturnReply.requisitionNumber =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                        headerCreateReturnReply.processedDate =
                            DateTime.ParseExact(_cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value), "yyyyMMdd", CultureInfo.InvariantCulture);
                        headerCreateReturnReply.processedDateSpecified = true;


                        var holding = new HoldingDetailsDTO();
                        var listHolding = new List<HoldingDetailsDTO>();

                        holding.quantitySpecified = true;
                        holding.quantity = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);
                        holding.stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value);

                        listHolding.Add(holding);
                        headerCreateReturnReply.holdingDetailsDTO = listHolding.ToArray();

                        var result = issueRequisitionItemStocklessService.immediateReturn(opRequisition, headerCreateReturnReply);

                        if (result.errors.Length == 0)
                        {
                            _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Success;
                            _cells.GetCell(ResultColumn01, currentRow).Value2 = "OK";
                        }
                        else
                        {
                            _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                            foreach (var e in result.errors)
                            {
                                _cells.GetCell(ResultColumn01, currentRow).Value2 += " " + e.messageText;
                            }
                        }


                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Value2 += "ERROR: " + ex.Message;
                        _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }

                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true); //Se restaura por 'Pooled Connection Request Timed Out'
                issueRequisitionItemStocklessService?.Dispose();
            }
        }

        private void ManualCreditRequisitionExtended()
        {
            //instancia del Servicio
            var issueRequisitionItemStocklessService = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();

            var currentRow = TitleRow01Ext + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01Ext);

                

                //Header
                var opRequisition = new OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                issueRequisitionItemStocklessService.Url = urlService + "/IssueRequisitionItemStocklessService";
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                opRequisition.district = _frmAuth.EllipseDstrct;
                opRequisition.maxInstances = 100;
                opRequisition.position = _frmAuth.EllipsePost;
                opRequisition.returnWarnings = false;

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var headerCreateReturnReply = new ImmediateReturnStocklessDTO();


                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value) != null ||
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
                {

                    try
                    {
                        headerCreateReturnReply.workOrderx1 = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                        headerCreateReturnReply.equipmentReferencex1 = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                        headerCreateReturnReply.projectNumberx1 = _cells.GetNullOrTrimmedValue(_cells.GetCell(12, currentRow).Value);
                        headerCreateReturnReply.costCodex1 = _cells.GetNullOrTrimmedValue(_cells.GetCell(13, currentRow).Value);

                        headerCreateReturnReply.districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDstrct)
                            ? "ICOR"
                            : _frmAuth.EllipseDstrct;
                        headerCreateReturnReply.processedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.percentageAllocatedx1 = 100;
                        headerCreateReturnReply.requestedByEmployee =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.requestedByPositionId = _frmAuth.EllipsePost;
                        headerCreateReturnReply.warehouseId =
                            _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value);
                        headerCreateReturnReply.authorisedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.transactionType =
                            MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                        headerCreateReturnReply.requisitionNumber =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                        headerCreateReturnReply.processedDate =
                            DateTime.ParseExact(_cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value), "yyyyMMdd", CultureInfo.InvariantCulture);
                        headerCreateReturnReply.processedDateSpecified = true;


                        var holding = new HoldingDetailsDTO();
                        var listHolding = new List<HoldingDetailsDTO>();

                        holding.quantitySpecified = true;
                        holding.quantity = Convert.ToDecimal(_cells.GetCell(20, currentRow).Value);
                        holding.stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(18, currentRow).Value);

                        listHolding.Add(holding);
                        headerCreateReturnReply.holdingDetailsDTO = listHolding.ToArray();

                        var result = issueRequisitionItemStocklessService.immediateReturn(opRequisition, headerCreateReturnReply);

                        if (result.errors.Length == 0)
                        {
                            _cells.GetCell(ResultColumn01Ext, currentRow).Style = StyleConstants.Success;
                            _cells.GetCell(ResultColumn01Ext, currentRow).Value2 = "OK";
                        }
                        else
                        {
                            _cells.GetCell(ResultColumn01Ext, currentRow).Style = StyleConstants.Error;
                            foreach (var e in result.errors)
                            {
                                _cells.GetCell(ResultColumn01Ext, currentRow).Value2 += " " + e.messageText;
                            }
                        }


                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn01Ext, currentRow).Value2 += "ERROR: " + ex.Message;
                        _cells.GetCell(ResultColumn01Ext, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }

                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true); //Se restaura por 'Pooled Connection Request Timed Out'
                issueRequisitionItemStocklessService?.Dispose();
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }



        private void ReviewRequisitionList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName02);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A5").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            try
            {
                var sqlQuery = Queries.GetRequisitionListQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
                var drRequisitions = _eFunctions.GetQueryResult(sqlQuery);

                if (drRequisitions == null || drRequisitions.IsClosed) return;

                var i = TitleRow02 + 1;

                while (drRequisitions != null && drRequisitions.Read())
                {
                    try
                    {
                        _cells.GetCell(1, i).Value = "" + drRequisitions["DSTRCT_CODE"].ToString().Trim();
                        _cells.GetCell(2, i).Value = "" + drRequisitions["WORK_GROUP"].ToString().Trim();
                        _cells.GetCell(3, i).Value = "" + drRequisitions["EQUIP_NO"].ToString().Trim();
                        _cells.GetCell(4, i).Value = "" + drRequisitions["WORK_ORDER"].ToString().Trim();
                        _cells.GetCell(5, i).Value = "" + drRequisitions["WO_DESC"].ToString().Trim();
                        _cells.GetCell(6, i).Value = "" + drRequisitions["PROJECT_NO"].ToString().Trim();
                        _cells.GetCell(7, i).Value = "" + drRequisitions["GL_ACCOUNT"].ToString().Trim();
                        _cells.GetCell(8, i).Value = "" + drRequisitions["IREQ_NO"].ToString().Trim();
                        _cells.GetCell(9, i).Value = "" + drRequisitions["NUM_OF_ITEMS"].ToString().Trim();
                        _cells.GetCell(10, i).Value = "" + drRequisitions["WO_RAISED_DATE"].ToString().Trim();
                        _cells.GetCell(11, i).Value = "" + drRequisitions["WO_PLAN_STR_DATE"].ToString().Trim();
                        _cells.GetCell(12, i).Value = "" + drRequisitions["CREATION_DATE"].ToString().Trim();
                        _cells.GetCell(13, i).Value = "" + drRequisitions["REQ_BY_DATE"].ToString().Trim();
                        _cells.GetCell(14, i).Value = "" + drRequisitions["AUTHSD_DATE"].ToString().Trim();
                        _cells.GetCell(15, i).Value = "" + drRequisitions["CREATED_BY"].ToString().Trim();
                        _cells.GetCell(16, i).Value = "" + drRequisitions["REQUESTED_BY"].ToString().Trim();
                        _cells.GetCell(17, i).Value = "" + drRequisitions["REQ_BY_POS"].ToString().Trim();
                        _cells.GetCell(18, i).Value = "" + drRequisitions["AUTHSD_BY"].ToString().Trim();
                        _cells.GetCell(19, i).Value = "" + drRequisitions["AUTHSD_POSITION"].ToString().Trim();
                        _cells.GetCell(20, i).Value = "" + drRequisitions["REQ_STATUS"].ToString().Trim();
                        _cells.GetCell(21, i).Value = "" + drRequisitions["AUTHSD_STATUS"].ToString().Trim();
                        _cells.GetCell(22, i).Value = "" + drRequisitions["IREQ_TYPE"].ToString().Trim();
                        _cells.GetCell(23, i).Value = "" + drRequisitions["ISS_TRAN_TYPE"].ToString().Trim();
                        _cells.GetCell(24, i).Value = "" + drRequisitions["ORIG_WHOUSE_ID"].ToString().Trim();
                        _cells.GetCell(25, i).Value = "" + drRequisitions["PRIORITY_CODE"].ToString().Trim();
                        _cells.GetCell(26, i).Value = "" + drRequisitions["EQUIP_GRP_ID"].ToString().Trim();
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionList()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionList(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void ReviewRequisitionControlList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var cp = new ExcelStyleCells(_excelApp, SheetName02); //cells parameters
            var cr = new ExcelStyleCells(_excelApp, SheetName03); //cells results

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = cp.GetEmptyIfNull(cp.GetCell("B3").Value);
            var searchCriteriaKey1Text = cp.GetEmptyIfNull(cp.GetCell("A4").Value);
            var searchCriteriaValue1 = cp.GetEmptyIfNull(cp.GetCell("B4").Value);
            var searchCriteriaKey2Text = cp.GetEmptyIfNull(cp.GetCell("A5").Value);
            var searchCriteriaValue2 = cp.GetEmptyIfNull(cp.GetCell("B5").Value);
            var statusKey = cp.GetEmptyIfNull(cp.GetCell("B6").Value);
            var dateCriteriaKeyText = cp.GetEmptyIfNull(cp.GetCell("D3").Value);
            var startDate = cp.GetEmptyIfNull(cp.GetCell("D4").Value);
            var endDate = cp.GetEmptyIfNull(cp.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;
            try
            {
                //Elimino los registros anteriores
                _excelApp.ActiveWorkbook.Sheets[3].Select(Type.Missing);
                cr.ClearTableRange(TableName03);
                cr.DeleteTableRange(TableName03);

                var sqlQuery = Queries.GetRequisitionControlListQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
                var drRequisitions = _eFunctions.GetQueryResult(sqlQuery);

                if (drRequisitions == null || drRequisitions.IsClosed) return;

                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < drRequisitions.FieldCount; i++)
                    cr.GetCell(i + 1, TitleRow03).Value2 = "'" + drRequisitions.GetName(i);

                _cells.FormatAsTable(cr.GetRange(1, TitleRow03, drRequisitions.FieldCount, TitleRow03 + 1), TableName03);

                //cargo los datos 
                var currentRow = TitleRow03 + 1;

                while (drRequisitions.Read())
                {
                    try
                    {
                        for (var i = 0; i < drRequisitions.FieldCount; i++)
                            cr.GetCell(i + 1, currentRow).Value2 = "'" + drRequisitions[i].ToString().Trim();
                    }
                    catch (Exception ex)
                    {
                        cr.GetCell(1, currentRow).Style = StyleConstants.Error;
                        cr.GetCell(1, currentRow).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionControlList()", ex.Message);
                    }
                    finally
                    {
                        cr.GetCell(1, currentRow).Select();
                        currentRow++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRequisitionControlList(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void ReReviewRequisitionControlList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            try
            {
                var cp = new ExcelStyleCells(_excelApp, SheetName02); //cells parameters
                var cr = new ExcelStyleCells(_excelApp, SheetName03); //cells results
                //Elimino los registros anteriores
                _excelApp.ActiveWorkbook.Sheets[3].Select(Type.Missing);
                cr.ClearTableRange(TableName03);
                cr.DeleteTableRange(TableName03);

                cp.SetFixedWorkingWorkSheet(false);

                var currentParam = TitleRow02 + 1; //itera según cada estándar
                var currentRow = TitleRow03 + 1; //itera la celda para cada tarea

                //criterios generales
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();
                var dateCriteriaKeyText = cp.GetEmptyIfNull(cp.GetCell("D3").Value);
                var startDate = cp.GetEmptyIfNull(cp.GetCell("D4").Value);
                var endDate = cp.GetEmptyIfNull(cp.GetCell("D5").Value);
                var statusKey = cp.GetEmptyIfNull(cp.GetCell("B6").Value);

                //Convierto los nombres de las opciones a llaves


                //mientras haya un registro con orden o con número de requisición
                while (!string.IsNullOrEmpty("" + cp.GetCell(4, currentParam).Value) || !string.IsNullOrEmpty("" + cp.GetCell(8, currentParam).Value))
                {
                    try
                    {
                        var district = cp.GetNullIfTrimmedEmpty(cp.GetCell(1, currentParam).Value2) ?? cp.GetNullIfTrimmedEmpty(cp.GetCell("B3").Value) ?? "ICOR";
                        var equipment = cp.GetEmptyIfNull(cp.GetCell(3, currentParam).Value2);
                        var workOrder = cp.GetEmptyIfNull(cp.GetCell(4, currentParam).Value2);
                        var reqNo = cp.GetEmptyIfNull(cp.GetCell(8, currentParam).Value2);
                        int searchKey1;
                        var searchValue1 = "";
                        var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;
                        var searchKey2 = SearchFieldCriteriaType.None.Key;
                        var searchValue2 = "";

                        if (!string.IsNullOrWhiteSpace(reqNo))
                        {
                            searchKey1 = SearchFieldCriteriaType.Requisition.Key;
                            searchValue1 = reqNo;
                            dateCriteriaKey = SearchDateCriteriaType.IgnoreDate.Key;
                        }
                        else if (!string.IsNullOrWhiteSpace(workOrder))
                        {
                            searchKey1 = SearchFieldCriteriaType.WorkOrder.Key;
                            searchValue1 = workOrder;
                            dateCriteriaKey = SearchDateCriteriaType.IgnoreDate.Key;
                        }
                        else if (!string.IsNullOrWhiteSpace(equipment))
                        {
                            searchKey1 = SearchFieldCriteriaType.EquipmentReference.Key;
                            searchValue1 = equipment;
                        }
                        else
                        {
                            searchKey1 = SearchFieldCriteriaType.WorkOrder.Key;
                            searchValue1 = workOrder;
                        }

                        var sqlQuery = Queries.GetRequisitionControlListQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchKey1, searchValue1, searchKey2, searchValue2, dateCriteriaKey, startDate, endDate, statusKey);
                        var drRequisitions = _eFunctions.GetQueryResult(sqlQuery);

                        if (drRequisitions == null || drRequisitions.IsClosed || !drRequisitions.HasRows) return;

                        //Cargo el encabezado de la tabla y doy formato
                        if (currentRow == TitleRow03 + 1)
                        {
                            for (var i = 0; i < drRequisitions.FieldCount; i++)
                                cr.GetCell(i + 1, TitleRow03).Value2 = "'" + drRequisitions.GetName(i);
                            cr.FormatAsTable(cr.GetRange(1, TitleRow03, drRequisitions.FieldCount, TitleRow03 + 1), TableName03);
                        }

                        

                        //cargo los datos 
                        while (drRequisitions.Read())
                        {
                            try
                            {
                                for (var i = 0; i < drRequisitions.FieldCount; i++)
                                    cr.GetCell(i + 1, currentRow).Value2 = "'" + drRequisitions[i].ToString().Trim();
                            }
                            catch (Exception ex)
                            {
                                cr.GetCell(1, currentRow).Style = StyleConstants.Error;
                                cr.GetCell(1, currentRow).Value = "ERROR: " + ex.Message;
                                Debugger.LogError("RibbonEllipse.cs:ReReviewRequisitionControlList(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                            }
                            finally
                            {
                                cr.GetCell(1, currentRow).Select();
                                currentRow++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse.cs:ReReviewRequisitionControlList(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                        cp.GetCell(2, currentParam).Style = StyleConstants.Error;
                        cp.GetCell(ResultColumn02, currentParam).Style = StyleConstants.Error;
                        cp.GetCell(ResultColumn02, currentParam).Value = "ERROR: " + ex.Message;
                    }
                    finally
                    {
                        currentParam++;
                    }
                }



            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReReviewRequisitionControlList(ignoreError = false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnConfigAssuranceSettings_Click(object sender, RibbonControlEventArgs e)
        {
            new AssuranceSettingsBox().ShowDialog();
        }

        private void cbMaxItems_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("maxItemValue", cbMaxItems.Checked.ToString());
            Settings.CurrentSettings.SaveCustomSettings();

        }

        private void cbSortItems_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("autoSortElements", cbSortItems.Checked.ToString());
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbAssuranceProcess_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("assuranceProcessTrack", cbAssuranceProcess.Checked.ToString());
            Settings.CurrentSettings.SaveCustomSettings();
        }
    }
}

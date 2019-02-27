using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseRequisitionServiceExcelAddIn.IssueRequisitionItemStocklessService;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseRequisitionServiceExcelAddIn
{
    public partial class RibbonEllipse
    {
        Excel.Application _excelApp;
        ExcelStyleCells _cells;
        readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        readonly FormAuthenticate _frmAuth = new FormAuthenticate();

        private const int TitleRow01 = 5;
        private const int ResultColumn01 = 19;

        private const int TitleRow01Ext = 5;
        private const int ResultColumn01Ext = 21;

        private const string SheetName01 = "RequisitionService";
        private const string TableName01 = "RequisitionServiceTable";
        private const string ValidationSheet = "ValidationRequisition";

        private bool _ignoreItemError;

        private Thread _thread;
        public List<RequisitionClassLibrary.SpecialRestriction.SpecialRestrictionItem> RestrictionList;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environmentList = Environments.GetEnviromentList();
            foreach (var item in environmentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpEnviroment.Items.Add(drpItem);
            }
        }

        private void btnFormatNewSheet_Click(object sender, RibbonControlEventArgs e)
        {
            RequisitionServiceFormat();
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
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                        
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(CreateRequisitionServiceExtended);
                    else
                        _thread = new Thread(CreateRequisitionService);
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
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 + "Ext")
                        _thread = new Thread(CreateRequisitionServiceExtended);
                    else
                        _thread = new Thread(CreateRequisitionService);

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
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if(_cells != null) _cells.SetCursorDefault();
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
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheet);
                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01;

                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;

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
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), optionReqTypeList, ValidationSheet, 1, false);
                _cells.GetCell(6, titleRow).Value = "Transaction Type";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), optionTransTypeList, ValidationSheet, 2, false);
                _cells.GetCell(7, titleRow).Value = "Required By Date";
                _cells.GetCell(7, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "Original Warehouse";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "Priority Code";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var itemList = _eFunctions.GetItemCodes("PI");
                var optionPriorList = MyUtilities.GetCodeList(itemList);
                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), optionPriorList, ValidationSheet, 3, false);

                _cells.GetCell(10, titleRow).Value = "Reference Type";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionRefTypeList = new List<string> {"Work Order", "Equipment No.", "Project No.", "Account Code"};
                _cells.SetValidationList(_cells.GetCell(10, titleRow + 1), optionRefTypeList, ValidationSheet, 4);

                _cells.GetCell(11, titleRow).Value = "Reference";
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(12, titleRow).Value = "Delivery Instructions"; //120 caracteres (60/60)
                _cells.GetCell(12, titleRow).AddComment("120 caracteres");
                _cells.GetCell(12, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(13, titleRow).Value = "Return Cause";
                _cells.GetCell(13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(14, titleRow).Value = "Issue Question";
                _cells.GetCell(14, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                var optionIssueList = new List<string> {"A - VENTAS", "B - RUBROS"};
                _cells.SetValidationList(_cells.GetCell(14, titleRow + 1), optionIssueList, ValidationSheet, 5, false);

                _cells.GetCell(15, titleRow).Value = "Partial Allowed";
                _cells.GetCell(15, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var partialAllowedList = new List<string> {"Y - YES", "N - No"};
                _cells.SetValidationList(_cells.GetCell(15, titleRow + 1), partialAllowedList, ValidationSheet, 5, false);

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
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheet);
                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01 + "Ext";

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
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), optionReqTypeList, ValidationSheet, 1, false);
                _cells.GetCell(6, titleRow).Value = "Transaction Type";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), optionTransTypeList, ValidationSheet, 2, false);
                _cells.GetCell(7, titleRow).Value = "Required By Date";
                _cells.GetCell(7, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "Original Warehouse";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "Priority Code";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var itemList = _eFunctions.GetItemCodes("PI");
                var optionPriorList = MyUtilities.GetCodeList(itemList);
                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), optionPriorList, ValidationSheet, 3, false);

                _cells.GetCell(10, titleRow).Value = "Work Order";
                _cells.GetCell(11, titleRow).Value = "Equipment No.";
                _cells.GetCell(12, titleRow).Value = "Project No.";
                _cells.GetCell(13, titleRow).Value = "Account Code";
                _cells.GetRange(10, titleRow, 13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(14, titleRow).Value = "Delivery Instructions"; //120 caracteres (60/60)
                _cells.GetCell(14, titleRow).AddComment("120 caracteres");
                _cells.GetCell(14, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(15, titleRow).Value = "Return Cause";
                _cells.GetCell(15, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(16, titleRow).Value = "Issue Question";
                _cells.GetCell(16, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                var optionIssueList = new List<string> { "A - VENTAS", "B - RUBROS" };
                _cells.SetValidationList(_cells.GetCell(16, titleRow + 1), optionIssueList, ValidationSheet, 4, false);

                _cells.GetCell(17, titleRow).Value = "Partial Allowed";
                _cells.GetCell(17, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var partialAllowedList = new List<string> { "Y - YES", "N - No" };
                _cells.SetValidationList(_cells.GetCell(17, titleRow + 1), partialAllowedList, ValidationSheet, 5, false);

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

                ((Excel.Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
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

        public RequisitionClassLibrary.RequisitionHeader PopulateRequisitionHeader(int currentRow, bool isExtended)
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
            bool partIssueIndicator;
            bool protectedIndicator;
            string deliveryInstructionsA;
            string deliveryInstructionsB;
            string answerB;
            string answerD;
            string workOrderAllocation;
            string workProjectIndicatorAllocation;
            string equipmentAllocation;
            string projectAllocation;
            string costCentreAllocation;
            string requisitionNumber; 

            if (isExtended)
            {
                allocPcA = "100";
                districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct;
                costDistrictAllocation = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct;
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
                deliveryInstructionsB = deliveryInstructionsA.Length > 80 ? deliveryInstructionsA.Substring(80) : null;
                answerB = _cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value).Length >= 2
                    ? _cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value).Substring(0, 2).Trim() : null;
                answerD = _cells.GetEmptyIfNull(_cells.GetCell(16, currentRow).Value).Length >= 2
                    ? _cells.GetEmptyIfNull(_cells.GetCell(16, currentRow).Value).Substring(0, 2).Trim() : null;

                partIssueIndicator = true;//general indicator
                protectedIndicator = false;
            }
            else
            {
                allocPcA = "100";
                districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct;
                costDistrictAllocation = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct;
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
                deliveryInstructionsB = deliveryInstructionsA.Length > 80 ? deliveryInstructionsA.Substring(80) : null;
                answerB = _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Length >= 2
                    ? _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Substring(0, 2).Trim() : null;
                answerD = _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Length >= 2
                    ? _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Substring(0, 2).Trim() : null;

                partIssueIndicator = true;//general indicator
                protectedIndicator = false;

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


            var requisitionHeader = new RequisitionClassLibrary.RequisitionHeader
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

        public RequisitionClassLibrary.RequisitionItem PopulateRequisitionItem(int currentRow, int indexList, bool isExtended)
        {
            bool partialAllowed;
            string stockCode;
            string unitOfMeasure;
            decimal quantityRequired;

            if (isExtended)
            {
                partialAllowed = MyUtilities.IsTrue(MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value)), true);
                stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(18, currentRow).Value);
                unitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(19, currentRow).Value);
                quantityRequired = Convert.ToDecimal(_cells.GetCell(20, currentRow).Value);
            }
            else
            {
                partialAllowed = MyUtilities.IsTrue(MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)), true);
                stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value);
                unitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(17, currentRow).Value);
                quantityRequired = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);
            }


            var item = new RequisitionClassLibrary.RequisitionItem
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

            _eFunctions.CloseConnection();

            return item;
        }

        private void CreateRequisitionServiceExtended()
        {
            CreateRequisitionService(true);
        }

        private void CreateRequisitionService()
        {
            CreateRequisitionService(false);
        }
        /// <summary>
        /// Recorre y Crea los vales de a tabla de Excel
        /// </summary>
        private void CreateRequisitionService(bool isExtended)
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

                //instancia del Servicio
                var proxyRequisition = new RequisitionService.RequisitionService();

                //Header
                var opRequisition = new RequisitionService.OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false);//Se asigna por 'Pooled Connection Request Timed Out'
                proxyRequisition.Url = urlService + "/RequisitionService";

                opRequisition.district = _frmAuth.EllipseDsct;
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
                RestrictionList = RequisitionClassLibrary.SpecialRestriction.GetPositionRestrictions(_eFunctions);

                var itemList = new List<RequisitionClassLibrary.RequisitionItem>();
                RequisitionService.RequisitionServiceCreateHeaderReplyDTO headerCreateReply = null;

                RequisitionClassLibrary.RequisitionHeader prevReqHeader = null;
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
                            headerCreateReply = proxyRequisition.createHeader(opRequisition, headerCreateRequest);
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
                                    proxyRequisition.createItem(opRequisition, itemRequest);
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
                                    DeleteHeader(proxyRequisition, headerCreateReply, opRequisition);
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
                                    proxyRequisition.finalise(opRequisition, finaliseRequest);
                                }
                                catch (TimeoutException ex)
                                {
                                    _cells.GetCell(resultColumn, currentRow - 1).Value2 = _cells.GetCell(resultColumn, currentRow - 1).Value2 + " " + ex.Message;
                                    _cells.GetCell(resultColumn, currentRow - 1).Style = StyleConstants.Warning;
                                    _cells.GetCell(requisitionNoColumn, currentRow - 1).Style = StyleConstants.Warning;
                                }
                            }

                            //creo el nuevo encabezado y reinicio variables
                            prevReqHeader = null;//no es una línea inservible. Es necesaria por si se produce una excepción al momento de creación de un nuevo encabezado
                            currentRowHeader = currentRow;
                            abortRequisition = false;
                            itemList = new List<RequisitionClassLibrary.RequisitionItem>();
                            var headerCreateRequest = curReqHeader.GetCreateHeaderRequest();
                            headerCreateReply = proxyRequisition.createHeader(opRequisition, headerCreateRequest);
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
                            curItem.StockCode = "";//Se vacía el campo para conservar la estructura del vale, pero para que indique el error
                        }

                        if (curItem.DirectOrderIndicator)
                        {
                            abortRequisition = true;

                            _cells.GetCell(resultColumn, currentRow).Value2 += curItem.StockCode + " ITEM DE ORDEN DIRECTA. DEBE CREAR EL VALE CON OTRO MÉTODO";
                            _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                            curItem.StockCode = "";//Se vacía el campo para conservar la estructura del vale, pero para que indique el error
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
                        proxyRequisition.createItem(opRequisition, itemRequest);

                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 = "OK";
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Value2 = itemRequest.ireqNo;
                        _cells.GetCell(requisitionNoColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRowHeader + item.Index).Value2 +=  "ERROR: " + ex.Message;
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
                        DeleteHeader(proxyRequisition, headerCreateReply, opRequisition);
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
                        proxyRequisition.finalise(opRequisition, finaliseRequest);
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
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
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
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false);//Se asigna por 'Pooled Connection Request Timed Out'
                
                //ScreenService Opción en reemplazo de los servicios
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
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

                var itemList = new List<RequisitionClassLibrary.RequisitionItem>();

                const int seriesIndicatorColumn = 3;
                var itemIndicatorColumn = isExtended ? 18 : 16;
                const int requisitionNoColumn = 4;

                RequisitionClassLibrary.RequisitionHeader prevReqHeader = null;
                RequisitionClassLibrary.RequisitionHeader curReqHeader;
                RestrictionList = RequisitionClassLibrary.SpecialRestriction.GetPositionRestrictions(_eFunctions);
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

                                    if (parItemIndex%2 == 0)
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
                                itemList = new List<RequisitionClassLibrary.RequisitionItem>();
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
                if (itemList.Count>0)
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
                        catch(Exception ex2)
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
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
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

        ///// <summary>
        ///// Borra el header de un vale cuando este no se puede finalizar usando el MSO140.
        ///// </summary>
        ///// <param name="position"></param>
        ///// <param name="requisitionHeader"></param>
        ///// <param name="urlService"></param>
        ///// <param name="district"></param>
        //private static void DeleteHeader(string urlService, string district, string position, RequisitionHeader requisitionHeader)
        //{
        //    if (requisitionHeader == null)
        //        return;
        //    //instancia del Servicio
        //    var proxyRequisition = new RequisitionService.RequisitionService();

        //    //Header
        //    var opRequisition = new RequisitionService.OperationContext();

        //    proxyRequisition.Url = urlService + "/RequisitionService";
        //    //El client conversation se ejecutó previamente en el proceso que hace llamado a este método
        //    //ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
        //    opRequisition.district = district;
        //    opRequisition.maxInstances = 100;
        //    opRequisition.position = position;
        //    opRequisition.returnWarnings = false;


        //    //new RequisitionService.RequisitionServiceDeleteHeaderReplyDTO();
        //    var deleteHeaderRequest = CreateDeleteRequestDto(requisitionHeader.GetCreateReplyHeader());

        //    proxyRequisition.deleteHeader(opRequisition, deleteHeaderRequest);
        //}

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
                authsdBy = createHeaderReply.authsdBy,
                authsdByName = createHeaderReply.authsdByName,
                authsdDate = createHeaderReply.authsdDate,
                authsdItmAmt = createHeaderReply.authsdItmAmt,
                authsdPosition = createHeaderReply.authsdPosition,
                authsdPositionDesc = createHeaderReply.authsdPositionDesc,
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
            var currentRow = TitleRow01 + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

                //instancia del Servicio
                var proxyRequisition = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();

                //Header
                var opRequisition = new OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                proxyRequisition.Url = urlService + "/IssueRequisitionItemStocklessService";

                opRequisition.district = _frmAuth.EllipseDsct;
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

                        headerCreateReturnReply.districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct)
                            ? "ICOR"
                            : _frmAuth.EllipseDsct;
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

                        var result = proxyRequisition.immediateReturn(opRequisition, headerCreateReturnReply);

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
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
            }
        }
        private void ManualCreditRequisitionExtended()
        {
            var currentRow = TitleRow01Ext + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01Ext);

                //instancia del Servicio
                var proxyRequisition = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();

                //Header
                var opRequisition = new OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                proxyRequisition.Url = urlService + "/IssueRequisitionItemStocklessService";
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                opRequisition.district = _frmAuth.EllipseDsct;
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

                        headerCreateReturnReply.districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct)
                            ? "ICOR"
                            : _frmAuth.EllipseDsct;
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

                        var result = proxyRequisition.immediateReturn(opRequisition, headerCreateReturnReply);

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
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        

        

        

    }

    

    public static class Queries
    {
        public static string GetItemUnitOfIssue(string stockCode)
        {
            var query = "SELECT UNIT_OF_ISSUE FROM ELLIPSE.MSF100 SC WHERE SC.STOCK_CODE = '" + stockCode + "' ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetItemDirectOrder(string stockCode)
        {
            var query = "SELECT SCI.DIRECT_ORDER_IND FROM ELLIPSE.MSF170 SCI WHERE STOCK_CODE = '" + stockCode + "' ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

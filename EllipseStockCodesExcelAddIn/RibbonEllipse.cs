using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseRequisitionClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService; 
// ReSharper disable FieldCanBeMadeReadOnly.Local

namespace EllipseStockCodesExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Application _excelApp;

        private const string SheetName0101 = "ListaStockCodes";//Format Requisition
        private const string SheetName0102 = "ResultadosStockCodes";
        private const string SheetName0301 = "ListaBusqueda";//Format
        private const string SheetName0302 = "Resultados";
        private const string ValidationSheetName01 = "ValidationSheetSC";
        private const string ValidationSheetName03 = "ValidationSheet";

        private const int TitleRow0101 = 6;
        private const int TitleRow0102 = 5;
        private const int TitleRow0301 = 6;
        private const int TitleRow0302 = 5;
        private const int ResultColumn0101 = 3;
        private const int ResultColumn0102 = 15;//aplica como indicador de ultimo registro
        private const int ResultColumn0301 = 3;
        private const int ResultColumn0302 = 37;//aplica como indicador de ultimo registro
        private const string TableName0101 = "StockCodesTable";
        private const string TableName0102 = "ReviewReqSCTable";
        private const string TableName0301 = "SearchListTable";
        private const string TableName0302 = "ReviewTable";

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            _eFunctions.DebugQueries = false;
            _eFunctions.DebugErrors = false;
            _eFunctions.DebugWarnings = false;
            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }
        private void btnFormatRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            FormatRequisitionSheet();
        }
        private void btnFormatPurchaseOrdersExtended_Click(object sender, RibbonControlEventArgs e)
        {
            FormatPurchaseOrderExtendedSheet();
        }
        private void btnReviewStockCodesRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0101) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0102))
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewReqScList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        
        private void btnReviewPurchaseOrders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0301) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0302))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewPurchaseOrderExtendedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        public void FormatRequisitionSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0101;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName01);
                
                //CONSTRUYO LA HOJA 0101
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE TRANSACCIONES DE STOCK CODES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A4").Value = "ESTADO REQ.";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("C3").Value = "FECHA INICIAL";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("C4").Value = "FECHA FINAL";
                _cells.GetCell("C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                _cells.GetCell("E3").Value = "TIPO REQ.";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("E4").Value = "TIPO TRANS.";//TABLE_TYPE 'IT'
                _cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("G3").Value = "PRIORIDAD";//TABLE_TYPE 'PI'
                _cells.GetCell("G3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName01, 1);

                var reqStatusList = Requisition.RequisitionStatus.GetRequisitionStatusList().Select(status => status.Value).ToList();
                reqStatusList.Add("UNCOMPLETED");
                _cells.SetValidationList(_cells.GetCell("B4"), reqStatusList, ValidationSheetName01, 2);

                var reqTypeList = Requisition.RequisitionType.GetRequisitionTypeList().Select(type => type.Key + " - " + type.Value).ToList();
                reqTypeList.Sort();
                _cells.SetValidationList(_cells.GetCell("F3"), reqTypeList, ValidationSheetName01, 3);

                var transTypeList = Requisition.TransactionType.GetTransactionTypeList(_eFunctions).Select(type => type.Key + " - " + type.Value).ToList();
                transTypeList.Sort();
                _cells.SetValidationList(_cells.GetCell("F4"), transTypeList, ValidationSheetName01, 4);

                var reqPriorityList = Requisition.PriorityCodes.GetPrioriyCodesList(_eFunctions).Select(type => type.Key + " - " + type.Value).ToList();
                reqPriorityList.Sort();
                _cells.SetValidationList(_cells.GetCell("H3"), reqPriorityList, ValidationSheetName01, 5);

                _cells.GetCell(1, TitleRow0101).Value = "STOCK_CODE";
                _cells.GetCell(1, TitleRow0101).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0101 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow0101).Value = "EVENTO";
                _cells.GetCell(2, TitleRow0101).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn0101, TitleRow0101).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0101, TitleRow0101).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0101, ResultColumn0101, TitleRow0101 + 1), TableName0101);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 0102
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0102;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS STOCK CODE - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                //GENERAL
                _cells.GetRange(1, TitleRow0102, ResultColumn0102 - 1, TitleRow0102).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0102).Value = "STOCK CODE";
                _cells.GetCell(2, TitleRow0102).Value = "DISTRITO";
                _cells.GetCell(3, TitleRow0102).Value = "REQ. NO";
                _cells.GetCell(4, TitleRow0102).Value = "REQ. ITEM";
                _cells.GetCell(5, TitleRow0102).Value = "REQ. TYPE";
                _cells.GetCell(6, TitleRow0102).Value = "TRAN. TYPE";
                _cells.GetCell(7, TitleRow0102).Value = "PRIORITY";
                _cells.GetCell(8, TitleRow0102).Value = "W/H";
                _cells.GetCell(9, TitleRow0102).Value = "REQD. BY.";
                _cells.GetCell(10, TitleRow0102).Value = "REQD. DATE";
                _cells.GetCell(11, TitleRow0102).Value = "QTY REQD";
                _cells.GetCell(12, TitleRow0102).Value = "PO ITEM";
                _cells.GetCell(13, TitleRow0102).Value = "ITEM STATUS";
                _cells.GetCell(14, TitleRow0102).Value = "CREATION DATE";
                _cells.GetCell(15, TitleRow0102).Value = "DELIVERY INST.";
                _cells.GetRange(1, TitleRow0102 + 1, ResultColumn0102, TitleRow0102 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0102, ResultColumn0102, TitleRow0102 + 1), TableName0102);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        public void FormatPurchaseOrderExtendedSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.Sheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0301;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName03);

                //CONSTRUYO LA HOJA 0301
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE STOCKCODES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A4").Value = "BUSCAR POR:";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Value = "PURCHASE ORDER";
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("C3").Value = "FECHA CREACIÓN INICIAL";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("C4").Value = "FECHA CREACIÓN FINAL";
                _cells.GetCell("C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                _cells.GetCell("E3").Value = "ESTADO";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F3").Style = _cells.GetStyle(StyleConstants.Select);
                //_cells.GetCell("E4").Value = "CONDICIÓN 3";
                //_cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetCell("F4").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName03, 1);
                var searchType = new List<string> { "PURCHASE ORDER", "STOCK CODE", "CONSULTAR TODO" };
                _cells.SetValidationList(_cells.GetCell("B4"), searchType, ValidationSheetName03, 2);
                //listas de validación
                var itemList1 = PurchaseOrderActions.OrderStatus.GetStatusList();
                var poStatusLust = itemList1.Select(item => item.Key + " - " + item.Value).ToList();
                _cells.SetValidationList(_cells.GetCell("F3"), poStatusLust, ValidationSheetName03, 3);

                _cells.GetCell(1, TitleRow0301).Value = "PURCH.ORD./STOCK_CODE";
                _cells.GetCell(1, TitleRow0301).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0301 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow0301).Value = "EVENTO";
                _cells.GetCell(2, TitleRow0301).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn0301, TitleRow0301).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0301, TitleRow0301).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0301, ResultColumn0301, TitleRow0301 + 1), TableName0301);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO HOJA 0302
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0302;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS PURCHASE ORDERS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetRange(1, TitleRow0302, ResultColumn0302 - 1, TitleRow0302).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0302).Value = "PO_NO";
                _cells.GetCell(2, TitleRow0302).Value = "PO_ITEM_NO";
                _cells.GetCell(3, TitleRow0302).Value = "PREQ_STK_CODE";
                _cells.GetCell(4, TitleRow0302).Value = "PART_NO";
                _cells.GetCell(5, TitleRow0302).Value = "MNEMONIC";
                _cells.GetCell(6, TitleRow0302).Value = "ITEM_NAME";
                _cells.GetCell(7, TitleRow0302).Value = "DESCRIPCIÓN";
                _cells.GetCell(8, TitleRow0302).Value = "CREATION_DATE";
                _cells.GetCell(9, TitleRow0302).Value = "ORDER_DATE";
                _cells.GetCell(10, TitleRow0302).Value = "ORIG_DUE_DATE";
                _cells.GetCell(11, TitleRow0302).Value = "OFST_RCPT_DATE";
                _cells.GetCell(12, TitleRow0302).Value = "ONST_RCPT_DATE";
                _cells.GetCell(8, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(9, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(10, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(11, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(12, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(13, TitleRow0302).Value = "ORIG_NET_PR_I";
                _cells.GetCell(14, TitleRow0302).Value = "CURR_NET_PR_I";
                _cells.GetCell(15, TitleRow0302).Value = "GROSS_PRICE_P";
                _cells.GetCell(16, TitleRow0302).Value = "UNIT_OF_PURCHASE";
                _cells.GetCell(17, TitleRow0302).Value = "CONV_FACTOR";
                _cells.GetCell(18, TitleRow0302).Value = "ORIG_QTY_I";
                _cells.GetCell(19, TitleRow0302).Value = "CURR_QTY_I";
                _cells.GetCell(20, TitleRow0302).Value = "QTY_RCV_OFST_I";
                _cells.GetCell(21, TitleRow0302).Value = "QTY_RCV_DIR_I";
                _cells.GetCell(22, TitleRow0302).Value = "FREIGHT_CODE";
                _cells.GetCell(23, TitleRow0302).Value = "DELIV_LOCATION";
                _cells.GetCell(24, TitleRow0302).Value = "EXPEDITE_CODE";
                _cells.GetCell(25, TitleRow0302).Value = "SUPPLIER_NO";
                _cells.GetCell(26, TitleRow0302).Value = "SUPPLIER_NAME";
                _cells.GetCell(27, TitleRow0302).Value = "UNIT OF ISSUE";
                _cells.GetCell(28, TitleRow0302).Value = "BODEGA PPAL";
                _cells.GetCell(29, TitleRow0302).Value = "SOH";
                _cells.GetCell(30, TitleRow0302).Value = "IN TRANSIT";
                _cells.GetCell(31, TitleRow0302).Value = "DUES IN";
                _cells.GetCell(32, TitleRow0302).Value = "DUES OUT";
                _cells.GetCell(33, TitleRow0302).Value = "RESERVED";
                _cells.GetCell(34, TitleRow0302).Value = "ROP";
                _cells.GetCell(35, TitleRow0302).Value = "ROQ";
                _cells.GetCell(36, TitleRow0302).Value = "USO12_UNSCH";
                _cells.GetCell(37, TitleRow0302).Value = "CURRENT_UNSCH";
                _cells.GetRange(1, TitleRow0302 + 1, ResultColumn0302, TitleRow0302 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0302, ResultColumn0302, TitleRow0302 + 1), TableName0302);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();


                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja." + ex.Message);
            }
        }

        public void ReviewReqScList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var scCells = new ExcelStyleCells(_excelApp, SheetName0101);
            scCells.SetAlwaysActiveSheet(false);

            var resultCells = new ExcelStyleCells(_excelApp, SheetName0102);
            resultCells.SetAlwaysActiveSheet(false);
            resultCells.ClearTableRange(TableName0102);

            var districtCode = _cells.GetEmptyIfNull(scCells.GetCell(2, 3).Value2);
            var scStatus = _cells.GetEmptyIfNull(scCells.GetCell(2, 4).Value2);
            var startDate = _cells.GetEmptyIfNull(scCells.GetCell(4, 3).Value2);
            var endDate = _cells.GetEmptyIfNull(scCells.GetCell(4, 4).Value2);
            var reqType = _cells.GetEmptyIfNull(scCells.GetCell(6, 3).Value2);
            var transType = _cells.GetEmptyIfNull(scCells.GetCell(6, 4).Value2);
            var priorityCode = _cells.GetEmptyIfNull(scCells.GetCell(8, 3).Value2);

            if (reqType != null && reqType.Contains(" - "))
                reqType = reqType.Substring(0, reqType.IndexOf(" - ", StringComparison.Ordinal));
            if (transType != null && transType.Contains(" - "))
                transType = transType.Substring(0, transType.IndexOf(" - ", StringComparison.Ordinal));
            if (priorityCode != null && priorityCode.Contains(" - "))
                priorityCode = priorityCode.Substring(0, priorityCode.IndexOf(" - ", StringComparison.Ordinal));
            var j = TitleRow0101 + 1;//itera según cada stock code
            var i = TitleRow0102 + 1;//itera la celda para cada req-sc

            while (!string.IsNullOrEmpty("" + scCells.GetCell(1, j).Value))
            {
                try
                {
                    var stockCode = _cells.GetEmptyIfNull(scCells.GetCell(1, j).Value2);
                    stockCode = (stockCode != null && stockCode.Length < 9) ? stockCode.PadLeft(9, '0') : stockCode;

                    if (!string.IsNullOrWhiteSpace(scStatus) && !scStatus.Equals("UNCOMPLETED"))
                        scStatus = Requisition.ItemStatus.GetStatusCode(scStatus);

                    var sqlQuery = Queries.GetFetchRequisitionStockCodeQuery(_eFunctions.dbReference, _eFunctions.dbLink, districtCode, stockCode, scStatus, startDate, endDate, reqType, transType, priorityCode);
                    if (_eFunctions.DebugQueries)
                        _cells.GetCell("L1").Value = sqlQuery;

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {

                        resultCells.GetCell(1, i).Value = "" + odr["STOCK_CODE"];
                        resultCells.GetCell(2, i).Value = "" + odr["DSTRCT_CODE"];
                        resultCells.GetCell(3, i).Value = "" + odr["IREQ_NO"];
                        resultCells.GetCell(4, i).Value = "" + odr["IREQ_ITEM"];
                        resultCells.GetCell(5, i).Value = "" + odr["IREQ_TYPE"];
                        resultCells.GetCell(6, i).Value = "" + odr["ISS_TRAN_TYPE"];
                        resultCells.GetCell(7, i).Value = "" + odr["PRIORITY_CODE"];
                        resultCells.GetCell(8, i).Value = "" + odr["WHOUSE_ID"];
                        resultCells.GetCell(9, i).Value = "" + odr["REQUESTED_BY"];
                        resultCells.GetCell(10,i).Value = "" + odr["REQ_BY_DATE"];
                        resultCells.GetCell(11,i).Value = "" + odr["QTY_REQ"];
                        resultCells.GetCell(12,i).Value = "" + odr["PO_ITEM_NO"];
                        resultCells.GetCell(13,i).Value = "" + odr["ITEM_141_STAT"];
                        resultCells.GetCell(14, i).Value = "" + odr["CREATION_DATE"];
                        resultCells.GetCell(15,i).Value = "" + odr["DELIV_INSTR_A"] + odr["DELIV_INSTR_B"];
                        i++;
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0102))
                            resultCells.GetCell(3, i).Select();
                    }
                    scCells.GetCell(ResultColumn0101, j).Style = StyleConstants.Success;
                    scCells.GetCell(ResultColumn0101, j).Value = "SUCCESS";
                }
                catch (Exception ex)
                {
                    scCells.GetCell(ResultColumn0101, j).Style = StyleConstants.Error;
                    scCells.GetCell(ResultColumn0101, j).Value += "ERROR:" + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewReqScList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    scCells.GetCell(2, j).Value = "CONSULTA DE VALES-ITEMS";
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0101))
                        scCells.GetCell(1, j).Select();
                    _eFunctions.CloseConnection();
                    j++;//aumenta SC
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }
        public void ReviewPurchaseOrderExtendedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);


            var poCells = new ExcelStyleCells(_excelApp, SheetName0301);
            poCells.SetAlwaysActiveSheet(false);

            var resultCells = new ExcelStyleCells(_excelApp, SheetName0302);
            resultCells.SetAlwaysActiveSheet(false);
            resultCells.ClearTableRange(TableName0302);

            var fullSearch = false;//para realizar búsquedas completas que no dependan de un PO dado
            var districtCode = _cells.GetEmptyIfNull(poCells.GetCell(2, 3).Value2);
            var searchType = _cells.GetEmptyIfNull(poCells.GetCell(2, 4).Value2);
            var startDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 3).Value2);
            var endDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 4).Value2);
            var poStatus = _cells.GetEmptyIfNull(poCells.GetCell(6, 3).Value2);

            if (poStatus != null && poStatus.Contains(" - "))
                poStatus = poStatus.Substring(0, poStatus.IndexOf(" - ", StringComparison.Ordinal));

            var j = TitleRow0301 + 1;//itera según cada stock code
            var i = TitleRow0302 + 1;//itera la celda para cada req-sc

            if (string.IsNullOrWhiteSpace(searchType) || searchType.Equals("CONSULTAR TODO"))
            {
                if (string.IsNullOrWhiteSpace(startDate) && string.IsNullOrWhiteSpace(poStatus)) //TO DO
                    throw new NullReferenceException("Debe seleccionar una fecha inicial o un estado de orden para esta búsqueda");
                fullSearch = true;
            }

            while (!string.IsNullOrEmpty("" + poCells.GetCell(1, j).Value) || fullSearch)
            {
                try
                {
                    string purchaseOrder = null;
                    string stockCode = null;
                    if (searchType.Equals("PURCHASE ORDER"))
                        purchaseOrder = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);
                    if (searchType.Equals("STOCK CODE"))
                        stockCode = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);

                    stockCode = (stockCode != null && stockCode.Length < 9) ? stockCode.PadLeft(9, '0') : stockCode;

                    var sqlQuery = Queries.GetFetchPurchaseOrderQuery(_eFunctions.dbReference, _eFunctions.dbLink, districtCode, purchaseOrder, stockCode, startDate, endDate, poStatus);
                    if (_eFunctions.DebugQueries)
                        _cells.GetCell("L1").Value = sqlQuery;

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {

                        resultCells.GetCell(01, i).Value = "" + odr["PO_NO"];
                        resultCells.GetCell(02, i).Value = "" + odr["PO_ITEM_NO"];
                        resultCells.GetCell(03, i).Value = "" + odr["PREQ_STK_CODE"];
                        resultCells.GetCell(04, i).Value = "" + odr["PART_NO"];
                        resultCells.GetCell(05, i).Value = "" + odr["MNEMONIC"];
                        resultCells.GetCell(06, i).Value = "" + odr["ITEM_NAME"];
                        resultCells.GetCell(07, i).Value = "" + resultCells.GetEmptyIfNull(odr["DESC_LINEX1"]) +
                            resultCells.GetEmptyIfNull(odr["DESC_LINEX2"]) +
                            resultCells.GetEmptyIfNull(odr["DESC_LINEX3"]) +
                            resultCells.GetEmptyIfNull(odr["DESC_LINEX4"]);
                        resultCells.GetCell(08, i).Value = "" + odr["CREATION_DATE"];
                        resultCells.GetCell(09, i).Value = "" + odr["ORDER_DATE"];
                        resultCells.GetCell(10, i).Value = "" + odr["ORIG_DUE_DATE"];
                        resultCells.GetCell(11, i).Value = "" + odr["OFST_RCPT_DATE"];
                        resultCells.GetCell(12, i).Value = "" + odr["ONST_RCPT_DATE"];
                        resultCells.GetCell(13, i).Value = "" + odr["ORIG_NET_PR_I"];
                        resultCells.GetCell(14, i).Value = "" + odr["CURR_NET_PR_I"];
                        resultCells.GetCell(15, i).Value = "" + odr["GROSS_PRICE_P"];
                        resultCells.GetCell(16, i).Value = "" + odr["UNIT_OF_PURCH"];
                        resultCells.GetCell(17, i).Value = "" + odr["CONV_FACTOR"];
                        resultCells.GetCell(18, i).Value = "" + odr["ORIG_QTY_I"];
                        resultCells.GetCell(19, i).Value = "" + odr["CURR_QTY_I"];
                        resultCells.GetCell(20, i).Value = "" + odr["QTY_RCV_OFST_I"];
                        resultCells.GetCell(21, i).Value = "" + odr["QTY_RCV_DIR_I"];
                        resultCells.GetCell(22, i).Value = "" + odr["FREIGHT_CODE"];
                        resultCells.GetCell(23, i).Value = "" + odr["DELIV_LOCATION"];
                        resultCells.GetCell(24, i).Value = "" + odr["EXPEDITE_CODE"];
                        resultCells.GetCell(25, i).Value = "" + odr["SUPPLIER_NO"];
                        resultCells.GetCell(26, i).Value = "" + odr["SUPPLIER_NAME"];
                        resultCells.GetCell(27, i).Value = "" + odr["UNIT_OF_ISSUE"];
                        resultCells.GetCell(28, i).Value = "" + odr["BODEGA_PRINCIPAL"];
                        resultCells.GetCell(29, i).Value = "" + odr["SOH"];
                        resultCells.GetCell(30, i).Value = "" + odr["IN_TRANSIT"];
                        resultCells.GetCell(31, i).Value = "" + odr["DUES_IN"];
                        resultCells.GetCell(32, i).Value = "" + odr["DUES_OUT"];
                        resultCells.GetCell(33, i).Value = "" + odr["RESERVED"];
                        resultCells.GetCell(34, i).Value = "" + odr["ROP"];
                        resultCells.GetCell(35, i).Value = "" + odr["ROQ"];
                        resultCells.GetCell(36, i).Value = "" + odr["USO12_UNSCH"];
                        resultCells.GetCell(37, i).Value = "" + odr["CURRENT_UNSCH"];

                        i++;
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0302))
                            resultCells.GetCell(3, i).Select();
                    }
                    poCells.GetCell(ResultColumn0301, j).Style = StyleConstants.Success;
                    poCells.GetCell(ResultColumn0301, j).Value = "SUCCESS";
                }
                catch (Exception ex)
                {
                    poCells.GetCell(ResultColumn0301, j).Style = StyleConstants.Error;
                    poCells.GetCell(ResultColumn0301, j).Value += "ERROR:" + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewPurchaseOrderList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    poCells.GetCell(2, j).Value = "CONSULTA DE PO-ITEMS";
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0301))
                        poCells.GetCell(1, j).Select();
                    j++;//aumenta SC
                    if (fullSearch)
                        fullSearch = false;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();

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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn(Assembly.GetExecutingAssembly()).ShowDialog();
        }
        
    }

    public class PurchaseOrder
    {
        public string PurchaseNumber;
        public string OrderType;
        public string NumberOfItems;
        public string OrderStatus;
        public string TotalEstimatedValue;
        public string AuthorizedStatus;
        public string SupplierNumber;
        public string SupplierName;
        public string DeliveryLocation;
        public string FreightCode;
        public string OrderDate;
        public string PurchaseOfficer;
        public string PurchaseTeam;
        public string Medium;
        public string OriginCode;
        public string Currency;

        public List<PurchaseOrderItem> Items;
    }

    public class PurchaseOrderItem
    {
        public string Index;
        public string Quantity;
        public string DueDate;
        public string ExpediteCode;
        public string Discount1;
        public string Discount2;
        public string Surcharge1;
        public string Surcharge2;
        public string GrossPrice;
        public string UnitOfPurchase;
        public string ConversionFactor;
    }

    public static class PurchaseOrderActions
    {
        public static class OrderStatus
        {
            public static string UnprintedCode = "0";
            public static string Unprinted = "UNPRINTED";
            public static string PrintedCode = "1";
            public static string Printed = "PRINTED";
            public static string ModifiedCode = "2";
            public static string Modified = "MODIFIED";
            public static string CancelledCode = "3";
            public static string Cancelled = "CANCELLED";
            public static string CompletedCode = "9";
            public static string Completed = "COMPLETED";
            public static string UncompletedCode = "U";
            public static string Uncompleted = "UNCOMPLETED";

            public static Dictionary<string, string> GetStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {UnprintedCode, Unprinted},
                    {PrintedCode, Printed},
                    {ModifiedCode, Modified},
                    {CancelledCode, Cancelled},
                    {CompletedCode, Completed},
                    {UncompletedCode, Uncompleted}
                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
    }
    public static class Queries
    {
        public static string GetFetchRequisitionStockCodeQuery(string dbReference, string dbLink, string districtCode, string stockCode, string scStatus, string startDate, string finishDate, string reqType, string transType, string priorityCode)
        {
            if (!string.IsNullOrWhiteSpace(districtCode))
                districtCode = " AND SC.DSTRCT_CODE = '" + districtCode + "'";
            if (!string.IsNullOrWhiteSpace(scStatus))
            {
                if (scStatus.Equals("UNCOMPLETED"))
                    scStatus = " AND SC.ITEM_141_STAT <> '" + Requisition.ItemStatus.CompleteCode + "'";
                else
                    scStatus = " AND SC.ITEM_141_STAT = '" + scStatus + "'";
            }
            if (!string.IsNullOrWhiteSpace(startDate))
                startDate = " AND RQ.CREATION_DATE >= " + startDate;
            if (!string.IsNullOrWhiteSpace(finishDate))
                finishDate = " AND RQ.CREATION_DATE <= " + finishDate;
            if (!string.IsNullOrWhiteSpace(reqType))
                reqType = " AND RQ.IREQ_TYPE = '" + reqType + "'";
            if (!string.IsNullOrWhiteSpace(transType))
                transType = " AND RQ.ISS_TRAN_TYPE = '" + transType + "'";
            if (!string.IsNullOrWhiteSpace(priorityCode))
                priorityCode = " AND RQ.PRIORITY_CODE = '" + priorityCode + "'";

            var sqlQuery = "" +
                           " SELECT " +
                           "   SC.DSTRCT_CODE, SC.IREQ_NO, RQ.IREQ_TYPE, RQ.ISS_TRAN_TYPE, SC.STOCK_CODE, SC.IREQ_ITEM," +
                           "   RQ.AUTHSD_STATUS, RQ.HDR_140_STATUS, SC.ITEM_141_STAT," +
                           "   RQ.PRIORITY_CODE, SC.WHOUSE_ID, RQ.REQUESTED_BY, RQ.CREATION_DATE, RQ.REQ_BY_DATE, RQ.DELIV_INSTR_A, RQ.DELIV_INSTR_B," +
                           "   SC.QTY_REQ, SC.PO_ITEM_NO" +
                           " FROM ELLIPSE.MSF141 SC" +
                           " JOIN ELLIPSE.MSF140 RQ" +
                           " ON SC.IREQ_NO         = RQ.IREQ_NO" +
                           " WHERE STOCK_CODE      = '" + stockCode + "'" +
                           districtCode +
                           scStatus +
                           startDate +
                           finishDate +
                           reqType +
                           transType +
                           priorityCode;
            return sqlQuery;
        }

        public static string GetFetchPurchaseOrderQuery(string dbReference, string dbLink, string districtCode, string purchaseOrder, string stockCode, string startDate, string finishDate, string poStatus)
        {
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                districtCode = " AND PO.DSTRCT_CODE = '" + districtCode + "'";// + " AND PN.DSTRCT_CODE = '" + districtCode + "'";
            if (!string.IsNullOrWhiteSpace(startDate))
                startDate = " AND PO.CREATION_DATE >= " + startDate;
            if (!string.IsNullOrWhiteSpace(finishDate))
                finishDate = " AND PO.CREATION_DATE <= " + finishDate;
            if(!string.IsNullOrWhiteSpace(purchaseOrder))
                purchaseOrder =  " PO.PO_NO = '" + purchaseOrder + "'";
            if (!string.IsNullOrWhiteSpace(stockCode))
            {
                stockCode = " POI.PREQ_STK_CODE = '" + stockCode + "'";
                if (!string.IsNullOrWhiteSpace(purchaseOrder))
                    stockCode = " AND " + stockCode;
            }
            if (!string.IsNullOrWhiteSpace(poStatus))
            {
                if (poStatus.Equals("U"))
                    poStatus = " AND PO.STATUS_220 NOT IN ('3', '9')";
                else
                    poStatus = " AND PO.STATUS_220 = '" + poStatus + "'";
            }


            var sqlQuery = "" +
                           " WITH POITEMS AS(" +
                           "  SELECT" +
                           "    POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE, SC.ITEM_NAME, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, PN.PART_NO, PN.MNEMONIC, "+
                           "    POI.GROSS_PRICE_P, POI.UNIT_OF_PURCH, POI.CONV_FACTOR, "+
                           "    PN.PREF_PART_IND, MIN(PN.PREF_PART_IND) OVER (PARTITION BY POI.PREQ_STK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE ORDER BY POI.PREQ_STK_CODE, PN.PREF_PART_IND ASC) ROWPPI, " +
                           "    PO.STATUS_220, PO.CREATION_DATE, PO.ORDER_DATE, POI.ORIG_DUE_DATE, POI.ORIG_NET_PR_I, POI.CURR_NET_PR_I, POI.ORIG_QTY_I, POI.CURR_QTY_I, POI.QTY_RCV_OFST_I, POI.OFST_RCPT_DATE, POI.QTY_RCV_DIR_I, POI.ONST_RCPT_DATE, PO.FREIGHT_CODE, PO.DELIV_LOCATION, POI.EXPEDITE_CODE, PO.SUPPLIER_NO, SUP.SUPPLIER_NAME, PO.PO_MEDIUM_IND, PO.ORIGIN_CODE, PO.PURCH_OFFICER, PO.TEAM_ID" +
                           "  FROM" +
                           "    ELLIPSE.MSF220 PO JOIN ELLIPSE.MSF221 POI ON PO.PO_NO = POI.PO_NO LEFT JOIN ELLIPSE.MSF100 SC ON POI.PREQ_STK_CODE = SC.STOCK_CODE LEFT JOIN ELLIPSE.MSF110 PN ON POI.PREQ_STK_CODE = PN.STOCK_CODE LEFT JOIN ELLIPSE.MSF200 SUP ON PO.SUPPLIER_NO = SUP.SUPPLIER_NO" +
                           "  WHERE" +
                           purchaseOrder +
                           stockCode +
                           districtCode +
                           startDate +
                           finishDate +
                           poStatus +
                           "    AND PN.STATUS_CODES = 'V'" +
                           "  ORDER BY POI.PO_ITEM_NO" +
                           "  )," +
                           " SCSTAT AS(" +
                           " SELECT STAT.DSTRCT_CODE, SC.stock_code, STAT.creation_date, STAT.last_mod_date, SC.stk_desc, SC.unit_of_issue, STAT.class, STAT.raf as algoritmo, STAT.invent_cost_pr as price, STAT.home_whouse as bodega_principal, ellipse.get_soh('ICOR',SC.stock_code) as soh, " +
                           "  STAT.in_transit, STAT.dues_in, STAT.dues_out, STAT.reserved, STAT.rop, STAT.REORDER_QTY roq, STAT.exp_element as detalle_gasto, STAT.restrict_rule as restr, STAT.direct_order_ind as do_ind, STAT.purch_officer as purchaser, " +
                           "  (select sum(unsched_usage) from ellipse.msf175 where dstrct_code=STAT.dstrct_code and stock_code=SC.stock_code " +
                           "  and full_acct_per between (select to_char(to_date(max(full_acct_per),'yyyymm')-365,'yyyymm') from ellipse.msf175 where dstrct_code=STAT.dstrct_code and stock_code=SC.stock_code) and (select max(full_acct_per) from ellipse.msf175 where dstrct_code=STAT.dstrct_code and stock_code=SC.stock_code) "+
                           "  and TRIM(WHOUSE_ID) IS NOT NULL ) as uso12_unsch, " +
                           "  (select sum(unsched_usage) from ellipse.msf175 where dstrct_code=STAT.dstrct_code and stock_code=SC.stock_code " +
                           "  and full_acct_per=(select max(full_acct_per) from ellipse.msf175 where dstrct_code=STAT.dstrct_code and stock_code=SC.stock_code) " +
                           "  and TRIM(WHOUSE_ID) IS NOT NULL ) as current_unsch, " +
                           "  STAT.invt_controllr as adi from ellipse.msf100 SC LEFT JOIN ellipse.msf170 STAT ON SC.stock_code = STAT.stock_code)" +
                           "  SELECT * FROM POITEMS LEFT JOIN SCSTAT ON POITEMS.PREQ_STK_CODE = SCSTAT.STOCK_CODE AND SCSTAT.DSTRCT_CODE = 'ICOR' WHERE POITEMS.PREF_PART_IND = POITEMS.MINPPI AND ROWPPI = 1";
            sqlQuery = sqlQuery.Replace("WHERE AND", "WHERE");
            return sqlQuery;
        }
    }


}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
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

        private const string SheetName01 = "ListaStockCodes";//Format Requisition
        private const string SheetName02 = "ResultadosStockCodes";
        private const string ValidationSheetName01 = "ValidationSheet";

        private const int TitleRow01 = 6;
        private const int TitleRow02 = 5;
        private const int ResultColumn01 = 3;
        private const string TableName01 = "SearchTable";
        private const string TableName02 = "ResultsTable";
        private string _searchType = SearchType.Inventory;

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
            FormatSheet();
        }
        private void btnReviewStockCodesRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _searchType = SearchType.Inventory;
                _thread = new Thread(GetReviewResult);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        
        
        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName01);
                
                //CONSTRUYO LA HOJA 0101
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE DE STOCK CODES - INVENTARIO & TRANSACCIONES - ELLIPSE 8";
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
                _cells.GetCell("A4").Value = "BUSCAR POR";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Value = "StockCode";
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName01, 1);

                var searchCriteriaList = new List<string>();
                searchCriteriaList.Add(SearchCriteria.StockCode);
                searchCriteriaList.Add(SearchCriteria.PartNumber);
                searchCriteriaList.Add(SearchCriteria.RequisitionNo);
                searchCriteriaList.Add(SearchCriteria.PurchaseOrder);
                _cells.SetValidationList(_cells.GetCell("B4"), searchCriteriaList, ValidationSheetName01, 2);

                _cells.GetCell(1, TitleRow01).Value = "SC/PN/REQ";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow01).Value = "EVENTO";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 02
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        public void GetReviewResult()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var cp = new ExcelStyleCells(_excelApp, SheetName01); //cells parameters
                var cr = new ExcelStyleCells(_excelApp, SheetName02); //cells results

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(SheetName02).Activate();

                //Elimino los registros anteriores
                cr.ClearTableRange(TableName02);
                cr.DeleteTableRange(TableName02);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                //Obtengo los parámetros de búsqueda
                var district = cp.GetEmptyIfNull(cp.GetCell("B3").Value);
                var searchCriteriaKey = cp.GetEmptyIfNull(cp.GetCell("B4").Value);

                var rowParam = TitleRow01 + 1;
                var rowResult = TitleRow02 + 1;
                var validOnly = cbValidOnly.Checked;
                var preferedOnly = cbPreferedOnly.Checked;

                while (!string.IsNullOrEmpty("" + cp.GetCell(1, rowParam).Value))
                {
                    try
                    {
                        var searchCriteriaValue = cp.GetEmptyIfNull(cp.GetCell(1, rowParam).Value);
                        var sqlQuery = "";
                        if (_searchType.Equals(SearchType.Inventory))
                        {
                            sqlQuery = Queries.GetFetchInventoryStockCodeQuery(_eFunctions.dbReference, _eFunctions.dbLink, district, searchCriteriaKey, searchCriteriaValue, validOnly, preferedOnly);
                        }
                        else
                        {
                            throw new Exception("Debe seleccionar una opción de búsqueda válida");
                        }

                        var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                        if (dataReader == null)
                            return;

                        //Cargo el encabezado de la tabla y doy formato
                        if (rowParam == TitleRow01 + 1)
                        {
                            for (var k = 0; k < dataReader.FieldCount; k++)
                                cr.GetCell(k + 1, TitleRow02).Value2 = "'" + dataReader.GetName(k);

                            _cells.FormatAsTable(cr.GetRange(1, TitleRow02, dataReader.FieldCount, TitleRow02 + 1),
                                TableName02);
                        }
                        //cargo los datos de cada consulta
                        if (dataReader.IsClosed || !dataReader.HasRows) return;

                        while (dataReader.Read())
                        {
                            for (var k = 0; k < dataReader.FieldCount; k++)
                                cr.GetCell(k + 1, rowResult).Value2 = "'" + dataReader[k].ToString().Trim();
                            rowResult++;
                        }
                    }
                    catch (Exception ex)
                    {
                        cp.GetCell(1, rowParam).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, rowParam).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:GetReviewResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                    }
                    finally
                    {
                        rowParam++;
                        _eFunctions.CloseConnection();
                    }
                }

                

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetReviewResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
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

    public static class SearchCriteria
    {
        public static string StockCode = "StockCode";
        public static string PartNumber = "PartNumber";
        public static string RequisitionNo = "RequisitionNo";
        public static string PurchaseOrder = "PurchaseOrder";
    }

    public static class SearchType
    {
        public static string Inventory = "Inventory";
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
        public static string GetFetchInventoryStockCodeQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, bool validOnly, bool preferedOnly)
        {
            if (!string.IsNullOrWhiteSpace(districtCode))
                districtCode = " AND (PN.DSTRCT_CODE = '" + districtCode + "' OR TRIM(PN.DSTRCT_CODE) IS NULL)";
            var paramSearch = "";
            if (searchCriteriaKey.Equals(SearchCriteria.StockCode))
                paramSearch = " AND SC.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') +"'";
            else if (searchCriteriaKey.Equals(SearchCriteria.PartNumber))
                paramSearch = " AND TRIM(PN.PART_NO) = '" + searchCriteriaValue + "'";

            var paramValidOnly = "";
            if (validOnly)
                paramValidOnly = "AND PN.STATUS_CODES = 'V'";

            var sqlQuery = "" +
                           "WITH SCINV AS(" +
                           "    SELECT SC.STOCK_CODE, PN.PART_NO, PN.MNEMONIC, PN.DSTRCT_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, SC.CLASS STOCK_CLASS, SC.STOCK_TYPE," +
                           "        INV.CREATION_DATE, INV.LAST_MOD_DATE, INV.CLASS, INV.RAF, INV.INVENT_COST_PR AS PRICE, INV.HOME_WHOUSE, ELLIPSE.GET_SOH(PN.DSTRCT_CODE, SC.STOCK_CODE) AS SOH," +
                           "        INV.IN_TRANSIT, INV.DUES_IN, INV.DUES_OUT, INV.RESERVED, INV.ROP, INV.ROQ, INV.REORDER_QTY, INV.EXP_ELEMENT, INV.RESTRICT_RULE, INV.DIRECT_ORDER_IND, INV.PURCH_OFFICER, INV.INVT_CONTROLLR," +
                           "        PN.PREF_PART_IND, PN.STATUS_CODES" +
                           "    FROM ELLIPSE.MSF100 SC" +
                           "        LEFT JOIN ELLIPSE.MSF110 PN ON SC.STOCK_CODE = PN.STOCK_CODE" +
                           "        LEFT JOIN ELLIPSE.MSF170 INV ON SC.STOCK_CODE = INV.STOCK_CODE" +
                           "    WHERE " +
                           "" + paramValidOnly +
                           "" + districtCode +
                           "" + paramSearch +
                           ")" +
                           "SELECT * FROM SCINV";
            sqlQuery = Utils.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE");
            return sqlQuery;
        }
        public static string GetFetchRequisitionStockCodeQuery(string dbReference, string dbLink, string districtCode, string stockCode, string scStatus, string startDate, string finishDate, string reqType, string transType, string priorityCode)
        {
            if (!string.IsNullOrWhiteSpace(districtCode))
                districtCode = " AND SC.DSTRCT_CODE = '" + districtCode + "'";
            //if (!string.IsNullOrWhiteSpace(scStatus))
            //{
            //    if (scStatus.Equals("UNCOMPLETED"))
            //        scStatus = " AND SC.ITEM_141_STAT <> '" + Requisition.ItemStatus.CompleteCode + "'";
            //    else
            //        scStatus = " AND SC.ITEM_141_STAT = '" + scStatus + "'";
            //}
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

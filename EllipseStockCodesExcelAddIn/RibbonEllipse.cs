using System;
using System.Collections.Generic;
using System.Linq;
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

        private const string SheetName01 = "SearchOptions";//Format Requisition
        private const string SheetName02 = "Results";
        private const string ValidationSheetName = "ValidationSheet";

        private const int TitleRow01 = 7;
        private const int TitleRow02 = 5;
        private const int ResultColumn01 = 3;
        private const string TableName01 = "SearchTable";
        private const string TableName02 = "ResultsTable";

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

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
        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
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
                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
                
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


                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();
                var issueCriteriaList = SearchCriteriaIssues.GetSearchCriteriaIssues().Select(g => g.Value).ToList();
                var typeCriteriaList = SearchCriteriaType.GetSearchCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "OPCIÓN";
                _cells.GetCell("B4").Value = SearchCriteriaIssues.Inventory.Value;
                _cells.GetCell("A5").Value = "BUSCAR POR";
                _cells.GetCell("B5").Value = SearchCriteriaType.StockCode.Value;
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Raised.Value;
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell("B4"), issueCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B5"), typeCriteriaList, ValidationSheetName, 3);
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 4);
                _cells.GetCell(1, TitleRow01).Value = "SC/PN/Item";
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
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
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
                var searchIssueKey = cp.GetEmptyIfNull(cp.GetCell("B4").Value);
                var searchCriteriaKey = cp.GetEmptyIfNull(cp.GetCell("B5").Value);

                var searchDateCriteriaKey = cp.GetEmptyIfNull(cp.GetCell("D3").Value);
                var startDate = cp.GetEmptyIfNull(cp.GetCell("D4").Value);
                var endDate = cp.GetEmptyIfNull(cp.GetCell("D5").Value);

                var rowParam = TitleRow01 + 1;
                var rowResult = TitleRow02 + 1;
                var validOnly = cbValidOnly.Checked;
                var preferedOnly = cbPreferedOnly.Checked;

                while (!string.IsNullOrEmpty("" + cp.GetCell(1, rowParam).Value))
                {
                    try
                    {
                        var searchCriteriaValue = cp.GetEmptyIfNull(cp.GetCell(1, rowParam).Value);
                        string sqlQuery;
                        if (searchIssueKey.Equals(SearchCriteriaIssues.Inventory.Value))
                            sqlQuery = Queries.GetFetchInventoryStockCodeQuery(_eFunctions.dbReference, _eFunctions.dbLink, district, searchCriteriaKey, searchCriteriaValue, validOnly, preferedOnly);
                        else if (searchIssueKey.Equals(SearchCriteriaIssues.PurchaseOrder.Value))
                            sqlQuery = Queries.GetFetchPurchaseOrderQuery(_eFunctions.dbReference, _eFunctions.dbLink, district, searchCriteriaKey, searchCriteriaValue, searchDateCriteriaKey, startDate, endDate, validOnly, preferedOnly);
                        else if (searchIssueKey.Equals(SearchCriteriaIssues.Requisition.Value))
                            sqlQuery = Queries.GetFetchRequisitionQuery(_eFunctions.dbReference, _eFunctions.dbLink, district, searchCriteriaKey, searchCriteriaValue, searchDateCriteriaKey, startDate, endDate, validOnly, preferedOnly);
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
                        if (dataReader.IsClosed || !dataReader.HasRows)
                        {
                            cp.GetCell(ResultColumn01, rowParam).Style = StyleConstants.Warning;
                            cp.GetCell(ResultColumn01, rowParam).Value = "No se encontraron datos. Intente la búsqueda desactivando la opción de sólo Parte Número válido y/o preferido";
                        }
                        else
                        {
                            while (dataReader.Read())
                            {
                                for (var k = 0; k < dataReader.FieldCount; k++)
                                    cr.GetCell(k + 1, rowResult).Value2 = "'" + dataReader[k].ToString().Trim();
                                rowResult++;
                            }
                            cp.GetCell(ResultColumn01, rowParam).Style = StyleConstants.Success;
                            cp.GetCell(2, rowParam).Value = "Consulta";
                            cp.GetCell(ResultColumn01, rowParam).Value = "OK";                          
                        }
  
                    }
                    catch (Exception ex)
                    {
                        cp.GetCell(ResultColumn01, rowParam).Style = StyleConstants.Error;
                        cp.GetCell(2, rowParam).Value = "Consulta";
                        cp.GetCell(ResultColumn01, rowParam).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:GetReviewResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                Debugger.LogError("RibbonEllipse:GetReviewResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
            new AboutBoxExcelAddIn().ShowDialog();
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

    public static class SearchCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> StockCode = new KeyValuePair<int, string>(1, "StockCode");
        public static KeyValuePair<int, string> PartNumber = new KeyValuePair<int, string>(2, "PartNumber");
        public static KeyValuePair<int, string> ItemCode = new KeyValuePair<int, string>(3, "ItemCode");

        public static List<KeyValuePair<int, string>> GetSearchCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> {None, StockCode, PartNumber, ItemCode};

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
    public static class SearchCriteriaIssues
    {
        public static KeyValuePair<int, string> Inventory = new KeyValuePair<int, string>(0, "Inventory");
        public static KeyValuePair<int, string> PurchaseOrder = new KeyValuePair<int, string>(1, "PurchaseOrder");
        public static KeyValuePair<int, string> Requisition = new KeyValuePair<int, string>(2, "Requisition");

        public static List<KeyValuePair<int, string>> GetSearchCriteriaIssues(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { Inventory, PurchaseOrder, Requisition};

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
    public static class SearchDateCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Raised = new KeyValuePair<int, string>(1, "Raised");
        //public static KeyValuePair<int, string> Closed = new KeyValuePair<int, string>(2, "Closed");
        //public static KeyValuePair<int, string> PlannedStart = new KeyValuePair<int, string>(3, "PlannedStart");
        //public static KeyValuePair<int, string> PlannedFinnish = new KeyValuePair<int, string>(4, "PlannedFinnish");
        //public static KeyValuePair<int, string> RequiredStart = new KeyValuePair<int, string>(5, "RequiredStart");
        //public static KeyValuePair<int, string> RequiredBy = new KeyValuePair<int, string>(6, "RequiredBy");
        //public static KeyValuePair<int, string> Modified = new KeyValuePair<int, string>(7, "Modified");
        //public static KeyValuePair<int, string> NotFinalized = new KeyValuePair<int, string>(8, "NotFinalized");
        //public static KeyValuePair<int, string> LastModified = new KeyValuePair<int, string>(9, "LastModified");
        //public static KeyValuePair<int, string> Finalized = new KeyValuePair<int, string>(10, "Finalized");

        public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
        {
            //var list = new List<KeyValuePair<int, string>> { None, Raised, Closed, PlannedStart, PlannedFinnish, RequiredStart, RequiredBy, Modified, NotFinalized };
            var list = new List<KeyValuePair<int, string>> { None, Raised};
            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
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

            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))
                paramDistrict = " AND (PN.DSTRCT_CODE = '" + districtCode + "' OR TRIM(PN.DSTRCT_CODE) IS NULL)";
            string paramSearch;
            if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
                paramSearch = " AND SC.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') +"'";
            else if (searchCriteriaKey.Equals(SearchCriteriaType.PartNumber.Value))
                paramSearch = " AND TRIM(PN.PART_NO) = '" + searchCriteriaValue + "'";
            else
                paramSearch = " AND SC.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE PREF_PART_IND = MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           "WITH SCINV AS(" +
                           "    SELECT SC.STOCK_CODE, PN.PART_NO, PN.MNEMONIC, PN.DSTRCT_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, SC.CLASS STOCK_CLASS, SC.STOCK_TYPE," +
                           "        INV.CREATION_DATE, INV.LAST_MOD_DATE, INV.CLASS, INV.RAF, INV.INVENT_COST_PR AS PRICE, INV.HOME_WHOUSE, ELLIPSE.GET_SOH('" + districtCode + "', SC.STOCK_CODE) AS OWNED_SOH, ELLIPSE.GET_CONSIGN_SOH('" + districtCode + "', SC.STOCK_CODE) AS CONSIGN_SOH," +
                           "        INV.IN_TRANSIT, INV.DUES_IN, INV.DUES_OUT, INV.RESERVED, INV.ROP, INV.ROQ, INV.REORDER_QTY, INV.EXP_ELEMENT, INV.RESTRICT_RULE, INV.DIRECT_ORDER_IND, INV.PURCH_OFFICER, INV.INVT_CONTROLLR," +
                           "        PN.PREF_PART_IND, PN.STATUS_CODES," +
                           "        MIN(PN.PREF_PART_IND) OVER (PARTITION BY SC.STOCK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY SC.STOCK_CODE ORDER BY SC.STOCK_CODE, PN.PREF_PART_IND ASC) ROWPPI" +
                           "    FROM ELLIPSE.MSF100 SC" +
                           "        LEFT JOIN ELLIPSE.MSF110 PN ON SC.STOCK_CODE = PN.STOCK_CODE" +
                           "        LEFT JOIN ELLIPSE.MSF170 INV ON SC.STOCK_CODE = INV.STOCK_CODE" +
                           "    WHERE " +
                           " " + paramValidOnly +
                           " " + paramDistrict +
                           " " + paramSearch +
                           ")" +
                           "SELECT * FROM SCINV" +
                           " " + paramPreferedOnly;
            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetFetchRequisitionQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, string dateCriteria, string startDate, string finishDate, bool validOnly, bool preferedOnly)
        {
            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                paramDistrict = " AND RQ.DSTRCT_CODE = '" + districtCode + "'";// + " AND PN.DSTRCT_CODE = '" + districtCode + "'";
            if (dateCriteria.Equals(SearchDateCriteriaType.Raised.Value))
            {
                if (!string.IsNullOrWhiteSpace(startDate))
                    startDate = " AND RQ.CREATION_DATE >= " + startDate;
                if (!string.IsNullOrWhiteSpace(finishDate))
                    finishDate = " AND RQ.CREATION_DATE <= " + finishDate;
            }

            var paramReqNo = "";
            var paramStockCode = "";
            if (searchCriteriaKey.Equals(SearchCriteriaType.ItemCode.Value))
            {
                paramReqNo = " AND RQI.IREQ_NO = '" + searchCriteriaValue + "'";
            }
            else if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
            {
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }
            else
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE PREF_PART_IND = MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           "WITH REQSC AS (" +
                           " SELECT " +
                           "   RQI.DSTRCT_CODE, RQI.IREQ_NO, RQ.IREQ_TYPE, RQ.ISS_TRAN_TYPE, RQI.STOCK_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, PN.PART_NO, PN.MNEMONIC, RQI.IREQ_ITEM," +
                           "   RQ.AUTHSD_STATUS, RQ.HDR_140_STATUS, RQI.ITEM_141_STAT," +
                           "   RQ.PRIORITY_CODE, RQI.WHOUSE_ID, RQ.REQUESTED_BY, RQ.CREATION_DATE, RQ.REQ_BY_DATE, RQ.DELIV_INSTR_A, RQ.DELIV_INSTR_B," +
                           "   RQI.QTY_REQ, RQI.PO_ITEM_NO," +
                           "   PN.PREF_PART_IND, PN.STATUS_CODES," +
                           "   MIN(PN.PREF_PART_IND) OVER (PARTITION BY RQI.STOCK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY RQI.IREQ_NO, RQI.IREQ_ITEM, RQI.STOCK_CODE ORDER BY RQI.STOCK_CODE, PN.PREF_PART_IND ASC) ROWPPI" +
                           " FROM ELLIPSE.MSF141 RQI" +
                           " JOIN ELLIPSE.MSF140 RQ ON RQI.IREQ_NO = RQ.IREQ_NO" +
                           " LEFT JOIN ELLIPSE.MSF100 SC ON RQI.STOCK_CODE = SC.STOCK_CODE" +
                           " LEFT JOIN ELLIPSE.MSF110 PN ON RQI.STOCK_CODE = PN.STOCK_CODE" +
                           " WHERE" +
                           " " + paramValidOnly +
                           " " + paramReqNo +
                           " " + paramStockCode +
                           " " + paramDistrict +
                           " " + startDate +
                           " " + finishDate +
                           ")" +
                           "SELECT * FROM REQSC" +
                           " " + paramPreferedOnly;
            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchPurchaseOrderQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, string dateCriteria, string startDate, string finishDate, bool validOnly, bool preferedOnly)
        {
            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                paramDistrict = " AND PO.DSTRCT_CODE = '" + districtCode + "'";
            if (dateCriteria.Equals(SearchDateCriteriaType.Raised.Value))
            {
                if (!string.IsNullOrWhiteSpace(startDate))
                    startDate = " AND PO.CREATION_DATE >= " + startDate;
                if (!string.IsNullOrWhiteSpace(finishDate))
                    finishDate = " AND PO.CREATION_DATE <= " + finishDate;
            }

            var paramPurchaseOrder = "";
            var paramStockCode = "";
            if (searchCriteriaKey.Equals(SearchCriteriaType.ItemCode.Value))
            {
                paramPurchaseOrder = " PO.PO_NO = '" + searchCriteriaValue + "'";
            }
            else if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
            {
                paramStockCode = " POI.PREQ_STK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }
            else
            {
                paramStockCode = " POI.PREQ_STK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE POITEMS.PREF_PART_IND = POITEMS.MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           " WITH POITEMS AS(" +
                           "  SELECT" +
                           "    POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE, SC.ITEM_NAME, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, PN.PART_NO, PN.MNEMONIC, " +
                           "    POI.GROSS_PRICE_P, POI.UNIT_OF_PURCH, POI.CONV_FACTOR, " +
                           "    PN.PREF_PART_IND, MIN(PN.PREF_PART_IND) OVER (PARTITION BY POI.PREQ_STK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE ORDER BY POI.PREQ_STK_CODE, PN.PREF_PART_IND ASC) ROWPPI, " +
                           "    PO.STATUS_220, PO.CREATION_DATE, PO.ORDER_DATE, POI.ORIG_DUE_DATE, POI.ORIG_NET_PR_I, POI.CURR_NET_PR_I, POI.ORIG_QTY_I, POI.CURR_QTY_I, POI.QTY_RCV_OFST_I, POI.OFST_RCPT_DATE, POI.QTY_RCV_DIR_I, POI.ONST_RCPT_DATE, PO.FREIGHT_CODE, PO.DELIV_LOCATION, POI.EXPEDITE_CODE, PO.SUPPLIER_NO, SUP.SUPPLIER_NAME, PO.PO_MEDIUM_IND, PO.ORIGIN_CODE, PO.PURCH_OFFICER, PO.TEAM_ID" +
                           "  FROM" +
                           "    ELLIPSE.MSF220 PO JOIN ELLIPSE.MSF221 POI ON PO.PO_NO = POI.PO_NO LEFT JOIN ELLIPSE.MSF100 SC ON POI.PREQ_STK_CODE = SC.STOCK_CODE LEFT JOIN ELLIPSE.MSF110 PN ON POI.PREQ_STK_CODE = PN.STOCK_CODE LEFT JOIN ELLIPSE.MSF200 SUP ON PO.SUPPLIER_NO = SUP.SUPPLIER_NO" +
                           "  WHERE" +
                           " " + paramPurchaseOrder +
                           " " + paramStockCode +
                           " " + paramDistrict +
                           " " + startDate +
                           " " + finishDate +
                           " " + paramValidOnly +
                           "  ORDER BY POI.PO_ITEM_NO" +
                           "  )," +
                           " SCSTAT AS(" +
                           " SELECT STAT.DSTRCT_CODE, SC.STOCK_CODE, STAT.CREATION_DATE, STAT.LAST_MOD_DATE, SC.STK_DESC, SC.UNIT_OF_ISSUE, STAT.CLASS, STAT.RAF AS ALGORITMO, STAT.INVENT_COST_PR AS PRICE, STAT.HOME_WHOUSE AS BODEGA_PRINCIPAL, ELLIPSE.GET_SOH('" + districtCode + "',SC.STOCK_CODE) AS OWNED_SOH, ELLIPSE.GET_CONSIGN_SOH('" + districtCode + "', SC.STOCK_CODE) AS CONSIGN_SOH," +
                           "  STAT.IN_TRANSIT, STAT.DUES_IN, STAT.DUES_OUT, STAT.RESERVED, STAT.ROP, STAT.REORDER_QTY ROQ, STAT.EXP_ELEMENT AS DETALLE_GASTO, STAT.RESTRICT_RULE AS RESTR, STAT.DIRECT_ORDER_IND AS DO_IND, STAT.PURCH_OFFICER AS PURCHASER, " +
                           "  (SELECT SUM(UNSCHED_USAGE) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE " +
                           "  AND FULL_ACCT_PER BETWEEN (SELECT TO_CHAR(TO_DATE(MAX(FULL_ACCT_PER),'YYYYMM')-365,'YYYYMM') FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) AND (SELECT MAX(FULL_ACCT_PER) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) " +
                           "  AND TRIM(WHOUSE_ID) IS NOT NULL ) AS USO12_UNSCH, " +
                           "  (SELECT SUM(UNSCHED_USAGE) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE " +
                           "  AND FULL_ACCT_PER=(SELECT MAX(FULL_ACCT_PER) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) " +
                           "  AND TRIM(WHOUSE_ID) IS NOT NULL ) AS CURRENT_UNSCH, " +
                           "  STAT.INVT_CONTROLLR AS ADI FROM ELLIPSE.MSF100 SC LEFT JOIN ELLIPSE.MSF170 STAT ON SC.STOCK_CODE = STAT.STOCK_CODE)" +
                           "  SELECT * FROM POITEMS LEFT JOIN SCSTAT ON POITEMS.PREQ_STK_CODE = SCSTAT.STOCK_CODE AND SCSTAT.DSTRCT_CODE = 'ICOR'" +
                           " " + paramPreferedOnly;
            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }


}

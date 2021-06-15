using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Web.Services.Ellipse;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;

// ReSharper disable FieldCanBeMadeReadOnly.Local

namespace EllipseStockCodesExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

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
        private void btnFormatRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }
        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
            {
                //si ya hay un thread corriendo que no se ha detenido
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK || _thread != null && _thread.IsAlive) return;
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

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

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
                            sqlQuery = Queries.GetFetchInventoryStockCodeQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey, searchCriteriaValue, validOnly, preferedOnly);
                        else if (searchIssueKey.Equals(SearchCriteriaIssues.PurchaseOrder.Value))
                            sqlQuery = Queries.GetFetchPurchaseOrderQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey, searchCriteriaValue, searchDateCriteriaKey, startDate, endDate, validOnly, preferedOnly);
                        else if (searchIssueKey.Equals(SearchCriteriaIssues.Requisition.Value))
                            sqlQuery = Queries.GetFetchRequisitionQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey, searchCriteriaValue, searchDateCriteriaKey, startDate, endDate, validOnly, preferedOnly);
                        else if (searchIssueKey.Equals(SearchCriteriaIssues.RequisitionDetailed.Value))
                            sqlQuery = Queries.GetFetchRequisitioneDetailedQuery(_eFunctions.DbReference, _eFunctions.DbLink, district, searchCriteriaKey, searchCriteriaValue, searchDateCriteriaKey, startDate, endDate, validOnly, preferedOnly);
                        else
                        {
                            throw new Exception("Debe seleccionar una opción de búsqueda válida");
                        }

                        var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                        if (dataReader == null || dataReader.IsClosed)
                            return;

                        //Cargo el encabezado de la tabla y doy formato
                        if (rowParam == TitleRow01 + 1)
                        {
                            for (var k = 0; k < dataReader.FieldCount; k++)
                                cr.GetCell(k + 1, TitleRow02).Value2 = "'" + dataReader.GetName(k);

                            cr.GetCell(dataReader.FieldCount + 1, TitleRow02).Value2 = "Inventory Lead Time";
                            cr.GetCell(dataReader.FieldCount + 2, TitleRow02).Value2 = "Purchase Lead Time";
                            cr.GetCell(dataReader.FieldCount + 3, TitleRow02).Value2 = "Supplier Lead Time";
                            cr.GetCell(dataReader.FieldCount + 4, TitleRow02).Value2 = "Freight Lead Time";
                            cr.GetCell(dataReader.FieldCount + 5, TitleRow02).Value2 = "Total Lead Time";

                            _cells.FormatAsTable(cr.GetRange(1, TitleRow02, dataReader.FieldCount + 5, TitleRow02 + 1), TableName02);
                        }
                        //cargo los datos de cada consulta
                        while (dataReader.Read())
                        {
                            for (var k = 0; k < dataReader.FieldCount; k++)
                            {
                                cr.GetCell(k + 1, rowResult).Value2 = "'" + dataReader[k].ToString().Trim();
                            }

                            var stockCode = "";
                            if (searchIssueKey.Equals(SearchCriteriaIssues.Inventory.Value))
                                stockCode = dataReader["STOCK_CODE"].ToString();
                            else if (searchIssueKey.Equals(SearchCriteriaIssues.PurchaseOrder.Value))
                                stockCode = dataReader["PREQ_STK_CODE"].ToString();
                            else if (searchIssueKey.Equals(SearchCriteriaIssues.Requisition.Value))
                                stockCode = dataReader["STOCK_CODE"].ToString();
                            else if (searchIssueKey.Equals(SearchCriteriaIssues.RequisitionDetailed.Value))
                                stockCode = dataReader["STOCK_CODE"].ToString();
                            GetLeadTime(stockCode, cr, dataReader.FieldCount, rowResult);
                            rowResult++;
                        }
                        cp.GetCell(ResultColumn01, rowParam).Style = StyleConstants.Success;
                        cp.GetCell(2, rowParam).Value = "Consulta";
                        cp.GetCell(ResultColumn01, rowParam).Value = "OK";
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
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
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

        private void GetLeadTime(string searchCriteriaValue, ExcelStyleCells cr, int column, int rowResult)
        {
            var requestSheet = new ScreenSubmitRequestDTO();
            var proxySheet = new ScreenService();

            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = true,
                returnWarningsSpecified = true
            };

            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);

            var replySheet = proxySheet.executeScreen(opContext, "MSO179");

            if (replySheet.mapName != "MSM179A")
                throw new Exception("NO SE PUEDE INGRESAR AL PROGRAMA MSO179");

            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("STOCK_CODE1I", searchCriteriaValue.PadLeft(9, '0'));
            requestSheet.screenFields = arrayFields.ToArray();

            requestSheet.screenKey = "1";

            replySheet = proxySheet.submit(opContext, requestSheet);

            if (replySheet == null)
                throw new Exception("SE HA PRODUCIDO UN ERROR AL INTENTAR CREAR EL CÓDIGO " + searchCriteriaValue.PadLeft(9, '0'));
            if (_eFunctions.CheckReplyError(replySheet) || _eFunctions.CheckReplyWarning(replySheet))
                throw new Exception(replySheet.message);
            if (replySheet.mapName != "MSM179A")
                throw new Exception("NO SE HA PODIDO CONTINUAR CON EL SIGUIENTE PASO MSM179A");

            var arrayScreenNameValue2 = new ArrayScreenNameValue(replySheet.screenFields);
            cr.GetCell(column + 1, rowResult).Value2 = "'" + arrayScreenNameValue2.GetField("INV_LEAD_DAY_C1I").value;
            cr.GetCell(column + 2, rowResult).Value2 = "'" + arrayScreenNameValue2.GetField("PUR_LEAD_DAY_C1I").value;
            cr.GetCell(column + 3, rowResult).Value2 = "'" + arrayScreenNameValue2.GetField("LEAD_TIME_B1I").value;
            cr.GetCell(column + 4, rowResult).Value2 = "'" + arrayScreenNameValue2.GetField("LEAD_TIME_D1I").value;
            cr.GetCell(column + 5, rowResult).Value2 = "'" + arrayScreenNameValue2.GetField("LEAD_TIME_A1I").value;
        }
    }
    
}

using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCatalogueStockInstExcelAddIn.Properties;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Oracle.ManagedDataAccess.Client;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;
using WorksheetTools = Microsoft.Office.Tools.Excel.Worksheet;
using WorksheetInterop = Microsoft.Office.Interop.Excel.Worksheet;

namespace EllipseCatalogueStockInstExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "Catalog Stock INST";
        private const int TittleRow = 3;
        private const int ResultColumn = 10;
        private const int MaxRows = 10000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private OracleDataReader _drContractItems;
        private Application _excelApp;
        private WorksheetTools _worksheet;

        ListObject _excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviromentList = Environments.GetEnviromentList();
            foreach (var item in enviromentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpEnviroment.Items.Add(drpItem);
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            WorksheetInterop excelSheet = excelBook.ActiveSheet;

            excelSheet.Name = SheetName01;


            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
            
            _cells.GetCell(1,1).Value = "CERREJÓN";
            _cells.GetCell(2,1).Value = "CREATE STOCKCODES INST";

            _cells.GetCell(1,2).Value = "Contract No";
            _cells.GetCell(1, TittleRow).Value = "Contract No";
            _cells.GetCell(2, TittleRow).Value = "Portion No";
            _cells.GetCell(3, TittleRow).Value = "Element No";
            _cells.GetCell(4, TittleRow).Value = "Category No";
            _cells.GetCell(5, TittleRow).Value = "Stock Code";
            _cells.GetCell(6, TittleRow).Value = "Stock Description";
            _cells.GetCell(7, TittleRow).Value = "Home Warehouse";
            _cells.GetCell(8, TittleRow).Value = "Warehouse ID";
            _cells.GetCell(9, TittleRow).Value = "Option List";
            _cells.GetCell(ResultColumn, TittleRow).Value = "Result";

            WarehouseValidationList();

            ScreenOptionValidationList();

            _cells.GetRange(5, TittleRow + 1, 5, MaxRows).NumberFormat = "@";

            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.GetCell("A2").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.GetCell("B2").Style = _cells.GetStyle(StyleConstants.Select);
            _cells.GetRange(1, TittleRow, ResultColumn - 1, TittleRow).Style =
                _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell(ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();


            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            var contractRange = _worksheet.Controls.AddNamedRange(_worksheet.Range["B2"], "ContractRange");
            contractRange.Change += changesContractPrefixRange_Change;

            _cells.GetRange(1, TittleRow + 1, _excelSheetItems.ListRows.Count + TittleRow, MaxRows).NumberFormat = "@";
        }

        private void ScreenOptionValidationList()
        {
            var optionList = new List<string>
            {
                "1. Add a Stock Item to this District - Long Form",
                "6. Modify Operational Information",
                "7. Modify Warehousing Information",
                "C. Delete Warehousing Information"
            };
            _cells.SetValidationList(_cells.GetRange(9, TittleRow + 1, 9, MaxRows), optionList);
        }

        private void WarehouseValidationList()
        {
            var optionList = new List<string>
            {
                "INP - Instalaciones Puerto",
                "INM - Instalaciones Mina",
                "IPA - Aires Puerto",
                "IMA - Aires Mina"
            };
            _cells.SetValidationList(_cells.GetRange(7, TittleRow + 1, 7, MaxRows), optionList);
            _cells.SetValidationList(_cells.GetRange(8, TittleRow + 1, 8, MaxRows), optionList);
        }

        /// <summary>
        ///     Esta funcion se ejecuta despues de que se detecte un cambio en las celdas B2, luego de lo cual se consulta el
        ///     contrato.
        /// </summary>
        /// <param name="target">Rango que cambio</param>
        private void changesContractPrefixRange_Change(Range target)
        {
            GetContractData(target);
        }

        private void GetContractData(Range target)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            _cells.GetRange(1, TittleRow + 1, _excelSheetItems.ListRows.Count + TittleRow, MaxRows).Clear();
            _cells.GetRange(1, TittleRow + 1, _excelSheetItems.ListRows.Count + TittleRow, MaxRows).NumberFormat = "@";

            WarehouseValidationList();
            ScreenOptionValidationList();

            var contractNo = _cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row).Value);

            if (string.IsNullOrEmpty(contractNo)) return;
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var sqlQuery = Queries.GetContractData(contractNo, _eFunctions.dbReference, _eFunctions.dbLink);
            _drContractItems = _eFunctions.GetQueryResult(sqlQuery);

            if (_drContractItems != null && !_drContractItems.IsClosed && _drContractItems.HasRows)
            {
                var currentRow = TittleRow + 1;
                while (_drContractItems.Read())
                {
                    _cells.GetCell(1, currentRow).Select();
                    _cells.GetCell(1, currentRow).Value = _drContractItems["CONTRACT_NO"].ToString();
                    _cells.GetCell(2, currentRow).Value = _drContractItems["PORTION_NO"].ToString();
                    _cells.GetCell(3, currentRow).Value = _drContractItems["ELEMENT_NO"].ToString();
                    _cells.GetCell(4, currentRow).Value = _drContractItems["CATEGORY_NO"].ToString();
                    _cells.GetCell(5, currentRow).Value = _drContractItems["STOCK_CODE"].ToString();
                    _cells.GetCell(6, currentRow).Value = _drContractItems["CATEG_DESC"].ToString();
                    _cells.GetCell(7, currentRow).Value = _drContractItems["HOME_WHOUSE"].ToString();
                    _cells.GetCell(8, currentRow).Value = _drContractItems["WHOUSE_ID"].ToString();

                    currentRow++;
                }
            }
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnCatStockless_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                    ExecuteCatalog();
            }
            else
                MessageBox.Show(Resources.RibbonEllipse_btnCatStockless_Click_La_hoja_de_Excel_seleccionada_no_tiene_el_formato_válido_para_realizar_la_acción);
        }

        private void ExecuteCatalog()
        {
            var arrayFields = new ArrayScreenNameValue();
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);


            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();

            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

            var currentRow = TittleRow + 1;

            while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
            {
                try
                {
                    var stockCode = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                    var homeWhouse = (_cells.GetEmptyIfNull((_cells.GetCell(7, currentRow).Value)).Length >= 2) ? _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value).Substring(0, 3) : null;
                    var whouseId = (_cells.GetEmptyIfNull((_cells.GetCell(8, currentRow).Value)).Length >= 2) ? _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value).Substring(0, 3) : null;
                    string option1I = (_cells.GetEmptyIfNull((_cells.GetCell(9, currentRow).Value)).Length >= 1) ? _cells.GetEmptyIfNull(_cells.GetCell(9, currentRow).Value).Substring(0, 1) : null;

                    if (_cells.GetEmptyIfNull(_cells.GetCell(9, currentRow).Value) == "") return;
                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO170");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        switch (option1I)
                        {
                            case "1":
                                ScreenMSM170A_O1(stockCode, homeWhouse, ref arrayFields, opSheet, proxySheet,
                                    requestSheet, ref replySheet, currentRow);
                                break;
                            case "6":
                                ScreenMSM170A_O6(stockCode, homeWhouse, ref arrayFields, opSheet, proxySheet,
                                    requestSheet, ref replySheet, currentRow);
                                break;
                            case "7":
                                ScreenMSM170A_O7(stockCode, homeWhouse, ref arrayFields, opSheet, proxySheet,
                                    requestSheet, ref replySheet, currentRow);
                                break;
                            case "C":
                                ScreenMSM170A_OC(stockCode, whouseId, ref arrayFields, opSheet, proxySheet,
                                    requestSheet, ref replySheet, currentRow);
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                    //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void ScreenMSM170A_O6(string stockCode, string homeWhouse, ref ArrayScreenNameValue arrayFields,
            screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet,
            ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM170A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("OPTION1I", "6");
                arrayFields.Add("STOCK_CODE1I", stockCode);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm170C(homeWhouse, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet,
                        currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMSM170A_OC(string stockCode, string whouseId, ref ArrayScreenNameValue arrayFields,
            screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet,
            ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM170A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("OPTION1I", "C");
                arrayFields.Add("STOCK_CODE1I", stockCode);
                arrayFields.Add("WHOUSE_ID1I", whouseId);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    //Procedimiento de Borrar
                    arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("DELETE_CONF2I", "Y");

                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMSM170A_O7(string stockCode, string homeWhouse, ref ArrayScreenNameValue arrayFields,
            screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet,
            ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM170A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("OPTION1I", "7");
                arrayFields.Add("STOCK_CODE1I", stockCode);
                arrayFields.Add("WHOUSE_ID1I", homeWhouse);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    
                    ScreenMsm180B(ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMSM170A_O1(string stockCode, string homeWhouse, ref ArrayScreenNameValue arrayFields,
            screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet,
            ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM170A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("OPTION1I", "1");
                arrayFields.Add("STOCK_CODE1I", stockCode);
                arrayFields.Add("WHOUSE_ID1I", homeWhouse);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm170B(homeWhouse, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet,
                        currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm170B(string homeWhouse, ref ArrayScreenNameValue arrayFields,
            screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet,
            ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM170B") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("EXPEDITE_CODE2I", "OI");
                arrayFields.Add("ORIGIN_CODE2I", "G");
                arrayFields.Add("ROP2I", "0");
                arrayFields.Add("ROP_ROQ_UPDATE2I", "N");
                arrayFields.Add("REQ_DET_IN_EXT_SYS2I", "N");
                arrayFields.Add("REORDER_QTY2I", "0");
                arrayFields.Add("MIN_STOCK_LVL2I", "0");
                arrayFields.Add("MIN_STK_UPDATE2I", "N");
                arrayFields.Add("RAF2I", "00");
                arrayFields.Add("ALPHA2I", "0.10");
                arrayFields.Add("INV_REV_DAYS2I", "1");
                arrayFields.Add("PUR_REV_DAYS2I", "1");
                arrayFields.Add("PLANNED_T_OVER2I", "0.01");
                arrayFields.Add("PLANNED_SLEVEL2I", "50.00");
                arrayFields.Add("CONSIGN_WS_IND2I", "C");
                arrayFields.Add("INVENTORY_FLAG2I", "N");
                arrayFields.Add("DIRECT_ORDER2I", "N");
                arrayFields.Add("REC_PRO_ONLINE2I", "N");
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm170C(homeWhouse, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet,
                        currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm170C(string whouse, ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet,
            screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet,
            int currentRow)
        {
            if (replySheet.mapName != "MSM170C") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("INVENT_COSTING3I", "S");
                arrayFields.Add("EXP_ELEMENT3I", "521");
                arrayFields.Add("PRICE3I", "0.000000");
                arrayFields.Add("FREIGHT_CHGE3I", "N");
                arrayFields.Add("AVAIL3I", "0");
                arrayFields.Add("SOH3I", "0");
                arrayFields.Add("DUES_OUT3I", "0");
                arrayFields.Add("RESERVED3I", "0");
                arrayFields.Add("IN_TRANSIT3I", "0");
                arrayFields.Add("DUES_IN3I", "0");
                arrayFields.Add("ROQ3I", "0");
                arrayFields.Add("HOME_WHOUSE3I", whouse);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm17Da(ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm17Da(ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet,
            screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet,
            int currentRow)
        {
            if (replySheet.mapName != "MSM17DA") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm180B(ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm180B(ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet,
            screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet,
            int currentRow)
        {
            if (replySheet.mapName != "MSM180B") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm175A(ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm175A(ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet,
            screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet,
            int currentRow)
        {
            if (replySheet.mapName != "MSM175A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                requestSheet.screenFields = arrayFields.ToArray();
                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    ScreenMsm213A(ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void ScreenMsm213A(ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet,
            screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet,
            int currentRow)
        {
            if (replySheet.mapName != "MSM213A") return;
            try
            {
                arrayFields = new ArrayScreenNameValue();
                requestSheet.screenFields = arrayFields.ToArray();
                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", Ex.Message, _eFunctions.DebugErrors);
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    public static class Queries
    {
        public static string GetContractData(string contractNo, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT DISTINCT" +
                           "  CON.CONTRACT_NO, " +
                           "  CON.PORTION_NO, " +
                           "  CON.ELEMENT_NO, " +
                           "  CON.CATEGORY_NO, " +
                           "  PN.STOCK_CODE, " +
                           "  CON.CATEG_DESC, " +
                           "  CAT.HOME_WHOUSE, " +
                           "  WH.WHOUSE_ID " +
                           "FROM " +
                           "  ELLIPSE.MSF387 CON " +
                           "INNER JOIN " + dbReference + ".MSF110" + dbLink + " PN " +
                           "ON " +
                           "  PN.PART_NO LIKE CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '%' || CON.CONTRACT_NO || '%' " +
                           "LEFT JOIN " + dbReference + ".MSF170" + dbLink + " CAT " +
                           "ON " +
                           "  PN.STOCK_CODE = CAT.STOCK_CODE " +
                           "AND CAT.DSTRCT_CODE = 'INST' " +
                           "LEFT JOIN " + dbReference + ".MSF180" + dbLink + " WH " +
                           "ON " +
                           "    CAT.STOCK_CODE = WH.STOCK_CODE " +
                           "AND CAT.DSTRCT_CODE = WH.DSTRCT_CODE " +
                           "WHERE " +
                           "  CON.CONTRACT_NO = '" + contractNo + "' " +
                           "ORDER BY " +
                           "  CON.CONTRACT_NO, " +
                           "  CON.PORTION_NO, " +
                           "  CON.ELEMENT_NO, " +
                           "  CON.CATEGORY_NO ";
            return sqlQuery;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseMSO210ExcelAddin.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Oracle.ManagedDataAccess.Client;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseMSO210ExcelAddin
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "Purchase Information";
        private const int TittleRow = 3;
        private const int ResultColumn = 22;
        private const int MaxRows = 10000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private OracleDataReader _drContractItems;
        private Application _excelApp;
        ListObject _excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            var environmentList = Environments.GetEnvironmentList();
            foreach (var item in environmentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpEnvironment.Items.Add(drpItem);
            }

            drpEnvironment.SelectedItem.Label = Resources.RibbonEllipse_RibbonEllipse_Load_DefaultEnvironment;
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                Worksheet excelSheet = excelBook.ActiveSheet;

                _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

                Microsoft.Office.Tools.Excel.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                excelSheet.Name = SheetName01;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(2, 1).Value = "Purchase Information for StockCodes, Distrito de Instalaciones";

                _cells.GetCell(1, 2).Value = "Contract No";

                _cells.GetRange(1, 1, 8, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.GetRange(2, 1, 8, 1).Merge();
                _cells.GetCell(1, 2).Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell(2, 2).Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell(2, 2).NumberFormat = "@";

                _cells.GetCell(1, TittleRow).Value = "Supplier No";
                _cells.GetCell(2, TittleRow).Value = "Contract No";
                _cells.GetCell(3, TittleRow).Value = "Portion No";
                _cells.GetCell(4, TittleRow).Value = "Element No";
                _cells.GetCell(5, TittleRow).Value = "Category No";
                _cells.GetCell(6, TittleRow).Value = "Category Desc";
                _cells.GetCell(7, TittleRow).Value = "Stock Code";
                _cells.GetCell(8, TittleRow).Value = "Category Base Unit";
                _cells.GetCell(9, TittleRow).Value = "Category Base Rate";
                _cells.GetCell(10, TittleRow).Value = "Convertion Factor";
                _cells.GetCell(11, TittleRow).Value = "Standard Pack";
                _cells.GetCell(12, TittleRow).Value = "Lead Time Indicator";
                _cells.GetCell(13, TittleRow).Value = "Freight Code";
                _cells.GetCell(14, TittleRow).Value = "Delivery Location";
                _cells.GetCell(15, TittleRow).Value = "Currency Type";
                _cells.GetCell(16, TittleRow).Value = "Price Eff Date";
                _cells.GetCell(17, TittleRow).Value = "Gross Price";
                _cells.GetCell(18, TittleRow).Value = "Nuevo Precio";
                _cells.GetCell(19, TittleRow).Value = "Order Description";
                _cells.GetCell(20, TittleRow).Value = "Price Code";
                _cells.GetCell(21, TittleRow).Value = "Action";
                _cells.GetCell(ResultColumn, TittleRow).Value = "Result";

                _cells.GetRange(1, TittleRow, ResultColumn - 1, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, TittleRow + 1, _excelSheetItems.ListRows.Count + TittleRow, MaxRows).NumberFormat = "@";
                _cells.GetRange(9, TittleRow + 1, 9, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
                _cells.GetRange(17, TittleRow + 1, 18, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

                var contractRange = worksheet.Controls.AddNamedRange(worksheet.Range["B2"], "ContractRange");
                contractRange.Change += changesContractPrefixRange_Change;

                var optionList = new List<string>
                {
                    "1. Modify Preferred  Supplier Information",
                    "3. Modify Stock Code/Supplier Information",
                    "4. Delete Stock Code/Supplier Information"
                };

                _cells.SetValidationList(_cells.GetRange(21, TittleRow + 1, 21, _excelSheetItems.ListRows.Count + TittleRow), optionList);

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, TittleRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, TittleRow).Value = "ERROR:  " + ex.Message;
            }
        }

        private void changesContractPrefixRange_Change(Range tarGet)
        {
            GetContractData(tarGet);
        }

        private void GetContractData(Range tarGet)
        {
            _excelApp = Globals.ThisAddIn.Application;
            var excelBook = _excelApp.ActiveWorkbook;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            Worksheet excelSheet = excelBook.ActiveSheet;

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            
            var optionList = new List<string>
            {
                "1. Modify Preferred  Supplier Information",
                "3. Modify Stock Code/Supplier Information",
                "4. Delete Stock Code/Supplier Information"
            };

            _cells.SetValidationList(_cells.GetRange(21, TittleRow + 1, 21, MaxRows), optionList);

            var contractNo = _cells.GetNullOrTrimmedValue(_cells.GetCell(tarGet.Column, tarGet.Row).Value);

            if (string.IsNullOrEmpty(contractNo)) return;
            var sqlQuery = Queries.GetContractData(contractNo, _eFunctions.dbReference, _eFunctions.dbLink);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            _drContractItems = _eFunctions.GetQueryResult(sqlQuery);

            if (_drContractItems != null && !_drContractItems.IsClosed && _drContractItems.HasRows)
            {
                var currentRow = TittleRow + 1;
                while (_drContractItems.Read())
                {
                    _cells.GetCell(1, currentRow).Value = _drContractItems["SUPPLIER_NO"].ToString();
                    _cells.GetCell(2, currentRow).Value = _drContractItems["CONTRACT_NO"].ToString();
                    _cells.GetCell(3, currentRow).Value = _drContractItems["PORTION_NO"].ToString();
                    _cells.GetCell(4, currentRow).Value = _drContractItems["ELEMENT_NO"].ToString();
                    _cells.GetCell(5, currentRow).Value = _drContractItems["CATEGORY_NO"].ToString();
                    _cells.GetCell(6, currentRow).Value = _drContractItems["CATEG_DESC"].ToString();
                    _cells.GetCell(7, currentRow).Value = _drContractItems["STOCK_CODE"].ToString();
                    _cells.GetCell(8, currentRow).Value = _drContractItems["CATEG_BASE_UN"].ToString();
                    _cells.GetCell(9, currentRow).Value = _drContractItems["CATEG_BASE_RT"].ToString();
                    _cells.GetCell(10, currentRow).Value = _drContractItems["CONV_FACTOR"].ToString();
                    _cells.GetCell(11, currentRow).Value = _drContractItems["STD_PACK"].ToString();
                    _cells.GetCell(12, currentRow).Value = _drContractItems["SUPP_ACT_IND"].ToString();
                    _cells.GetCell(13, currentRow).Value = _drContractItems["FREIGHT_CODE"].ToString();
                    _cells.GetCell(14, currentRow).Value = _drContractItems["DELIV_LOCATION"].ToString();
                    _cells.GetCell(15, currentRow).Value = _drContractItems["CURRENCY_TYPE"].ToString();
                    _cells.GetCell(16, currentRow).Value = _drContractItems["PRICE_EFF_DATE"].ToString();
                    _cells.GetCell(17, currentRow).Value = _drContractItems["GROSS_PRICE_I"].ToString();
                    _cells.GetCell(18, currentRow).Value = _drContractItems["NUEVO_PRECIO"].ToString();
                    _cells.GetCell(19, currentRow).Value = _drContractItems["PO_DESC_IND"].ToString();
                    _cells.GetCell(20, currentRow).Value = _drContractItems["PRICE_CODE"].ToString();
                    _cells.GetCell(21, currentRow).Value = _drContractItems["ACTION"].ToString();
                    currentRow++;
                }
            }

            _cells.GetRange(9, TittleRow + 1, 9, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(17, TittleRow + 1, 18, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                    LoadSheet();
            }
            else
                MessageBox.Show(Resources.RibbonEllipse_btnLoad_Click_La_hoja_de_Excel_seleccionada_no_tiene_el_formato_válido_para_realizar_la_acción);
        }

        private void LoadSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;
            var excelBook = _excelApp.ActiveWorkbook;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            Worksheet excelSheet = excelBook.ActiveSheet;

            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
            _cells.GetRange(ResultColumn, TittleRow + 1, ResultColumn, MaxRows).ClearContents();

            var optionList = new List<string>
            {
                "1. Modify Preferred  Supplier Information",
                "3. Modify Stock Code/Supplier Information",
                "4. Delete Stock Code/Supplier Information"
            };

            _cells.SetValidationList(_cells.GetRange(21, TittleRow + 1, 21, MaxRows), optionList);


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

            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

            var currentRow = TittleRow + 1;

            while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
            {
                try
                {
                    _cells.GetCell(7, currentRow).Select();
                    var supplierNo = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                    var stockCode = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);
                    var categBaseUn = _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value);
                    var categBaseRt = _cells.GetEmptyIfNull(_cells.GetCell(9, currentRow).Value);
                    var convFactor = _cells.GetEmptyIfNull(_cells.GetCell(10, currentRow).Value);
                    var stdPack = _cells.GetEmptyIfNull(_cells.GetCell(11, currentRow).Value);
                    var suppActInd = _cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value);
                    var freightCode = _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value);
                    var delivLocation = _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value);
                    var currencyType = _cells.GetEmptyIfNull(_cells.GetCell(15, currentRow).Value);
                    var priceEffDate = _cells.GetEmptyIfNull(_cells.GetCell(16, currentRow).Value);
                    var grossPriceI = _cells.GetEmptyIfNull(_cells.GetCell(17, currentRow).Value);
                    var nuevoPrecio = _cells.GetEmptyIfNull(_cells.GetCell(18, currentRow).Value);
                    var poDescInd = _cells.GetEmptyIfNull(_cells.GetCell(19, currentRow).Value);
                    var priceCode = _cells.GetEmptyIfNull(_cells.GetCell(20, currentRow).Value);
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(21, currentRow).Value);

                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO210");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        ArrayScreenNameValue arrayFields;
                        switch (action)
                        {
                            case "1. Modify Preferred  Supplier Information":

                                if (replySheet.mapName == "MSM210A")
                                {
                                    try
                                    {
                                        //Opcion 1, selecciona la opcion de creacion de Supplier preferido.
                                        arrayFields = new ArrayScreenNameValue();
                                        arrayFields.Add("OPTION1I", "1");
                                        arrayFields.Add("DISTRICT_CODE1I", _frmAuth.EllipseDsct);
                                        arrayFields.Add("STOCK_CODE1I", stockCode);
                                        requestSheet.screenFields = arrayFields.ToArray();

                                        requestSheet.screenKey = "1";
                                        replySheet = proxySheet.submit(opSheet, requestSheet);

                                        while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                                            replySheet = proxySheet.submit(opSheet, requestSheet);

                                        if (_eFunctions.CheckReplyError(replySheet))
                                        {
                                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                                        }
                                        else
                                        {
                                            Msm210C(supplierNo, priceCode, categBaseUn, categBaseRt, convFactor, stdPack, suppActInd, freightCode, delivLocation, priceEffDate, grossPriceI, nuevoPrecio, currencyType, poDescInd, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                                        _cells.GetCell(ResultColumn, currentRow).Value = "ERROR:  " + ex.Message;
                                    }
                                }

                                break;
                            case "3. Modify Stock Code/Supplier Information":

                                if (replySheet.mapName == "MSM210A")
                                {
                                    try
                                    {
                                        //Opcion 1, selecciona la opcion de creacion de Supplier preferido.
                                        arrayFields = new ArrayScreenNameValue();
                                        arrayFields.Add("OPTION1I", "3");
                                        arrayFields.Add("DISTRICT_CODE1I", _frmAuth.EllipseDsct);
                                        arrayFields.Add("STOCK_CODE1I", stockCode);
                                        arrayFields.Add("SUPPLIER_NO1I", supplierNo);
                                        arrayFields.Add("PRICE_CODE1I", priceCode);
                                        requestSheet.screenFields = arrayFields.ToArray();

                                        requestSheet.screenKey = "1";
                                        replySheet = proxySheet.submit(opSheet, requestSheet);

                                        while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                                            replySheet = proxySheet.submit(opSheet, requestSheet);

                                        if (_eFunctions.CheckReplyError(replySheet))
                                        {
                                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                                        }
                                        else
                                        { Msm210B(categBaseUn, convFactor, stdPack, suppActInd, freightCode, delivLocation, priceEffDate, grossPriceI, nuevoPrecio, currencyType, poDescInd, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow); }
                                    }
                                    catch (Exception)
                                    {
                                        // ignored
                                    }
                                }
                                break;
                            case "4. Delete Stock Code/Supplier Information":
                                if (replySheet.mapName == "MSM210A")
                                {
                                    try
                                    {
                                        //Opcion 1, selecciona la opcion de creacion de Supplier preferido.
                                        arrayFields = new ArrayScreenNameValue();
                                        arrayFields.Add("OPTION1I", "4");
                                        arrayFields.Add("DISTRICT_CODE1I", _frmAuth.EllipseDsct);
                                        arrayFields.Add("STOCK_CODE1I", stockCode);
                                        arrayFields.Add("SUPPLIER_NO1I", supplierNo);
                                        arrayFields.Add("PRICE_CODE1I", priceCode);
                                        requestSheet.screenFields = arrayFields.ToArray();

                                        requestSheet.screenKey = "1";
                                        replySheet = proxySheet.submit(opSheet, requestSheet);

                                        while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                                            replySheet = proxySheet.submit(opSheet, requestSheet);

                                        if (_eFunctions.CheckReplyError(replySheet))
                                        {
                                            _cells.GetCell(ResultColumn, currentRow).Select();
                                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                                        }
                                        else
                                        {
                                            if (replySheet.mapName == "MSM210B")
                                            {
                                                try
                                                {
                                                    arrayFields = new ArrayScreenNameValue();
                                                    arrayFields.Add("ANSWER_FLD2I", "Y");

                                                    requestSheet.screenFields = arrayFields.ToArray();

                                                    requestSheet.screenKey = "1";
                                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                                    while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                                                        replySheet = proxySheet.submit(opSheet, requestSheet);

                                                    if (_eFunctions.CheckReplyError(replySheet))
                                                    {
                                                        _cells.GetCell(ResultColumn, currentRow).Select();
                                                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                                                    }
                                                    else
                                                    {
                                                        _cells.GetCell(ResultColumn, currentRow).Select();
                                                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                                                        _cells.GetCell(ResultColumn, currentRow).Value = "Supplier Information Deleted.";
                                                    }
                                                }
                                                catch (Exception)
                                                {
                                                    //ignored
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        // ignored
                                    }
                                }
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn, currentRow).Select();
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = "ERROR:  " + ex.Message;
                }
                finally { currentRow++; }
            }
        }

        private void Msm210C(string supplierNo, string priceCode, string categBaseUn, string categBaseRt, string convFactor, string stdPack, string suppActInd, string freightCode, string delivLocation, string priceEffDate, string grossPriceI, string nuevoPrecio, string currencyType, string poDescInd, ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM210C") return;
            try
            {
                //asigna el suplier preferido
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("SUPPLIER_NO3I", supplierNo);
                arrayFields.Add("PRICE_CODE3I", priceCode);
                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    //Si es la primera vez que se entra a crear el supplier preferido, envia a la pantalla 210B para complementar la informacion de terminos de compra.
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn, currentRow).Value = "Preferred Supplier Modified.";
                    Msm210B(categBaseUn, convFactor, stdPack, suppActInd, freightCode, delivLocation, priceEffDate, grossPriceI, nuevoPrecio, currencyType, poDescInd, ref arrayFields, opSheet, proxySheet, requestSheet, ref replySheet, currentRow);
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR:  " + ex.Message;
            }
        }

        private void Msm210B(string categBaseUn, string convFactor, string stdPack, string suppActInd, string freightCode, string delivLocation, string priceEffDate, string grossPriceI, string nuevoPrecio, string currencyType, string poDescInd, ref ArrayScreenNameValue arrayFields, screen.OperationContext opSheet, screen.ScreenService proxySheet, screen.ScreenSubmitRequestDTO requestSheet, ref screen.ScreenDTO replySheet, int currentRow)
        {
            if (replySheet.mapName != "MSM210B") return;
            try
            {
                //Informacion correspondiente a la pantalla de terminos de compra, que se toma de la consulta de SQL del contrato.
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("UNIT_OF_PURCH2I", categBaseUn);
                arrayFields.Add("CONV_FACTOR2I", convFactor);
                arrayFields.Add("STD_PACK2I", stdPack);
                arrayFields.Add("SUPP_ACT_IND2I", suppActInd);
                arrayFields.Add("FREIGHT_CODE2I", freightCode);
                arrayFields.Add("DELIV_LOCATION2I", delivLocation);
                arrayFields.Add("CURRENCY_TYPE2I", currencyType);
                arrayFields.Add("PRICE_EFF_DATE2I", priceEffDate);
                arrayFields.Add("GROSS_PRICE_P2I", nuevoPrecio);
                arrayFields.Add("PO_DESC_IND2I", poDescInd);

                requestSheet.screenFields = arrayFields.ToArray();

                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (replySheet.mapName == "MSM111A")
                {
                    arrayFields = new ArrayScreenNameValue();
                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);
                    if (replySheet.mapName == "MSM210B") replySheet = proxySheet.submit(opSheet, requestSheet);
                }


                while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn, currentRow).Value += "Supplier Parameters Updated";
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR:  " + ex.Message;
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
            var sqlQuery = "  " +
            "SELECT " +
            "  SUP.SUPPLIER_NO, " +
            "  CON.CONTRACT_NO, " +
            "  CON.PORTION_NO, " +
            "  CON.ELEMENT_NO, " +
            "  CON.CATEGORY_NO, " +
            "  CON.CATEG_DESC, " +
            "  PN.STOCK_CODE, " +
            "  SC.UNIT_OF_ISSUE CATEG_BASE_UN, " +
            "  CON.CATEG_BASE_PRC_RT CATEG_BASE_RT, " +
            "  DECODE ( SCS.CONV_FACTOR, NULL, '1', SCS.CONV_FACTOR ) CONV_FACTOR, " +
            "  DECODE ( SCS.STD_PACK, NULL, '1', SCS.STD_PACK ) STD_PACK, " +
            "  DECODE ( SCS.SUPP_ACT_IND, NULL, 'A', SCS.SUPP_ACT_IND ) SUPP_ACT_IND, " +
            "  DECODE ( SCS.FREIGHT_CODE, NULL, 'NA', SCS.FREIGHT_CODE ) FREIGHT_CODE, " +
            "  DECODE ( SCS.DELIV_LOCATION, NULL, 'AC', SCS.DELIV_LOCATION ) DELIV_LOCATION, " +
            "  DECODE ( SCS.CURRENCY_TYPE, NULL, 'USD', SCS.CURRENCY_TYPE ) CURRENCY_TYPE, " +
            "  DECODE ( SCS.PRICE_EFF_DATE, NULL, TO_CHAR ( SYSDATE, 'YYYYMMDD' ), SCS.PRICE_EFF_DATE ) PRICE_EFF_DATE, " +
            "  SCS.GROSS_PRICE_I, " +
            "  DECODE ( CON.CONTRACT_NO , '22232012', CON.CATEG_BASE_PRC_RT / 100, CON.CATEG_BASE_PRC_RT / 10 ) NUEVO_PRECIO," +
            "  DECODE ( SCS.PO_DESC_IND, NULL, 'N', SCS.PO_DESC_IND ) PO_DESC_IND, " +
            "  TRIM ( SCS.PRICE_CODE ) PRICE_CODE, " +
            "  DECODE ( SCS.STOCK_CODE, NULL, '1. Modify Preferred  Supplier Information', '3. Modify Stock Code/Supplier Information' ) ACTION " +
            "FROM " +
            "  ELLIPSE.MSF387 CON " +
            "INNER JOIN ELLIPSE.MSF384 SUP " +
            "ON " +
            "  CON.CONTRACT_NO = SUP.CONTRACT_NO " +
            "INNER JOIN ELLIPSE.MSF110 PN " +
            "ON " +
            "  PN.PART_NO = CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '-MI-' || CON.CONTRACT_NO " +
            "  OR PN.PART_NO = CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '-AA-' || CON.CONTRACT_NO " +
            "INNER JOIN ELLIPSE.MSF100 SC " +
            "ON " +
            "  PN.STOCK_CODE = SC.STOCK_CODE " +
            "LEFT JOIN ELLIPSE.MSF210 SCS " +
            "ON " +
            "  SCS.STOCK_CODE = PN.STOCK_CODE " +
            "AND SUP.SUPPLIER_NO = SCS.SUPPLIER_NO " +
            "AND SCS.DSTRCT_CODE = 'INST' " +
            "WHERE " +
            "  CON.CONTRACT_NO = '" + contractNo + "' " +
            "ORDER BY " +
            "  CON.CONTRACT_NO, " +
            "  CON.PORTION_NO, " +
            "  CON.ELEMENT_NO, " +
            "  CON.CATEGORY_NO";
            return sqlQuery;
        }
    }
}
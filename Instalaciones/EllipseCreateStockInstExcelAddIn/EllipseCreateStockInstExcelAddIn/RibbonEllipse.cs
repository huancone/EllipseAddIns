using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse.Connections;
using EllipseCreateStockInstExcelAddIn.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using CatProdService = EllipseCreateStockInstExcelAddIn.CatalogueProductService;
using CatService = EllipseCreateStockInstExcelAddIn.CatalogueService;
using WorksheetTools = Microsoft.Office.Tools.Excel.Worksheet;
using WorksheetInterop = Microsoft.Office.Interop.Excel.Worksheet;

namespace EllipseCreateStockInstExcelAddIn
{
    /// <summary>
    /// 
    /// </summary>
    public partial class RibbonEllipse
    {
        private const int ResultColumn = 5;
        private const string SheetName01 = "Create Stock INST";
        private const int TittleRow = 5;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private WorksheetTools _worksheet;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        /// <summary>
        ///     Establece la configuración inicial del AddIn
        /// </summary>
        public void LoadSettings()
        {
            var settings = new SharedClassLibrary.Ellipse.Settings();
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
        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        /// <summary>
        ///     Da formato a la hoja para cargar los items contractuales y los StockCodes asociados a estas.
        /// </summary>
        private void FormatSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            WorksheetInterop excelSheet = excelBook.ActiveSheet;

            excelSheet.Name = SheetName01;
            
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("B1").Value = "CREATE STOCKCODES INST";

            _cells.GetCell("A3").Value = "Contract No";
            _cells.GetCell("A4").Value = "Contract Prefix";
            _cells.GetCell("A5").Value = "Stock Code";
            _cells.GetCell("B5").Value = "Stock Description";
            _cells.GetCell("C5").Value = "Unit Of Issue";
            _cells.GetCell("D5").Value = "Part No";
            _cells.GetCell("E5").Value = "Result";

            //AA y MI fueron establecidos en la estrategia de Instalaciones como los prefijos para identificar los contratos
            var optionList = new List<string> {"AA", "MI"};
            _cells.SetValidationList(_cells.GetCell("B4"), optionList);

            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
            _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
            _cells.GetRange("A5", "D5").Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell("E5").Style = _cells.GetStyle(StyleConstants.TitleInformation);
            _cells.MergeCells("A1", "A2");


            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            var contractRange = _worksheet.Controls.AddNamedRange(_worksheet.Range["B3:B4"], "ContractRange");
            contractRange.Change += changesCotractPrefixRange_Change;

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        /// <summary>
        ///     Esta funciona se ejecuta despues de que se detecte un cambio en las celdas B3 y B4, luego de lo cual se consulta el
        ///     contrato.
        /// </summary>
        /// <param name="target">Rango que cambio</param>
        private void changesCotractPrefixRange_Change(Range target)
        {
            string contractNo;
            string contractPrefix;

            _excelApp = Globals.ThisAddIn.Application;
            var excelBook = _excelApp.ActiveWorkbook;
            WorksheetInterop excelSheet = excelBook.ActiveSheet;

            _cells.GetRange("A6", "D10000").Clear();
            _cells.GetRange("A6", "D10000").NumberFormat = "@";

            if (target.Row == 3)
            {
                contractNo = _cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row).Value);
                contractPrefix = _cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row + 1).Value);
            }
            else
            {
                contractNo = _cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row - 1).Value);
                contractPrefix = _cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row).Value);
            }

            if (!(!string.IsNullOrEmpty(contractNo) & !string.IsNullOrEmpty(contractPrefix))) return;
            var sqlQuery = Queries.GetContractData(contractNo, contractPrefix, _eFunctions.DbReference,
                _eFunctions.DbLink);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var drContractItems = _eFunctions.GetQueryResult(sqlQuery);

            if (drContractItems != null && !drContractItems.IsClosed)
            {
                var currentRow = TittleRow + 1;
                while (drContractItems.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = drContractItems["STOCK_CODE"].ToString();
                    _cells.GetCell("B" + currentRow).Value = drContractItems["DESCRIPCION"].ToString();
                    _cells.GetCell("C" + currentRow).Value = drContractItems["UNIT_OF_ISSUE"].ToString();
                    _cells.GetCell("D" + currentRow).Value = drContractItems["PART_NO"].ToString();
                    currentRow++;
                }
            }

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnCreateStock_Click(object sender, RibbonControlEventArgs e)
        {
            //Si la Hoja Activa es la que tiene el formato adecuado
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    CreateStock();
                }
            }
            else
                MessageBox.Show(Resources.sheet_format_error);
        }

        /// <summary>
        ///     Funcion para crear, editar y asociar al Parte Numero al Stock Code.
        /// </summary>
        private void CreateStock()
        {
            var catalogueServiceProxy = new CatService.CatalogueService();
            var catalogueOp = new CatService.OperationContext();
            catalogueServiceProxy.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CatalogueService";

            var productServiceProxy = new CatProdService.CatalogueProductService();
            var product = new CatProdService.CatalogueProductDTO();
            var productServiceOp = new CatProdService.OperationContext();
            productServiceProxy.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CatalogueProductService";

            _cells.GetRange("E6", "E10000").Clear();
            _cells.GetRange("A6", "E10000").NumberFormat = "@";


            var currentRow = TittleRow + 1;

            catalogueOp.district = _frmAuth.EllipseDsct;
            catalogueOp.position = _frmAuth.EllipsePost;
            catalogueOp.maxInstances = 100;
            catalogueOp.returnWarnings = Debugger.DebugWarnings;

            productServiceOp.district = _frmAuth.EllipseDsct;
            productServiceOp.position = _frmAuth.EllipsePost;
            productServiceOp.maxInstances = 100;
            productServiceOp.returnWarnings = Debugger.DebugWarnings;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while ( _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    var catalogue = new CatService.CatalogueDTO
                    {
                        stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(1, currentRow).Value),
                        isItemNameEnabled = true,
                        stockStatus = "D",
                        trackingIndicator = "N",
                        externallyManagedIndicator = false,
                        isExistedInventory = false,
                        itemName = "98008",
                        assetInd = false,
                        isExcludedFromRequirementDetermination = false,
                        isGloballyMaintained = false,
                        consumeAtSupplyCustomerAccount = false,
                        isEntitlementCheckingRequired = false,
                        isInspectionRequired = false,
                        unitOfIssue = _cells.GetNullOrTrimmedValue(_cells.GetCell(3, currentRow).Value),
                        batchLotManagementIndicator = false,
                        isRelifeAllowed = false,
                        isQualityDocumentationRequired = false,
                        isPackagingOrTreatmentRequired = false,
                        repairOrderFlag = false,
                        isAutoGenerateRepairOrder = false
                    };

                    var catalogueServiceReply = catalogueServiceProxy.create(catalogueOp, catalogue);

                    if (catalogueServiceReply.errors.Length > 0)
                    {
                        foreach (var t in catalogueServiceReply.errors)
                        {
                            _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText;
                        }
                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    }
                    else
                    {
                        _cells.GetCell(1, currentRow).Value = catalogueServiceReply.catalogueDTO.stockCode;

                        catalogue = new CatService.CatalogueDTO
                        {
                            stockCode = catalogueServiceReply.catalogueDTO.stockCode,
                            unitOfIssue = _cells.GetNullOrTrimmedValue(_cells.GetCell(3, currentRow).Value),
                            volume = Convert.ToDecimal(0),
                            weight = Convert.ToDecimal(0),
                            isExistedInventory = false,
                            externallyManagedIndicator = false,
                            productServiceCategoryId = "9999",
                            description = _cells.GetNullOrTrimmedValue(_cells.GetCell(2, currentRow).Value),
                            classification = "T",
                            stockType = "T",
                            assetInd = false,
                            isExcludedFromRequirementDetermination = false,
                            isGloballyMaintained = false,
                            consumeAtSupplyCustomerAccount = false,
                            trackingIndicator = "N",
                            isEntitlementCheckingRequired = false,
                            isInspectionRequired = false,
                            batchLotManagementIndicator = false,
                            isRelifeAllowed = false,
                            isQualityDocumentationRequired = false,
                            isPackagingOrTreatmentRequired = false,
                            repairOrderFlag = false,
                            isAutoGenerateRepairOrder = false
                        };


                        try
                        {
                            catalogueServiceReply = new CatService.CatalogueServiceResult();
                            catalogueServiceReply = catalogueServiceProxy.update(catalogueOp, catalogue);

                            //Si existe algun error en la actualizacion del Stock se muestran los mensajes de Error y se procede a Eliminar el Registro
                            if (catalogueServiceReply.errors.Length > 0)
                            {
                                foreach (var t in catalogueServiceReply.errors)
                                {
                                    _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText;
                                }
                                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;

                                catalogueServiceReply = DeleteStock(catalogueServiceProxy, catalogueServiceReply,
                                    catalogue, catalogueOp, currentRow);
                            }
                            else
                            {
                                try
                                {
                                    product.stockCode = catalogue.stockCode;
                                    product.partNumber =
                                        _cells.GetNullOrTrimmedValue(_cells.GetCell(4, currentRow).Value);
                                    product.manufacturerMnemonic = "MINSTAL";
                                    product.preferredPartIndicator = "01";
                                    product.partStatus1 = "V";

                                    productServiceProxy.create(productServiceOp, product);

                                    //Activar Stock
                                    try
                                    {
                                        catalogueServiceReply = new CatService.CatalogueServiceResult();
                                        catalogueServiceReply = catalogueServiceProxy.activate(catalogueOp, catalogue);

                                        if (catalogueServiceReply.errors.Length > 0)
                                        {
                                            foreach (var t in catalogueServiceReply.errors)
                                            {
                                                _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText;
                                            }
                                            _cells.GetCell(1, currentRow).Style = StyleConstants.Error;

                                            catalogueServiceReply = DeleteStock(catalogueServiceProxy,
                                                catalogueServiceReply, catalogue, catalogueOp, currentRow);
                                        }
                                        else
                                        {
                                            _cells.GetCell(1, currentRow).Style = StyleConstants.Success;
                                        }
                                    }
                                    catch (Exception errorActivate)
                                    {
                                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                                        _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " +
                                                                                         errorActivate.Message;
                                        //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", errorActivate.Message,_eFunctions.DebugErrors);

                                        catalogueServiceReply = DeleteStock(catalogueServiceProxy, catalogueServiceReply,
                                            catalogue, catalogueOp, currentRow);
                                    }
                                }
                                catch (Exception errorCreatePn)
                                {
                                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + errorCreatePn.Message;
                                    //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", errorCreatePn.Message,_eFunctions.DebugErrors);

                                    catalogueServiceReply = DeleteStock(catalogueServiceProxy, catalogueServiceReply,
                                        catalogue, catalogueOp, currentRow);
                                }
                            }
                        }
                        catch (Exception errorUpdate)
                        {
                            _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + errorUpdate.Message;
                            //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", errorUpdate.Message,_eFunctions.DebugErrors);

                            DeleteStock(catalogueServiceProxy, catalogueServiceReply, catalogue, catalogueOp, currentRow);
                        }
                    }
                }
                catch (Exception errorCreate)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + errorCreate.Message;
                    //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", errorCreate.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private CatService.CatalogueServiceResult DeleteStock(CatService.CatalogueService catalogueServiceProxy,
            CatService.CatalogueServiceResult catalogueServiceReply, CatService.CatalogueDTO catalogue,
            CatService.OperationContext catalogueOp, int currentRow)
        {
            try
            {
                catalogueServiceReply = new CatService.CatalogueServiceResult();
                catalogueServiceReply = catalogueServiceProxy.delete(catalogueOp, catalogue);
            }
            catch (Exception errordelete)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + errordelete.Message;
                //ErrorLogger.LogError("RibbonEllipse.cs:CreateStock()", errordelete.Message, _eFunctions.DebugErrors);
            }
            return catalogueServiceReply;
        }

        private void drpEnvironment_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    /// <summary>
    /// Consultas SQL
    /// </summary>
    internal static class Queries
    {
        /// <summary>
        /// Trae los items contractuales del contracNo de la tabla msf387 y calcula el parte numero recomendado para la creacion del Stock.
        /// </summary>
        /// <param name="contractNo"></param>
        /// <param name="contractPrefix"></param>
        /// <param name="dbReference"></param>
        /// <param name="dbLink"></param>
        /// <returns></returns>
        public static string GetContractData(string contractNo, string contractPrefix, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT " +
                           "  CAT.STOCK_CODE, " +
                           "  CON.CATEG_DESC DESCRIPCION, " +
                           "  CON.CATEG_BASE_UN UNIT_OF_ISSUE, " +
                           "  CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '-" + contractPrefix +
                           "-' || CON.CONTRACT_NO PART_NO " +
                           "FROM " +
                           "  " + dbReference + ".MSF387" + dbLink + " CON " +
                           "LEFT JOIN ELLIPSE.MSF110 PN " +
                           "ON " +
                           "  CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '-" + contractPrefix +
                           "-' || CON.CONTRACT_NO = PN.PART_NO " +
                           "AND PN.STATUS_CODES = 'V' " +
                           "LEFT JOIN ELLIPSE.MSF100 CAT " +
                           "ON " +
                           "  PN.STOCK_CODE = CAT.STOCK_CODE " +
                           "WHERE " +
                           "  CON.CONTRACT_NO = '" + contractNo + "' " +
                           "AND CATEG_CODE = 'TRFA'";
            return sqlQuery;
        }
    }
}
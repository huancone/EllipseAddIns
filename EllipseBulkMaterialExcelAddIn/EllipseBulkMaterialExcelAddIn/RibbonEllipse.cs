using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseBulkMaterialExcelAddIn.Properties;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Connections;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using BMUSheet = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;
using BMUSheetItem = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;
using EllipseEquipmentClassLibrary;
using ListService = EllipseEquipmentClassLibrary.EquipmentListService;
using System.Threading;

namespace EllipseBulkMaterialExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "BulkMaterialSheet";
        private const string SheetName02 = "EquipmentsLists";
        private const string TableName01 = "ExcelSheetItems";
        private const string TableName02 = "FuelListItems";
        private const int TitleRow01 = 7;
        private const int TitleRow02 = 7;
        private const int ResultColumn01 = 18;
        private const int ResultColumn02 = 20;
        private const int MaxRows = 5000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;

        private List<string> _optionList;
        private const string ValidationSheetName = "ValidationListSheet";
        private Thread _thread;

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

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(BulkMaterialExcecute);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:BulkMaterialExcecute()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnImport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ImportFile);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ImportFile()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnBulkMaterialFormatMultiple_Click(object sender, RibbonControlEventArgs e)
        {
            BulkMaterialFormatMultiple();
        }

        private void BulkMaterialFormatMultiple()
        {
            try
            {
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).Clear();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = "@";


                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "Bulk Material Usage Sheet";

                _cells.GetRange("A1", "B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.GetRange("B1", "D1").Merge();

                _cells.GetCell(1, TitleRow01).Value = "Usage Sheet Id";
                _cells.GetCell(2, TitleRow01).Value = "District";
                _cells.GetCell(3, TitleRow01).Value = "Warehouse";
                _cells.GetCell(4, TitleRow01).Value = "Usage Date";
                _cells.GetCell(5, TitleRow01).Value = "Usage Time";
                _cells.GetCell(6, TitleRow01).Value = "General Account Code";

                _cells.GetCell(7, TitleRow01).Value = "Usage Item Id";

                _cells.GetCell(8, TitleRow01).Value = "Equipment Reference";
                _cells.GetCell(9, TitleRow01).Value = "Component Code";
                _cells.GetCell(10, TitleRow01).Value = "Modifier Code";
                _cells.GetCell(11, TitleRow01).Value = "Bulk Material Type";
                _cells.GetCell(12, TitleRow01).Value = "Condition Monitoring Action";
                _cells.GetCell(13, TitleRow01).Value = "Quantity";
                _cells.GetCell(14, TitleRow01).Value = "Transaction Date";
                _cells.GetCell(15, TitleRow01).Value = "Statistic Time";
                _cells.GetCell(16, TitleRow01).Value = "Statistic Type";
                _cells.GetCell(17, TitleRow01).Value = "Statistic Meter";
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "Result";

                #region Styles

                _cells.GetCell(1, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(2, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(6, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(8, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(10, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(11, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(12, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(13, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(14, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(15, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(16, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(17, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                #endregion

                #region Instructions

                _cells.GetCell("E1").Value = "OBLIGATORIO";
                _cells.GetCell("E1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("E2").Value = "OPCIONAL";
                _cells.GetCell("E2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("E3").Value = "INFORMATIVO";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("E4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("E5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("E5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                #endregion

                _optionList = new List<string>
                {
                    "    Fuel/Diesel",
                    "B - Condition Monitoring Fitment",
                    "L - Condition Monitoring Rebuild in Situ",
                    "O - Oil Changed",
                    "C - Condition Monitoring Defitment",
                    "A - Oil Added",
                    "F - Filter Changed"
                };

                _cells.SetValidationList(_cells.GetCell(12, TitleRow01 + 1), _optionList, ValidationSheetName, 1);
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).NumberFormat = "@";
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                OrderAndSort(_excelApp.ActiveWorkbook.ActiveSheet);

                //Hoja 2
                #region Hoja de Listas
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "EQUIPMENT LIST CHECKER - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = EquipListSearchFieldCriteria.ListType.Value;
                _cells.GetCell("B3").Value = "PCOMBU";
                _cells.GetCell("A4").Value = EquipListSearchFieldCriteria.ListId.Value;
                _cells.SetValidationList(_cells.GetCell("B4"), GetListIdList("PCOMBU"), ValidationSheetName, 2);
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                var statusCodeList = _eFunctions.GetItemCodes("ES").Select(item => item.code + " - " + item.description).ToList();
                var equipClassCodeList = _eFunctions.GetItemCodes("EC").Select(item => item.code + " - " + item.description).ToList();
                var equipTypeCodeList = _eFunctions.GetItemCodes("ET").Select(item => item.code + " - " + item.description).ToList();
                var compCodeList = _eFunctions.GetItemCodes("CO").Select(item => item.code + " - " + item.description).ToList();
                var mnemonicCodeList = _eFunctions.GetItemCodes("AA").Select(item => item.code + " - " + item.description).ToList();
                var classTypeCodeList = _eFunctions.GetItemCodes("E0").Select(item => item.code + " - " + item.description).ToList();
                var fuelTypeCodeList = _eFunctions.GetItemCodes("E2").Select(item => item.code + " - " + item.description).ToList();

                _cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02).Style = StyleConstants.TitleInformation;

                _cells.GetCell(1, TitleRow02).Value = "Equipment Number";
                _cells.GetCell(1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow02).Value = "Description 1";
                _cells.GetCell(3, TitleRow02).Value = "Description 2";
                _cells.GetCell(4, TitleRow02).Value = "Status";
                _cells.GetCell(5, TitleRow02).Value = "List Type";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, TitleRow02).Value = "List Id";
                _cells.GetCell(6, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, TitleRow02).Value = "Equipment Class";
                _cells.GetCell(8, TitleRow02).Value = "Equipment Type";
                _cells.GetCell(9, TitleRow02).Value = "EGI";
                _cells.GetCell(10, TitleRow02).Value = "Serial Number";
                _cells.GetCell(11, TitleRow02).Value = "Operator Id/Pos";
                _cells.GetCell(12, TitleRow02).Value = "Input By";
                _cells.GetCell(13, TitleRow02).Value = "Account Code";
                _cells.GetCell(14, TitleRow02).Value = "Component Code";
                _cells.GetCell(15, TitleRow02).Value = "Mnemonic";
                _cells.GetCell(16, TitleRow02).Value = "Stock Code";
                _cells.GetCell(17, TitleRow02).Value = "Part Number";
                _cells.GetCell(18, TitleRow02).Value = "E0. Class Type";
                _cells.GetCell(19, TitleRow02).Value = "E2. Fuel Type";

                _cells.SetValidationList(_cells.GetCell(4, TitleRow02 + 1), statusCodeList, ValidationSheetName, 3);
                _cells.SetValidationList(_cells.GetCell(7, TitleRow02 + 1), equipClassCodeList, ValidationSheetName, 4);
                _cells.SetValidationList(_cells.GetCell(8, TitleRow02 + 1), equipTypeCodeList, ValidationSheetName, 5);
                _cells.SetValidationList(_cells.GetCell(14, TitleRow02 + 1), compCodeList, ValidationSheetName, 6);
                _cells.SetValidationList(_cells.GetCell(15, TitleRow02 + 1), mnemonicCodeList, ValidationSheetName, 7);
                _cells.SetValidationList(_cells.GetCell(18, TitleRow02 + 1), classTypeCodeList, ValidationSheetName, 8);
                _cells.SetValidationList(_cells.GetCell(19, TitleRow02 + 1), fuelTypeCodeList, ValidationSheetName, 9);


                _cells.GetCell(20, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(20, TitleRow02).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);

                ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                #endregion
                ((Worksheet)_excelApp.ActiveWorkbook.Sheets[1]).Select(Type.Missing);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void OrderAndSort(Worksheet excelSheet)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            var tableSheetItems = _cells.GetRange(TableName01).ListObject;
            tableSheetItems.Sort.SortFields.Clear();
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            tableSheetItems.Sort.Apply();
        }

        private void ImportFile()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();
            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearContents();
            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).NumberFormat = "@";

            var openFileDialog1 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"Seleccione un archivo de Texto",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\\"
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog1.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var bulkMaterials = cc.Read<BulkMaterial>(filePath, inputFileDescription);

            var currentRow = TitleRow01 + 1;
            foreach (var bulkMaterial in bulkMaterials)
            {
                try
                {
                    _cells.GetCell(3, currentRow).Value = bulkMaterial.WarehouseId;
                    _cells.GetCell(4, currentRow).Value = DateTime.ParseExact(bulkMaterial.DefaultUsageDate, @"MM/dd/yy", CultureInfo.CurrentCulture).ToString("yyyyMMdd");
                    _cells.GetCell(8, currentRow).Value = bulkMaterial.EquipmentReference;
                    _cells.GetCell(11, currentRow).Value = bulkMaterial.BulkMaterialTypeId;
                    _cells.GetCell(13, currentRow).Value = bulkMaterial.Quantity;
                }
                catch (Exception error)
                {
                    _cells.GetCell(ResultColumn01, currentRow).Value = "Error: " + error.Message;
                }
                finally { currentRow++; }
            }

            OrderAndSort(excelSheet);
        }

        /// <summary>
        ///     Crea las instancias a los servicios BulkMaterialUsageSheetService y BulkMaterialUsageSheetItemService
        /// </summary>
        private void BulkMaterialExcecute()
        {
            try
            {
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.ActiveWorkbook;
                Worksheet excelSheet = excelBook.ActiveSheet;

                if (excelSheet.Name != SheetName01) return;
                var proxySheet = new BMUSheet.BulkMaterialUsageSheetService();
                var opSheet = new BMUSheet.OperationContext();

                var proxyItem = new BMUSheetItem.BulkMaterialUsageSheetItemService();
                var opItem = new BMUSheetItem.OperationContext();


                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals("")) return;
                proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/BulkMaterialUsageSheet";
                proxyItem.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/BulkMaterialUsageSheetItem";
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                opSheet.district = _frmAuth.EllipseDsct;
                opSheet.maxInstances = 100;
                opSheet.position = _frmAuth.EllipsePost;
                opSheet.returnWarnings = false;

                opItem.district = _frmAuth.EllipseDsct;
                opItem.maxInstances = 100;
                opItem.position = _frmAuth.EllipsePost;
                opItem.returnWarnings = false;

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                try
                {
                    var tableSheetItems = _cells.GetRange(TableName01).ListObject;
                    tableSheetItems.Sort.SortFields.Clear();
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.Apply();

                    var currentRow = TitleRow01 + 1;

                    while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value)) != null)
                    {
                        DateTime usageDate;
                        if (DateTime.TryParseExact(_cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value), "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, out usageDate))
                        {

                            var currentHeader = currentRow;
                            var requestSheet = new BMUSheet.BulkMaterialUsageSheetDTO();
                            var requestItemList = new List<BMUSheetItem.BulkMaterialUsageSheetItemDTO>();
                            var allRequestItemList = new List<BMUSheetItem.BulkMaterialUsageSheetItemDTO>();

                            _cells.GetCell(1, currentRow).Select();

                            //llenado de variables del encabezado de la hoja
                            requestSheet.bulkMaterialUsageSheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null ? _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) : null;
                            requestSheet.districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) ?? "ICOR";
                            requestSheet.warehouseId = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);


                            requestSheet.defaultUsageDate = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                            requestSheet.defaultAccountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value) != null ? _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value) : null;
                            requestSheet.defaultAccountCode = requestSheet.defaultAccountCode ?? GetBulkAccountCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                            //Crea el encabezado
                            var replySheet = proxySheet.create(opSheet, requestSheet);


                            //valida si el encabezado tiene errores
                            if (replySheet.errors.Length > 0)
                            {
                                foreach (var t in replySheet.errors)
                                    _cells.GetCell(ResultColumn01, currentRow).Value += " - " + t.messageText;

                                _cells.GetRange(1, currentHeader, 6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                                _cells.GetRange(1, currentHeader, 6, currentRow).Select();
                                currentRow++;
                            }
                            else
                            {
                                //si el encabezado no tiene errores empueza a agregar los items a la coleccion.
                                requestSheet.bulkMaterialUsageSheetId = replySheet.bulkMaterialUsageSheetDTO.bulkMaterialUsageSheetId;
                                _cells.GetCell(1, currentRow).Value = replySheet.bulkMaterialUsageSheetDTO.bulkMaterialUsageSheetId;

                                //mientras que el encabezado sea el mismo, llene la lista de items
                                var sheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                                var warehouseId = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                                var defaultUsageDate = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                                var defaultAccountCode = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value) != null ? _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value) : null);
                                defaultAccountCode = defaultAccountCode ?? GetBulkAccountCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));

                                while (
                                        (
                                            requestSheet.bulkMaterialUsageSheetId == sheetId ||
                                            (
                                                sheetId == null &&
                                                requestSheet.warehouseId == warehouseId &&
                                                requestSheet.defaultUsageDate == defaultUsageDate &&
                                                requestSheet.defaultAccountCode == defaultAccountCode
                                            )
                                        )
                                      )
                                {
                                    ItemListAdd(currentRow, requestSheet, requestItemList, allRequestItemList, excelSheet);
                                    currentRow++;

                                    sheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                                    warehouseId = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                                    defaultUsageDate = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                                    defaultAccountCode = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value) != null ? _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value) : null);
                                    defaultAccountCode = defaultAccountCode ?? GetBulkAccountCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                                }

                                try
                                {
                                    if (requestItemList.Count > 0)
                                    {
                                        //esta operacion agrega la lista de items al encabezado
                                        var replyItem = proxyItem.multipleCreate(opItem, requestItemList.ToArray());

                                        //recorre el resultado de la ejecucion de la operacion multipleCreate donde hubo errores.
                                        var errorCounter = 0;

                                        foreach (var rItem in replyItem.Where(rItem => rItem.errors.Length > 0))
                                        {
                                            errorCounter++;
                                            var errorMessage = rItem.errors.Aggregate("", (current, error) => current + (error.messageText + ", "));

                                            var currentItem = 0;
                                            foreach (var item in allRequestItemList)
                                            {
                                                if (_cells.GetEmptyIfNull(item.bulkMaterialUsageSheetId).ToUpper() == _cells.GetEmptyIfNull(rItem.bulkMaterialUsageSheetItemDTO.bulkMaterialUsageSheetId) &
                                                    _cells.GetEmptyIfNull(item.bulkMaterialUsageSheetItemId).ToUpper() == _cells.GetEmptyIfNull(rItem.bulkMaterialUsageSheetItemDTO.bulkMaterialUsageSheetItemId)
                                                    )
                                                {
                                                    requestItemList.Remove(item);
                                                    _cells.GetRange(8, currentHeader + currentItem, 13, currentHeader + currentItem).Style = _cells.GetStyle(StyleConstants.Error);
                                                    _cells.GetCell(ResultColumn01, currentHeader + currentItem).Value += errorMessage;
                                                    _cells.GetCell(ResultColumn01, currentHeader + currentItem).Select();
                                                }
                                                currentItem++;
                                            }
                                        }

                                        if (errorCounter > 0 & requestItemList.Count > 0)
                                        {
                                            try
                                            {
                                                var deleteHeader = false;
                                                replyItem = proxyItem.multipleCreate(opItem, requestItemList.ToArray());
                                                foreach (var rItem in replyItem.Where(item => item.errors.Length > 0)) { deleteHeader = true; }
                                                if (deleteHeader)
                                                {
                                                    DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow - 1);
                                                }
                                                else
                                                {
                                                    ApplyHeader(proxySheet, opSheet, requestSheet, currentRow - 1, currentHeader);
                                                }
                                            }
                                            catch (Exception error)
                                            {
                                                MessageBox.Show(error.Message);
                                                DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow - 1);
                                            }
                                        }
                                        else if (errorCounter == 0 & requestItemList.Count > 0)
                                        {
                                            ApplyHeader(proxySheet, opSheet, requestSheet, currentRow - 1, currentHeader);
                                        }
                                        else
                                        {
                                            DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow - 1);
                                        }
                                    }
                                    else
                                    {
                                        _cells.GetCell(ResultColumn01, currentRow - 1).Value += "No hay Items para Aplicar en esta hoja!";
                                        DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow - 1);
                                    }
                                }
                                catch (Exception error)
                                {
                                    MessageBox.Show(error.Message);
                                }
                            }
                        }
                        else
                        {
                            _cells.GetCell(4, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                            _cells.GetCell(ResultColumn01, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                            _cells.GetCell(ResultColumn01, currentRow).Value += "Fecha Errada";
                            currentRow++;
                        }
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void ApplyHeader(BMUSheet.BulkMaterialUsageSheetService proxySheet, BMUSheet.OperationContext opSheet, BMUSheet.BulkMaterialUsageSheetDTO requestSheet, int currentRow, int currentHeader)
        {
            try
            {
                var replySheet = proxySheet.apply(opSheet, requestSheet);
                if (replySheet.errors.Length > 0)
                {
                    foreach (var t in replySheet.errors)
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Value += " - " + t.messageText;
                    }
                    _cells.GetRange(1, currentHeader, ResultColumn01 - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow);
                }
                else
                {
                    _cells.GetRange(1, currentHeader, ResultColumn01 - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Success); _cells.GetRange(1, currentHeader, 6, currentRow).Select();
                }
            }
            catch (Exception)
            {
                DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow);
            }
        }

        private void ItemListAdd(int currentRow, BMUSheet.BulkMaterialUsageSheetDTO requestSheet, List<BMUSheetItem.BulkMaterialUsageSheetItemDTO> requestItemList, List<BMUSheetItem.BulkMaterialUsageSheetItemDTO> allRequestItemList, Worksheet excelSheet)
        {
            _cells.GetCell(1, currentRow).Select();
            _cells.GetCell(1, currentRow).Value = requestSheet.bulkMaterialUsageSheetId;

            var requestItem = new BMUSheetItem.BulkMaterialUsageSheetItemDTO
            {
                bulkMaterialUsageSheetId = requestSheet.bulkMaterialUsageSheetId,
                equipmentReference = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                componentCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                modifier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                bulkMaterialTypeId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                conditionMonitoringAction = (_cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value) == "Fuel/Diesel") || (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value) == null) ? null : _cells.GetCell(12, currentRow).Value.ToString().Substring(0, 1),
                quantity = decimal.Round(Convert.ToDecimal(_cells.GetCell(13, currentRow).Value)),
                quantitySpecified = (_cells.GetNullIfTrimmedEmpty(_cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value)) != null),
                usageDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                usageTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                operationStatisticType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                meterReading = Convert.ToDecimal(_cells.GetCell(17, currentRow).Value),
                meterReadingSpecified = (_cells.GetNullIfTrimmedEmpty(_cells.GetEmptyIfNull(_cells.GetCell(17, currentRow).Value)) != null),
            };

            //consulta la base de datos y obtiene la capacidad maxima de combustible a cargar al equipo, si no tiene coloca 0.
            try
            {
                _cells.GetCell(8, currentRow).Select();

                allRequestItemList.Add(requestItem);

                var profile = GetFuelCapacity(requestItem.equipmentReference, requestItem.bulkMaterialTypeId);

                if (requestItem.bulkMaterialTypeId == profile.FuelType && requestItem.quantity > profile.Capacity)
                {
                    _cells.GetCell(ResultColumn01, currentRow).Value = "Este valor supera la capacidad del Equipo!";
                    _cells.GetRange(8, currentRow, 13, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }
                else
                {
                    //agrega el item a la coleccion
                    requestItemList.Add(requestItem);
                    _cells.GetRange(8, currentRow, 13, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                }

            }
            catch (Exception error)
            {
                _cells.GetCell(ResultColumn01, currentRow).Value = error.Message;
                _cells.GetCell(13, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell(13, currentRow).Select();
            }
        }

        private void DeleteHeader(BMUSheet.BulkMaterialUsageSheetService proxySheet, BMUSheet.OperationContext opSheet, BMUSheet.BulkMaterialUsageSheetDTO requestSheet, int currentHeader, int currentRow)
        {
            try
            {
                var replySheet = proxySheet.delete(opSheet, requestSheet);

                if (replySheet.errors.Length > 0)
                {
                    foreach (var t in replySheet.errors)
                    {
                        _cells.GetCell(ResultColumn01, (currentHeader + t.fieldIndex)).Value += " - " + t.messageText;
                    }
                }
                else
                {
                    _cells.GetCell(ResultColumn01, currentRow).Value += " - Hoja " + replySheet.bulkMaterialUsageSheetDTO.bulkMaterialUsageSheetId + " Borrada";
                    _cells.GetRange(1, currentHeader, ResultColumn01 - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }
            }
            catch (Exception err)
            {
                _cells.GetCell(ResultColumn01, currentRow).Value += err.Message;
            }
        }

        private string GetBulkAccountCode(string equipNo)
        {
            try
            {
                if (string.IsNullOrEmpty(equipNo)) return "";

                var sqlQuery = Queries.GetBulkAccountCode(equipNo, _eFunctions.dbReference, _eFunctions.dbLink);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var drEquipCapacity = _eFunctions.GetQueryResult(sqlQuery);

                if (!drEquipCapacity.Read()) return "";

                if (!drEquipCapacity.IsClosed && drEquipCapacity.HasRows)
                {
                    return drEquipCapacity["BULK_ACCOUNT"].ToString();
                }
                else
                    return "";
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                _eFunctions.CloseConnection();
            }
        }

        private Profile GetFuelCapacity(string equipNo, string fuelType)
        {
            try
            {
                var profile = new Profile();

                if (string.IsNullOrEmpty(equipNo))
                {
                    Profile.Error = "Defina un Equipo";
                    return profile;
                }

                var sqlQuery = Queries.GetFuelCapacity(equipNo, _eFunctions.dbReference, _eFunctions.dbLink);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var drEquipCapacity = _eFunctions.GetQueryResult(sqlQuery);

                if (!drEquipCapacity.Read())
                {
                    Profile.Error = "No Tiene Perfil";

                    return profile;
                }

                if (!drEquipCapacity.IsClosed && drEquipCapacity.HasRows)
                {
                    profile.Equipo = drEquipCapacity["EQUIP_NO"].ToString();
                    profile.Egi = drEquipCapacity["EQUIP_GRP_ID"].ToString();
                    profile.FuelType = drEquipCapacity["FUEL_OIL_TYPE"].ToString();
                    profile.Capacity = Convert.ToDecimal(drEquipCapacity["FUEL_CAPACITY"].ToString());
                    return profile;
                }
                else
                {
                    Profile.Error = "No Tiene Perfil";
                    return profile;
                }
            }
            finally
            {
                _eFunctions.CloseConnection();
            }
        }

        private void btnUnApplyDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(Unapply);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:Unapply()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void Unapply()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;
            var proxySheet = new BMUSheet.BulkMaterialUsageSheetService();
            var opSheet = new BMUSheet.OperationContext();


            if (drpEnviroment.Label == null || drpEnviroment.Label.Equals("")) return;
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/BulkMaterialUsageSheet";
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;

            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            opSheet.district = _frmAuth.EllipseDsct;
            opSheet.maxInstances = 100;
            opSheet.position = _frmAuth.EllipsePost;
            opSheet.returnWarnings = false;


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TitleRow01 + 1;

            while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value)) != null)
            {
                var requestSheet = new BMUSheet.BulkMaterialUsageSheetDTO();
                _cells.GetCell(1, currentRow).Select();

                try
                {
                    requestSheet.bulkMaterialUsageSheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);

                    var replySheet = proxySheet.unapply(opSheet, requestSheet);

                    if (replySheet.errors.Length > 0)
                    {
                        foreach (var t in replySheet.errors) { _cells.GetCell(ResultColumn01, currentRow).Value += " - " + t.messageText; }

                        _cells.GetRange(1, currentRow, 6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        _cells.GetRange(1, currentRow, 6, currentRow).Select();
                    }
                    else
                    {
                        _cells.GetRange(1, currentRow, 6, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                        _cells.GetRange(1, currentRow, 6, currentRow).Select();
                        DeleteHeader(proxySheet, opSheet, requestSheet, currentRow, currentRow);

                    }
                }
                catch (Exception error)
                {
                    _cells.GetRange(1, currentRow, 6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(ResultColumn01, currentRow).Value = error.Message;
                    _cells.GetCell(ResultColumn01, currentRow).Select();
                }
                finally { currentRow++; }
            }
        }
        private void btnValidateStats_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ValidateStats);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidateStats()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private Stats GetLastStatistic(string equipNo, string statType, string statDate)
        {
            try
            {

                var stats = new Stats();
                if (string.IsNullOrEmpty(equipNo) || string.IsNullOrEmpty(statType)) stats.Error = "Error";

                var sqlQuery = Queries.GetLastStatistic(equipNo, statType, statDate, _eFunctions.dbReference, _eFunctions.dbLink);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var drLastStat = _eFunctions.GetQueryResult(sqlQuery);

                if (!drLastStat.Read()) stats.Error = "Error";

                if (!drLastStat.IsClosed && drLastStat.HasRows)
                {
                    stats.MeterValue = Convert.ToDecimal(drLastStat["METER_VALUE"].ToString());
                    stats.EquipNo = drLastStat["EQUIP_NO"].ToString();
                    stats.StatType = drLastStat["STAT_TYPE"].ToString();
                    stats.StatDate = drLastStat["STAT_DATE"].ToString();
                }
                else
                    stats.Error = "Error";

                return stats;
            }
            finally { _eFunctions.CloseConnection(); }
        }

        private void ValidateStats()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;

            if (drpEnviroment.Label == null || drpEnviroment.Label.Equals("")) return;

            var currentRow = TitleRow01 + 1;
            while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value)) != null)
            {
                var statType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
                var equipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                var stat = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
                var statDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);

                if (equipNo != null & statType != null & stat != null)
                {
                    var lastStat = GetLastStatistic(equipNo, statType, statDate);
                    _cells.GetCell(17, currentRow).AddComment(Convert.ToString(lastStat.StatDate + " - " + lastStat.MeterValue, CultureInfo.InvariantCulture));
                    _cells.GetCell(17, currentRow).Style = _cells.GetStyle(Convert.ToDecimal(stat) < lastStat.MeterValue ? StyleConstants.Error : StyleConstants.Success);
                }
                currentRow++;
            }
        }

        private List<string> GetListIdList(string listType)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var sqlQuery = Queries.GetListIdList(_eFunctions.dbReference, _eFunctions.dbLink, listType);
            var drItem = _eFunctions.GetQueryResult(sqlQuery);

            var list = new List<string>();

            if (drItem == null || drItem.IsClosed || !drItem.HasRows) return list;

            while (drItem.Read())
            {
                list.Add("" + drItem["LIST_ID"].ToString().Trim());
            }

            return list;
        }

        private static class Queries
        {
            public static string GetBulkAccountCode(string equipNo, string dbReference, string dbLink)
            {
                var query = "" +
                    "WITH " +
                    "  REFERENCE AS " +
                    "  ( " +
                    "    SELECT " +
                    "      RC.REF_NO, " +
                    "      RC.SCREEN_LITERAL, " +
                    "      RCD.ENTITY_VALUE EQUIP_NO, " +
                    "      RCD.REF_CODE BULK_ACCOUNT, " +
                    "      RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM FECHA, " +
                    "      MAX ( RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM ) OVER ( PARTITION BY RCD.REF_NO, RCD.ENTITY_VALUE ) MAX_FECHA " +
                    "    FROM " +
                    "      " + dbReference + ".MSF071" + dbLink + " RCD " +
                    "    INNER JOIN " + dbReference + ".MSF070" + dbLink + " RC " +
                    "    ON " +
                    "      RCD.ENTITY_TYPE = RC.ENTITY_TYPE " +
                    "    AND RC.REF_NO = RCD.REF_NO " +
                    "    WHERE " +
                    "      RCD.ENTITY_TYPE = 'EQP' " +
                    "    AND RCD.REF_NO = '003' " +
                    "  ) " +
                    "SELECT " +
                    "  EQUIP_NO, " +
                    "  BULK_ACCOUNT " +
                    "FROM " +
                    "  REFERENCE " +
                    "WHERE " +
                    "  FECHA = MAX_FECHA " +
                    "AND EQUIP_NO = '" + equipNo + "'";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFuelCapacity(string equipNo, string dbReference, string dbLink)
            {
                var query = "" +
                    "WITH   " +
                    "  EQUIPO AS   " +
                    "  (   " +
                    "    SELECT   " +
                    "      EQ.EQUIP_NO   " +
                    "    FROM   " +
                    "      " + dbReference + ".MSF600" + dbLink + " EQ   " +
                    "    WHERE   " +
                    "      EQ.EQUIP_NO = '" + equipNo + "'   " +
                    "  )   " +
                    "  ,   " +
                    "  BASE AS   " +
                    "  (   " +
                    "    SELECT   " +
                    "      1 PESO,   " +
                    "      PROFILES.EQUIP_GRP_ID,   " +
                    "      PROFILES.FUEL_OIL_TYPE,   " +
                    "      PROFILES.FUEL_CAPACITY   " +
                    "    FROM   " +
                    "      " + dbReference + ".MSF617_GENERAL" + dbLink + "  PROFILES   " +
                    "    WHERE   " +
                    "      PROFILES.EGI_REC_TYPE = 'E'   " +
                    "    AND TRIM ( PROFILES.FUEL_OIL_TYPE ) IS NOT NULL   " +
                    "    UNION ALL   " +
                    "    SELECT   " +
                    "      0 PESO,   " +
                    "      PROFILES.EQUIP_GRP_ID,   " +
                    "      PROFILES.FUEL_OIL_TYPE,   " +
                    "      PROFILES.FUEL_CAPACITY   " +
                    "    FROM   " +
                    "      " + dbReference + ".MSF617_GENERAL" + dbLink + "  PROFILES   " +
                    "    WHERE   " +
                    "      PROFILES.EGI_REC_TYPE = 'G'   " +
                    "    AND TRIM ( PROFILES.FUEL_OIL_TYPE ) IS NOT NULL   " +
                    "  )   " +
                    "  ,   " +
                    "  EQUIPOS AS   " +
                    "  (   " +
                    "    SELECT   " +
                    "      BASE.PESO,   " +
                    "      EQ.EQUIP_NO,   " +
                    "      EQ.EQUIP_GRP_ID,   " +
                    "      BASE.FUEL_OIL_TYPE,   " +
                    "      BASE.FUEL_CAPACITY   " +
                    "    FROM   " +
                    "      " + dbReference + ".MSF600" + dbLink + "  EQ   " +
                    "    LEFT JOIN BASE   " +
                    "    ON   " +
                    "      EQ.EQUIP_NO = BASE.EQUIP_GRP_ID   " +
                    "    OR EQ.EQUIP_GRP_ID = BASE.EQUIP_GRP_ID   " +
                    "    WHERE   " +
                    "      EQ.DSTRCT_CODE = 'ICOR'   " +
                    "  )   " +
                    "  ,   " +
                    "  PROFILES AS   " +
                    "  (   " +
                    "    SELECT   " +
                    "      EQUIPOS.PESO,   " +
                    "      MAX ( EQUIPOS.PESO ) OVER ( PARTITION BY EQUIPOS.EQUIP_NO, EQUIPOS.EQUIP_GRP_ID ) MAX_PESO,   " +
                    "      EQUIPOS.EQUIP_NO,   " +
                    "      EQUIPOS.EQUIP_GRP_ID,   " +
                    "      EQUIPOS.FUEL_OIL_TYPE,   " +
                    "      EQUIPOS.FUEL_CAPACITY   " +
                    "    FROM   " +
                    "      EQUIPOS   " +
                    "  )   " +
                    "SELECT   " +
                    "  EQUIPO.EQUIP_NO,   " +
                    "  DECODE ( PROFILES.EQUIP_GRP_ID, NULL, 'NO TIENE', TRIM(PROFILES.EQUIP_GRP_ID) ) EQUIP_GRP_ID,   " +
                    "  DECODE ( PROFILES.FUEL_OIL_TYPE, NULL, 'NO TIENE', TRIM(PROFILES.FUEL_OIL_TYPE) ) FUEL_OIL_TYPE,   " +
                    "  DECODE ( PROFILES.FUEL_CAPACITY, NULL, 0, PROFILES.FUEL_CAPACITY ) FUEL_CAPACITY   " +
                    "FROM   " +
                    "  EQUIPO   " +
                    "LEFT JOIN PROFILES   " +
                    "ON   " +
                    "  EQUIPO.EQUIP_NO = PROFILES.EQUIP_NO   " +
                    "AND PROFILES.PESO = PROFILES.MAX_PESO   " +
                    "ORDER BY   " +
                    "  PROFILES.PESO   ";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetLastStatistic(string equipNo, string statType, string statDate, string dbReference, string dbLink)
            {
                var query = "" +
                    "SELECT " +
                    "  STAT.EQUIP_NO, " +
                    "  STAT.ITEM_NAME_1, " +
                    "  STAT.METER_VALUE, " +
                    "  STAT.STAT_TYPE, " +
                    "  STAT.STAT_DATE " +
                    "FROM " +
                    "  ( " +
                    "    SELECT " +
                    "      EQ.EQUIP_NO, " +
                    "      EQ.ITEM_NAME_1, " +
                    "      STAT.STAT_TYPE, " +
                    "      STAT.STAT_DATE, " +
                    "      STAT.STAT_VALUE, " +
                    "      STAT.CUM_VALUE, " +
                    "      STAT.METER_VALUE, " +
                    "      MAX ( STAT.STAT_DATE ) OVER ( PARTITION BY STAT.EQUIP_NO, STAT.STAT_TYPE ) MAX_DATE, " +
                    "      EQ.DSTRCT_CODE " +
                    "    FROM " +
                    "      " + dbReference + ".MSF600" + dbLink + " EQ " +
                    "    LEFT JOIN " + dbReference + ".MSF400" + dbLink + " STAT " +
                    "    ON " +
                    "      EQ.EQUIP_NO = STAT.EQUIP_NO " +
                    "    WHERE " +
                    "      EQ.EQUIP_NO = '" + equipNo + "' " +
                    "    AND EQ.DSTRCT_CODE = 'ICOR' " +
                    "    AND STAT.STAT_TYPE = '" + statType + "' " +
                    "    AND STAT_DATE <= '" + statDate + "' " +
                    "  ) " +
                    "  STAT " +
                    "WHERE " +
                    "  STAT.MAX_DATE = STAT.STAT_DATE " +
                    "OR STAT.STAT_DATE IS NULL ";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetListIdList(string dbReference, string dbLink, string listType)
            {
                var query = "" +
                            "SELECT EQL.LIST_TYP, EQL.LIST_ID FROM " + dbReference + ".MSF606" + dbLink + " EQL " +
                            "WHERE EQL.LIST_TYP = '" + listType + "'";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }

        private class BulkMaterial
        {
            [CsvColumn(FieldIndex = 1)]
            public string WarehouseId { get; set; }

            [CsvColumn(FieldIndex = 2, OutputFormat = "yyyyMMdd")]
            public string DefaultUsageDate { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string UserId { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string EquipmentReference { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string BulkMaterialTypeId { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string Quantity { get; set; }
        }

        private class Profile
        {
            public string Equipo { get; set; }
            public string Egi { get; set; }
            public string FuelType { get; set; }
            public decimal Capacity { get; set; }
            public static string Error { get; set; }
        }

        private class Stats
        {
            public string EquipNo { get; set; }
            public string StatType { get; set; }
            public decimal MeterValue { get; set; }
            public string StatDate { get; set; }

            public string Error { get; set; }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnReviewEquipList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewListEquipmentsList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewListEquipmentsList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ReviewListEquipmentsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName02);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            //Obtengo los valores de las opciones de búsqueda
            var searchCriteriaKey1 = EquipListSearchFieldCriteria.ListType.Key;
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey2 = EquipListSearchFieldCriteria.ListId.Key;
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var previousEquipment = new Equipment { EquipmentNo = "" };

            var listeq = ListActions.FetchListEquipmentsList(_eFunctions, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, null);
            var i = TitleRow02 + 1;
            foreach (var eql in listeq)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                    _cells.GetCell(1, i).Value = "'" + eql.EquipNo;
                    _cells.GetCell(5, i).Value = "'" + eql.ListType;
                    _cells.GetCell(6, i).Value = "'" + eql.ListId;

                    var eq = eql.EquipNo.Trim().Equals(previousEquipment.EquipmentNo.Trim()) ? previousEquipment : EquipmentActions.FetchEquipmentData(_eFunctions, eql.EquipNo);

                    _cells.GetCell(2, i).Value = "'" + eq.EquipmentNoDescription1;
                    _cells.GetCell(3, i).Value = "'" + eq.EquipmentNoDescription2;
                    _cells.GetCell(4, i).Value = "'" + eq.EquipmentStatus;
                    _cells.GetCell(7, i).Value = "'" + eq.EquipmentClass;
                    _cells.GetCell(8, i).Value = "'" + eq.EquipmentType;
                    _cells.GetCell(9, i).Value = "'" + eq.EquipmentGrpId;
                    _cells.GetCell(10, i).Value = "'" + eq.SerialNumber;
                    _cells.GetCell(11, i).Value = "'" + eq.OperatorId + "/" + eq.OperatorPosition;
                    _cells.GetCell(12, i).Value = "'" + eq.InputBy;
                    _cells.GetCell(13, i).Value = "'" + eq.AccountCode;
                    _cells.GetCell(14, i).Value = "'" + eq.CompCode;
                    _cells.GetCell(15, i).Value = "'" + eq.Mnemonic;
                    _cells.GetCell(16, i).Value = "'" + eq.StockCode;
                    _cells.GetCell(17, i).Value = "'" + eq.PartNo;
                    _cells.GetCell(18, i).Value = "'" + eq.ClassCodes.EquipmentClassif0;
                    _cells.GetCell(19, i).Value = "'" + eq.ClassCodes.EquipmentClassif2;

                    previousEquipment = eq;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewListEquipmentsList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();

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

        private void btnReviewFromBulkSheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01 || ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewFromEquipmentList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewListEquipmentsList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void ReviewFromEquipmentList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var celleq = new ExcelStyleCells(_excelApp, SheetName01);
            var cellli = new ExcelStyleCells(_excelApp, SheetName02);
            _cells.SetCursorWait();
            cellli.ClearTableRange(TableName02);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);


            var k = TitleRow01 + 1;
            var i = TitleRow02 + 1;
            while (!string.IsNullOrEmpty("" + celleq.GetCell(8, k).Value))
            {
                var equipmentNo = _cells.GetEmptyIfNull(celleq.GetCell(8, k).Value);
                try
                {
                    var eq = EquipmentActions.FetchEquipmentData(_eFunctions, equipmentNo);
                    var listeq = ListActions.FetchListEquipmentsList(_eFunctions, equipmentNo);

                    if (listeq != null && listeq.Count > 0)
                    {
                        foreach (var eql in listeq)
                        {
                            try
                            {
                                //Para resetear el estilo
                                cellli.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                                cellli.GetCell(1, i).Value = "'" + eq.EquipmentNo;
                                cellli.GetCell(2, i).Value = "'" + eq.EquipmentNoDescription1;
                                cellli.GetCell(3, i).Value = "'" + eq.EquipmentNoDescription2;
                                cellli.GetCell(4, i).Value = "'" + eq.EquipmentStatus;
                                cellli.GetCell(5, i).Value = "'" + eql.ListType;
                                cellli.GetCell(6, i).Value = "'" + eql.ListId;
                                cellli.GetCell(7, i).Value = "'" + eq.EquipmentClass;
                                cellli.GetCell(8, i).Value = "'" + eq.EquipmentType;
                                cellli.GetCell(9, i).Value = "'" + eq.EquipmentGrpId;
                                cellli.GetCell(10, i).Value = "'" + eq.SerialNumber;
                                cellli.GetCell(11, i).Value = "'" + eq.OperatorId + "/" + eq.OperatorPosition;
                                cellli.GetCell(12, i).Value = "'" + eq.InputBy;
                                cellli.GetCell(13, i).Value = "'" + eq.AccountCode;
                                cellli.GetCell(14, i).Value = "'" + eq.CompCode;
                                cellli.GetCell(15, i).Value = "'" + eq.Mnemonic;
                                cellli.GetCell(16, i).Value = "'" + eq.StockCode;
                                cellli.GetCell(17, i).Value = "'" + eq.PartNo;
                                cellli.GetCell(18, i).Value = "'" + eq.ClassCodes.EquipmentClassif0;
                                cellli.GetCell(19, i).Value = "'" + eq.ClassCodes.EquipmentClassif2;
                            }
                            catch (Exception ex)
                            {
                                cellli.GetCell(1, i).Style = StyleConstants.Error;
                                cellli.GetCell(ResultColumn02, i).Value = "ERRORLIST: " + ex.Message;
                                Debugger.LogError("RibbonEllipse.cs:ReviewFromEquipmentList()", ex.Message);
                            }
                            finally
                            {
                                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                                    cellli.GetCell(2, i).Select();
                                i++;
                            }
                        }
                    }
                    else
                    {
                        //Para resetear el estilo
                        cellli.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                        cellli.GetCell(1, i).Value = "'" + eq.EquipmentNo;
                        cellli.GetCell(1, i).Style = StyleConstants.Warning;
                        cellli.GetCell(2, i).Value = "'" + eq.EquipmentNoDescription1;
                        cellli.GetCell(3, i).Value = "'" + eq.EquipmentNoDescription2;
                        cellli.GetCell(4, i).Value = "'" + eq.EquipmentStatus;
                        cellli.GetCell(5, i).Value = "'" + "-";
                        cellli.GetCell(6, i).Value = "'" + "-";
                        cellli.GetCell(7, i).Value = "'" + eq.EquipmentClass;
                        cellli.GetCell(8, i).Value = "'" + eq.EquipmentType;
                        cellli.GetCell(9, i).Value = "'" + eq.EquipmentGrpId;
                        cellli.GetCell(10, i).Value = "'" + eq.SerialNumber;
                        cellli.GetCell(11, i).Value = "'" + eq.OperatorId + "/" + eq.OperatorPosition;
                        cellli.GetCell(12, i).Value = "'" + eq.InputBy;
                        cellli.GetCell(13, i).Value = "'" + eq.AccountCode;
                        cellli.GetCell(14, i).Value = "'" + eq.CompCode;
                        cellli.GetCell(15, i).Value = "'" + eq.Mnemonic;
                        cellli.GetCell(16, i).Value = "'" + eq.StockCode;
                        cellli.GetCell(17, i).Value = "'" + eq.PartNo;
                        cellli.GetCell(18, i).Value = "'" + eq.ClassCodes.EquipmentClassif0;
                        cellli.GetCell(19, i).Value = "'" + eq.ClassCodes.EquipmentClassif2;
                        cellli.GetCell(ResultColumn02, i).Value = "Equipo no existe en ninguna lista ";
                        cellli.GetCell(ResultColumn02, i).Style = StyleConstants.Warning;

                        if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                            cellli.GetCell(2, i).Select();
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, k).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, k).Value = "ERRORLIST: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewFromEquipmentList()", ex.Message);
                }
                finally
                {
                    if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                        celleq.GetCell(1, k).Select();
                    k++;
                }
            }

            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();

        }

        private void btnAddToList_Click(object sender, RibbonControlEventArgs e)
        {
            if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(AddListEquipmentsList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void AddListEquipmentsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var i = TitleRow02 + 1;

            var opSheet = new ListService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    var equiplist = new EquipListItem()
                    {
                        EquipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                        ListType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                        ListId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                    };

                    ListActions.AddEquipmentToList(opSheet, urlService, equiplist);

                    _cells.GetCell(ResultColumn02, i).Value = "AGREGADO A LA LISTA";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn02, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:AddListEquipmentsList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                }
            }
            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnRemoveFromList_Click(object sender, RibbonControlEventArgs e)
        {
            if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeleteListEquipmentsList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void DeleteListEquipmentsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var i = TitleRow02 + 1;

            var opSheet = new ListService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    var equiplist = new EquipListItem()
                    {
                        EquipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                        ListType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                        ListId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                    };

                    ListActions.DeleteEquipmentFromList(opSheet, urlService, equiplist);

                    _cells.GetCell(ResultColumn02, i).Value = "ELIMINADO DE LISTA";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn02, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:DeleteListEquipmentsList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                }
            }
            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
    }
}
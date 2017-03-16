using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseBulkMaterialExcelAddIn.Properties;
using EllipseCommonsClassLibrary;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using BMUSheet = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;
using BMUSheetItem = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;

namespace EllipseBulkMaterialExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "BulkMaterialSheet";
        private const string SheetName02 = "BulkMaterialSheetErrors";
        private const int TittleRow = 7;
        private const int ResultColumn = 18;
        private const int MaxRows = 5000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private List<string> _optionList;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            _eFunctions.DebugErrors = false;
            _eFunctions.DebugQueries = false;
            _eFunctions.DebugWarnings = false;

            var enviromentList = EnviromentConstants.GetEnviromentList();
            foreach (var item in enviromentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpBulkMaterialEnv.Items.Add(drpItem);
            }

            drpBulkMaterialEnv.SelectedItem.Label = Resources.RibbonEllipse_RibbonEllipse_Load_Productivo;
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            BulkMaterialExcecute();
        }

        private void btnImport_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFile();
        }

        private void btnBulkMaterialFormatMultiple_Click(object sender, RibbonControlEventArgs e)
        {
            BulkMaterialFormatMultiple();
        }

        private void BulkMaterialFormatMultiple()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                excelSheet.Name = SheetName01;

                _cells = new ExcelStyleCells(_excelApp);

                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Clear();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, TittleRow + 1).NumberFormat = "@";


                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "Bulk Material Usage Sheet";

                _cells.GetRange("A1", "B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.GetRange("B1", "D1").Merge();

                _cells.GetCell(1, TittleRow).Value = "Usage Sheet Id";
                _cells.GetCell(2, TittleRow).Value = "District";
                _cells.GetCell(3, TittleRow).Value = "Warehouse";
                _cells.GetCell(4, TittleRow).Value = "Usage Date";
                _cells.GetCell(5, TittleRow).Value = "Usage Time";
                _cells.GetCell(6, TittleRow).Value = "General Account Code";

                _cells.GetCell(7, TittleRow).Value = "Usage Item Id";

                _cells.GetCell(8, TittleRow).Value = "Equipment Reference";
                _cells.GetCell(9, TittleRow).Value = "Component Code";
                _cells.GetCell(10, TittleRow).Value = "Modifier Code";
                _cells.GetCell(11, TittleRow).Value = "Bulk Material Type";
                _cells.GetCell(12, TittleRow).Value = "Condition Monitoring Action";
                _cells.GetCell(13, TittleRow).Value = "Quantity";
                _cells.GetCell(14, TittleRow).Value = "Transaction Date";
                _cells.GetCell(15, TittleRow).Value = "Statistic Time";
                _cells.GetCell(16, TittleRow).Value = "Statistic Type";
                _cells.GetCell(17, TittleRow).Value = "Statistic Meter";
                _cells.GetCell(ResultColumn, TittleRow).Value = "Result";

                #region Styles

                _cells.GetCell(1, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(2, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(6, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(8, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(10, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(11, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(12, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(13, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(14, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(15, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(16, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(17, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

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
                _cells.SetValidationList(_cells.GetCell(12, TittleRow + 1), _optionList);

                _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                    _cells.GetRange(1, TittleRow, ResultColumn, MaxRows), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                _excelSheetItems.Name = "ExcelSheetItems";

                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";

                OrderAndSort(excelSheet);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void OrderAndSort(Worksheet excelSheet)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            _excelSheetItems.Sort.SortFields.Clear();
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
            _excelSheetItems.Sort.Apply();
        }

        private void ImportFile()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;

            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearContents();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";

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

            var currentRow = TittleRow + 1;
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
                    _cells.GetCell(ResultColumn, currentRow).Value = "Error: " + error.Message;
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
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.ActiveWorkbook;
                Worksheet excelSheet = excelBook.ActiveSheet;

                if (excelSheet.Name != SheetName01) return;
                var proxySheet = new BMUSheet.BulkMaterialUsageSheetService();
                var opSheet = new BMUSheet.OperationContext();

                var proxyItem = new BMUSheetItem.BulkMaterialUsageSheetItemService();
                var opItem = new BMUSheetItem.OperationContext();


                if (drpBulkMaterialEnv.Label == null || drpBulkMaterialEnv.Label.Equals("")) return;
                proxySheet.Url = _eFunctions.GetServicesUrl(drpBulkMaterialEnv.SelectedItem.Label) + "/BulkMaterialUsageSheet";
                proxyItem.Url = _eFunctions.GetServicesUrl(drpBulkMaterialEnv.SelectedItem.Label) + "/BulkMaterialUsageSheetItem";
                _frmAuth.SelectedEnviroment = drpBulkMaterialEnv.SelectedItem.Label;
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
                    _excelSheetItems.Sort.SortFields.Clear();
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TittleRow), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    _excelSheetItems.Sort.Apply();

                    var currentRow = TittleRow + 1;

                    while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value)) != null)
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
                                _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText;

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
                                                _cells.GetCell(ResultColumn, currentHeader + currentItem).Value += errorMessage;
                                                _cells.GetCell(ResultColumn, currentHeader + currentItem).Select();
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
                                    _cells.GetCell(ResultColumn, currentRow - 1).Value += "No hay Items para Aplicar en esta hoja!";
                                    DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow - 1);
                                }
                            }
                            catch (Exception error)
                            {
                                MessageBox.Show(error.Message);
                            }
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
                        _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText;
                    }
                    _cells.GetRange(1, currentHeader, ResultColumn - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    DeleteHeader(proxySheet, opSheet, requestSheet, currentHeader, currentRow);
                }
                else
                {
                    _cells.GetRange(1, currentHeader, ResultColumn - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Success);                    _cells.GetRange(1, currentHeader, 6, currentRow).Select();
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

                if (requestItem.bulkMaterialTypeId == profile.FuelType && requestItem.quantity > profile.capacity)
                {
                    _cells.GetCell(ResultColumn, currentRow).Value = "Este valor supera la capacidad del Equipo!";
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
                _cells.GetCell(ResultColumn, currentRow).Value = error.Message;
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
                        _cells.GetCell(ResultColumn, (currentHeader + t.fieldIndex)).Value += " - " + t.messageText;
                    }
                }
                else
                {
                    _cells.GetCell(ResultColumn, currentRow).Value += " - Hoja " + replySheet.bulkMaterialUsageSheetDTO.bulkMaterialUsageSheetId + " Borrada";
                    _cells.GetRange(1, currentHeader, ResultColumn - 1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }
            }
            catch (Exception err)
            {
                _cells.GetCell(ResultColumn, currentRow).Value += err.Message;
            }
        }

        private string GetBulkAccountCode(string equipNo)
        {
            try
            {
                if (string.IsNullOrEmpty(equipNo)) return "";

                var sqlQuery = Queries.GetBulkAccountCode(equipNo, _eFunctions.dbReference, _eFunctions.dbLink);

                _eFunctions.SetDBSettings(drpBulkMaterialEnv.SelectedItem.Label);

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

                _eFunctions.SetDBSettings(drpBulkMaterialEnv.SelectedItem.Label);

                var drEquipCapacity = _eFunctions.GetQueryResult(sqlQuery);

                if (!drEquipCapacity.Read())
                {
                    Profile.Error = "No Tiene Perfil";

                    return profile;
                }

                if (!drEquipCapacity.IsClosed && drEquipCapacity.HasRows)
                {
                    profile.equipo = drEquipCapacity["EQUIP_NO"].ToString();
                    profile.egi = drEquipCapacity["EQUIP_GRP_ID"].ToString();
                    profile.FuelType = drEquipCapacity["FUEL_OIL_TYPE"].ToString();
                    profile.capacity = Convert.ToDecimal(drEquipCapacity["FUEL_CAPACITY"].ToString());
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
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;
            var proxySheet = new BMUSheet.BulkMaterialUsageSheetService();
            var opSheet = new BMUSheet.OperationContext();


            if (drpBulkMaterialEnv.Label == null || drpBulkMaterialEnv.Label.Equals("")) return;
            proxySheet.Url = _eFunctions.GetServicesUrl(drpBulkMaterialEnv.SelectedItem.Label) + "/BulkMaterialUsageSheet";
            _frmAuth.SelectedEnviroment = drpBulkMaterialEnv.SelectedItem.Label;
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;

            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            opSheet.district = _frmAuth.EllipseDsct;
            opSheet.maxInstances = 100;
            opSheet.position = _frmAuth.EllipsePost;
            opSheet.returnWarnings = false;


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TittleRow + 1;

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
                        foreach (var t in replySheet.errors) { _cells.GetCell(ResultColumn, currentRow).Value += " - " + t.messageText; }

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
                    _cells.GetCell(ResultColumn, currentRow).Value = error.Message;
                    _cells.GetCell(ResultColumn, currentRow).Select();
                }
                finally { currentRow++; }
            }
        }

        private void btnValidateStats_Click(object sender, RibbonControlEventArgs e)
        {
            ValidateStats();
        }

        private Stats GetLastStatistic(string equipNo, string statType, string statDate)
        {
            try
            {

                var stats = new Stats();
                if (string.IsNullOrEmpty(equipNo) || string.IsNullOrEmpty(statType)) stats.Error = "Error";

                var sqlQuery = Queries.GetLastStatistic(equipNo, statType, statDate, _eFunctions.dbReference, _eFunctions.dbLink);

                _eFunctions.SetDBSettings(drpBulkMaterialEnv.SelectedItem.Label);

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

            if (drpBulkMaterialEnv.Label == null || drpBulkMaterialEnv.Label.Equals("")) return;

            var currentRow = TittleRow + 1;
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

        private static class Queries
        {
            public static string GetBulkAccountCode(string equipNo, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
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

                return sqlQuery;
            }

            public static string GetFuelCapacity(string equipNo, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
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
                    "  DECODE ( PROFILES.EQUIP_GRP_ID, NULL, 'NO TIENE', PROFILES.EQUIP_GRP_ID ) EQUIP_GRP_ID,   " +
                    "  DECODE ( PROFILES.FUEL_OIL_TYPE, NULL, 'NO TIENE', PROFILES.FUEL_OIL_TYPE ) FUEL_OIL_TYPE,   " +
                    "  DECODE ( PROFILES.FUEL_CAPACITY, NULL, 0, PROFILES.FUEL_CAPACITY ) FUEL_CAPACITY   " +
                    "FROM   " +
                    "  EQUIPO   " +
                    "LEFT JOIN PROFILES   " +
                    "ON   " +
                    "  EQUIPO.EQUIP_NO = PROFILES.EQUIP_NO   " +
                    "AND PROFILES.PESO = PROFILES.MAX_PESO   " +
                    "ORDER BY   " +
                    "  PROFILES.PESO   ";

                return sqlQuery;
            }

            public static string GetLastStatistic(string equipNo, string statType, string statDate,  string dbReference, string dbLink)
            {
                var sqlQuery = "" +
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

                return sqlQuery;
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
            public string equipo { get; set; }
            public string egi { get; set; }
            public string FuelType { get; set; }
            public decimal capacity { get; set; }
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
    }
}
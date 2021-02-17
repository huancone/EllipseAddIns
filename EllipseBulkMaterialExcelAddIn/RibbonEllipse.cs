using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using BMUService = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;
using BMUItemService = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;
using EllipseEquipmentClassLibrary;
using ListService = EllipseEquipmentClassLibrary.EquipmentListService;
using System.Threading;
using BulkMaterialClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;

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
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private ExcelStyleCells _cells;
        private Application _excelApp;

        private List<string> _optionList;
        private const string ValidationSheetName = "ValidationListSheet";
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

            settings.SetDefaultCustomSettingValue("AutoSort", "Y");
            settings.SetDefaultCustomSettingValue("OverrideAccountCode", "Maintenance");
            settings.SetDefaultCustomSettingValue("IgnoreItemError", "N");
            settings.SetDefaultCustomSettingValue("AllowBackgroundWork", "N");

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            var overrideAccountCode = settings.GetCustomSettingValue("OverrideAccountCode");
            if (overrideAccountCode.Equals("Maintenance"))
                cbAccountElementOverrideMntto.Checked = true;
            else if (overrideAccountCode.Equals("Disable"))
                cbAccountElementOverrideDisable.Checked = true;
            else if (overrideAccountCode.Equals("Always"))
                cbAccountElementOverrideAlways.Checked = true;
            else if (overrideAccountCode.Equals("Default"))
                cbAccountElementOverrideDefault.Checked = true;
            else
                cbAccountElementOverrideDefault.Checked = true;
            cbAutoSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AutoSort"));
            cbIgnoreItemError.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreItemError"));
            cbAllowBackgroundWork.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AllowBackgroundWork"));

            //
            settings.SaveCustomSettings();
        }
        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(BulkMaterialExecute);

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
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(@"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
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

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).Style = StyleConstants.Normal;
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).Clear();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = "@";


                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "Bulk Material Usage Sheet";

                _cells.GetRange("A1", "B1").Style = StyleConstants.HeaderDefault;
                _cells.GetRange("B1", "D1").Merge();

                _cells.GetCell(1, TitleRow01).Value = "Usage Sheet Id";
                _cells.GetCell(2, TitleRow01).Value = "District";
                _cells.GetCell(3, TitleRow01).Value = "Warehouse";
                _cells.GetCell(4, TitleRow01).Value = "Usage Date";
                _cells.GetCell(4, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(5, TitleRow01).Value = "Usage Time";
                _cells.GetCell(5, TitleRow01).AddComment("HHmmss");
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

                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(9, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(12, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(15, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(17, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleInformation;

                #endregion

                #region Instructions

                _cells.GetCell("E1").Value = "OBLIGATORIO";
                _cells.GetCell("E1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("E2").Value = "OPCIONAL";
                _cells.GetCell("E2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("E3").Value = "INFORMATIVO";
                _cells.GetCell("E3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("E4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("E4").Style = StyleConstants.TitleAction;
                _cells.GetCell("E5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("E5").Style = StyleConstants.TitleAdditional;

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
                _cells.ActiveSheet.Cells.Columns.AutoFit();

                OrderAndSort(_excelApp.ActiveWorkbook.ActiveSheet, TableName01);

                //Hoja 2
                #region Hoja de Listas
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "EQUIPMENT LIST CHECKER - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;

                _cells.GetCell("A3").Value = EquipListSearchFieldCriteria.ListType.Value;
                _cells.GetCell("B3").Value = "PCOMBU";
                _cells.GetCell("A4").Value = EquipListSearchFieldCriteria.ListId.Value;
                _cells.SetValidationList(_cells.GetCell("B4"), GetListIdList("PCOMBU"), ValidationSheetName, 2);
                _cells.GetRange("A3", "A4").Style = StyleConstants.Option;
                _cells.GetRange("B3", "B4").Style = StyleConstants.Select;

                var statusCodeList = _eFunctions.GetItemCodes("ES").Select(item => item.Code + " - " + item.Description).ToList();
                var equipClassCodeList = _eFunctions.GetItemCodes("EC").Select(item => item.Code + " - " + item.Description).ToList();
                var equipTypeCodeList = _eFunctions.GetItemCodes("ET").Select(item => item.Code + " - " + item.Description).ToList();
                var compCodeList = _eFunctions.GetItemCodes("CO").Select(item => item.Code + " - " + item.Description).ToList();
                var mnemonicCodeList = _eFunctions.GetItemCodes("AA").Select(item => item.Code + " - " + item.Description).ToList();
                var classTypeCodeList = _eFunctions.GetItemCodes("E0").Select(item => item.Code + " - " + item.Description).ToList();
                var fuelTypeCodeList = _eFunctions.GetItemCodes("E2").Select(item => item.Code + " - " + item.Description).ToList();

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

                _cells.ActiveSheet.Cells.Columns.AutoFit();

                #endregion
                _cells.GetWorksheet(1).Select(Type.Missing);
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

        private void OrderAndSort(Worksheet excelSheet, string tableName)
        {
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            var tableSheetItems = _cells.GetRange(tableName).ListObject;
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
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetWorkingWorksheet(_cells.ActiveSheet);
                _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);

                #region AlertaDecimal

                if (!_cells.IsDecimalDotSeparator())
                {
                    const string message =
                        @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse";

                    var dialogResult = MessageBox.Show(message + "\n\n¿Está seguro que desea continuar?",
                        "ADVERTENCIA - Separador Decimal", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2);
                    if (dialogResult == DialogResult.No)
                        return;
                }

                #endregion

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

                var bulkMaterials = cc.Read<BulkMaterialItem>(filePath, inputFileDescription);

                var currentRow = TitleRow01 + 1;
                foreach (var bulkMaterial in bulkMaterials)
                {
                    try
                    {
                        _cells.GetCell(3, currentRow).Value = "'" + bulkMaterial.WarehouseId;
                        _cells.GetCell(4, currentRow).Value = DateTime
                            .ParseExact(bulkMaterial.DefaultUsageDate, @"MM/dd/yy", CultureInfo.CurrentCulture)
                            .ToString("yyyyMMdd");
                        _cells.GetCell(8, currentRow).Value = "'" + bulkMaterial.EquipmentReference;
                        _cells.GetCell(11, currentRow).Value = bulkMaterial.BulkMaterialTypeId;
                        _cells.GetCell(13, currentRow).Value = bulkMaterial.Quantity;
                    }
                    catch (Exception error)
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Value = "Error: " + error.Message;
                    }
                    finally
                    {
                        currentRow++;
                    }
                }

                OrderAndSort(_cells.GetWorkingWorksheet(), TableName01);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ImportFile()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        /// <summary>
        ///     Crea las instancias a los servicios BulkMaterialUsageSheetService y BulkMaterialUsageSheetItemService
        /// </summary>
        private void BulkMaterialExecute()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetWorkingWorksheet(_cells.ActiveSheet);
                _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);

                #region AlertaDecimal
                if (!_cells.IsDecimalDotSeparator())
                {
                    const string message = @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse";

                    var dialogResult = MessageBox.Show(message + "\n\n¿Está seguro que desea continuar?", "ADVERTENCIA - Separador Decimal", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (dialogResult == DialogResult.No)
                        return;
                }
                #endregion 

                _cells.SetCursorWait();
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                //
                //var urlServicePost = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label, ServiceType.PostService);
                //_eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlServicePost);
                //
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

                var sheetService = new BMUService.BulkMaterialUsageSheetService {Url = urlService + "/BulkMaterialUsageSheet"};

                var opContext = new BMUService.OperationContext()
                {
                    district = _frmAuth.EllipseDsct,
                    maxInstances = 100,
                    position = _frmAuth.EllipsePost,
                    returnWarnings = false
                };

                var itemService = new BMUItemService.BulkMaterialUsageSheetItemService();
                itemService.Url = urlService + "/BulkMaterialUsageSheetItem";
                var opItem = new BMUItemService.OperationContext()
                {
                    district = _frmAuth.EllipseDsct,
                    maxInstances = 100,
                    position = _frmAuth.EllipsePost,
                    returnWarnings = false
                };
                
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var currentRow = TitleRow01 + 1;
                var currentHeaderRow = currentRow;
                
                if (cbAutoSortItems.Checked)
                {
                    var tableSheetItems = _cells.GetRange(TableName01).ListObject;
                    tableSheetItems.Sort.SortFields.Clear();
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(8, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TitleRow01), XlSortOn.xlSortOnValues, XlOrder.xlDownThenOver, Type.Missing, Type.Missing);
                    tableSheetItems.Sort.Apply();
                }

                BulkMaterialUsageSheet currentSheetHeader = null;

                var itemList = new List<BulkMaterialUsageSheetItem>();
                
                while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value)) != null)
                {
                    try
                    {
                        if(_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(1, currentRow).Select();

                        //llenado de variables del encabezado de la hoja
                        var newSheetHeader = new BulkMaterialUsageSheet
                        {
                            BulkMaterialUsageSheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                            DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) ?? "ICOR",
                            WarehouseId = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value),
                            DefaultUsageDate = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value),
                            DefaultAccountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value)
                        };


                        string itemAccountCode = null;
                        var materialTypeId = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull("" + _cells.GetCell(11, currentRow).Value));
                        var equipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                        itemAccountCode = GetItemAccountCode(_eFunctions, newSheetHeader.DefaultAccountCode, equipNo, materialTypeId);
                        newSheetHeader.DefaultAccountCode = newSheetHeader.DefaultAccountCode ?? itemAccountCode;

                        if (currentSheetHeader == null)
                        {
                            currentSheetHeader = newSheetHeader;
                            currentHeaderRow = currentRow;
                        }

                        //Crea el encabezado cuando los ids sean diferente o si el encabezado es diferente en caso de ids automáticos
                        var isNullIds = string.IsNullOrWhiteSpace(newSheetHeader.BulkMaterialUsageSheetId) && string.IsNullOrWhiteSpace(currentSheetHeader.BulkMaterialUsageSheetId);
                        var isEqualIds = newSheetHeader.BulkMaterialUsageSheetId == currentSheetHeader.BulkMaterialUsageSheetId;

                        if (!isEqualIds || (isNullIds && !newSheetHeader.Equals(currentSheetHeader)))
                        {
                            CreateBulkMaterialSheet(sheetService, opContext, itemService, opItem, currentSheetHeader, itemList, currentHeaderRow, currentRow - 1);
                            currentSheetHeader = newSheetHeader;
                            currentHeaderRow = currentRow;
                        }

                        var requestItem = new BulkMaterialUsageSheetItem();

                        requestItem.BulkMaterialUsageSheetId = "" + currentSheetHeader.BulkMaterialUsageSheetId;
                        requestItem.EquipmentReference = equipNo;
                        requestItem.ComponentCode = _cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(9, currentRow).Value);
                        requestItem.Modifier = _cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(10, currentRow).Value);
                        requestItem.BulkMaterialTypeId = materialTypeId;
                        requestItem.ConditionMonitoringAction = (_cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value) == "Fuel/Diesel") || (_cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(12, currentRow).Value) == null) ? null : "" + _cells.GetCell(12, currentRow).Value.ToString().Substring(0, 1);
                        requestItem.Quantity = "" + _cells.GetCell(13, currentRow).Value;
                        requestItem.UsageDate = _cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(14, currentRow).Value);
                        requestItem.UsageTime = _cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(15, currentRow).Value);
                        requestItem.OperationStatisticType = _cells.GetNullIfTrimmedEmpty("" + _cells.GetCell(16, currentRow).Value);
                        requestItem.MeterReading = "" + _cells.GetCell(17, currentRow).Value;
                        requestItem.AccountCode = itemAccountCode;
                        

                        itemList.Add(requestItem);

                        //Si es el último registro
                        if (string.IsNullOrWhiteSpace(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow + 1).Value)))
                        {
                            //Para control de estilos en caso de falla
                            currentRow++;
                            //Creo la hoja si es el último registro
                            CreateBulkMaterialSheet(sheetService, opContext, itemService, opItem, currentSheetHeader, itemList, currentHeaderRow, currentRow - 1);
                            //Reajuste de control de estilo si no hay fallas
                            currentRow--;
                        }
                    }
                    catch (Exception ex)
                    {
                        var exceptionMessage = "";
                        var exceptionType = StyleConstants.Error;
                        Debugger.LogError("RibbonEllipse:BulkMaterialExcecute(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        if (ex.Message.Equals("The operation has timed out"))
                        {
                            exceptionMessage = " - Hoja " + currentSheetHeader.BulkMaterialUsageSheetId + " Completada. No se recibió respuesta de Ellipse, por lo que se recomienda verificar la completación";
                            exceptionType = StyleConstants.Warning;
                        }
                        else if(ex.Message.Contains("record with the same key already exists in table [BulkMaterialUsageSheet]."))
                        {
                            exceptionMessage = "El id ingresado ya existe. " + currentSheetHeader.BulkMaterialUsageSheetId;
                            exceptionType = StyleConstants.Error;
                        }
                        else if(ex.Message.Contains("reference not set to an instance of an object"))
                        {
                            exceptionMessage = ex.Message;
                            exceptionType = StyleConstants.Error;
                        }
                        else
                        {
                            if(!string.IsNullOrWhiteSpace(currentSheetHeader.BulkMaterialUsageSheetId))
                                BulkMaterialActions.DeleteHeader(sheetService, opContext, currentSheetHeader.ToDto());
                            exceptionMessage = " - Hoja " + currentSheetHeader.BulkMaterialUsageSheetId + " Borrada. " + ex.Message;
                            exceptionType = StyleConstants.Error;
                        }
                        //Agrego el mensaje para el resultado de la excepción
                        for (int i = currentHeaderRow; i < currentRow; i++)
                            _cells.GetCell(ResultColumn01, i).Value += exceptionMessage;
                        _cells.GetRange(1, currentHeaderRow, ResultColumn01, currentRow - 1).Style = exceptionType;

                        //Cuando ocurre un error no se agrega el elemento actual que estaba leyendo por lo que se fuerza a que vuelva a procesar la línea actual
                        currentRow--;

                        currentSheetHeader = null;
                        itemList.Clear();
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:BulkMaterialExecute()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }

        }

        private void CreateBulkMaterialSheet(BMUService.BulkMaterialUsageSheetService sheetService, BMUService.OperationContext opContext, BMUItemService.BulkMaterialUsageSheetItemService itemService, BMUItemService.OperationContext opItem, BulkMaterialUsageSheet sheetHeader, List<BulkMaterialUsageSheetItem> itemList, int headerRow, int currentRow)
        {
            DateTime usageDate;

            if (!DateTime.TryParseExact(sheetHeader.DefaultUsageDate, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, out usageDate))
                throw new Exception("Se ha ingresado una fecha inválida");

            if (itemList.Count <= 0)
                throw new Exception("No hay items para agregar en esta hoja");

            var newSheetId = "";
            if(!string.IsNullOrWhiteSpace(sheetHeader.BulkMaterialUsageSheetId) && sheetHeader.BulkMaterialUsageSheetId.Length > 32)
                throw new Exception("El Id de la hoja no puede tener más de 32 caracteres");

            var replySheet = BulkMaterialActions.CreateHeader(sheetService, opContext, sheetHeader.ToDto());

            //valido que no haya errores en la creación del encabezado
            if (replySheet.errors != null && replySheet.errors.Length > 0)
            {
                var errorMessage = "";
                foreach (var t in replySheet.errors)
                    errorMessage += " - " + t.messageText;

                throw new Exception(errorMessage);
            }

            newSheetId = replySheet.bulkMaterialUsageSheetDTO.bulkMaterialUsageSheetId;
            
            sheetHeader.BulkMaterialUsageSheetId = newSheetId;
            
            _cells.GetRange(1, headerRow, 1, currentRow).Value = sheetHeader.BulkMaterialUsageSheetId;
            _cells.GetRange(1, headerRow, 6, currentRow).Style = StyleConstants.Success;

            var existItemError = false;

            foreach (var item in itemList)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(item.BulkMaterialUsageSheetId))
                        item.BulkMaterialUsageSheetId = sheetHeader.BulkMaterialUsageSheetId;

                    var replyItem = BulkMaterialActions.AddItemToHeader(_eFunctions, itemService, opItem, item.ToDto());

                    //valido que no haya errores en la creación del ítem
                    if (replyItem.errors != null && replyItem.errors.Length > 0)
                    {
                        var errorMessage = "";
                        foreach (var t in replyItem.errors)
                            errorMessage += " - " + t.messageText;

                        throw new Exception(errorMessage);
                    }
                    
                    _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Value = "OK";
                    _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Style = StyleConstants.Success;
                    if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                        _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Select();
                }
                catch (Exception ex)
                {
                    existItemError = true;
                    _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Value = ex.Message;
                    _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Style = StyleConstants.Error;
                    if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                        _cells.GetCell(ResultColumn01, headerRow + itemList.IndexOf(item)).Select();
                }
                finally
                {
                    //valido si hay un error y debe ser ignorado
                    if (existItemError && !cbIgnoreItemError.Checked)
                        throw new Exception("Se ha cancelado la creación de la hoja por un error al intentar agregar uno de sus ítems.");
                }
            }

            BulkMaterialActions.ApplyHeader(sheetService, opContext, sheetHeader.ToDto());
            _cells.GetRange(1, headerRow, ResultColumn01 - 1, currentRow).Style = StyleConstants.Success;
            if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                _cells.GetRange(1, headerRow, 6, currentRow).Select();

            itemList.Clear();
        }

        private void btnUnApplyDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
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
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetWorkingWorksheet(_cells.ActiveSheet);
                _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);


                _cells.SetCursorWait();
                var service = new BMUService.BulkMaterialUsageSheetService();
                var opContext = new BMUService.OperationContext()
                {
                    district = _frmAuth.EllipseDsct,
                    maxInstances = 100,
                    position = _frmAuth.EllipsePost,
                    returnWarnings = false,
                };
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

                if (drpEnvironment.Label == null || drpEnvironment.Label.Equals("")) return;
                service.Url = urlService + "/BulkMaterialUsageSheet";

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var currentRow = TitleRow01 + 1;

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, MaxRows).ClearComments();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
                
                while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value)) != null)
                {
                    var requestSheet = new BMUService.BulkMaterialUsageSheetDTO();

                    try
                    {
                        requestSheet.bulkMaterialUsageSheetId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);

                        var replySheet = BulkMaterialActions.UnApplyHeader(service, opContext, requestSheet, true);

                        BulkMaterialActions.DeleteHeader(service, opContext, requestSheet);
                        _cells.GetRange(1, currentRow, ResultColumn01, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn01, currentRow).Value2 = "HOJA ELIMINADA";
                    }
                    catch (Exception error)
                    {
                        _cells.GetRange(1, currentRow, ResultColumn01, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, currentRow).Value = error.Message;
                    }
                    finally
                    {
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn01, currentRow).Select();
                        currentRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:Unapply()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
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

            var stats = new Stats();
            if (string.IsNullOrEmpty(equipNo) || string.IsNullOrEmpty(statType)) stats.Error = "Error";

            var sqlQuery = Queries.GetLastStatistic(equipNo, statType, statDate, _eFunctions.DbReference,
                _eFunctions.DbLink);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var drLastStat = _eFunctions.GetQueryResult(sqlQuery);

            if (drLastStat != null && !drLastStat.IsClosed && drLastStat.Read())
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

        private void ValidateStats()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetWorkingWorksheet(_cells.ActiveSheet);
                _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                const int resultColumn = ResultColumn01;

                _cells?.SetCursorDefault();
                var currentRow = TitleRow01 + 1;
                while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value)) != null)
                {
                    try
                    {
                        var statType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
                        var equipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                        var stat = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
                        var statDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                        if(string.IsNullOrWhiteSpace(statDate))
                            statDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);

                        if (string.IsNullOrWhiteSpace(equipNo))
                            throw new ArgumentNullException(nameof(equipNo), "Se necesita un número de equipo para poder validar la estadística");
                        if (string.IsNullOrWhiteSpace(stat))
                            stat = "0";
                        if (string.IsNullOrWhiteSpace(statType))
                            throw new ArgumentNullException(nameof(statType), "Se necesita el tipo de estadística para poder validar la estadística");
                        
                        
                        var lastStat = GetLastStatistic(equipNo, statType, statDate);
                        _cells.GetCell(17, currentRow)
                            .AddComment(Convert.ToString(lastStat.StatDate + " - " + lastStat.MeterValue,
                                CultureInfo.InvariantCulture));
                        _cells.GetCell(17, currentRow).Style = _cells.GetStyle(
                            Convert.ToDecimal(stat) < lastStat.MeterValue
                                ? StyleConstants.Error
                                : StyleConstants.Success);
                        
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, currentRow).Value = "VALIDADO";
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = ex.Message;
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidateStats()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private List<string> GetListIdList(string listType)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var sqlQuery = Queries.GetListIdList(_eFunctions.DbReference, _eFunctions.DbLink, listType);
            var drItem = _eFunctions.GetQueryResult(sqlQuery);

            var list = new List<string>();

            if (drItem == null || drItem.IsClosed) return list;

            while (drItem.Read())
            {
                list.Add("" + drItem["LIST_ID"].ToString().Trim());
            }

            return list;
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
                    _thread = new Thread(ReviewEquipmentsList);
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

        private void ReviewEquipmentsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetWorkingWorksheet(_cells.ActiveSheet);
            _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            try
            {
                _cells.SetCursorWait();
                _cells.ClearTableRange(TableName02);


                //Obtengo los valores de las opciones de búsqueda
                var searchCriteriaKey1 = EquipListSearchFieldCriteria.ListType.Key;
                var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                var searchCriteriaKey2 = EquipListSearchFieldCriteria.ListId.Key;
                var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
                var previousEquipment = new Equipment {EquipmentNo = ""};

                var listeq = ListActions.FetchListEquipmentsList(_eFunctions, searchCriteriaKey1, searchCriteriaValue1,
                    searchCriteriaKey2, searchCriteriaValue2, null);
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

                        var eq = eql.EquipNo.Trim().Equals(previousEquipment.EquipmentNo.Trim())
                            ? previousEquipment
                            : EquipmentActions.FetchEquipmentData(_eFunctions, eql.EquipNo);

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
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _cells.WorkingWorksheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewEquipmentList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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

        private void btnReviewFromBulkSheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01 || ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewEquipmentListFromBulkList);
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

        private void ReviewEquipmentListFromBulkList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var celleq = new ExcelStyleCells(_excelApp, SheetName01);
            var cellli = new ExcelStyleCells(_excelApp, SheetName02);
            celleq.SetCursorWait();
            cellli.ClearTableRange(TableName02);
            const int resultColumn = ResultColumn02;
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            try
            {
                var k = TitleRow01 + 1;
                var i = TitleRow02 + 1;
                var previousEquipment = new Equipment {EquipmentNo = ""};
                while (!string.IsNullOrEmpty("" + celleq.GetCell(8, k).Value))
                {
                    var equipmentNo = celleq.GetEmptyIfNull(celleq.GetCell(8, k).Value);

                    try
                    {
                        var eqlists = ListActions.FetchListEquipmentsList(_eFunctions,
                            EquipListSearchFieldCriteria.EquipmentNo.Key, equipmentNo,
                            EquipListSearchFieldCriteria.None.Key, null, null);

                        foreach (var list in eqlists)
                        {
                            var listedEq = ListActions.FetchListEquipmentsList(_eFunctions,
                                EquipListSearchFieldCriteria.ListType.Key, list.ListType,
                                EquipListSearchFieldCriteria.ListId.Key, list.ListId, null);
                            foreach (EquipListItem eql in listedEq)
                            {
                                try
                                {
                                    //Para resetear el estilo
                                    cellli.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                                    cellli.GetCell(1, i).Value = "'" + eql.EquipNo;
                                    cellli.GetCell(5, i).Value = "'" + eql.ListType;
                                    cellli.GetCell(6, i).Value = "'" + eql.ListId;

                                    var eq = eql.EquipNo.Trim().Equals(previousEquipment.EquipmentNo.Trim())
                                        ? previousEquipment
                                        : EquipmentActions.FetchEquipmentData(_eFunctions, eql.EquipNo);

                                    cellli.GetCell(2, i).Value = "'" + eq.EquipmentNoDescription1;
                                    cellli.GetCell(3, i).Value = "'" + eq.EquipmentNoDescription2;
                                    cellli.GetCell(4, i).Value = "'" + eq.EquipmentStatus;
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

                                    previousEquipment = eq;
                                }
                                catch (Exception ex)
                                {
                                    cellli.GetCell(1, i).Style = StyleConstants.Error;
                                    cellli.GetCell(resultColumn, i).Value = "ERRORLIST: " + ex.Message;
                                    Debugger.LogError("RibbonEllipse.cs:ReviewFromEquipmentList()", ex.Message);
                                }
                                finally
                                {
                                    if (_cells.ActiveSheet.Name == SheetName02)
                                        cellli.GetCell(2, i).Select();
                                    i++;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        celleq.GetCell(1, k).Style = StyleConstants.Error;
                        celleq.GetCell(ResultColumn01, k).Value = "ERRORLIST: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ReviewFromEquipmentList()", ex.Message);
                    }
                    finally
                    {
                        if (_cells.ActiveSheet.Name == SheetName01)
                            celleq.GetCell(1, k).Select();
                        else if (_cells.ActiveSheet.Name == SheetName02)
                            cellli.GetCell(1, i).Select();
                        k++;
                    }
                }

                cellli.WorkingWorksheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewEquipmentListFromBulkList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnAddToList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(AddEquipmentsToList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AddEquipmentToList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void AddEquipmentsToList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetWorkingWorksheet(_cells.ActiveSheet);
            _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);

            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            try
            {
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
                        var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                        var equiplist = new EquipListItem()
                        {
                            EquipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            ListType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                            ListId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                        };

                        ListActions.AddEquipmentToList(opSheet, urlService, equiplist);

                        _cells.GetCell(ResultColumn02, i).Value = "AGREGADO A LA LISTA";
                        _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                        Debugger.LogError("RibbonEllipse.cs:AddListEquipmentsList()", ex.Message);
                    }
                    finally
                    {
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                        i++;
                    }
                }

                _cells.WorkingWorksheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AddEquipmentsToList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnRemoveFromList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName02)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(RemoveEquipmentsFromList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:RemoveEquipmentsFromList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void RemoveEquipmentsFromList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetWorkingWorksheet(_cells.ActiveSheet);
            _cells.SetFixedWorkingWorkSheet(cbAllowBackgroundWork.Checked);

            try
            {
                _cells.SetCursorWait();
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                
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
                        var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                        var equiplist = new EquipListItem()
                        {
                            EquipNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            ListType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                            ListId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                        };

                        ListActions.DeleteEquipmentFromList(opSheet, urlService, equiplist);

                        _cells.GetCell(ResultColumn02, i).Value = "ELIMINADO DE LISTA";
                        _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                        Debugger.LogError("RibbonEllipse.cs:RemoveEquipmentsFromList()", ex.Message);
                    }
                    finally
                    {
                        if (_cells.ActiveSheet.Equals(_cells.WorkingWorksheet))
                            _cells.GetCell(ResultColumn02, i).Select();
                        i++;
                    }
                }

                _cells.WorkingWorksheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:RemoveEquipmentsFromList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private string GetItemAccountCode(EllipseFunctions ef, string defaultAccountCode, string equipNo, string materialTypeId)
        {
            
            if (cbAccountElementOverrideDisable.Checked)
                return defaultAccountCode;
            if (cbAccountElementOverrideDefault.Checked && !string.IsNullOrWhiteSpace(defaultAccountCode))
                return defaultAccountCode;

            var bulkItem = BulkMaterialActions.GetEquipmentBulkItem(_eFunctions, equipNo, materialTypeId);

            //Si EquipClassCode19 es nulo es porque no pudo hacer la conexión o xq no encontró el item o reference code
            if (bulkItem == null || bulkItem.EquipClassCode19 == null) return null;

            if (cbAccountElementOverrideMntto.Checked && !bulkItem.EquipClassCode19.Equals("MT") && !string.IsNullOrWhiteSpace(defaultAccountCode))
                return defaultAccountCode;

            return bulkItem.PreferredAccountCode;

        }

        private bool CheckOverrideAccountCheckBoxes()
        {
            return (cbAccountElementOverrideDisable.Checked || 
                   cbAccountElementOverrideAlways.Checked ||
                   cbAccountElementOverrideDefault.Checked||
                   cbAccountElementOverrideMntto.Checked);

        }
        private void cbAccountElementOverrideMntto_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("OverrideAccountCode", "Maintenance");
            Settings.CurrentSettings.SaveCustomSettings();
            cbAccountElementOverrideDisable.Checked = false;
            cbAccountElementOverrideAlways.Checked = false;
            cbAccountElementOverrideDefault.Checked = false;
            if (!CheckOverrideAccountCheckBoxes())
                cbAccountElementOverrideMntto.Checked = true;
        }

        private void cbAccountElementOverrideDisable_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("OverrideAccountCode", "Disable");
            Settings.CurrentSettings.SaveCustomSettings();
            cbAccountElementOverrideAlways.Checked = false;
            cbAccountElementOverrideDefault.Checked = false;
            cbAccountElementOverrideMntto.Checked = false;
            if (!CheckOverrideAccountCheckBoxes())
                cbAccountElementOverrideDisable.Checked = true;
        }

        private void cbAccountElementOverrideDefault_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("OverrideAccountCode", "Default");
            Settings.CurrentSettings.SaveCustomSettings();
            cbAccountElementOverrideDisable.Checked = false;
            cbAccountElementOverrideAlways.Checked = false;
            cbAccountElementOverrideMntto.Checked = false;
            if (!CheckOverrideAccountCheckBoxes())
                cbAccountElementOverrideDefault.Checked = true;
        }

        private void cbAccountElementOverrideAlways_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("OverrideAccountCode", "Always");
            Settings.CurrentSettings.SaveCustomSettings();
            cbAccountElementOverrideDisable.Checked = false;
            cbAccountElementOverrideDefault.Checked = false;
            cbAccountElementOverrideMntto.Checked = false;
            if (!CheckOverrideAccountCheckBoxes())
                cbAccountElementOverrideAlways.Checked = true;
        }

        private void cbAutoSortItems_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("AutoSort", MyUtilities.ToString(cbAutoSortItems.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbIgnoreItemError_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("IgnoreItemError", MyUtilities.ToString(cbIgnoreItemError.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbAllowBackgroundWork_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("AllowBackgroundWork", MyUtilities.ToString(cbAllowBackgroundWork.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }
    }
}
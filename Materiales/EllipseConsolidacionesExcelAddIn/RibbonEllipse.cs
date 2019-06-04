using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseConsolidacionesExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        private Excel.Application _excelApp;

        private const int TitleRow01 = 5;
        private const int ResultColumn01 = 3;
        private const string SheetName01 = "Consolidaciones";
        private const string TableName01 = "ConsolidacionesTable";

        private const int TitleRow02 = 5;
        private const int ResultColumn02 = 3;
        private const string SheetName02 = "Categoría de Servicios";
        private const string TableName02 = "ServiceCategoryTable";

        private const string ValidationSheetName = "ValidationSheet";

        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                formatSheet();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al formatear la hoja: " + ex.Message);
            }
        }

        
        private void btnConsolidations_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ConsolidateProducts);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ConsolidateProducts()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnServiceCategory_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateCategoryService);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:UpdateCategoryService()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnClean_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
            _cells.ClearTableRange(TableName02);
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
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

        private void formatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);


                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.Worksheets.Add();
                _cells.CreateNewWorksheet(ValidationSheetName);

                //CONSTRUYO LA HOJA 1
                #region Hoja1
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;
                var tableName = TableName01;
                var sheetName = SheetName01;

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A3");

                _cells.GetCell("B1").Value = "CONSOLIDACIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "B3");

                _cells.GetCell("C1").Value = "OBLIGATORIO";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("C2").Value = "OPCIONAL";
                _cells.GetCell("C2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("C3").Value = "INFORMATIVO";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(01, titleRow).Value = "Antiguo StockCode";
                _cells.GetCell(02, titleRow).Value = "Nuevo StockCode";
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion
                #region Hoja 2
                titleRow = TitleRow02;
                resultColumn = ResultColumn02;
                tableName = TableName02;
                sheetName = SheetName02;
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A3");

                _cells.GetCell("B1").Value = "CATEGORÍA DE SERVICIOS - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "B3");

                _cells.GetCell("C1").Value = "OBLIGATORIO";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("C2").Value = "OPCIONAL";
                _cells.GetCell("C2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("C3").Value = "INFORMATIVO";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(01, titleRow).Value = "StockCode";
                _cells.GetCell(02, titleRow).Value = "Categoría de Servicio de Producto";
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion  
                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void ConsolidateProducts()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //ScreenService Opción en reemplazo de los servicios
                var opContext = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                var screenService = new Screen.ScreenService
                {
                    Url = urlService + "/ScreenService"
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var tableName = TableName01;
                var resultColumn = ResultColumn01;

                _cells.ClearTableRangeColumn(tableName, resultColumn);
                var i = TitleRow01 + 1;
                while ("" + _cells.GetCell(1, i).Value != "")
                {
                    try
                    {
                        _eFunctions.RevertOperation(opContext, screenService);


                        const string program = "MSB109";

                        var oldStockCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) == null ? null : _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value).PadLeft(9, '0');
                        var newStockCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value) == null ? null : _cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value).PadLeft(9, '0');

                        //ejecutamos el programa
                        var reply = screenService.executeScreen(opContext, "MSO080");
                        //Validamos el ingreso
                        if (reply.mapName != "MSM080A")
                            throw new Exception("No se ha podido ingresar al programa " + program);
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("RESTART1I", program);
                        var request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = screenService.submit(opContext, request);

                        if (reply.mapName != "MSM080A")
                            throw new Exception("No se ha podido ingresar al programa " + program);
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("SKLITEM1I", "1");


                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = screenService.submit(opContext, request);
                        //Enviamos la primera pantalla
                        if (reply.mapName != "MSM080B") continue;
                        var screenFields = new ArrayScreenNameValue(reply.screenFields);
                        if (!screenFields.GetField("REPORT2I").value.Equals(program))
                            throw new Exception("No se ha podido ingresar al programa " + program);

                        arrayFields = new ArrayScreenNameValue();
                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = screenService.submit(opContext, request);

                        if (reply.mapName != "MSM109A")
                            throw new Exception("Error al enviar información del programa " + program);

                        //se adicionan los valores a los campos
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OLD_STOCK_CODE1I", oldStockCode);
                        arrayFields.Add("NEW_STOCK_CODE1I", newStockCode);
                        arrayFields.Add("REQ_BY1I", _frmAuth.EllipseUser);

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = screenService.submit(opContext, request);

                        //Confirmaciones, Validaciones y Advertencias
                        while(reply.functionKeys.ToUpper() == "XMIT-CONFIRM" || reply.functionKeys.ToUpper().StartsWith("XMIT-WARNING") || reply.functionKeys.ToUpper() == "XMIT-VALIDATE")
                        {
                            request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenFields = arrayFields.ToArray(),
                                screenKey = "1"
                            };
                            reply = screenService.submit(opContext, request);
                        }

                        _cells.GetCell(resultColumn, i).Value2 = reply.message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Value2 = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        i++;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:ConsolidateProducts()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void UpdateCategoryService()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Inicio de elementos del servicio
                var catalogueService = new CatalogueService.CatalogueService()
                {
                    Url = urlService + "/CatalogueService"
                };

                var opContext = new CatalogueService.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };


                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var tableName = TableName01;
                var resultColumn = ResultColumn01;

                _cells.ClearTableRangeColumn(tableName, resultColumn);
                var i = TitleRow01 + 1;
                while ("" + _cells.GetCell(1, i).Value != "")
                {
                    try
                    {

                        var stockCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) == null ? null : _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value).PadLeft(9, '0');
                        var categoryId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);

                        var request = new CatalogueService.CatalogueDTO();
                        request.stockCode = stockCode;
                        request.productServiceCategoryId = categoryId;
                        
                        //ejecutamos el programa
                        var reply = catalogueService.update(opContext, request);

                        //si hay errores al final
                        if (reply.errors != null && reply.errors.Length > 0)
                        {
                            foreach (var t in reply.errors)
                                _cells.GetCell(resultColumn, i).Value += " - " + t.messageText;

                            _cells.GetCell(resultColumn, i).Value = "ERROR" + _cells.GetCell(resultColumn, i).Value;
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        }
                        //si hay advertencias
                        else if (reply.warnings != null && reply.warnings.Length > 0)
                        {
                            foreach (var t in reply.warnings)
                                _cells.GetCell(resultColumn, i).Value += " - " + t.messageText;

                            _cells.GetCell(resultColumn, i).Value = "WARNING" + _cells.GetCell(resultColumn, i).Value;
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Warning;
                        }
                        else
                        {
                            _cells.GetCell(resultColumn, i).Value2 = "ACTUALIZADO";
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Value2 = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        i++;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:UpdateCategoryService()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
    }
}

using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Web.Services.Ellipse;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseEqOperStatisticsExcelAddIn.EquipmentOperatingStatisticsService;
using Microsoft.Office.Tools.Excel;
using System.Threading;
using EllipseEqOperStatisticsExcelAddIn.EllipseEqOperStatisticsClassLibrary;

namespace EllipseEqOperStatisticsExcelAddIn
{
    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;

        private const string SheetName01 = "OperationStatistics";
        private const int TitleRow01 = 6;
        private const int ResultColumn01 = 10;
        private const string TableName01 = "OperationStatisticsTable";
        private const string ValidationSheetName = "ValidationSheet";

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            //Office 2013 requiere no ejecutar esta sentencia al iniciar porque no se cuenta con un libro activo vacío. Se debe ejecutar obligatoriamente al formatear las hojas
            //adcionalmente validar la cantidad de hojas a utilizar al momento de dar formato
            //if (_cells == null)
            //    _cells = new ExcelStyleCells(_excelApp);
            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }
        private void btnLoadStatistics_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(LoadStatistics);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:LoadStatistics()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewStatisticList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:LoadStatistics()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        /// <summary>
        /// Establece el formato de la hoja para el cargue de estadísticas de operación
        /// </summary>
        public void FormatSheetHeaderData()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName);


                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "EQUIPMENT OPERATION STATISTICS - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "EQUIPO";
                _cells.GetCell("A4").Value = "ESTADÍSTICA";
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                var statsList = _eFunctions.GetItemCodes("SS").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell("B4"), statsList, ValidationSheetName, 1, false);//TIPO ESTAD


                _cells.GetRange(1, TitleRow01, 7, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, TitleRow01).Value = "FECHA";
                _cells.GetCell(1, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(2, TitleRow01).Value = "TURNO";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Value = "EQUIPO";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPCIÓN";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(5, TitleRow01).Value = "TIPO ESTAD.";
                _cells.GetCell(5, TitleRow01).AddComment("Ej: HR - Horas" +
                                                         "\nR1 - HR MOTOR BABOR          " +
                                                         "\nR2 - HR MOTOR ESTRIBOR       " +
                                                         "\nR3 - HR GENERADOR BABOR      " +
                                                         "\nR4 - HR GENERADOR ESTRIBOR   " +
                                                         "\nR5 - HR COMPRESOR POPA       " +
                                                         "\nR6 - HR COMPRESOR PROA       " +
                                                         "\nR7 - HR MOTOR MONITOR		");
                _cells.GetCell(6, TitleRow01).Value = "TIPO ENTRADA";
                _cells.GetCell(6, TitleRow01).AddComment("Ej: D - DAYLY, M - METER. Predeterminado M");
                _cells.GetCell(7, TitleRow01).Value = "FECHA ÚLTIMA EST.";
                _cells.GetCell(7, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow01).Value = "MEDIDOR ÚLT. EST.";
                _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(1, TitleRow01 + 1, 8, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(9, TitleRow01).Value = "VALOR ESTAD.";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);

                var shiftList = _eFunctions.GetItemCodes("SH").Select(item => item.code + " - " + item.description).ToList();

                var entryList = new List<string>();
                entryList.Add("D - Diario/Daily");
                entryList.Add("M - Medidor/Meter");

                _cells.SetValidationList(_cells.GetCell(2, TitleRow01 + 1), shiftList, ValidationSheetName, 2, false);//TURNO
                _cells.SetValidationList(_cells.GetCell(5, TitleRow01 + 1), ValidationSheetName, 1, false);//TIPO ESTAD
                _cells.SetValidationList(_cells.GetCell(6, TitleRow01 + 1), entryList, ValidationSheetName, 3, false);//TIPO ENTRADA

                var table = _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                //búsquedas especiales de tabla
                var tableObject = Globals.Factory.GetVstoObject(table);
                tableObject.Change += GetTableChangedValue;

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        /// <summary>
        /// Carga la estadística de la hoja especificada
        /// </summary>
        public void LoadStatistics()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                var statService = new EquipmentOperatingStatisticsService.EquipmentOperatingStatisticsService();
                statService.Url = urlService + "/EquipmentOperatingStatistics";
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var opContext = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var currentRow = TitleRow01 + 1;
                while ("" + _cells.GetCell(1, currentRow).Value != "")
                {
                    try
                    {

                        var request = new List<EquipmentOperatingStatisticsDTO>();

                        var reqItem = new EquipmentOperatingStatisticsDTO();


                        reqItem.statisticDate = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        reqItem.shiftCode = MyUtilities.GetCodeKey(_cells.GetCell(2, currentRow).Value);
                        reqItem.equipmentNumber = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                        reqItem.operationStatisticType = MyUtilities.GetCodeKey(_cells.GetCell(5, currentRow).Value);
                        var entryType = "" + MyUtilities.GetCodeKey(_cells.GetCell(6, currentRow).Value);
                        if (entryType.Equals("D"))
                        {
                            var statRegister = EqOperStatisticsActions.GetEquipmentLastStat(_eFunctions, reqItem.equipmentNumber, reqItem.operationStatisticType, reqItem.statisticDate);
                            var lastMeter = statRegister.MeterValue;
                            reqItem.meterReading = Convert.ToDecimal(_cells.GetCell(9, currentRow).Value) + Convert.ToDecimal(lastMeter);
                        }
                        else
                        {
                            reqItem.meterReading = Convert.ToDecimal(_cells.GetCell(9, currentRow).Value);
                        }

                        reqItem.meterReadingSpecified = true;

                        request.Add(reqItem);
                        var replySheet = statService.multipleAdjust(opContext, request.ToArray());
                        foreach (var reply in replySheet)
                        {
                            if (reply.errors.Length > 0)
                            {
                                var errors = "";
                                foreach (var er in reply.errors)
                                    errors = errors + "/" + er.messageText;
                                throw new Exception(errors);
                            }

                            _cells.GetCell(ResultColumn01, currentRow).Value = "ENVIADO";
                            _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Success;
                        }

                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Select();
                        currentRow++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:LoadStatistics()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        /// <summary>
        /// Establece el resultado de búsqueda de la descripción de un equipo después de que este es escrita
        /// </summary>
        /// <param name="target"></param>
        /// <param name="changedRanges"></param>
        void GetTableChangedValue(Excel.Range target, ListRanges changedRanges)//Excel.Range target)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            switch (target.Column)
            {
                case 3://Equipo
                    try
                    {
                        if (string.IsNullOrWhiteSpace("" + target.Value))
                        {
                            _cells.GetCell(target.Column + 1, target.Row).Value = "";
                            break;
                        }

                        _cells.GetCell(target.Column + 1, target.Row).Value = "Buscando Equipo...";
                        string description = EqOperStatisticsActions.GetEquipmentDescription(_eFunctions, "" + target.Value);

                        _cells.GetCell(target.Column + 1, target.Row).Value = !string.IsNullOrWhiteSpace(description) ? description.Trim() : "Equipo no encontrado";
                        _cells.GetCell(target.Column + 1, target.Row).Columns.AutoFit();
                    }
                    catch (NullReferenceException ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(target.Column + 1, target.Row).Value = "No fue Posible Obtener Informacion!";
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(target.Column + 1, target.Row).Value = "No fue Posible Obtener Informacion!";
                    }
                    break;
                case 5://Estadística
                    try
                    {
                        var equipNo = ""  + _cells.GetCell(3, target.Row).Value;
                        var statType = "" + MyUtilities.GetCodeKey(_cells.GetCell(5, target.Row).Value);

                        var statDate = "" + _cells.GetCell(1, target.Row).Value;

                        if (string.IsNullOrWhiteSpace(equipNo) || string.IsNullOrWhiteSpace(statType) || string.IsNullOrWhiteSpace(statDate))
                        {
                            _cells.GetCell(7, target.Row).Value = "No fue Posible Obtener Información";
                            _cells.GetCell(8, target.Row).Value = "No fue Posible Obtener Información";
                        }
                        else
                        {
                            var lastStatReg = EqOperStatisticsActions.GetEquipmentLastStat(_eFunctions, equipNo, statType, statDate);

                            _cells.GetCell(7, target.Row).Value = !string.IsNullOrWhiteSpace(lastStatReg.StatDate) ? lastStatReg.StatDate.Trim() : "";
                            _cells.GetCell(8, target.Row).Value = !string.IsNullOrWhiteSpace(lastStatReg.MeterValue) ? lastStatReg.MeterValue.Trim() : "";
                            
                        }
                        _cells.GetCell(7, target.Row).Columns.AutoFit();
                        _cells.GetCell(8, target.Row).Columns.AutoFit();
                    }
                    catch (NullReferenceException ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(7, target.Row).Value = "";
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(7, target.Row).Value = "Se ha producido un error";
                    }
                    break;
            }
        }

        

        

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(DeleteStatistics);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:DeleteStatistics()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            
        }

        private void ReviewStatisticList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.ClearTableRange(TableName01);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);


                _excelApp.EnableEvents = false;
                var resultColumn = ResultColumn01;

                //Obtengo los valores de las opciones de búsqueda
                var equipNo = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                var statType = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
                var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
                var finishDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);

                if (string.IsNullOrWhiteSpace(equipNo) || (string.IsNullOrWhiteSpace(startDate) && string.IsNullOrWhiteSpace(finishDate)))
                {
                    MessageBox.Show("Debe escribir un equipo y por lo menos una fecha de búsqueda", "Error de Consulta");
                    if (_cells != null) _cells.SetCursorDefault();
                    return;
                }

                List<StatRegister> statsList = EqOperStatisticsActions.ReviewEquipmentOperStatistics(_eFunctions, equipNo, statType, startDate, finishDate);
                var i = TitleRow01 + 1;
                foreach (var os in statsList)
                {
                    try
                    {
                        //Para resetear el estilo
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var selectedValue = "";
                        if (os.EntryType.Equals("C"))
                            selectedValue = os.CumValue;
                        else if (os.EntryType.Equals("D"))
                            selectedValue = os.StatValue;
                        else if (os.EntryType.Equals("M"))
                            selectedValue = os.MeterValue;
                        else
                            selectedValue = os.MeterValue;

                        _cells.GetCell(1, i).Value = "" + os.StatDate;
                        _cells.GetCell(2, i).Value = "'" + os.Shift;
                        _cells.GetCell(3, i).Value = "'" + os.EquipNo;
                        _cells.GetCell(4, i).Value = "" + os.EquipDesc1 + " " +os.EquipDesc2;
                        _cells.GetCell(5, i).Value = "'" + os.StatType;
                        _cells.GetCell(6, i).Value = "'" + os.EntryType;
                        _cells.GetCell(7, i).Value = "" + "CONSULTA MEDIDOR/VALOR";
                        _cells.GetCell(8, i).Value = "" + os.MeterValue;
                        _cells.GetCell(9, i).Value = "" + selectedValue;

                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ReviewStatisticList()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error", ex.Message);
                Debugger.LogError("RibbonEllipse.cs:ReviewStatisticList()", ex.Message);
            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
                _excelApp.EnableEvents = true;
            }
        }

        private void DeleteStatistics()
        {
            try
            {
                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    _cells.SetCursorWait();
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

                    var i = TitleRow01 + 1;
                    while ("" + _cells.GetCell(1, i).Value != "")
                    {
                        try
                        {
                            var opContext = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                maxInstancesSpecified = true,
                                returnWarnings = Debugger.DebugWarnings,
                                returnWarningsSpecified = true
                            };

                            var screenService = new Screen.ScreenService();

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            screenService.Url = urlService + "/ScreenService";
                            _eFunctions.RevertOperation(opContext, screenService);
                            //ejecutamos el programa
                            Screen.ScreenDTO reply = screenService.executeScreen(opContext, "MSO400");
                            //Validamos el ingreso
                            if (reply.mapName != "MSM400A") continue;

                            var statisticDate = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                            var shift = "" + "" + MyUtilities.GetCodeKey(_cells.GetCell(2, i).Value);
                            var equipmentNumber = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                            var operationStatisticType = "" + MyUtilities.GetCodeKey(_cells.GetCell(5, i).Value);
                            

                            var arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("OPTION1I", "3");
                            arrayFields.Add("STAT_DATE1I", statisticDate);
                            arrayFields.Add("STAT_TYPE1I", operationStatisticType);
                            arrayFields.Add("PLANT_NO1I", equipmentNumber);
                            arrayFields.Add("SHIFT1I", shift);

                            var request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenFields = arrayFields.ToArray(),
                                screenKey = "1"
                            };
                            reply = screenService.submit(opContext, request);

                            if (reply != null && !_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply))
                            {
                                arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("DELETE3I", "Y");

                                request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };

                                reply = screenService.submit(opContext, request);

                                if (reply != null && (_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM400A"))
                                    throw new ArgumentException(reply.message);

                                _cells.GetCell(ResultColumn01, i).Value = "ELIMINADO";
                                _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                            }
                            else if (reply != null) throw new Exception(reply.message);
                            else throw new Exception(@"No se ha podido obtener respuesta del servidor");
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                        }
                        finally
                        {
                            _cells.GetCell(ResultColumn01, i).Select();
                            i++;
                        }
                    } //--while de registros
                } //---if no se está en un ambiente válido
                else
                {
                    MessageBox.Show(@"\nSeleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeleteStatistics()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
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

        private void btnRestoreEvents_Click(object sender, RibbonControlEventArgs e)
        {
            RestoreEvents();
        }

        public void RestoreEvents()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var table = _cells.GetRange(TableName01).Worksheet.ListObjects[TableName01];
            var tableObject = Globals.Factory.GetVstoObject(table);
            try
            {
                tableObject.Change -= GetTableChangedValue;
            }
            catch
            {
                //ignored
            }
            tableObject.Change += GetTableChangedValue;

        }
    }
}

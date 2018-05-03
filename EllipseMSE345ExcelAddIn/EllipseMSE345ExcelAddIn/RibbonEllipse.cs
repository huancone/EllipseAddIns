using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMSE345ExcelAddIn.CondMeasurementService;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseStdTextClassLibrary;
using System.Threading;
using EllipseCommonsClassLibrary.Utilities;
// ReSharper disable FieldCanBeMadeReadOnly.Local

namespace EllipseMSE345ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;
        private const string SheetName01 = "MSE345";
        private const int TitleRow01 = 13;
        private const int ResultColumn01 = 12;
        private const string TableName01 = "CondMonitoringTable01";

        private const string SheetNameMtto01 = "MSE345_MTTO";
        private const int TitleRowMtto01 = 14;
        private const int ResultColumnMtto01 = 9;
        private const string TableNameMtto01 = "CondMonitoringTable01Mtto";

        private Worksheet _worksheet;
        private Microsoft.Office.Tools.Excel.Controls.DateTimePicker _fechaCalendario;
        private const string ValidationSheetName = "ValidationSheet";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = Environments.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void btnFormatGeneral_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnFormatMntto_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetMntto();
        }

        private void btnCreate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(LoadInfo);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameMtto01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(LoadInfoMntto);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoList()",
                    "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.CreateNewWorksheet(ValidationSheetName); //hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");
                _cells.MergeCells("C6", "L11");
                _cells.MergeCells("A12", "L12");

                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("M1", "XFD1048576").Columns.Hidden = true;

                
                _cells.GetRange("A6", "A11").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B6", "B11").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("A6").Value = "MONITOREO";
                var monitoreosCodeList = _eFunctions.GetItemCodes("OI").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell("B6"), monitoreosCodeList, ValidationSheetName, 1, false);

                _cells.GetCell("A7").Value = "EQUIPO";
                _cells.GetCell("B7").NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell("A8").Value = "FECHA";
                _cells.GetCell("B8").Value = DateTime.Now.ToString("yyyyMMdd");

                var inspectoresCodeList =_eFunctions.GetItemCodes("VI", "AND SUBSTR(TABLE_DESC,1,6)<='999999'").Select(item => item.code + " - " + item.description).ToList();
                _cells.GetCell("A9").Value = "INSPECTOR 1";
                _cells.SetValidationList(_cells.GetCell("B9"), inspectoresCodeList, ValidationSheetName, 2, false);
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A10").Value = "INSPECTOR 2";
                _cells.SetValidationList(_cells.GetCell("B10"), ValidationSheetName, 2, false);
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A11").Value = "INSPECTOR 3";
                _cells.SetValidationList(_cells.GetCell("B11"), ValidationSheetName, 2, false);
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(1, TitleRow01).Value = "COMPONENTE";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow01).Value = "MODIFICADOR";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Value = "POSICION";
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow01).Value = "CODIGO";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01).Value = "DESCRIPCION";
                _cells.GetCell(6, TitleRow01).Value = "CAUTION LOW";
                _cells.GetCell(7, TitleRow01).Value = "CAUTION";
                _cells.GetCell(8, TitleRow01).Value = "DANGER LOW";
                _cells.GetCell(9, TitleRow01).Value = "DANGER";
                _cells.GetCell(10, TitleRow01).Value = "VALOR ENCONTRADO";
                _cells.GetCell(10, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(11, TitleRow01).Value = "COMENTARIO";
                _cells.GetCell(11, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("B6").Select();

                //Changers
                _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var equipParamRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B7"), "equipParam");
                equipParamRange.Change += CondMonParam_Changed;

                var monTypeParamRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B6"), "monTypeParam");
                monTypeParamRange.Change += CondMonParam_Changed;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:FormatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void LoadInfo()
        {
            var monitorType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B6").Value);
            var monitorEquipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value);
            var monitorDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B8").Value);

            var inspector1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B9").Value);
            var inspector2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B10").Value);
            var inspector3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B11").Value);
            inspector1 = MyUtilities.GetCodeKey(inspector1);
            inspector2 = MyUtilities.GetCodeKey(inspector2);
            inspector3 = MyUtilities.GetCodeKey(inspector3);

            if (string.IsNullOrWhiteSpace(monitorType) || string.IsNullOrWhiteSpace(monitorDate) || string.IsNullOrWhiteSpace(monitorEquipment))
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Revíselos e Intente Nuevamente");
                return;
            }

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new CondMeasurementService.CondMeasurementService { Url = urlService + "/CondMeasurementService" };
            var stdTextOpContext = StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, false);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value))
            {
                try
                {
                    var componentCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var modifierCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var modifierPosition = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    componentCode = MyUtilities.GetCodeKey(componentCode);
                    modifierCode = MyUtilities.GetCodeKey(modifierCode);
                    modifierPosition = MyUtilities.GetCodeKey(modifierPosition);

                    var measurementCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(4, i).Value);
                    var comment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    var value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);


                    var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                    requestParamsSheet.equipmentRef = monitorEquipment;
                    requestParamsSheet.condMonType = monitorType;
                    requestParamsSheet.measureDate = monitorDate;
                    requestParamsSheet.visInsCode1 = inspector1;
                    requestParamsSheet.visInsCode2 = inspector2;
                    requestParamsSheet.visInsCode3 = inspector3;

                    requestParamsSheet.condMonMeas = measurementCode;
                    requestParamsSheet.compCode = componentCode;
                    requestParamsSheet.compModCode = modifierCode;
                    requestParamsSheet.condMonPos = modifierPosition;
                    requestParamsSheet.measureValue = Convert.ToDecimal(value);
                    requestParamsSheet.measureValueSpecified = true;

                    var reply = proxySheet.create(opSheet, requestParamsSheet);

                    if (!string.IsNullOrEmpty(comment))
                    {
                        var narrativeNoStdText = reply.stdTxtKey;//Prefix: ME
                        if (string.IsNullOrWhiteSpace(narrativeNoStdText))
                            throw new Exception("No se ha podido ingresar el comentario");

                        StdText.SetText(urlService, stdTextOpContext, narrativeNoStdText, comment);
                    }

                    _cells.GetCell(ResultColumn01, i).Value = "OK";
                    _cells.GetCell(ResultColumn01, i).Style = _cells.GetStyle(StyleConstants.Success);
                    _cells.GetCell(ResultColumn01, i).Select();
                }

                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Value = ex.Message;
                    _cells.GetCell(ResultColumn01, i).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(ResultColumn01, i).Select();
                    Debugger.LogError("RibbonEllipse:LoadInfo()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                }
                finally
                {
                    i++;
                }
            }
        }
        private void LoadInfoMntto()
        {
            var monitorType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value);
            var monitorEquipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B8").Value);
            var monitorDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B9").Value);

            var inspector1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B10").Value);
            var inspector2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B11").Value);
            var inspector3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B12").Value);
            inspector1 = MyUtilities.GetCodeKey(inspector1);
            inspector2 = MyUtilities.GetCodeKey(inspector2);
            inspector3 = MyUtilities.GetCodeKey(inspector3);

            if (string.IsNullOrWhiteSpace(monitorType) || string.IsNullOrWhiteSpace(monitorDate) || string.IsNullOrWhiteSpace(monitorEquipment))
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Revíselos e Intente Nuevamente");
                return;
            }

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new CondMeasurementService.CondMeasurementService { Url = urlService + "/CondMeasurementService" };
            var stdTextOpContext = StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, false);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };

            var i = TitleRowMtto01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var measurementCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value);
                    var value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    var comment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value);


                    var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                    requestParamsSheet.equipmentRef = monitorEquipment;
                    requestParamsSheet.condMonType = monitorType;
                    requestParamsSheet.measureDate = monitorDate;
                    requestParamsSheet.visInsCode1 = inspector1;
                    requestParamsSheet.visInsCode2 = inspector2;
                    requestParamsSheet.visInsCode3 = inspector3;

                    requestParamsSheet.condMonMeas = measurementCode;
                    requestParamsSheet.measureValue = Convert.ToDecimal(value);
                    requestParamsSheet.measureValueSpecified = true;

                    var reply = proxySheet.create(opSheet, requestParamsSheet);

                    if (!string.IsNullOrEmpty(comment))
                    {
                        var narrativeNoStdText = reply.stdTxtKey;//Prefix: ME
                        if (string.IsNullOrWhiteSpace(narrativeNoStdText))
                            throw new Exception("No se ha podido ingresar el comentario");

                        StdText.SetText(urlService, stdTextOpContext, narrativeNoStdText, comment);
                    }

                    _cells.GetCell(ResultColumnMtto01, i).Value = "OK";
                    _cells.GetCell(ResultColumnMtto01, i).Style = _cells.GetStyle(StyleConstants.Success);
                    _cells.GetCell(ResultColumnMtto01, i).Select();
                }

                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumnMtto01, i).Value = ex.Message;
                    _cells.GetCell(ResultColumnMtto01, i).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(ResultColumnMtto01, i).Select();
                    Debugger.LogError("RibbonEllipse:LoadInfo()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                }
                finally
                {
                    i++;
                }
            }
        }
        public List<string> GetFlotas()
        {
            var currentEnviroment = _eFunctions.GetCurrentEnviroment();
            try
            {
                _eFunctions.SetDBSettings(Environments.SigcorProductivo);

                const string sqlQuery = "SELECT DISTINCT TRIM(FLOTA_ELLIPSE) AS FLOTA " +
                                        "FROM SIGMAN.EQMTLIST WHERE FLOTA_ELLIPSE IS NOT NULL AND ACTIVE_FLG = 'Y' ORDER BY 1";

                var odr = _eFunctions.GetQueryResult(sqlQuery);
                var getFlotas = new List<string>();

                while (odr.Read())
                    getFlotas.Add("" + odr["FLOTA"]);
                return getFlotas;
            }
            finally
            {
                _eFunctions.SetDBSettings(currentEnviroment);
            }
        }

        public List<string> GetEquipos()
        {
            var currentEnviroment = _eFunctions.GetCurrentEnviroment();
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                const string sqlQuery = "SELECT EQUIP_NO FROM ELLIPSE.MSF600 " +
                                        "WHERE EQUIP_NO BETWEEN '0220701' AND '0220999' AND EQUIP_NO NOT IN ( '02209       ','02208       ') ORDER BY EQUIP_NO";

                var odr = _eFunctions.GetQueryResult(sqlQuery);
                var getEquipos = new List<string>();

                while (odr.Read())
                    getEquipos.Add("" + odr["EQUIP_NO"]);
                return getEquipos;
            }
            finally
            {
                _eFunctions.SetDBSettings(currentEnviroment);
            }
        }

        public void CondMonParam_Changed(Excel.Range target)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var currentRow = TitleRow01 + 1;
            try
            {

                _cells.SetCursorWait();

                var monitoringType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell("B6").Value));
                var equipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value);

                if (string.IsNullOrWhiteSpace(monitoringType) || string.IsNullOrWhiteSpace(equipment))
                    return;

                var list = GetMonitoringConditionList(equipment, monitoringType);
                _cells.ClearTableRange(TableName01);
                foreach (var item in list)
                {
                    _cells.GetCell(1, currentRow).Value = item.ComponentCode + (string.IsNullOrWhiteSpace(item.ComponentCode) ? "" : " - " + item.ComponentDescription);
                    _cells.GetCell(2, currentRow).Value = item.ModifierCode + (string.IsNullOrWhiteSpace(item.ModifierCode) ? "" : " - " + item.ModifierDescription);
                    _cells.GetCell(3, currentRow).Value = item.PositionCode + (string.IsNullOrWhiteSpace(item.PositionCode) ? "" : " - " + item.PositionDescription);
                    _cells.GetCell(4, currentRow).Value = item.MeassureCode;
                    _cells.GetCell(5, currentRow).Value = item.MeassureDescription;
                    _cells.GetCell(6, currentRow).Value = item.CautionLow;
                    _cells.GetCell(7, currentRow).Value = item.CautionUpper;
                    _cells.GetCell(8, currentRow).Value = item.DangerLow;
                    _cells.GetCell(9, currentRow).Value = item.DangerUpper;
                    currentRow++;
                }
            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void CondMonParamMntto_Changed(Excel.Range target)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var currentRow = TitleRowMtto01 + 1;
            try
            {
                _cells.SetCursorWait();

                var monitoringType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value));
                var equipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B8").Value);


                if (string.IsNullOrWhiteSpace(monitoringType) || string.IsNullOrWhiteSpace(equipment))
                    return;

                var list = GetMonitoringConditionList(equipment, monitoringType);
                _cells.ClearTableRange(TableNameMtto01);
                foreach (var item in list)
                {
                    _cells.GetCell(1, currentRow).Value = item.MeassureCode;
                    _cells.GetCell(2, currentRow).Value = item.MeassureDescription;
                    _cells.GetCell(3, currentRow).Value = item.CautionLow;
                    _cells.GetCell(4, currentRow).Value = item.CautionUpper;
                    _cells.GetCell(5, currentRow).Value = item.DangerLow;
                    _cells.GetCell(6, currentRow).Value = item.DangerUpper;
                    currentRow++;
                }
            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void FleetParam_Changed(Excel.Range target)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var currentEnviroment = _eFunctions.GetCurrentEnviroment();
            try
            {
                _cells.SetCursorWait();
                _eFunctions.SetDBSettings(Environments.SigcorProductivo);

                string sqlQuery = "SELECT EQU FROM SIGMAN.EQMTLIST WHERE FLOTA_ELLIPSE = '" + target.Value + "'" +
                                  " AND ACTIVE_FLG = 'Y' ORDER BY EQU ";

                var odr = _eFunctions.GetQueryResult(sqlQuery);
                var equipList = new List<string>();
                while (odr.Read())
                    equipList.Add("" + odr["EQU"]);

                _cells.SetValidationList(_cells.GetCell("B8"), equipList, ValidationSheetName, 3, false);
            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A15").Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            finally
            {
                _eFunctions.SetCurrentEnviroment(currentEnviroment);
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void FormatSheetMntto()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameMtto01;

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.CreateNewWorksheet(ValidationSheetName); //hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "H5");
                _cells.MergeCells("C6", "I12");
                _cells.MergeCells("A13", "I13");

                _cells.GetCell("I1").Value = "OBLIGATORIO";
                _cells.GetCell("I1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("I1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("I2").Value = "OPCIONAL";
                _cells.GetCell("I2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("I3").Value = "INFORMATIVO";
                _cells.GetCell("I3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("I4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("I4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("I5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("I5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("M1", "XFD1048576").Columns.Hidden = true;


                _cells.GetRange("A6", "A12").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B6", "B12").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("A6").Value = "FLOTA";
                var flotasList = GetFlotas();
                _cells.SetValidationList(_cells.GetCell("B6"), flotasList, ValidationSheetName, 1, false);

                _cells.GetCell("A7").Value = "MONITOREO";
                var monitoreosCodeList = _eFunctions.GetItemCodes("OI", "AND TRIM(TABLE_CODE) IN ('IE','UT')").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell("B7"), monitoreosCodeList, ValidationSheetName, 2, false);

                _cells.GetCell("A8").Value = "EQUIPO";
                _cells.GetCell("B8").NumberFormat = NumberFormatConstants.Text;
                var equipmentList = GetEquipos();
                _cells.SetValidationList(_cells.GetCell("B8"), equipmentList, ValidationSheetName, 3, false);


                _cells.GetCell("A9").Value = "FECHA";
                _cells.GetCell("B9").Value = DateTime.Now.ToString("yyyyMMdd");

                var inspectoresCodeList = _eFunctions.GetItemCodes("VI", "AND SUBSTR(TABLE_DESC,1,6)<='999999'").Select(item => item.code + " - " + item.description).ToList();
                _cells.GetCell("A10").Value = "INSPECTOR 1";
                _cells.SetValidationList(_cells.GetCell("B10"), inspectoresCodeList, ValidationSheetName, 4, false);
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A11").Value = "INSPECTOR 2";
                _cells.SetValidationList(_cells.GetCell("B11"), ValidationSheetName, 4, false);
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A12").Value = "INSPECTOR 3";
                _cells.SetValidationList(_cells.GetCell("B12"), ValidationSheetName, 4, false);
                _cells.GetCell("A12").Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, TitleRowMtto01, ResultColumnMtto01 - 1, TitleRowMtto01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(1, TitleRowMtto01).Value = "CODIGO";
                _cells.GetCell(1, TitleRowMtto01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRowMtto01).Value = "DESCRIPCION";
                _cells.GetCell(3, TitleRowMtto01).Value = "CAUTION LOW";
                _cells.GetCell(4, TitleRowMtto01).Value = "CAUTION";
                _cells.GetCell(5, TitleRowMtto01).Value = "DANGER LOW";
                _cells.GetCell(6, TitleRowMtto01).Value = "DANGER";
                _cells.GetCell(7, TitleRowMtto01).Value = "VALOR ENCONTRADO";
                _cells.GetCell(7, TitleRowMtto01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, TitleRowMtto01 + 1).Validation.Add(Excel.XlDVType.xlValidateWholeNumber, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlGreaterEqual, "0");

                _cells.GetCell(8, TitleRowMtto01).Value = "COMENTARIO";
                _cells.GetCell(8, TitleRowMtto01).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(ResultColumnMtto01, TitleRowMtto01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnMtto01, TitleRowMtto01).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, TitleRowMtto01 + 1, ResultColumnMtto01 - 3, TitleRowMtto01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowMtto01, ResultColumnMtto01, TitleRowMtto01 + 1), TableNameMtto01);

                _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _fechaCalendario = _worksheet.Controls.AddDateTimePicker(_cells.GetCell("B9"), "Calendario");
                _fechaCalendario.Format = DateTimePickerFormat.Short;
                _fechaCalendario.ValueChanged += CambioFecha;

                _cells.GetCell("B9").Value = _fechaCalendario.Value.ToString("yyyyMMdd");
                
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _cells.GetCell("B7").Select();

                //Changers
                _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var fleetParamRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B6"), "fleetParam");
                fleetParamRange.Change += FleetParam_Changed;

                var monTypeParamRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B7"), "monTypeParam");
                monTypeParamRange.Change += CondMonParamMntto_Changed;

                var equipParamRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B8"), "equipParam");
                equipParamRange.Change += CondMonParamMntto_Changed;
                //
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:FormatSheetMntto()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void CambioFecha(object sender, EventArgs e)
        {
            var picker = (Microsoft.Office.Tools.Excel.Controls.DateTimePicker) sender;
            _cells.GetCell("B9").Value = picker.Value.ToString("yyyyMMdd");


        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private List<MonitoringCondition> GetMonitoringConditionList(string equipment, string monitoringType)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var monitoringList = new List<MonitoringCondition>();

            var sqlQuery = "SELECT" +
                           "   MON.COND_MON_TYPE," +
                           "   SUBSTR(MON.COMP_MOD_DATA,1,4) COMP_CODE," +
                           "   CCOMP.TABLE_DESC COMP_CODE_DESC," +
                           "   SUBSTR(MON.COMP_MOD_DATA,5,2) MOD_CODE," +
                           "   CMODI.TABLE_DESC MOD_CODE_DESC," +
                           "   MON.COND_MON_POS," +
                           "   CPOSI.TABLE_DESC COND_MON_POS_DESC," +
                           "   MON.COND_MON_MEAS," +
                           "   TRIM(CMEAS.TABLE_DESC) AS COND_MON_MEAS_DESC," +
                           "   MON.MEAS_CAUT_LOWR," +
                           "   MON.MEAS_CAUT_UPPR," +
                           "   MON.MEAS_DANG_LOWR," +
                           "   MON.MEAS_DANG_UPPR" +
                           " FROM" +
                           "   ELLIPSE.MSF341 MON JOIN ELLIPSE.MSF340_SET_DEF SETM ON (SETM.TYPE_REFERENCE = MON.TYPE_REFERENCE " +
                           "   AND SETM.COND_MON_TYPE = MON.COND_MON_TYPE " +
                           "   AND SETM.COMP_MOD_DATA = MON.COMP_MOD_DATA " +
                           "   AND SETM.COND_MON_POS = MON.COND_MON_POS)" +
                           "   LEFT JOIN ELLIPSE.MSF010 CMEAS ON (CMEAS.TABLE_TYPE = 'MS' AND MON.COND_MON_MEAS = CMEAS.TABLE_CODE)" +
                           "   LEFT JOIN ELLIPSE.MSF010 CCOMP ON (CCOMP.TABLE_TYPE = 'CO' AND TRIM(SUBSTR(MON.COMP_MOD_DATA,1,4)) = TRIM(CCOMP.TABLE_CODE))" +
                           "   LEFT JOIN ELLIPSE.MSF010 CMODI ON (CMODI.TABLE_TYPE = 'MO' AND TRIM(SUBSTR(MON.COMP_MOD_DATA,5,2)) = TRIM(CMODI.TABLE_CODE))" +
                           "   LEFT JOIN ELLIPSE.MSF010 CPOSI ON (CPOSI.TABLE_TYPE = 'PM' AND TRIM(MON.COND_MON_POS) = TRIM(CPOSI.TABLE_CODE))" +
                           " WHERE" +
                           "   MON.COND_MON_TYPE = '" + monitoringType + "'" +
                           "   AND (MON.TYPE_REFERENCE = 'G' || (SELECT EQUIP_GRP_ID FROM ELLIPSE.MSF600 WHERE EQUIP_NO = '" + equipment + "') OR MON.TYPE_REFERENCE = 'E' ||'" + equipment + "')" +
                           " ORDER BY MON.COND_MON_TYPE, MON.COMP_MOD_DATA, MON.COND_MON_POS, MON.COND_MON_MEAS";
  

            var dr = _eFunctions.GetQueryResult(sqlQuery);

            while (dr.Read())
            {
                var item = new MonitoringCondition
                {
                    ComponentCode = dr["COMP_CODE"] + "",
                    ComponentDescription = dr["COMP_CODE_DESC"] + "",
                    ModifierCode = dr["MOD_CODE"] + "",
                    ModifierDescription = dr["MOD_CODE_DESC"] + "",
                    PositionCode = dr["COND_MON_POS"] + "",
                    PositionDescription = dr["COND_MON_POS_DESC"] + "",
                    MeassureCode = dr["COND_MON_MEAS"] + "",
                    MeassureDescription = dr["COND_MON_MEAS_DESC"] + "",
                    CautionLow = dr["MEAS_CAUT_LOWR"] + "",
                    CautionUpper = dr["MEAS_CAUT_UPPR"] + "",
                    DangerLow = dr["MEAS_DANG_LOWR"] + "",
                    DangerUpper = dr["MEAS_DANG_UPPR"] + ""
                };

                monitoringList.Add(item);
            }

            return monitoringList;
        }
        public class MonitoringCondition
        {
            public string Equipment;
            public string Egi;
            public string ComponentCode;
            public string ComponentDescription;
            public string ModifierCode;
            public string ModifierDescription;
            public string PositionCode;
            public string PositionDescription;
            public string MeassureCode;
            public string MeassureDescription;
            public string CautionLow;
            public string CautionUpper;
            public string DangerLow;
            public string DangerUpper;
            public string Type;
        }

    }
}
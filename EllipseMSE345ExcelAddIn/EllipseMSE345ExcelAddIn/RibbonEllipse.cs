using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
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
        private const string SheetName01Mtto = "MSE345_MTTO";

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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01Mtto)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(LoadManttoInfo);

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

        private void LoadManttoInfo()
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

            if (!string.IsNullOrEmpty("" + _cells.GetCell("B7").Value) &&
                !string.IsNullOrEmpty("" + _cells.GetCell("B8").Value) &&
                !string.IsNullOrEmpty("" + _cells.GetCell("B9").Value))
            {

                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    // Cells.getCell("A1").Value = "Conectado";

                    var proxySheet = new CondMeasurementService.CondMeasurementService();

                    var opSheet = new OperationContext();

                    var currentRow = 15;
                    string equipo = "" + _cells.GetCell("B8").Value;
                    string tipoMonitoreo = "" + _cells.GetCell("B7").Value;
                    string fecha = "" + _cells.GetCell("B9").Value;
                    string medida = "" + _cells.GetCell("A" + currentRow).Value;
                    string insp1 = "" + _cells.GetCell("B10").Value;
                    if (!string.IsNullOrEmpty(insp1))
                    {
                        insp1 = insp1.Substring(0, 2);
                    }
                    string insp2 = "" + _cells.GetCell("B11").Value;
                    if (!string.IsNullOrEmpty(insp2))
                    {
                        insp2 = insp2.Substring(0, 2);
                    }
                    string insp3 = "" + _cells.GetCell("B12").Value;
                    if (!string.IsNullOrEmpty(insp3))
                    {
                        insp3 = insp3.Substring(0, 2);
                    }
                    string comentario = "" + _cells.GetCell("H" + currentRow).Value;

                    while (!string.IsNullOrEmpty(medida))
                    {
                        if (string.IsNullOrEmpty(Convert.ToString(_cells.GetCell("G" + currentRow).Value)))
                        {
                            currentRow++;
                            medida = "" + _cells.GetCell("A" + currentRow).Value;
                            comentario = "" + _cells.GetCell("H" + currentRow).Value;
                        }
                        else
                        {
                            var value = Convert.ToDecimal(_cells.GetCell("G" + currentRow).Value);
                            try
                            {
                                var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                                proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) +
                                                 "/CondMeasurementService";

                                opSheet.district = _frmAuth.EllipseDsct;
                                opSheet.position = _frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                requestParamsSheet.equipmentRef = equipo;
                                requestParamsSheet.condMonType = tipoMonitoreo;
                                requestParamsSheet.measureDate = fecha;
                                requestParamsSheet.condMonMeas = medida;

                                requestParamsSheet.measureValue = Convert.ToDecimal(value);
                                requestParamsSheet.measureValueSpecified = true;

                                requestParamsSheet.visInsCode1 = insp1;

                                if (!string.IsNullOrEmpty(insp2))
                                {
                                    requestParamsSheet.visInsCode2 = insp2;
                                }

                                if (!string.IsNullOrEmpty(insp3))
                                {
                                    requestParamsSheet.visInsCode3 = insp3;
                                }

                                proxySheet.create(opSheet, requestParamsSheet);

                                if (!string.IsNullOrEmpty(comentario))
                                {
                                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                                    var sqlQuery =
                                        "select narrative_no from ellipse.msf345 where substr(99999999999999-rev_meas_data,1,8) = '" +
                                        fecha + "' and equip_no = '" + equipo + "' and trim(comp_pos_data) = '" +
                                        tipoMonitoreo + "' and trim(cond_mon_meas) = '" + medida + "'";

                                    var odr = _eFunctions.GetQueryResult(sqlQuery);
                                    var narrativeNo = "";
                                    if (odr.Read())
                                    {
                                        narrativeNo = odr["narrative_no"] + "";
                                    }

                                    StdText.SetText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label),
                                        StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100,
                                            false), "ME" + narrativeNo, comentario);
                                }

                                _cells.GetCell("I" + currentRow).Value = "OK";
                                _cells.GetCell("I" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                                _cells.GetCell("I" + currentRow).Select();
                            }
                            catch (Exception ex)
                            {
                                _cells.GetCell("I" + currentRow).Value = ex.Message;
                                _cells.GetCell("I" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                _cells.GetCell("I" + currentRow).Select();
                                Debugger.LogError("RibbonEllipse:startLabourCostLoad()",
                                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                    ex.StackTrace);

                            }
                            finally
                            {
                                currentRow++;
                                medida = "" + _cells.GetCell("A" + currentRow).Value;
                                comentario = "" + _cells.GetCell("H" + currentRow).Value;
                            }
                        }
                    }
                    MessageBox.Show(@"Proceso Finalizado Correctamente");
                }
            }
            else
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Reviselos e Intente Nuevamente");
            }
        }

        private void LoadInfo()
        {
            var monitorType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B6").Value);
            var monitorDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value);
            var monitorEquipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B11").Value);

            var inspector1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B8").Value);
            var inspector2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B9").Value);
            var inspector3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B10").Value);
            inspector1 = MyUtilities.GetCodeKey(inspector1);
            inspector2 = MyUtilities.GetCodeKey(inspector2);
            inspector3 = MyUtilities.GetCodeKey(inspector3);

            if (!string.IsNullOrWhiteSpace(monitorType) || !string.IsNullOrWhiteSpace(monitorDate) || !string.IsNullOrWhiteSpace(monitorEquipment))
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Reviselos e Intente Nuevamente");
                return;
            }

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new CondMeasurementService.CondMeasurementService {Url = urlService + "/CondMeasurementService"};
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

                    proxySheet.create(opSheet, requestParamsSheet);

                    if (!string.IsNullOrEmpty(comment))
                    {
                        _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                        var sqlQuery =
                            "SELECT NARRATIVE_NO FROM ELLIPSE.MSF345 " +
                            "WHERE SUBSTR(99999999999999-REV_MEAS_DATA,1,8) = '" + monitorDate + "'" +
                            " AND EQUIP_NO = '" + monitorEquipment + "'" +
                            " AND TRIM(COMP_POS_DATA) = '" + monitorType + "'" +
                            " AND TRIM(COND_MON_MEAS) = '" + measurementCode + "'";

                        var odr = _eFunctions.GetQueryResult(sqlQuery);

                        var narrativeNo = "";
                        if (odr.Read())
                        {
                            narrativeNo = odr["NARRATIVE_NO"] + "";
                        }

                        if (string.IsNullOrWhiteSpace(narrativeNo))
                            throw new Exception("No se ha podido ingresar el comentario");

                        StdText.SetText(urlService, stdTextOpContext, "ME" + narrativeNo, comment);
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

        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
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
                var monitoreosCodeList =
                    _eFunctions.GetItemCodes("OI").Select(item => item.code + " - " + item.description).ToList();
                ;
                _cells.SetValidationList(_cells.GetCell("B6"), monitoreosCodeList, ValidationSheetName, 1, false);
                _cells.GetCell("A7").Value = "FECHA";
                _cells.GetCell("B7").Value = DateTime.Now.ToString("yyyyMMdd");

                var inspectoresCodeList =
                    _eFunctions.GetItemCodes("VI", "AND SUBSTR(TABLE_DESC,1,6)<='999999'")
                        .Select(item => item.code + " - " + item.description)
                        .ToList();
                ;
                _cells.GetCell("A8").Value = "INSPECTOR 1";
                _cells.SetValidationList(_cells.GetCell("B8"), inspectoresCodeList, ValidationSheetName, 2, false);
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A9").Value = "INSPECTOR 2";
                _cells.SetValidationList(_cells.GetCell("B9"), ValidationSheetName, 2, false);
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A10").Value = "INSPECTOR 3";
                _cells.SetValidationList(_cells.GetCell("B10"), ValidationSheetName, 2, false);
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("A11").Value = "EQUIPO";
                _cells.GetCell("B11").NumberFormat = NumberFormatConstants.Text;
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos()); //TO DO


                _cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01).Style = StyleConstants.TitleRequired;

                _cells.GetCell(1, TitleRow01).Value = "COMPONENTE";
                _cells.GetCell(2, TitleRow01).Value = "MODIFICADOR";
                _cells.GetCell(3, TitleRow01).Value = "POSICION";
                _cells.GetCell(4, TitleRow01).Value = "CODIGO";
                _cells.GetCell(5, TitleRow01).Value = "DESCRIPCION";
                _cells.GetCell(6, TitleRow01).Value = "CAUTION LOW";
                _cells.GetCell(7, TitleRow01).Value = "CAUTION";
                _cells.GetCell(8, TitleRow01).Value = "DANGER LOW";
                _cells.GetCell(9, TitleRow01).Value = "DANGER";
                _cells.GetCell(10, TitleRow01).Value = "VALOR ENCONTRADO";
                _cells.GetCell(11, TitleRow01).Value = "COMENTARIO";
                _cells.GetCell(11, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(12, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(12, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("B6").Select();

                //Changers
                _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var groupRange = _worksheet.Controls.AddNamedRange(_cells.GetCell("B11"), "GroupRange");

                groupRange.Change += CondMonParam_Changed;

                var groupCells2 = _worksheet.Range["B6:B6"];
                var groupRange2 = _worksheet.Controls.AddNamedRange(groupCells2, "GroupRange2");

                groupRange2.Change += CondMonParam_Changed;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void Limpiar()
        {
            var row = 14;
            var max = 200;
            while (row <= max)
            {
                _cells.GetCell("A" + row).Clear();
                _cells.GetCell("B" + row).Clear();
                _cells.GetCell("C" + row).Clear();
                _cells.GetCell("D" + row).Clear();
                _cells.GetCell("E" + row).Clear();
                _cells.GetCell("F" + row).Clear();
                _cells.GetCell("G" + row).Clear();
                _cells.GetCell("H" + row).Clear();
                _cells.GetCell("I" + row).Clear();
                _cells.GetCell("J" + row).Clear();
                _cells.GetCell("K" + row).Clear();
                _cells.GetCell("L" + row).Clear();
                row++;
            }
        }

        private void LimpiarMtto()
        {
            var row = 15;
            var max = 200;
            while (row <= max)
            {
                _cells.GetCell("A" + row).Clear();
                _cells.GetCell("B" + row).Clear();
                _cells.GetCell("C" + row).Clear();
                _cells.GetCell("D" + row).Clear();
                _cells.GetCell("E" + row).Clear();
                _cells.GetCell("F" + row).Clear();
                _cells.GetCell("G" + row).Clear();
                _cells.GetCell("H" + row).Clear();
                _cells.GetCell("I" + row).Clear();
                _cells.GetCell("J" + row).Clear();
                _cells.GetCell("K" + row).Clear();
                _cells.GetCell("L" + row).Clear();
                row++;
            }
        }

        public List<string> GetFlotas()
        {
            _eFunctions.SetDBSettings(Environments.SigcorProductivo);

            const string sqlQuery =
                "SELECT DISTINCT TRIM(FLOTA_ELLIPSE) AS FLOTA FROM EQMTLIST WHERE FLOTA_ELLIPSE IS NOT NULL AND ACTIVE_FLG = 'Y' ORDER BY 1";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getFlotas = new List<string>();

            while (odr.Read())
            {
                getFlotas.Add("" + odr["FLOTA"]);
            }
            return getFlotas;
        }

        public List<string> GetEquipos()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery =
                "select equip_no from ellipse.msf600 where equip_no between '0220701' and '0220999' and equip_no not in ( '02209       ','02208       ') order by 1";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getEquipos = new List<string>();

            while (odr.Read())
            {
                getEquipos.Add("" + odr["equip_no"]);
            }
            return getEquipos;
        }

        public List<string> GetMonitoreosMtto()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery =
                "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'OI' and trim(table_code) in ('IE','UT')";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getMonitoreos = new List<string>();

            while (odr.Read())
            {
                getMonitoreos.Add("" + odr["table_code"]);
            }
            return getMonitoreos;
        }

        public void CondMonParam_Changed(Excel.Range target)
        {
            var currentRow = TitleRow01 + 1;
            try
            {
                var monitoringType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell("B6").Value));
                var date = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B7").Value);
                var inspector1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B8").Value);
                var inspector2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B9").Value);
                var inspector3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B10").Value);
                var equipment = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B11").Value);

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
            }
        }

        public void changesGroupRange_ChangeMTTO(Excel.Range target)
        {
            var currentRow = 15;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +
                                  "trim(m.cond_mon_meas) as codigo, " +
                                  "trim(t.table_desc) as descripcion, " +
                                  "trim(m.meas_caut_lowr) as caution_low, " +
                                  "trim(m.meas_caut_uppr) as caution, " +
                                  "trim(m.meas_dang_lowr) as danger_low, " +
                                  "trim(m.meas_dang_uppr) as danger " +
                                  "from " +
                                  "ellipse.msf341 m, " +
                                  "ellipse.msf340_set_def s, " +
                                  "ellipse.msf010 t " +
                                  "where " +
                                  "m.cond_mon_type = '" + _cells.GetCell(target.Column, target.Row - 1).Value + "'" +
                                  "and " +
                                  "( " +
                                  "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" +
                                  target.Value + "')" +
                                  "or m.type_reference = 'E'||'" + target.Value + "'" +
                                  ") " +
                                  "and t.table_type = 'MS' " +
                                  "and t.table_code = m.cond_mon_meas " +
                                  "and s.type_reference = m.type_reference " +
                                  "and s.cond_mon_type = m.cond_mon_type " +
                                  "and s.comp_mod_data = m.comp_mod_data " +
                                  "and s.cond_mon_pos = m.cond_mon_pos " +
                                  "and s.status_340 = 'A' " +
                                  "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                LimpiarMtto();

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["codigo"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["descripcion"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["caution_low"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["caution"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["danger_low"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["danger"] + "";

                    currentRow++;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                //TO DO Centrar
                ValidacionMtto();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_Change2MTTO(Excel.Range target)
        {
            var currentRow = 15;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +
                                  "trim(m.cond_mon_meas) as codigo, " +
                                  "trim(t.table_desc) as descripcion, " +
                                  "trim(m.meas_caut_lowr) as caution_low, " +
                                  "trim(m.meas_caut_uppr) as caution, " +
                                  "trim(m.meas_dang_lowr) as danger_low, " +
                                  "trim(m.meas_dang_uppr) as danger " +
                                  "from " +
                                  "ellipse.msf341 m, " +
                                  "ellipse.msf340_set_def s, " +
                                  "ellipse.msf010 t " +
                                  "where " +
                                  "m.cond_mon_type = '" + target.Value + "'" +
                                  "and " +
                                  "( " +
                                  "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" +
                                  _cells.GetCell(target.Column, target.Row + 1).Value + "')" +
                                  "or m.type_reference = 'E'||'" + _cells.GetCell(target.Column, target.Row + 1).Value +
                                  "'" +
                                  ") " +
                                  "and t.table_type = 'MS' " +
                                  "and t.table_code = m.cond_mon_meas " +
                                  "and s.type_reference = m.type_reference " +
                                  "and s.cond_mon_type = m.cond_mon_type " +
                                  "and s.comp_mod_data = m.comp_mod_data " +
                                  "and s.cond_mon_pos = m.cond_mon_pos " +
                                  "and s.status_340 = 'A' " +
                                  "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                LimpiarMtto();

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["codigo"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["descripcion"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["caution_low"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["caution"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["danger_low"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["danger"] + "";

                    currentRow++;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                //TO DO Centrar
                ValidacionMtto();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_FLOTA(Excel.Range target)
        {
            try
            {
                _eFunctions.SetDBSettings(Environments.SigcorProductivo);

                string sqlQuery = "SELECT EQU FROM EQMTLIST WHERE FLOTA_ELLIPSE = '" + target.Value + "'" +
                                  " AND ACTIVE_FLG = 'Y' ORDER BY 1 ";

                var odr = _eFunctions.GetQueryResult(sqlQuery);


                var getFlotas = new List<string>();

                while (odr.Read())
                {
                    getFlotas.Add("" + odr["EQU"]);
                }

                //Cells.getCell("B12").Value2 = "";
                _cells.SetValidationList(_cells.GetCell("B8"), getFlotas);

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
        }



        public void FormatSheetMntto()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                Excel.Worksheet excelSheet = excelBook.ActiveSheet;
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01Mtto;

                _worksheet = Globals.Factory.GetVstoObject(excelSheet);

                if (_cells == null)

                    _cells = new ExcelStyleCells(_excelApp);

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

                _cells.GetRange("J1", "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A14").Value = "CODIGO";
                _cells.GetCell("A14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("B14").Value = "DESCRIPCION";
                _cells.GetCell("B14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("C14").Value = "CAUTION LOW";
                _cells.GetCell("C14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("D14").Value = "CAUTION";
                _cells.GetCell("D14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("E14").Value = "DANGER LOW";
                _cells.GetCell("E14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("F14").Value = "DANGER";
                _cells.GetCell("F14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("G14").Value = "VALOR ENCONTRADO";
                _cells.GetCell("G14").Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("H14").Value = "COMENTARIO";
                _cells.GetCell("H14").Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("I14").Value = "RESULTADO";
                _cells.GetCell("I14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "H5");
                _cells.GetRange("B1", "H5").Borders.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", "H5").Borders.Weight = "2";
                _cells.MergeCells("C6", "I12");
                _cells.GetRange("C6", "I12").Borders.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("C6", "I12").Borders.Weight = "2";
                _cells.MergeCells("A13", "I13");
                _cells.GetRange("A13", "I13").Borders.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A13", "I13").Borders.Weight = "2";

                _cells.GetCell("A6").Value = "FLOTA";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationList(_cells.GetCell("B6"), GetFlotas());
                _cells.GetCell("B6").NumberFormat = "@";
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos());
                _cells.GetCell("B6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B6").Borders.Weight = "2";
                _cells.GetCell("B6").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B6").Font.Bold = true;

                _cells.GetCell("A7").Value = "MONITOREO";
                _cells.GetCell("A7").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationList(_cells.GetCell("B7"), GetMonitoreosMtto());
                _cells.GetCell("B7").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B7").Borders.Weight = "2";
                _cells.GetCell("B7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B7").Font.Bold = true;

                _cells.GetCell("A8").Value = "EQUIPO";
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B8").NumberFormat = "@";
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos());
                _cells.GetCell("B8").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B8").Borders.Weight = "2";
                _cells.GetCell("B8").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B8").Font.Bold = true;

                _cells.GetCell("A9").Value = "FECHA";
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B9").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                _cells.GetCell("B9").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B9").Borders.Weight = "2";
                _cells.GetCell("B9").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B9").Font.Bold = true;

                var inspectoresCodeList =
                    _eFunctions.GetItemCodes("VI", "AND SUBSTR(TABLE_DESC,1,6)<='999999'")
                        .Select(item => item.code + " - " + item.description)
                        .ToList();
                ;

                _cells.GetCell("A10").Value = "INSPECTOR 1";
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B10"), inspectoresCodeList, ValidationSheetName, 2, false);
                _cells.GetCell("B10").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B10").Borders.Weight = "2";
                _cells.GetCell("B10").Font.Bold = true;

                _cells.GetCell("A11").Value = "INSPECTOR 2";
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B11"), ValidationSheetName, 2, false);
                _cells.GetCell("B11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B11").Borders.Weight = "2";
                _cells.GetCell("B11").Font.Bold = true;

                _cells.GetCell("A12").Value = "INSPECTOR 3";
                _cells.GetCell("A12").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B12"), ValidationSheetName, 2, false);
                _cells.GetCell("B12").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B12").Borders.Weight = "2";
                _cells.GetCell("B12").Font.Bold = true;

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _fechaCalendario = _worksheet.Controls.AddDateTimePicker(_cells.GetCell("B9"), "Calendario");

                _fechaCalendario.ValueChanged += CambioFecha;

                _cells.GetCell("B9").Value = _fechaCalendario.Value.ToString("yyyyMMdd");

                _cells.GetCell("B9").Select();

                //Changers

                //TO DO Centrar
                ValidacionMtto();
                _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var groupCells = _worksheet.Range["B8:B8"];
                var groupRange = _worksheet.Controls.AddNamedRange(groupCells, "GroupRange");

                groupRange.Change += changesGroupRange_ChangeMTTO;

                var groupCells2 = _worksheet.Range["B7:B7"];
                var groupRange2 = _worksheet.Controls.AddNamedRange(groupCells2, "GroupRange2");

                groupRange2.Change += changesGroupRange_Change2MTTO;

                var groupCells4 = _worksheet.Range["B6:B6"];
                var groupRange4 = _worksheet.Controls.AddNamedRange(groupCells4, "GroupRange4");

                groupRange4.Change += changesGroupRange_FLOTA;

                //
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void CambioFecha(object sender, EventArgs e)
        {
            var picker = (Microsoft.Office.Tools.Excel.Controls.DateTimePicker) sender;
            _cells.GetCell("B9").Value = picker.Value.ToString("yyyyMMdd");


        }

        public void ValidacionMtto()
        {
            _cells.GetCell("G15:G200").Validation.Delete();
            _cells.GetCell("G15:G200").Validation.Add(
                Excel.XlDVType.xlValidateWholeNumber,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Excel.XlFormatConditionOperator.xlGreaterEqual,
                "0");
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
                           "   AND (MON.TYPE_REFERENCE = 'G' || (SELECT EQUIP_GRP_ID FROM ELLIPSE.MSF600 WHERE EQUIP_NO = '" + equipment + "') OR MON.TYPE_REFERENCE = 'E' ||'" + equipment + "')";
  

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
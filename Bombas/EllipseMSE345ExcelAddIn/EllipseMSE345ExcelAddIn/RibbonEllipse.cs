using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMSE345ExcelAddIn.CondMeasurementService;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseCommonsClassLibrary;
using EllipseStdTextClassLibrary;

namespace EllipseMSE345ExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        private const string SheetName01 = "MSE345";
        private const bool DebugErrors = false;
        string _narrativeNo;
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            _eFunctions.DebugQueries = false;
            _eFunctions.DebugErrors = false;
            _eFunctions.DebugWarnings = false;
            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void Crear_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(Cargar_Info_Estandar);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");


        }

        private void Cargar_Info_Estandar()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var currentRow = 7;


            while (!string.IsNullOrEmpty("" + _cells.GetCell("A" + currentRow).Value))
            {

                var proxySheet = new CondMeasurementService.CondMeasurementService();
                var opSheet = new OperationContext();

                var monitorType = _cells.GetEmptyIfNull(_cells.GetCell("A" + currentRow ).Value);
                var monitorEquipment = _cells.GetEmptyIfNull(_cells.GetCell("B" + currentRow).Value);
                var monitorDate = _cells.GetEmptyIfNull(_cells.GetCell("C" + currentRow).Value);

                var insp1 = "" + _cells.GetCell("D" + currentRow).Value;
                if (!string.IsNullOrEmpty(insp1))
                {
                    insp1 = insp1.Substring(0, 2);
                }
                var insp2 = "" + _cells.GetCell("E" + currentRow).Value;
                if (!string.IsNullOrEmpty(insp2))
                {
                    insp2 = insp2.Substring(0, 2);
                }
                var insp3 = "" + _cells.GetCell("F" + currentRow).Value;
                if (!string.IsNullOrEmpty(insp3))
                {
                    insp3 = insp3.Substring(0, 2);
                }

                var componentCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("G" + currentRow).Value);
                var modifierCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("H" + currentRow).Value);
                var modifierPosition = _cells.GetNullOrTrimmedValue(_cells.GetCell("I" + currentRow).Value);
                var measurementCode = "" + _cells.GetCell("J" + currentRow).Value;
                var valor = Convert.ToDecimal(_cells.GetCell("K" + currentRow).Value);
                var comentario = "" + _cells.GetCell("L" + currentRow).Value;

                componentCode = componentCode != null && componentCode.Length > 4 && componentCode.Contains("-") ? componentCode.Substring(0, 4) : componentCode;
                modifierCode = modifierCode != null && modifierCode.Length > 4 && modifierCode.Contains("-") ? modifierCode.Substring(0, 4) : modifierCode;
                modifierPosition = modifierPosition != null && modifierPosition.Length > 4 && modifierPosition.Contains("-") ? modifierPosition.Substring(0, 4) : modifierPosition;
                try
                {
                    if (string.IsNullOrEmpty(Convert.ToString(_cells.GetCell("J" + currentRow).Value))) continue;
                
                    var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                    proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                    opSheet.district = _frmAuth.EllipseDsct;
                    opSheet.position = _frmAuth.EllipsePost;
                    opSheet.maxInstances = 100;
                    opSheet.returnWarnings = _eFunctions.DebugWarnings;

                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                    requestParamsSheet.equipmentRef = monitorEquipment;
                    requestParamsSheet.condMonType = monitorType;
                    requestParamsSheet.measureDate = monitorDate;
                    requestParamsSheet.condMonMeas = measurementCode;
                    requestParamsSheet.compCode = componentCode;
                    requestParamsSheet.compModCode = modifierCode;
                    requestParamsSheet.condMonPos = modifierPosition;

                    requestParamsSheet.measureValue = Convert.ToDecimal(valor);
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

                        var sqlQuery = "select narrative_no from ellipse.msf345 where substr(99999999999999-rev_meas_data,1,8) = '" + monitorDate + "' and equip_no = '" + monitorEquipment + "' and trim(comp_pos_data) = '" + monitorType + "' and trim(cond_mon_meas) = '" + measurementCode + "'";

                        var odr = _eFunctions.GetQueryResult(sqlQuery);
                        if (odr.Read())
                        {
                            _narrativeNo = odr["narrative_no"] + "";
                        }

                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, false), "ME" + _narrativeNo, comentario);
                    }

                    _cells.GetCell("L" + currentRow).Value = "OK";
                    _cells.GetCell("L" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                    _cells.GetCell("L" + currentRow).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell("L" + currentRow).Value = ex.Message;
                    _cells.GetCell("L" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                    _cells.GetCell("L" + currentRow).Select();
                    Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);

                }
                finally
                {
                    currentRow++;
                }
            }
            MessageBox.Show(@"Proceso Finalizado Correctamente");
        }

        public void SetSheetHeaderData()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

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

                _cells.GetCell("A6").Value = "MONITOREO";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("B6").Value = "EQUIPO";
                _cells.GetCell("B6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B6").NumberFormat = "@";
                _cells.GetCell("B6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                _cells.GetCell("C6").Value = "FECHA";
                _cells.GetCell("C6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("C6").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                _cells.GetCell("C6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("C6").Borders.Weight = "2";

                _cells.GetCell("D6").Value = "INSPECTOR 1";
                _cells.GetCell("D6").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("D6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("D6").Borders.Weight = "2";

                _cells.GetCell("E6").Value = "INSPECTOR 2";
                _cells.GetCell("E6").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("E6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("E6").Borders.Weight = "2";

                _cells.GetCell("F6").Value = "INSPECTOR 3";
                _cells.GetCell("F6").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("F6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("F6").Borders.Weight = "2";

                _cells.GetCell("G6").Value = "COMPONENTE";
                _cells.GetCell("G6").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("H6").Value = "MODIFICADOR";
                _cells.GetCell("H6").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("I6").Value = "POSICION";
                _cells.GetCell("I6").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("J6").Value = "CODIGO MEDIDA";
                _cells.GetCell("J6").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K6").Value = "VALOR ENCONTRADO";
                _cells.GetCell("K6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L6").Value = "COMENTARIO";
                _cells.GetCell("L6").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("M6").Value = "RESULTADO";
                _cells.GetCell("M6").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");
                _cells.GetRange("B1", "K5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", "K5").Borders.Weight = "2";

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("A7").Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void btnStopProcess_Click(object sender, RibbonControlEventArgs e)
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

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                _excelApp.ActiveWorkbook.Worksheets.Add();

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

            SetSheetHeaderData();

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }
    }
}
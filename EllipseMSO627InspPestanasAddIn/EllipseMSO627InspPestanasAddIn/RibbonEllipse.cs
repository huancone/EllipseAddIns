using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseMSO627InspPestanasAddIn
{
    public partial class RibbonEllipse
    {
        private const int TittleRow = 7;
        private const int ResultColumn = 7;
        private const int ResultColumnLimpieza = 8;
        private const int MaxRows = 8;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private readonly string _sheetName01 = "Inspeccion Pestañas";
        private readonly string _sheetName02 = "Limpieza de Carbon";
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Format();
        }


        private void btnLimpiezaCarbon_Click(object sender, RibbonControlEventArgs e)
        {
            FormatLimpieza();
        }

        private void FormatLimpieza()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                excelSheet.Name = _sheetName02;

                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).ClearFormats();
                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).ClearComments();
                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).Clear();

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "REGISTRO DE LIMPIEZA DE CARBON";
                _cells.GetRange("B1", "D2").Merge();
                _cells.GetRange("B1", "D2").WrapText = true;

                _cells.GetCell(1, TittleRow).Value = "Fecha";
                _cells.GetCell(1, TittleRow).AddComment("YYYYMMDD");
                _cells.GetCell(2, TittleRow).Value = "Grupo";
                _cells.GetCell(3, TittleRow).Value = "Descripcion";
                _cells.GetCell(4, TittleRow).Value = "Cantidad";
                _cells.GetCell(5, TittleRow).Value = "Usuario";
                _cells.GetCell(6, TittleRow).Value = "Equipo";
                _cells.GetCell(7, TittleRow).Value = "Accion Correctiva";
                _cells.GetCell(ResultColumnLimpieza, TittleRow).Value = "Resultado";

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

                #region Styles

                _cells.GetCell(1, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumnLimpieza, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).NumberFormat = "@";

                #endregion

                _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumnLimpieza, MaxRows), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                _excelSheetItems.Name = "MSO627Limpieza";

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();

                #region Selection List Options

                #endregion

                _cells.GetRange(1, TittleRow + 1, ResultColumnLimpieza, MaxRows).NumberFormat = "@";
            }
            catch (Exception error)
            {
                _cells.GetCell(ResultColumnLimpieza, TittleRow).Value += " Error " + error.Message;
            }
        }

        private void Format()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                excelSheet.Name = _sheetName01;

                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Clear();

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "REGISTRO DE REVISION DE PESTAÑAS EN VAGONES Y LOCOMOTORAS";
                _cells.GetRange("B1", "D2").Merge();
                _cells.GetRange("B1", "D2").WrapText = true;

                _cells.GetCell(1, TittleRow).Value = "Fecha";
                _cells.GetCell(1, TittleRow).AddComment("YYYYMMDD");
                _cells.GetCell(2, TittleRow).Value = "Grupo";
                _cells.GetCell(3, TittleRow).Value = "Descripcion";
                _cells.GetCell(4, TittleRow).Value = "Conformidad";
                _cells.GetCell(5, TittleRow).Value = "Usuario";
                _cells.GetCell(6, TittleRow).Value = "Equipo";
                _cells.GetCell(ResultColumn, TittleRow).Value = "Resultado";

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

                #region Styles

                _cells.GetCell(1, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";

                #endregion

                _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,_cells.GetRange(1, TittleRow, ResultColumn, MaxRows), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                _excelSheetItems.Name = "MSO627Data";

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();

                #region Selection List Options

                var optionList = new List<string>
                {
                    "CO - CONFORME",
                    "NC - NO CONFORME"
                };
                _cells.SetValidationList(_cells.GetRange(4, TittleRow + 1, 4, MaxRows), optionList);

                #endregion

                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";
            }
            catch (Exception error)
            {
                _cells.GetCell(ResultColumn, TittleRow).Value += " Error " + error.Message;
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadInspPestanas);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadLimpieza);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void LoadLimpieza()
        {

            var opSheet = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

            var currentRow = TittleRow + 1;
            while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
            {
                try
                {
                    var fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                    var grupo = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                    var descripcion = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                    var cantidad = "" + _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                    var usuario = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                    var equipo = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);
                    var accionCorrectiva = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);

                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO627");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumnLimpieza, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumnLimpieza, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "1");
                        arrayFields.Add("WORK_GROUP1I", grupo);
                        arrayFields.Add("RAISED_DATE1I", fecha);
                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(ResultColumnLimpieza, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumnLimpieza, currentRow).Value = replySheet.message;
                        }
                        else if (replySheet.mapName == "MSM627B")
                        {
                            try
                            {
                                arrayFields = new ArrayScreenNameValue();

                                arrayFields.Add("RAISED_TIME2I1", "00:00");
                                arrayFields.Add("INCIDENT_DESC2I1", descripcion);
                                arrayFields.Add("MAINT_TYPE2I1", "NM");
                                arrayFields.Add("JOB_DUR_HOURS2I1", cantidad);
                                arrayFields.Add("ORIGINATOR_ID2I1", usuario);
                                arrayFields.Add("JOB_DUR_FINISH2I1", "00:00");
                                arrayFields.Add("EQUIP_REF2I1", equipo);
                                arrayFields.Add("CORRECT_DESC2I1", accionCorrectiva);
                                requestSheet.screenFields = arrayFields.ToArray();

                                requestSheet.screenKey = "1";
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (_eFunctions.CheckReplyWarning(replySheet) ||
                                       replySheet.functionKeys == "XMIT-Confirm")
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                if (_eFunctions.CheckReplyError(replySheet))
                                {
                                    _cells.GetCell(ResultColumnLimpieza, currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(ResultColumnLimpieza, currentRow).Value = replySheet.message;
                                }
                                else
                                    _cells.GetCell(ResultColumnLimpieza, currentRow).Style = StyleConstants.Success;
                            }
                            catch (Exception ex)
                            {
                                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(ResultColumnLimpieza, currentRow).Value = "ERROR: " + ex.Message;
                                Debugger.LogError("RibbonEllipse.cs:MSO627LoadLimpieza()", ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnLimpieza, currentRow).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:MSO627LoadLimpieza()", ex.Message);
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void LoadInspPestanas()
        {
            var opSheet = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

            var currentRow = TittleRow + 1;
            while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
            {
                try
                {
                    var fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                    var grupo = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                    var descripcion = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                    var conformidad = (_cells.GetEmptyIfNull((_cells.GetCell(4, currentRow).Value)).Length >= 2)
                        ? _cells.GetNullOrTrimmedValue(_cells.GetCell(4, currentRow).Value).Substring(0, 2)
                        : null;
                    var usuario = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                    var equipo = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);

                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO627");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "1");
                        arrayFields.Add("WORK_GROUP1I", grupo);
                        arrayFields.Add("RAISED_DATE1I", fecha);
                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                        }
                        else if (replySheet.mapName == "MSM627B")
                        {
                            try
                            {
                                arrayFields = new ArrayScreenNameValue();

                                arrayFields.Add("RAISED_TIME2I1", "00:00");
                                arrayFields.Add("INCIDENT_DESC2I1", descripcion);
                                arrayFields.Add("MAINT_TYPE2I1", conformidad);
                                arrayFields.Add("ORIGINATOR_ID2I1", usuario);
                                arrayFields.Add("JOB_DUR_FINISH2I1", "00:00");
                                arrayFields.Add("EQUIP_REF2I1", equipo);
                                requestSheet.screenFields = arrayFields.ToArray();

                                requestSheet.screenKey = "1";
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (_eFunctions.CheckReplyWarning(replySheet) ||
                                       replySheet.functionKeys == "XMIT-Confirm")
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                if (_eFunctions.CheckReplyError(replySheet))
                                {
                                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                                }
                                else
                                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                            }
                            catch (Exception ex)
                            {
                                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                                Debugger.LogError("RibbonEllipse.cs:MSO627LoadPestanas()", ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:MSO627Load()", ex.Message);
                }
                finally
                {
                    currentRow++;
                }
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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
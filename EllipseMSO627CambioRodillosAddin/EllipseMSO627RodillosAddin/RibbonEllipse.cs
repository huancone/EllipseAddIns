using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using System.Threading;

namespace EllipseMSO627RodillosAddin
{
    public partial class RibbonEllipse
    {
        private const int TitleRow = 7;
        private const int ResultColumnPbv = 11;
        private const int ResultColumnPcs = 13;
        private const int MaxRows = 10000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private const string PbvSheetName01 = "PBV Cambio Rodillos";
        private const string PcsSheetName01 = "PCSERVI Cambio Rodillos";
        private const string ValidationSheetName = "ValidationSheetLabour";
        private const string TableName01 = "RodillosTable";
        private ExcelStyleCells _cells;
        private Application _excelApp;

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
        private void btnFormatPbv_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetPbv();
        }

        private void btnFormatPcservi_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetPcservi();
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            //si ya hay un thread corriendo que no se ha detenido
            if (_thread != null && _thread.IsAlive) return;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            if (!_cells.IsDecimalDotSeparator())
                if (MessageBox.Show(@"El separador de decimales configurado actualmente no es el punto. Usar un separador de decimales diferente puede generar errores al momento de cargar valores numéricos. ¿Está seguro que desea continuar?", @"ALERTA DE SEPARADOR DE DECIMALES", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(PbvSheetName01))
                _thread = new Thread(LoadSheetPbv);
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(PcsSheetName01))
                _thread = new Thread(LoadSheetPcservi);
            else
            {
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para el cambio de rodillos");
                return;
            }
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }
        private void FormatSheetPbv()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _excelApp.ActiveWorkbook.ActiveSheet.Name = PbvSheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _excelApp.ActiveWorkbook.ActiveSheet.Name = PbvSheetName01;

                _cells.GetRange(1, TitleRow + 1, ResultColumnPbv, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TitleRow + 1, ResultColumnPbv, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow + 1, ResultColumnPbv, MaxRows).ClearComments();
                _cells.GetRange(1, TitleRow + 1, ResultColumnPbv, MaxRows).Clear();

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "REGISTRO DE CAMBIO DE RODILLOS";
                _cells.GetRange("B1", "D2").Merge();
                _cells.GetRange("B1", "D2").WrapText = true;

                //RAZON DEL CAMBIO	RODILLO MONTADO	CARGUE EN ELLIPSE


                _cells.GetCell(1, TitleRow).Value = "Fecha";
                _cells.GetCell(1, TitleRow).AddComment("YYYYMMDD");
                _cells.GetCell(2, TitleRow).Value = "Tipo de Rodillo";
                _cells.GetCell(3, TitleRow).Value = "Rodillo Desmontado";
                _cells.GetCell(4, TitleRow).Value = "Tipo de Cambio";
                _cells.GetCell(5, TitleRow).Value = "Usuario";
                _cells.GetCell(6, TitleRow).Value = "Equipo";
                _cells.GetCell(7, TitleRow).Value = "Estacion";
                _cells.GetCell(8, TitleRow).Value = "Posicion en la Estacion";
                _cells.GetCell(9, TitleRow).Value = "Razon del Cambio";
                _cells.GetCell(10, TitleRow).Value = "Rodillo Montado";
                _cells.GetCell(ResultColumnPbv, TitleRow).Value = "Resultado";

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

                _cells.GetCell(1, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumnPbv, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetRange(1, TitleRow + 1, ResultColumnPbv, MaxRows).NumberFormat = NumberFormatConstants.Text;

                #endregion
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow, ResultColumnPbv, TitleRow + 1), TableName01);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                #region Selection List Options

                var optionList = new List<string>
                {
                    "RC",
                    "RI",
                    "RR",
                    "RT"
                };

                _cells.SetValidationList(_cells.GetCell(2, TitleRow + 1), optionList, ValidationSheetName, 1);

                optionList = new List<string>
                {
                    "BUC",
                    "CAT",
                    "DIM",
                    "EXA",
                    "FMC",
                    "LUF",
                    "PPI",
                    "PYH",
                    "REP",
                    "REU",
                    "SAD",
                    "SAM",
                    "SAN",
                    "VAN"
                };
                _cells.SetValidationList(_cells.GetCell(3, TitleRow + 1), optionList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell(10, TitleRow + 1), optionList, ValidationSheetName, 3);

                optionList = new List<string>
                {
                    "BO",
                    "P"
                };
                _cells.SetValidationList(_cells.GetCell(4, TitleRow + 1), optionList, ValidationSheetName, 4);
                
                optionList = new List<string>
                {
                    "DEG",
                    "DES",
                    "PEG",
                    "REC",
                    "ROD",
                    "RUI",
                    "VIB"
                };
                _cells.SetValidationList(_cells.GetCell(9, TitleRow + 1), optionList, ValidationSheetName, 5);

                optionList = new List<string>
                {
                    "C",
                    "CD",
                    "CI",
                    "D",
                    "D2",
                    "I",
                    "I2"
                };
                _cells.SetValidationList(_cells.GetCell(8, TitleRow + 1), optionList, ValidationSheetName, 6);
                optionList = new List<string>
                {
                    "BC402",
                    "BC403",
                    "BC404A",
                    "BC404B",
                    "BC404C",
                    "BC405",
                    "BC407",
                    "BC407A",
                    "BC408",
                    "BC504",
                    "BC507",
                    "BC508",
                    "BC509",
                    "BF1",
                    "BF2",
                    "RAMSEY",
                    "RAMSEY2",
                    "SL",
                    "SL2",
                    "SL3",
                    "SR1",
                    "SR2",
                    "SR3"
                };
                _cells.SetValidationList(_cells.GetCell(6, TitleRow + 1), optionList, ValidationSheetName, 7);

                #endregion

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheetPbv()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al formatear la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void FormatSheetPcservi()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _excelApp.ActiveWorkbook.ActiveSheet.Name = PcsSheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetRange(1, TitleRow + 1, ResultColumnPcs, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TitleRow + 1, ResultColumnPcs, MaxRows).ClearFormats();
                _cells.GetRange(1, TitleRow + 1, ResultColumnPcs, MaxRows).ClearComments();
                _cells.GetRange(1, TitleRow + 1, ResultColumnPcs, MaxRows).Clear();

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "REGISTRO DE CAMBIO DE RODILLOS";
                _cells.GetRange("B1", "D2").Merge();
                _cells.GetRange("B1", "D2").WrapText = true;
                
                _cells.GetCell("A3").Value = "GRUPO:";
                _cells.GetCell("B3").Value = "PCSERVI";
                _cells.GetCell("A3").Style = StyleConstants.Option;
                _cells.GetCell("B3").Style = StyleConstants.Select;
                //RAZON DEL CAMBIO	RODILLO MONTADO	CARGUE EN ELLIPSE
                _cells.GetCell(1, TitleRow).Value = "Fecha";
                _cells.GetCell(1, TitleRow).AddComment("YYYYMMDD");
                _cells.GetCell(2, TitleRow).Value = "Hora Inicial";
                _cells.GetCell(2, TitleRow).AddComment("HH:MM");
                _cells.GetCell(3, TitleRow).Value = "Hora Final";
                _cells.GetCell(3, TitleRow).AddComment("HH:MM");
                _cells.GetCell(4, TitleRow).Value = "Tipo de Rodillo";
                _cells.GetCell(5, TitleRow).Value = "Rodillo Desmontado";
                _cells.GetCell(6, TitleRow).Value = "Tipo de Cambio";
                _cells.GetCell(7, TitleRow).Value = "Usuario";
                _cells.GetCell(8, TitleRow).Value = "Equipo";
                _cells.GetCell(9, TitleRow).Value = "Estacion";
                _cells.GetCell(10, TitleRow).Value = "Posicion en la Estacion";
                _cells.GetCell(11, TitleRow).Value = "Razon del Cambio";
                _cells.GetCell(12, TitleRow).Value = "Rodillo Montado";
                _cells.GetCell(ResultColumnPcs, TitleRow).Value = "Resultado";

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

                _cells.GetCell(1, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(4, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(11, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(12, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumnPcs, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetRange(1, TitleRow + 1, ResultColumnPcs, MaxRows).NumberFormat = NumberFormatConstants.Text;

                #endregion
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow, ResultColumnPcs, TitleRow + 1), TableName01);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                #region Selection List Options

                //TIPO
                var optionList = new List<string>
                {
                    "RC - CARGUE",
                    "RI - IMPACTO",
                    "RR - RETORNO"
                };

                _cells.SetValidationList(_cells.GetCell(4, TitleRow + 1), optionList, ValidationSheetName, 1, false);

                //TIPO DE CAMBIO
                optionList = new List<string>
                {
                    "METS - METSO",
                    "SAN - SANDVICK",
                    "METALPLAST"
                };
                _cells.SetValidationList(_cells.GetCell(5, TitleRow + 1), optionList, ValidationSheetName, 2, false);
                _cells.SetValidationList(_cells.GetCell(12, TitleRow + 1), ValidationSheetName, 2, false);

                optionList = new List<string>
                {
                    "P - PLANEADO",
                    "BO - BREAK DOWN/OPERATIVO",
                    "OT - OTROS"
                };
                _cells.SetValidationList(_cells.GetCell(6, TitleRow + 1), optionList, ValidationSheetName, 3, false);

                //RAZON DE CAMBIO - FALLA
                optionList = new List<string>
                {
                    "DEG - DESGASTE",
                    "DES - DESARMADO",
                    "PEG - PEGADO",
                    "REC - RECALENTADO",
                    "ROD - RODAMIENTO",
                    "RUI - RUIDO",
                    "VIB - VIBRACION",
                    "000 - ACCIDENTE",
                    "ICE - INCENDIO"
                };
                _cells.SetValidationList(_cells.GetCell(11, TitleRow + 1), optionList, ValidationSheetName, 4, false);

                //POSICIÓN
                optionList = new List<string>
                {
                    "C - CENTRAL",
                    "CD - CENTRAL DERECHA",
                    "CI - CENTRAL IZQUIERDA",
                    "D - DERECHA",
                    "DI - DERECHA INFERIOR",
                    "I - IAQUIERDA",
                    "LG - LAGUNA",
                    "TD - TRITURADORAS",
                    "PL - PLANO"
                };
                _cells.SetValidationList(_cells.GetCell(10, TitleRow + 1), optionList, ValidationSheetName, 5, false);

                //EQUIPO
                optionList = new List<string>
                {
                    "BC201",
                    "BC301",
                    "BANDA DE TRANSFERENCIA"
                };
                _cells.SetValidationList(_cells.GetCell(8, TitleRow + 1), optionList, ValidationSheetName, 6, false);

                #endregion

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheetPbv()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al formatear la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        private void LoadSheetPcservi()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != PcsSheetName01)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals(""))
                    throw new ArgumentException("Seleccione un ambiente válido");

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.GetRange(ResultColumnPcs, TitleRow + 1, ResultColumnPcs, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(ResultColumnPcs, TitleRow + 1, ResultColumnPcs, MaxRows).ClearFormats();
                _cells.GetRange(ResultColumnPcs, TitleRow + 1, ResultColumnPcs, MaxRows).ClearComments();
                _cells.GetRange(ResultColumnPcs, TitleRow + 1, ResultColumnPcs, MaxRows).Clear();

                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);



                var currentRow = TitleRow + 1;
                string grupo = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                {
                    try
                    {
                        string fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        string horainicial = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                        string horafinal = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                        string tipo = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                        string desmontado = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                        string cambio = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);
                        string usuario = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);
                        string equipo = _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value);
                        string estacion = _cells.GetEmptyIfNull(_cells.GetCell(9, currentRow).Value);
                        string posicion = _cells.GetEmptyIfNull(_cells.GetCell(10, currentRow).Value);
                        string razon = _cells.GetEmptyIfNull(_cells.GetCell(11, currentRow).Value);
                        string montado = _cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value);

                        var montaje = new MontajeRodillo
                        {
                            Grupo = grupo,
                            Tipo = Utils.GetCodeKey(tipo),
                            Fecha = fecha,
                            HoraInicial = string.IsNullOrWhiteSpace(horainicial) ? "00:00" : horainicial,
                            HoraFinal = string.IsNullOrWhiteSpace(horafinal) ? "00:00" : horafinal,
                            Desmontado = Utils.GetCodeKey(desmontado),
                            Cambio = Utils.GetCodeKey(cambio),
                            Usuario = usuario,
                            Equipo = Utils.GetCodeKey(equipo),
                            Estacion = Utils.GetCodeKey(estacion),
                            Posicion = Utils.GetCodeKey(posicion),
                            Razon = Utils.GetCodeKey(razon),
                            Montado = Utils.GetCodeKey(montado)
                        };

                        CreateMontaje(opSheet, montaje);
                        _cells.GetCell(ResultColumnPcs, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumnPcs, currentRow).Value2 = "CARGADO";


                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumnPcs, currentRow).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MSO627Load()", ex.Message);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadSheetPbv()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        private void LoadSheetPbv()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != PbvSheetName01)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals(""))
                    throw new ArgumentException("Seleccione un ambiente válido");

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.GetRange(ResultColumnPbv, TitleRow + 1, ResultColumnPbv, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(ResultColumnPbv, TitleRow + 1, ResultColumnPbv, MaxRows).ClearFormats();
                _cells.GetRange(ResultColumnPbv, TitleRow + 1, ResultColumnPbv, MaxRows).ClearComments();
                _cells.GetRange(ResultColumnPbv, TitleRow + 1, ResultColumnPbv, MaxRows).Clear();

                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                

                var currentRow = TitleRow + 1;
                while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                {
                    try
                    {
                        string fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        string tipo = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                        string desmontado = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                        string cambio = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                        string usuario = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                        string equipo = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);
                        string estacion = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);
                        string posicion = _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value);
                        string razon = _cells.GetEmptyIfNull(_cells.GetCell(9, currentRow).Value);
                        string montado = _cells.GetEmptyIfNull(_cells.GetCell(10, currentRow).Value);

                        var montaje = new MontajeRodillo
                        {
                            Grupo = "PTORODI",
                            Tipo = tipo,
                            Fecha = fecha,
                            HoraInicial = "00:00",
                            HoraFinal = "00:00",
                            Desmontado = desmontado,
                            Cambio = cambio,
                            Usuario = usuario,
                            Equipo = equipo,
                            Estacion = estacion,
                            Posicion = posicion,
                            Razon = razon,
                            Montado = montado
                        };

                        CreateMontaje(opSheet, montaje);
                        _cells.GetCell(ResultColumnPbv, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumnPbv, currentRow).Value2 = "CARGADO";


                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumnPbv, currentRow).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MSO627Load()", ex.Message);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadSheetPbv()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public bool CreateMontaje(Screen.OperationContext opSheet, MontajeRodillo montaje)
        {
            var proxySheet = new Screen.ScreenService { Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService" };
            _eFunctions.RevertOperation(opSheet, proxySheet);
            var replySheet = proxySheet.executeScreen(opSheet, "MSO627");

            if (_eFunctions.CheckReplyError(replySheet))
                throw new Exception("No se pudo ingresar al módulo MSO627. " + replySheet.message);

            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("WORK_GROUP1I", montaje.Grupo);
            arrayFields.Add("RAISED_DATE1I", montaje.Fecha);
            arrayFields.Add("SHIFT1I", montaje.Tipo);
            requestSheet.screenFields = arrayFields.ToArray();

            requestSheet.screenKey = "1";
            replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyWarning(replySheet))
                replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyError(replySheet) || replySheet.mapName != "MSM627B")
                throw new Exception("MSO627. " + replySheet.message);

            arrayFields = new ArrayScreenNameValue();

            arrayFields.Add("RAISED_TIME2I1", montaje.HoraInicial);
            arrayFields.Add("INCIDENT_DESC2I1", montaje.Desmontado);
            arrayFields.Add("MAINT_TYPE2I1", montaje.Cambio);
            arrayFields.Add("ORIGINATOR_ID2I1", montaje.Usuario);
            arrayFields.Add("JOB_DUR_FINISH2I1", montaje.HoraFinal);
            arrayFields.Add("EQUIP_REF2I1", montaje.Equipo);
            arrayFields.Add("COMP_CODE2I1", montaje.Estacion);
            arrayFields.Add("COMP_MOD_CODE2I1", montaje.Posicion);
            arrayFields.Add("JOB_DUR_CODE2I1", montaje.Razon);
            arrayFields.Add("CORRECT_DESC2I1", montaje.Montado);
            arrayFields.Add("ACTION2I1", "C");

            requestSheet.screenFields = arrayFields.ToArray();

            requestSheet.screenKey = "1";
            replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyError(replySheet))
                throw new Exception("MSM627B. " + replySheet.message);

            while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys == "XMIT-Confirm" || replySheet.functionKeys == "XMIT-Validate") //requiere confirmación
            {
                replySheet = proxySheet.submit(opSheet, requestSheet);
                if (_eFunctions.CheckReplyError(replySheet))
                    throw new Exception("MSM627B. " + replySheet.message);
            }
            return true;
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
        
    }

    public class MontajeRodillo
    {
        public string Grupo;
        public string Tipo;
        public string Desmontado;
        public string Cambio;
        public string Usuario;
        public string Fecha;
        public string HoraInicial;
        public string HoraFinal;
        public string Equipo;
        public string Estacion;
        public string Posicion;
        public string Razon;
        public string Montado;
    }
}
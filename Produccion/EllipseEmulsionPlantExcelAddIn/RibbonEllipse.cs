using System;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary.Utilities;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseEmulsionPlantExcelAddIn.LogSheet;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;

namespace EllipseEmulsionPlantExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string SheetName01 = "Emulsión";
        private const string SheetName02 = "Solución";

        private const int TitleRow01 = 7;
        private const int TitleRow02 = 7;

        private const int ResultColumn01 = 17;
        private const int ResultColumn02 = 20;

        private const string TableName01 = "EmulsionTable";
        private const string TableName02 = "SolutionTable";

        private const string ValidationSheetName = "ValidationSheet";
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

            //settings.SetDefaultCustomSettingValue("AutoSort", "Y");
            //settings.SetDefaultCustomSettingValue("OverrideAccountCode", "Maintenance");
            //settings.SetDefaultCustomSettingValue("IgnoreItemError", "N");
            //settings.SetDefaultCustomSettingValue("AllowBackgroundWork", "N");

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //var overrideAccountCode = settings.GetCustomSettingValue("OverrideAccountCode");
            //if (overrideAccountCode.Equals("Maintenance"))
            //    cbAccountElementOverrideMntto.Checked = true;
            //else if (overrideAccountCode.Equals("Disable"))
            //    cbAccountElementOverrideDisable.Checked = true;
            //else if (overrideAccountCode.Equals("Always"))
            //    cbAccountElementOverrideAlways.Checked = true;
            //else if (overrideAccountCode.Equals("Default"))
            //    cbAccountElementOverrideDefault.Checked = true;
            //else
            //    cbAccountElementOverrideDefault.Checked = true;
            //cbAutoSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AutoSort"));
            //cbIgnoreItemError.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreItemError"));
            //cbAllowBackgroundWork.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AllowBackgroundWork"));

            //
            settings.SaveCustomSettings();
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnLoadEmulsion_Click(object sender, RibbonControlEventArgs e)
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
                    _thread = new Thread(LoadEmulsionListData);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:LoadEmulsionListData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnLoadSolox_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(LoadSolutionListData);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:LoadSolutionListData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void FormatSheet()
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

                #region CONSTRUYO LA HOJA 1 - EMULSIÓN
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;
                var tableName = TableName01;

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "MODELO DE EMULSIÓN - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                //TO DO - Adicionar consulta de datos
                //_cells.GetCell("A3").Value = "DESDE";
                //_cells.GetCell("B3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                //_cells.GetCell("B3").AddComment("YYYYMMDD");
                //_cells.GetCell("A4").Value = "HASTA";
                //_cells.GetCell("B4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                //_cells.GetCell("B4").AddComment("YYYYMMDD");
                //_cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleRequired;

                //GENERAL
                _cells.GetCell(1, titleRow).Value = "Fecha (YYYYMMDD)";
                _cells.GetCell(2, titleRow).Value = "Turno";
                _cells.GetCell(3, titleRow).Value = "Operador";
                _cells.GetCell(4, titleRow).Value = "T.EM. Inicial";
                _cells.GetCell(5, titleRow).Value = "T.EM. Final";
                _cells.GetCell(6, titleRow).Value = "T.EM. Producida";
                _cells.GetCell(7, titleRow).Value = "T.EM. Despachada";
                _cells.GetCell(8, titleRow).Value = "S1 Inicial";
                _cells.GetCell(9, titleRow).Value = "S1 Final";
                _cells.GetCell(10, titleRow).Value = "S1 Producción";
                _cells.GetCell(11, titleRow).Value = "S2 Inicial";
                _cells.GetCell(12, titleRow).Value = "S2 Final";
                _cells.GetCell(13, titleRow).Value = "S2 Producción";
                _cells.GetCell(14, titleRow).Value = "S3 Inicial";
                _cells.GetCell(15, titleRow).Value = "S3 Final";
                _cells.GetCell(16, titleRow).Value = "S3 Producción";
                
                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 2 - SOLUCIÓN OXIDANTE
                titleRow = TitleRow02;
                resultColumn = ResultColumn02;
                tableName = TableName02;
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "MODELO DE SOLUCIÓN OXIDANTE - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                //TO DO - Adicionar consulta de datos
                //_cells.GetCell("A3").Value = "DESDE";
                //_cells.GetCell("B3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                //_cells.GetCell("B3").AddComment("YYYYMMDD");
                //_cells.GetCell("A4").Value = "HASTA";
                //_cells.GetCell("B4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                //_cells.GetCell("B4").AddComment("YYYYMMDD");
                //_cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleRequired;

                //GENERAL
                _cells.GetCell(1, titleRow).Value = "Fecha (YYYYMMDD)";
                _cells.GetCell(2, titleRow).Value = "Turno";
                _cells.GetCell(3, titleRow).Value = "Operador";
                _cells.GetCell(4, titleRow).Value = "T.SOL. Inicial";
                _cells.GetCell(5, titleRow).Value = "T.SOL. Final";
                _cells.GetCell(6, titleRow).Value = "T.SOL. Producida";
                _cells.GetCell(7, titleRow).Value = "T.SOL. Utilizada";
                _cells.GetCell(8, titleRow).Value = "TK1 Inicial";
                _cells.GetCell(9, titleRow).Value = "TK1 Final";
                _cells.GetCell(10, titleRow).Value = "TK1 Producción";
                _cells.GetCell(11, titleRow).Value = "TK2 Inicial";
                _cells.GetCell(12, titleRow).Value = "TK2 Final";
                _cells.GetCell(13, titleRow).Value = "TK2 Producción";
                _cells.GetCell(14, titleRow).Value = "TK3 Inicial";
                _cells.GetCell(15, titleRow).Value = "TK3 Final";
                _cells.GetCell(16, titleRow).Value = "TK3 Producción";
                _cells.GetCell(17, titleRow).Value = "TK4 Inicial";
                _cells.GetCell(18, titleRow).Value = "TK4 Final";
                _cells.GetCell(19, titleRow).Value = "TK4 Producción";

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        private void btnDelete_Click(object sender, RibbonControlEventArgs e)
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
                    _thread = new Thread(DeleteEmulsionListData);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(DeleteSolutionListData);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:LoadEmulsionListData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        public void LoadEmulsionListData()
        {
            var tableName = TableName01;
            var resultColumn = ResultColumn01;
            var titleRow = TitleRow01;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var opContext = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var date = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var shiftCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var operatorId = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var totalInicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                    var totalFinal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    var totalProducido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    var totalDespachado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    var s1Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value);
                    var s1Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                    var s1Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    var s2Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    var s2Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                    var s2Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);
                    var s3Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value);
                    var s3Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value);
                    var s3Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value);

                    var emulsionLog = new EmulsionModelLogSheet();
                    emulsionLog.Date = date;
                    emulsionLog.ShiftCode = shiftCode;
                    emulsionLog.Operator = operatorId;
                    emulsionLog.TotalDestination = "";
                    emulsionLog.TotalEmStartValue = totalInicial;
                    emulsionLog.TotalEmEndValue = totalFinal;
                    emulsionLog.TotalEmProduced = totalProducido;
                    emulsionLog.TotalEmDispatched = totalDespachado;

                    emulsionLog.Silo1StartValue = s1Inicial;
                    emulsionLog.Silo1EndValue = s1Final;
                    emulsionLog.Silo1Produced = s1Producido;
                    emulsionLog.Silo2StartValue = s2Inicial;
                    emulsionLog.Silo2EndValue = s2Final;
                    emulsionLog.Silo2Produced = s2Producido;
                    emulsionLog.Silo3StartValue = s3Inicial;
                    emulsionLog.Silo3EndValue = s3Final;
                    emulsionLog.Silo3Produced = s3Producido;

                    var result = LogSheetActions.CreateLogSheet(_eFunctions, opContext, urlService, emulsionLog.ToLogSheet());
                    var styleResult = StyleConstants.Success;
                    if (result.StartsWith("WARNING"))
                        styleResult = StyleConstants.Warning;

                    _cells.GetCell(resultColumn, i).Style = styleResult;
                    _cells.GetCell(resultColumn, i).Value = result;
                    _cells.GetCell(resultColumn, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:LoadEmulsionListData()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void LoadSolutionListData()
        {
            var tableName = TableName02;
            var resultColumn = ResultColumn02;
            var titleRow = TitleRow02;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var opContext = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var date = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var shiftCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var operatorId = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var totalInicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                    var totalFinal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    var totalProducido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    var totalUtilizado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    var tk1Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value);
                    var tk1Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                    var tk1Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    var tk2Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    var tk2Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                    var tk2Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);
                    var tk3Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value);
                    var tk3Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value);
                    var tk3Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value);
                    var tk4Inicial = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value);
                    var tk4Final = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, i).Value);
                    var tk4Producido = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, i).Value);

                    var solutionLog = new SolutionModelLogSheet();
                    solutionLog.Date = date;
                    solutionLog.ShiftCode = shiftCode;
                    solutionLog.Operator = operatorId;
                    solutionLog.TotalDestination = "";
                    solutionLog.TotalSolStartValue = totalInicial;
                    solutionLog.TotalSolEndValue = totalFinal;
                    solutionLog.TotalSolProduced = totalProducido;
                    solutionLog.TotalSolUsed = totalUtilizado;

                    solutionLog.Tank1StartValue = tk1Inicial;
                    solutionLog.Tank1EndValue = tk1Final;
                    solutionLog.Tank1Produced = tk1Producido;
                    solutionLog.Tank2StartValue = tk2Inicial;
                    solutionLog.Tank2EndValue = tk2Final;
                    solutionLog.Tank2Produced = tk2Producido;
                    solutionLog.Tank3StartValue = tk3Inicial;
                    solutionLog.Tank3EndValue = tk3Final;
                    solutionLog.Tank3Produced = tk3Producido;
                    solutionLog.Tank4StartValue = tk4Inicial;
                    solutionLog.Tank4EndValue = tk4Final;
                    solutionLog.Tank4Produced = tk4Producido;

                    var result = LogSheetActions.CreateLogSheet(_eFunctions, opContext, urlService, solutionLog.ToLogSheet());
                    var styleResult = StyleConstants.Success;
                    if (result.StartsWith("WARNING"))
                        styleResult = StyleConstants.Warning;

                    _cells.GetCell(resultColumn, i).Style = styleResult;
                    _cells.GetCell(resultColumn, i).Value = result;
                    _cells.GetCell(resultColumn, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:LoadSolutionListData()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }
        private void DeleteEmulsionListData()
        {
            var tableName = TableName01;
            var resultColumn = ResultColumn01;
            var titleRow = TitleRow01;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var opContext = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var date = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var shiftCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var emulsionLog = new LogSheetItem("EMULPLANT", date, shiftCode);

                    var result = LogSheetActions.DeleteLogSheet(_eFunctions, opContext, urlService, emulsionLog);
                    var styleResult = StyleConstants.Success;
                    if (result.StartsWith("WARNING"))
                        styleResult = StyleConstants.Warning;

                    _cells.GetCell(resultColumn, i).Style = styleResult;
                    _cells.GetCell(resultColumn, i).Value = result;
                    _cells.GetCell(resultColumn, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:LoadEmulsionListData()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }
        private void DeleteSolutionListData()
        {
            var tableName = TableName02;
            var resultColumn = ResultColumn02;
            var titleRow = TitleRow02;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var opContext = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var date = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var shiftCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var emulsionLog = new LogSheetItem("EMULPLANT2", date, shiftCode);

                    var result = LogSheetActions.DeleteLogSheet(_eFunctions, opContext, urlService, emulsionLog);
                    var styleResult = StyleConstants.Success;
                    if (result.StartsWith("WARNING"))
                        styleResult = StyleConstants.Warning;

                    _cells.GetCell(resultColumn, i).Style = styleResult;
                    _cells.GetCell(resultColumn, i).Value = result;
                    _cells.GetCell(resultColumn, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:LoadEmulsionListData()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }
        
        private void btnGetModuleEmulsion_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("No se ha definido una fuente de obtención automática");
        }

        private void btnGetModuleSolox_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("No se ha definido una fuente de obtención automática");
        }
    }
}

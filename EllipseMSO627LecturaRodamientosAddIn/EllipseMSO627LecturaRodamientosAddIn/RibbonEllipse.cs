using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Utilities;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using EllipseIncidentLogSheetClassLibraries;
using SharedClassLibrary.Ellipse.Connections;

namespace EllipseMSO627LecturaRodamientosAddIn
{
    public partial class RibbonEllipse
    {
        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 7;
        private const string SheetName01 = "Lectura de Rodamientos";
        private const string TableName01 = "LectRodamientosTable";
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        
        private ExcelStyleCells _cells;
        private Application _excelApp;
        
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

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }
        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatLectRodamientos();
        }


        private void FormatLectRodamientos()
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

                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;
                var tableName = TableName01;
                var sheetName = SheetName01;

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "LECTURA DE RODAMIENTOS PARA FFCC";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "D2");

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

                _cells.GetCell("B1").AddComment("Valores Predeterminados \nWorkGroup: CTC," +
                "MaintenanceType: PD," +
                "Shift: ");

                _cells.GetCell(1, titleRow).Value = "Fecha";
                _cells.GetCell(1, titleRow).AddComment("YYYYMMDD");
                _cells.GetCell(2, titleRow).Value = "Hora";
                _cells.GetCell(2, titleRow).AddComment("HHMM");
                _cells.GetCell(3, titleRow).Value = "Descripcion";

                _cells.GetCell(4, titleRow).Value = "Equipo";
                _cells.GetCell(5, titleRow).Value = "Componente";
                _cells.GetCell(6, titleRow).Value = "Correctivo";
                //= "Grupo";//CTC Default
                //= "Conformidad";//PD Default
                //= "Usuario";//Login Default
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";

                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();



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

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadLecturaRodamientos);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnDelete_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(DeleteLecturaRodamientos);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");

        }

        
        private void LoadLecturaRodamientos()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                var opContext = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);



                var currentRow = titleRow + 1;
                while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                {
                    try
                    {
                        var item = new IncidentItem();

                        var workGroup = "CTC";
                        var date = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        var shift = "";

                        item.RaisedTime = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                        item.IncidentDescription = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                        item.MaintenanceType = "PD";
                        item.Originator = _frmAuth.EllipseUser;
                        item.EquipmentReference = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                        item.ComponentCode = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value));
                        item.CorrectiveDescription = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value));

                        var reply = IncidentActions.CreateIncident(_eFunctions, opContext, urlService, workGroup, date, shift, item);
                        if (reply != null && reply.mapName != "MSM627A")
                            throw new Exception("Se ha producido un error al finalizar el proceso");
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, currentRow).Value = "CREADO";
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:LoadLecturaRodamientos()", ex.Message);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch(Exception ex)
            {
                Debugger.LogError("RibbonEllipse:LoadLecturaRodamientos()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        private void DeleteLecturaRodamientos()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                var opContext = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var currentRow = titleRow + 1;
                while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                {
                    try
                    {
                        var item = new IncidentItem();

                        var workGroup = "CTC";
                        var date = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        var shift = "";

                        item.RaisedTime = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                        item.IncidentDescription = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                        item.MaintenanceType = "PD";
                        //item.Originator = _frmAuth.EllipseUser; //En Blanco para que pueda ser borrado por cualquier usuario
                        item.EquipmentReference = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                        item.ComponentCode = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value));
                        item.CorrectiveDescription = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value));

                        var reply = IncidentActions.DeleteIncident(_eFunctions, opContext, urlService, workGroup, date, shift, item.Originator, item.EquipmentReference, item.IncidentStatus, item);
                        if (reply != null && reply.mapName != "MSM627A")
                            throw new Exception("Se ha producido un error al finalizar el proceso");
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, currentRow).Value = "REGISTRO ELIMINADO";
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:DeleteLecturaRodamientos()", ex.Message);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            catch(Exception ex)
            {
                Debugger.LogError("RibbonEllipse:LoadLecturaRodamientos()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.Message);
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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        
    }
}
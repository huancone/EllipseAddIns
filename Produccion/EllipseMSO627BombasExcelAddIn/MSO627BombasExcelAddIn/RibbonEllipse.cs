using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using EllipseEquipmentClassLibrary;
using LINQtoCSV;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;

namespace MSO627BombasExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;
        private Thread _thread;

        private const string SheetName01 = "MSO627 Bombas";
        private const int TittleRow01 = 5;
        private const int ResultColumn01 = 6;
        private const string TableName01 = "BombasTable";
        private const string ValidationSheetName = "ValidationSheetBombas";
        private List<Locations> _locationsList;
        

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
        private void btnFormato_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnCargar_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(CargarDatosBombas);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
        }

        private void FormatSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;
            _excelApp.Workbooks.Add();
            while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                _excelApp.ActiveWorkbook.Worksheets.Add();

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
            _cells.CreateNewWorksheet(ValidationSheetName);
            
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("A1", "A2");

            _cells.GetCell("B1").Value = "EQUIPMENT - ELLIPSE 8";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("B1", "J2");

            _cells.GetCell("K1").Value = "OBLIGATORIO";
            _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell("K2").Value = "OPCIONAL";
            _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
            _cells.GetCell("K3").Value = "INFORMATIVO";
            _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.GetCell(1, TittleRow01).Value = "Fecha";
            _cells.GetCell(1, TittleRow01).AddComment("YYYYMMDD");
            _cells.GetCell(2, TittleRow01).Value = "Origen";
            _cells.GetCell(3, TittleRow01).Value = "Usuario";
            _cells.GetCell(4, TittleRow01).Value = "Equipo";
            _cells.GetCell(5, TittleRow01).Value = "Destino";
            _cells.GetRange(1, TittleRow01, 5, TittleRow01).Style = StyleConstants.TitleRequired;
            _cells.GetCell(ResultColumn01, TittleRow01).Value = "Resultado";
            _cells.GetCell(ResultColumn01, TittleRow01).Style = _cells.GetStyle(StyleConstants.TitleInformation);

            LoadLocationsList();

            var listaNombres = new List<string>();
            if(_locationsList != null) 
                listaNombres.AddRange(_locationsList.Select(item => item.Nombre));


            _cells.SetValidationList(_cells.GetCell(2, TittleRow01 + 1), listaNombres, ValidationSheetName, 1, false);
            _cells.SetValidationList(_cells.GetCell(5, TittleRow01 + 1), ValidationSheetName, 1, false);

            _cells.FormatAsTable(_cells.GetRange(1, TittleRow01, ResultColumn01, TittleRow01 + 1), TableName01);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            var listaEquipos = new[] { "WPHV-VNM-10", "WPHH-DS-10", "WPHV-GNVN-8", "WPHS-SS-6", "WPHS-SS-10", "WPF-2201-4", "WPF-2250-6", "WPF-2250-10", "WPF-2400-6", "WPF-3201-6", "WPF-3230-8" }
                     .SelectMany(id => EquipmentActions.GetEgiEquipments(_eFunctions, id))
                     .ToList();

            _cells.SetValidationList(_cells.GetCell(4, TittleRow01 + 1), listaEquipos, ValidationSheetName, 2, false);

        }

        private void CargarDatosBombas()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            var opSheet = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

            var currentRow = TittleRow01 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, currentRow).Value))
            {
                try
                {
                    string origen = null;
                    string destino = null;
                    var fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);

                    var row = currentRow;
                    if(_locationsList != null)
                        foreach (var item in _locationsList.Where(item => item.Nombre == _cells.GetEmptyIfNull(_cells.GetCell(2, row).Value)))
                            origen = item.Nombre + "/" + item.X + "/" + item.Y + "/" + item.Z;

                    if (_locationsList != null)
                        foreach (var item in _locationsList.Where(item => item.Nombre == _cells.GetEmptyIfNull(_cells.GetCell(5, row).Value)))
                            destino = item.Nombre + "/" + item.X + "/" + item.Y + "/" + item.Z;
                    

                    var usuario = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                    var equipo = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);

                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO627");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "1");
                        arrayFields.Add("WORK_GROUP1I", "EOEAGUA");
                        arrayFields.Add("RAISED_DATE1I", fecha);
                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumn01, currentRow).Value = replySheet.message;
                        }
                        else if (replySheet.mapName == "MSM627B")
                        {
                            arrayFields = new ArrayScreenNameValue();

                            arrayFields.Add("RAISED_TIME2I1", "00:00");
                            arrayFields.Add("INCIDENT_DESC2I1", origen);
                            arrayFields.Add("MAINT_TYPE2I1", "NM");
                            arrayFields.Add("ORIGINATOR_ID2I1", usuario.ToUpper());
                            arrayFields.Add("JOB_DUR_FINISH2I1", "00:00");
                            arrayFields.Add("EQUIP_REF2I1", equipo);
                            arrayFields.Add("CORRECT_DESC2I1", destino);

                            requestSheet.screenFields = arrayFields.ToArray();

                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (_eFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(ResultColumn01, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                while (_eFunctions.CheckReplyWarning(replySheet))
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                if (replySheet.functionKeys.StartsWith("XMIT-Confirm"))
                                    proxySheet.submit(opSheet, requestSheet);

                                _cells.GetCell(ResultColumn01, currentRow).Value = "REGISTRADO";
                                _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Success;
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, currentRow).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn01, currentRow).Select();
                    Debugger.LogError("RibbonEllipse.cs:CargarDatosBombas()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, currentRow).Select();
                    currentRow++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        private void LoadLocationsList()
        {
            var openFileDialog1 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"Ubicaciones.csv",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\Data\Loaders\Parametros"
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            if (_locationsList == null)
                _locationsList = new List<Locations>();
            _locationsList.Clear();

            var filePath = openFileDialog1.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var archivoUbicaciones = cc.Read<Locations>(filePath, inputFileDescription);

            foreach (var p in archivoUbicaciones)
            {
                try
                {
                    _locationsList.Add(p);
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ImportFile()", ex.Message);
                }
            }
        }

        

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if(_cells != null) _cells.SetCursorDefault();
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

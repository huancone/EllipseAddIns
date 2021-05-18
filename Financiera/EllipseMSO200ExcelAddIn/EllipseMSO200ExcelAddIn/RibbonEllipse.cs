using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = SharedClassLibrary.Ellipse.ScreenService;

namespace EllipseMSO200ExcelAddIn
{
    
    public partial class RibbonEllipse
    {

        private Thread _thread;
        private const int TittleRow = 5;
        private static int _resultColumn = 14;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private string _sheetName01;
        public List<Bancos> ListaBancos;

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

            //Example of Default Custom Options
            //settings.SetDefaultCustomSettingValue("AutoSort", "Y");
            //settings.SetDefaultCustomSettingValue("OverrideAccountCode", "Maintenance");
            //settings.SetDefaultCustomSettingValue("IgnoreItemError", "N");

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //Example of Getting Custom Options from Save File
            //var overrideAccountCode = settings.GetCustomSettingValue("OverrideAccountCode");
            //if (overrideAccountCode.Equals("Maintenance"))
            //    cbAccountElementOverrideMntto.Checked = true;
            //else if (overrideAccountCode.Equals("Disable"))
            //    cbAccountElementOverrideDisable.Checked = true;
            //else if (overrideAccountCode.Equals("Alwats"))
            //    cbAccountElementOverrideAlways.Checked = true;
            //else if (overrideAccountCode.Equals("Default"))
            //    cbAccountElementOverrideDefault.Checked = true;
            //else
            //    cbAccountElementOverrideDefault.Checked = true;
            //cbAutoSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AutoSort"));
            //cbIgnoreItemError.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreItemError"));

            //
            settings.SaveCustomSettings();
        }
        private void btnInactivateSupplier_Click(object sender, RibbonControlEventArgs e)
        {
            FormatInactivateSupplierAddress();
        }

        private void FormatInactivateSupplierAddress()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;
            _sheetName01 = "Inactivar Supplier";
            excelSheet.Name = _sheetName01;
            _resultColumn = 4;

            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                _cells.GetRange(1, TittleRow, _resultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);


            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Inactivar Supplier Business";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";

            _cells.GetCell(1, TittleRow).Value = "Distrito";
            _cells.GetCell(2, TittleRow).Value = "Supplier";
            _cells.GetCell(3, TittleRow).Value = "Nombre";
            _cells.GetCell(_resultColumn, TittleRow).Value = "Resultado";

            _cells.GetRange(1, TittleRow, _resultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnChangeAccounts_Click(object sender, RibbonControlEventArgs e)
        {
            FormatCambioCuentas();
        }

        private void FormatCambioCuentas()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;
            _sheetName01 = "MSO200 Cambio Cuentas";
            excelSheet.Name = _sheetName01;
            _resultColumn = 14;

            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                _cells.GetRange(1, TittleRow, _resultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "CAMBIO DE CUENTA EMPLEADOS MENSUAL";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";

            _cells.GetCell(1, TittleRow).Value = "Supplier";
            _cells.GetCell(2, TittleRow).Value = "Cedula";
            _cells.GetCell(3, TittleRow).Value = "Nombre";
            _cells.GetCell(4, TittleRow).Value = "Codigo Banco";
            _cells.GetCell(5, TittleRow).Value = "Nombre Banco";
            _cells.GetCell(6, TittleRow).Value = "Tipo Cuenta";
            _cells.GetCell(7, TittleRow).Value = "Numero Cuenta";
            _cells.GetCell(8, TittleRow).Value = "BankAccount";
            _cells.GetCell(9, TittleRow).Value = "BankAccountName";
            _cells.GetCell(10, TittleRow).Value = "Default Bank Branch";
            _cells.GetCell(11, TittleRow).Value = "Default Bank Account";
            _cells.GetCell(12, TittleRow).Value = "Default Bank Account Name";
            _cells.GetCell(13, TittleRow).Value = "Actual Ellipse BankAccount";
            _cells.GetCell(_resultColumn, TittleRow).Value = "Resultado";

            _cells.GetRange(1, TittleRow, _resultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            ImportFileBancos();
        }

        private void ImportFileBancos()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != _sheetName01) return;

            var openFileDialog1 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"Bancos.csv",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\Data\Loaders\Parametros"
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog1.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true,
                EnforceCsvColumnAttribute = true
            };
            var cc = new CsvContext();
            var bancosfile = cc.Read<Bancos>(filePath, inputFileDescription);
            ListaBancos = new List<Bancos>();
            foreach (var c in bancosfile)
            {
                try
                {
                    var bancos = new Bancos
                    {
                        CodigoMims = c.CodigoMims,
                        CodigoNomina = c.CodigoNomina,
                        NombreInstitucion = c.NombreInstitucion
                    };
                    ListaBancos.Add(bancos);
                }
                catch (Exception error)
                {
                    MessageBox.Show(string.Format("Error: {0}", error.Message));
                }
            }
        }

        private void btnValidateAccounts_Click(object sender, RibbonControlEventArgs e)
        {
            ValidateBancos();
        }

        private void ValidateBancos()
        {
            Debugger.LogDebugging("Inicio del Proceso");
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var currentRow = TittleRow + 1;

            Debugger.LogDebugging("Preverificación de Bancos");
            if (ListaBancos == null)
                ListaBancos = new List<Bancos>();
            Debugger.LogDebugging("Pos Verificación de Bancos");
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                Debugger.LogDebugging("Iteración " + currentRow);
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var employee = new EmployeeInfo
                    {
                        Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        Cedula = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Nombre = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        CodigoBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        NombreBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        TipoCuenta = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        NumeroCuenta = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value)
                    };

                    if (employee.NumeroCuenta != null)
                    {
                        employee.BankAccount = employee.NumeroCuenta.Replace("-", "");
                        _cells.GetCell(8, currentRow).Value = employee.NumeroCuenta.Replace("-", "");
                    }

                    // ReSharper disable once UnusedVariable
                    if (ListaBancos.Any(banco => employee.CodigoBanco.PadLeft(2, '0') == banco.CodigoNomina.PadLeft(2, '0')))
                    {
                        employee.TipoCuenta = employee.TipoCuenta.ToUpper() == "AHORRO" ? "02" : "01";
                        _cells.GetCell(9, currentRow).Value = "2-" + employee.CodigoBanco.PadLeft(4, '0') + "-0000-" + employee.TipoCuenta;
                    }
                    else
                    {
                        _cells.GetCell(9, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        _cells.GetCell(9, currentRow).ClearComments();
                        _cells.GetCell(9, currentRow).AddComment("Verificar Banco");
                    }
                    
                    if(string.IsNullOrWhiteSpace("" + _cells.GetCell(10, currentRow).Value))
                        _cells.GetCell(10, currentRow).Value = "NM009";
                    if (string.IsNullOrWhiteSpace("" + _cells.GetCell(11, currentRow).Value))
                        _cells.GetCell(11, currentRow).Value = "1209";
                    if (string.IsNullOrWhiteSpace("" + _cells.GetCell(12, currentRow).Value))
                        _cells.GetCell(12, currentRow).Value = "47703371338";


                    var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", employee.Cedula, _eFunctions.DbReference, _eFunctions.DbLink);

                    var drSupplierInfo = _eFunctions.GetQueryResult(sqlQuery);

                    if (!drSupplierInfo.Read())
                    {
                        _cells.GetCell(8, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = "Supplier-Empleado Inactivo";
                    }
                    else
                    {
                        var cant = Convert.ToInt16(drSupplierInfo["CANTIDAD_REGISTROS"].ToString());
                        if (cant > 1)
                        {
                            _cells.GetCell(_resultColumn, currentRow).Value = "la Cedula tiene mas de un Supplier Activo";
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        }
                        else
                        {
                            _cells.GetCell(1, currentRow).Value = drSupplierInfo["SUPPLIER_NO"].ToString();
                            _cells.GetCell(13, currentRow).Value = drSupplierInfo["BANK_ACCT_NO"].ToString();
                            if (employee.BankAccount == null)
                                employee.BankAccount = "";
                            _cells.GetCell(_resultColumn, currentRow).Value = drSupplierInfo["BANK_ACCT_NO"].ToString().Contains(employee.BankAccount) ? "Cuentas Iguales":"Numero de Cuenta Diferente";
                            _cells.GetCell(_resultColumn, currentRow).Style = drSupplierInfo["BANK_ACCT_NO"].ToString().Contains(employee.BankAccount) ? StyleConstants.Success : StyleConstants.Error;
                        }
                        
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadCambioCuentas);
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnLoadData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }


        private void btnInactivateBussiness_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadInactivarSupplierBusiness);
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnLoadData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }


        private void btnInactivareAddress_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadInactivarSupplierAddress);
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnLoadData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void LoadInactivarSupplierBusiness()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO200");
                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM200A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDstrct);
                        arrayFields.Add("OPTION1I", "6");
                        arrayFields.Add("SUPPLIER_NO1I",_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("DELETE_CONF1I", "Y");
                            requestSheet.screenFields = arrayFields.ToArray();
                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            while (_eFunctions.CheckReplyWarning(replySheet))
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (replySheet.message.Contains("Confirm"))
                                replySheet = proxySheet.submit(opSheet, requestSheet);
                            if (_eFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Success;
                                _cells.GetCell(_resultColumn, currentRow).Value = "Borrado";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void LoadInactivarSupplierAddress()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO200");
                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM200A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "5");
                        arrayFields.Add("SUPPLIER_NO1I",_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("DELETE_CONF2I", "Y");
                            requestSheet.screenFields = arrayFields.ToArray();
                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            while (_eFunctions.CheckReplyWarning(replySheet))
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (replySheet.message.Contains("Confirm"))
                                replySheet = proxySheet.submit(opSheet, requestSheet);
                            if (_eFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Success;
                                _cells.GetCell(_resultColumn, currentRow).Value = "Borrado";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void LoadSuspenderSupplier()
        {
            //
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO200");
                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM200A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "B");
                        arrayFields.Add("SUPPLIER_NO1I", _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("DELETE_CONF2I", "Y");
                            requestSheet.screenFields = arrayFields.ToArray();
                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            while (_eFunctions.CheckReplyWarning(replySheet))
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (replySheet.message.Contains("Confirm"))
                                replySheet = proxySheet.submit(opSheet, requestSheet);
                            if (_eFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Success;
                                _cells.GetCell(_resultColumn, currentRow).Value = "Suspendido";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }

            //

        }

        private void LoadCambioCuentas()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var employee = new EmployeeInfo
                    {
                        Supplier = null,
                        Cedula = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Nombre = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        CodigoBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        NombreBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        TipoCuenta = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        NumeroCuenta = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        BankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        BankAccountName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        DefaultBankBranch = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        DefaultBankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        DefaultBankAccountName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value)
                    };
                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO200");
                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM200A") return;

                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDstrct);
                        arrayFields.Add("OPTION1I", "4");
                        arrayFields.Add("SUP_MNEMONIC1I", employee.Cedula);
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            if (replySheet.mapName == "MSM202A")
                            {
                                requestSheet.screenKey = "1";
                                arrayFields = new ArrayScreenNameValue();
                                requestSheet.screenFields = arrayFields.ToArray();
                                replySheet = proxySheet.submit(opSheet, requestSheet);
                            }
                            if (_eFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                if (replySheet.mapName == "MSM200A")
                                {
                                    var replyFields = new ArrayScreenNameValue(replySheet.screenFields);
                                    _cells.GetCell(1, currentRow).Value = replyFields.GetField("SUPPLIER_NO1I").value;
                                    employee.Supplier = replyFields.GetField("SUPPLIER_NO1I").value;
                                    requestSheet.screenKey = "1";
                                    arrayFields = new ArrayScreenNameValue();
                                    requestSheet.screenFields = arrayFields.ToArray();
                                    replySheet = proxySheet.submit(opSheet, requestSheet);
                                    while (_eFunctions.CheckReplyWarning(replySheet))
                                        replySheet = proxySheet.submit(opSheet, requestSheet);
                                    if (replySheet.message.Contains("Confirm"))
                                        replySheet = proxySheet.submit(opSheet, requestSheet);
                                    if (_eFunctions.CheckReplyError(replySheet))
                                    {
                                        _cells.GetCell(_resultColumn, currentRow).Style =
                                            _cells.GetStyle(StyleConstants.Error);
                                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                                    }
                                    else
                                    {
                                        if (replySheet.mapName != "MSM20DA") continue;
                                        arrayFields = new ArrayScreenNameValue();
                                        arrayFields.Add("BANK_ACCNT1I", employee.BankAccount);
                                        arrayFields.Add("ACCT_NAME1I", employee.BankAccountName);
                                        arrayFields.Add("DEF_BNK_BRNCH1I", employee.DefaultBankBranch);
                                        arrayFields.Add("DEF_BNK_ACCT1I", employee.DefaultBankAccount);
                                        arrayFields.Add("DEF_BNK_NAME1I", employee.DefaultBankAccountName);
                                        requestSheet.screenFields = arrayFields.ToArray();
                                        requestSheet.screenKey = "1";
                                        replySheet = proxySheet.submit(opSheet, requestSheet);
                                        while (_eFunctions.CheckReplyWarning(replySheet))
                                            replySheet = proxySheet.submit(opSheet, requestSheet);
                                        if (replySheet.message.Contains("Confirm"))
                                            replySheet = proxySheet.submit(opSheet, requestSheet);
                                        if (_eFunctions.CheckReplyError(replySheet))
                                        {
                                            _cells.GetCell(_resultColumn, currentRow).Style =
                                                _cells.GetStyle(StyleConstants.Error);
                                            
                                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                                        }
                                        else
                                        {
                                            arrayFields = new ArrayScreenNameValue();
                                            requestSheet.screenFields = arrayFields.ToArray();
                                            requestSheet.screenKey = "1";
                                            replySheet = proxySheet.submit(opSheet, requestSheet);
                                            while (_eFunctions.CheckReplyWarning(replySheet))
                                                replySheet = proxySheet.submit(opSheet, requestSheet);
                                            if (replySheet.message.Contains("Confirm"))
                                                replySheet = proxySheet.submit(opSheet, requestSheet);
                                            if (_eFunctions.CheckReplyError(replySheet))
                                            {
                                                _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                                            }
                                            else
                                            {
                                                _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                                                _cells.GetCell(_resultColumn, currentRow).Value = "Cambiado";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                                    _cells.GetCell(_resultColumn, currentRow).Value = "No se Actualizo";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        

        



        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnSuspender_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(LoadSuspenderSupplier);
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnLoadData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
    }
}
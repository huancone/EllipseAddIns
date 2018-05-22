using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseMSO200ExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        private const int TittleRow = 5;
        private static int _resultColumn = 13;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private string _sheetName01;
        public List<Bancos> ListaBancos;

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

        private void btnInactivateSupplier_Click(object sender, RibbonControlEventArgs e)
        {
            FormatInactivateSupplier();
        }

        private void btnFormatInactiveBusiness_Click(object sender, RibbonControlEventArgs e)
        {
            FormatInactivateSupplierBusiness();
        }

        private void FormatInactivateSupplierBusiness()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;
            _sheetName01 = "Inactivar Supplier Business";
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

        private void FormatInactivateSupplier()
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

        private void ValidateAccounts_Click(object sender, RibbonControlEventArgs e)
        {
            ValidateBancos();
        }

        private void ValidateBancos()
        {
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
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

                    var existe = false;
                    // ReSharper disable once UnusedVariable
                    foreach (var banco in ListaBancos.Where(banco => employee.CodigoBanco.PadLeft(2, '0') == banco.CodigoNomina.PadLeft(2, '0')))
                    {
                        existe = true;
                        employee.TipoCuenta = employee.TipoCuenta.ToUpper() == "AHORRO" ? "02" : "01";
                        _cells.GetCell(9, currentRow).Value = "2-" + employee.CodigoBanco.PadLeft(4, '0') + "-0000-" + employee.TipoCuenta;
                    }

                    _cells.GetCell(9, currentRow).Value = existe == false? "Verificar Banco": _cells.GetCell(9, currentRow).Value;
                    _cells.GetCell(9, currentRow).Style = existe == false? _cells.GetStyle(StyleConstants.Error): _cells.GetStyle(StyleConstants.Normal);
                    _cells.GetCell(10, currentRow).Value = "NM009";
                    _cells.GetCell(11, currentRow).Value = "1209";
                    _cells.GetCell(12, currentRow).Value = "47703371338";


                    var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", employee.Cedula, _eFunctions.dbReference, _eFunctions.dbLink);
                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
            LoadData();
        }

        private void LoadData()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            Worksheet excelSheet = excelBook.ActiveSheet;
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            switch (excelSheet.Name)
            {
                case "MSO200 Cambio Cuentas":
                    //LoadCesantias();
                    LoadCambioCuentas();
                    break;
                case "Inactivar Supplier Business":
                    //LoadCesantias();
                    LoadInactivarSupplierBusiness();
                    break;
                case "Inactivar Supplier":
                    //LoadCesantias();
                    LoadInactivarSupplier();
                    break;
            }
        }

        private void LoadInactivarSupplierBusiness()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
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
                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDsct);
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

        private void LoadInactivarSupplier()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
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
                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDsct);
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

        private void LoadCambioCuentas()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
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
                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDsct);
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
                                    _cells.GetCell(_resultColumn, currentRow).Style =
                                        _cells.GetStyle(StyleConstants.Error);
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

        public class Bancos
        {
            [CsvColumn(FieldIndex = 1)]
            public string CodigoNomina { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string CodigoMims { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string NombreInstitucion { get; set; }
        }

        public class EmployeeInfo
        {
            public string Supplier { get; set; }
            public string Nombre { get; set; }
            public string CodigoBanco { get; set; }
            public string NombreBanco { get; set; }
            public string TipoCuenta { get; set; }
            public string NumeroCuenta { get; set; }
            public string Cedula { get; set; }
            public string BankAccount { get; set; }
            public string BankAccountName { get; set; }
            public string DefaultBankBranch { get; set; }
            public string DefaultBankAccount { get; set; }
            public string DefaultBankAccountName { get; set; }
        }

        public static class Queries
        {
            public static string GetSupplierInvoiceInfo(string districtCode, string cedula, string dbReference, string dbLink)
            {
                var sqlQuery = "SELECT " +
                               "  TRIM(BI.BANK_ACCT_NO) BANK_ACCT_NO, " +
                               "  TRIM(BI.TAX_FILE_NO) TAX_FILE_NO, " +
                               "  TRIM(BI.SUPPLIER_NO) SUPPLIER_NO, " +
                               "  COUNT(BI.SUPPLIER_NO) OVER(PARTITION BY BI.TAX_FILE_NO) CANTIDAD_REGISTROS " +
                               "FROM " +
                               "  ELLIPSE.MSF203 BI " +
                               "WHERE " +
                               "  BI.TAX_FILE_NO = '" + cedula + "' " +
                               "  AND BI.DSTRCT_CODE = '" + districtCode + "' ";
                return sqlQuery;
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
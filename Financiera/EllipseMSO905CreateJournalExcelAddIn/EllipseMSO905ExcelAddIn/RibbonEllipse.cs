using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService;

namespace EllipseMSO905ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private ExcelStyleCells _cells;
        private Application _excelApp;
        const int ResultColumn01 = 9;
        const string SheetName01 = "MSO905";
        const int TitleRow = 11;
        const string TableName01 = "JournalTable";


        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            var settings = new SharedClassLibrary.Ellipse.Settings();
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
        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;

            excelSheet.Name = SheetName01;
            var titleRow = TitleRow;
            var resultColumn = ResultColumn01;

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Create Journal";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(2, 1, 7, 2);

            _cells.GetCell(1, 4).Value = "Accounting Period (MM/YY) *";
            _cells.GetCell(1, 5).Value = "Journal         *";
            _cells.GetCell(1, 6).Value = "Jnl Type";
            _cells.GetCell(1, 7).Value = "Description     *";
            _cells.GetCell(1, 8).Value = "Accrual Journal *";
            _cells.GetCell(1, 9).Value = "Rate";
            _cells.GetRange(1, 4, 1, 9).Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetRange(1, 4, 1, 9).NumberFormat = "@";

            _cells.GetCell(1, titleRow).Value = "Account Code";
            _cells.GetCell(2, titleRow).Value = "W/Order Or Project";
            _cells.GetCell(3, titleRow).Value = "W/P";
            _cells.GetCell(4, titleRow).Value = "Journal Description";
            _cells.GetCell(5, titleRow).Value = "Amount (+/-) Pesos";
            _cells.GetCell(6, titleRow).Value = "Document Ref";
            _cells.GetCell(7, titleRow).Value = "Foreign";
            _cells.GetCell(8, titleRow).Value = "Dolars";
            _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell(resultColumn, titleRow).Value = "Result";
            _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableName01);

            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
        }

        private void btnValidate_Click(object sender, RibbonControlEventArgs e)
        {
            if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ValidarDatos);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ValidarDatos()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var titleRow = TitleRow;
            var resultColumn = ResultColumn01;
            var currentRow = titleRow + 1;

            _cells.GetRange(1, currentRow, resultColumn, currentRow).Style = StyleConstants.Normal;
            _cells.ClearTableRangeColumn(TableName01, resultColumn);
            
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    //Valida Centro de Costo
                    var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    var projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) ?? "";
                    var projectInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) ?? "";

                    var accountCode = new AccountCode(_eFunctions, "ICOR", account);
                    _cells.GetCell(resultColumn, currentRow).Value += accountCode.Error;

                    if (accountCode.Error == null && accountCode.ActiveStatus == "A") continue;

                    _cells.GetRange(1, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Success);


                    if (accountCode.Error != null)
                    {
                        _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    }

                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
                        _cells.GetCell(2, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        _cells.GetRange(1, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    if (accountCode.WorkOrderEntryInd == "M" && (projectNo == "" || projectInd == "P"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Numero de Orden Requerido";
                        _cells.GetRange(1, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    //valido si se necesita Subledger
                    if (accountCode.SubLedgerInd == "M" && !projectNo.Contains(";"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Subledger Requerido";
                        _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    var nit = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);

                    if (nit.Contains("#") & nit.Contains("@"))
                    {

                        var startIndex = nit.IndexOf("#", 1, StringComparison.Ordinal);
                        var endIndex = nit.IndexOf("@", startIndex + 1, StringComparison.Ordinal);
                        nit = nit.Substring(startIndex + 1, endIndex - startIndex - 1);

                        var ellipseNit = new EllipseNit(_eFunctions, nit);

                        _cells.GetCell(4, currentRow).Style = (ellipseNit.Nit == null) ? _cells.GetStyle(StyleConstants.Error) : _cells.GetStyle(StyleConstants.Success);
                        _cells.GetCell(resultColumn + 1, currentRow).Value += ellipseNit.Error;

                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, currentRow).Value = "ERROR:  " + ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }

            ValidarDatos2();
        }

        private void ValidarDatos2()
        {
            var resultColumn = ResultColumn01;
            var currentRow = TitleRow + 1;
            try
            {
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var proxySheet = new Screen.ScreenService();
                var requestSheet = new Screen.ScreenSubmitRequestDTO();

                proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };
                _cells.GetCell(7, currentRow).Select();

                const string option = "3";
                var fullPeriod = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value);
                const string foreignInd = "Y";
                var journalNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value);
                var journalType = _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value);
                var journalDesc = _cells.GetEmptyIfNull(_cells.GetCell(2, 7).Value);
                var accrualJournal = _cells.GetEmptyIfNull(_cells.GetCell(2, 8).Value);

                _eFunctions.RevertOperation(opSheet, proxySheet);
                var replySheet = proxySheet.executeScreen(opSheet, "MSO905");

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    if (replySheet.mapName != "MSM905A") return;
                    var arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("OPTION1I", option);
                    arrayFields.Add("ACCT_PERIOD1I", fullPeriod);
                    arrayFields.Add("FOREIGN_IND1I", foreignInd);
                    arrayFields.Add("MAN_JNL_NO1I", journalNo);
                    requestSheet.screenFields = arrayFields.ToArray();

                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                    while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                    }
                    else if (replySheet.mapName == "MSM907A")
                    {
                        var currentMso = 1;

                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("JOURNAL_TYPE1I", journalType);
                        arrayFields.Add("JOURNAL_DESC1I", journalDesc);
                        arrayFields.Add("ACCRUAL_IND1I", accrualJournal);
                        arrayFields.Add("APPROVAL_STAT1I", "Y");
                        arrayFields.Add("ACCOUNTANT1I", _frmAuth.EllipseUser);

                        while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                        {
                            _cells.GetCell(1, currentRow).Select();
                            var accountCode = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                            var workOrderProject = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                            var workOrderProjectInd = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                            var journalDescItem = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                            var memoAmount = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                            var documentReference = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);
                            var foreingCurrency = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);
                            var tranAmount = _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value);

                            arrayFields.Add("ACCOUNT_CODE1I" + currentMso, accountCode);
                            arrayFields.Add("WORK_PROJ1I" + currentMso, workOrderProject);
                            arrayFields.Add("WORK_PROJ_IND1I" + currentMso, workOrderProjectInd);
                            arrayFields.Add("JNL_DESC1I" + currentMso, journalDescItem);
                            arrayFields.Add("TRAN_AMOUNT1I" + currentMso, tranAmount);
                            arrayFields.Add("DOCUMENT_REF1I" + currentMso, documentReference);
                            arrayFields.Add("FOREIGN_CURR1I" + currentMso, foreingCurrency);
                            arrayFields.Add("MEMO_AMOUNT1I" + currentMso, memoAmount);
                            requestSheet.screenFields = arrayFields.ToArray();

                            if (currentMso == 3)
                            {
                                requestSheet.screenKey = "1";
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (_eFunctions.CheckReplyWarning(replySheet))
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                if (_eFunctions.CheckReplyError(replySheet))
                                {
                                    _cells.GetRange(1, currentRow - currentMso + 1, resultColumn, currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                                }
                                else
                                {
                                    _cells.GetRange(1, currentRow - currentMso + 1, resultColumn, currentRow).Style = StyleConstants.Success;
                                }
                                currentMso = 1;
                            }
                            else
                                currentMso++;

                            currentRow++;
                        }

                        if (replySheet.mapName != "MSM907A") return;

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetRange(1, currentRow - currentMso, resultColumn, currentRow - 1).Style = StyleConstants.Error;
                            _cells.GetCell(resultColumn, currentRow - 1).Value = replySheet.message;
                        }
                        else
                        {
                            _cells.GetRange(1, currentRow - currentMso, resultColumn, currentRow - 1).Style = StyleConstants.Success;
                            _cells.GetCell(2, 5).Value = replySheet.message ?? _cells.GetCell(2, 5).Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(resultColumn, currentRow).Value = "ERROR:  " + ex.Message;
                //                ErrorLogger.LogError("RibbonEllipse.cs:LoadSheet()", ex.Message, EFunctions.DebugErrors);
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(LoadSheet);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void LoadSheet()
        {
            var titleRow = TitleRow;
            var resultColumn = ResultColumn01;
            var currentRow = titleRow + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.GetRange(1, currentRow, resultColumn, currentRow).Style = StyleConstants.Normal;
                _cells.ClearTableRangeColumn(TableName01, resultColumn);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var proxySheet = new Screen.ScreenService();
                var requestSheet = new Screen.ScreenSubmitRequestDTO();

                proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };
                _cells.GetCell(7, currentRow).Select();

                const string option = "3";
                var fullPeriod = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value);
                const string foreignInd = "Y";
                var journalNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value);
                var journalType = _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value);
                var journalDesc = _cells.GetEmptyIfNull(_cells.GetCell(2, 7).Value);
                var accrualJournal = _cells.GetEmptyIfNull(_cells.GetCell(2, 8).Value);

                _eFunctions.RevertOperation(opSheet, proxySheet);
                var replySheet = proxySheet.executeScreen(opSheet, "MSO905");

                if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                }
                else
                {
                    if (replySheet.mapName != "MSM905A") return;
                    var arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("OPTION1I", option);
                    arrayFields.Add("ACCT_PERIOD1I", fullPeriod);
                    arrayFields.Add("FOREIGN_IND1I", foreignInd);
                    arrayFields.Add("MAN_JNL_NO1I", journalNo);
                    requestSheet.screenFields = arrayFields.ToArray();

                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                    while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                    }
                    else if (replySheet.mapName == "MSM907A")
                    {
                        var currentMso = 1;

                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("JOURNAL_TYPE1I", journalType);
                        arrayFields.Add("JOURNAL_DESC1I", journalDesc);
                        arrayFields.Add("ACCRUAL_IND1I", accrualJournal);
                        arrayFields.Add("APPROVAL_STAT1I", "Y");
                        arrayFields.Add("ACCOUNTANT1I", _frmAuth.EllipseUser);

                        while (_cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value) != "")
                        {
                            _cells.GetCell(1, currentRow).Select();
                            var accountCode = _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                            var workOrderProject = _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value);
                            var workOrderProjectInd = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value);
                            var journalDescItem = _cells.GetEmptyIfNull(_cells.GetCell(4, currentRow).Value);
                            var memoAmount = _cells.GetEmptyIfNull(_cells.GetCell(5, currentRow).Value);
                            var documentReference = _cells.GetEmptyIfNull(_cells.GetCell(6, currentRow).Value);
                            var foreingCurrency = _cells.GetEmptyIfNull(_cells.GetCell(7, currentRow).Value);
                            var tranAmount = _cells.GetEmptyIfNull(_cells.GetCell(8, currentRow).Value);

                            arrayFields.Add("ACCOUNT_CODE1I" + currentMso, accountCode);
                            arrayFields.Add("WORK_PROJ1I" + currentMso, workOrderProject);
                            arrayFields.Add("WORK_PROJ_IND1I" + currentMso, workOrderProjectInd);
                            arrayFields.Add("JNL_DESC1I" + currentMso, journalDescItem);
                            arrayFields.Add("TRAN_AMOUNT1I" + currentMso, tranAmount);
                            arrayFields.Add("DOCUMENT_REF1I" + currentMso, documentReference);
                            arrayFields.Add("FOREIGN_CURR1I" + currentMso, foreingCurrency);
                            arrayFields.Add("MEMO_AMOUNT1I" + currentMso, memoAmount);
                            requestSheet.screenFields = arrayFields.ToArray();

                            if (currentMso == 3)
                            {
                                requestSheet.screenKey = "1";
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (_eFunctions.CheckReplyWarning(replySheet))
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (replySheet.functionKeys.Contains("XMIT-Confirm"))
                                {
                                    replySheet = proxySheet.submit(opSheet, requestSheet);
                                    _cells.GetCell(3, 5).Value = replySheet.message;
                                    arrayFields = new ArrayScreenNameValue();
                                    requestSheet.screenFields = arrayFields.ToArray();
                                }

                                if (_eFunctions.CheckReplyError(replySheet))
                                {
                                    _cells.GetRange(1, currentRow - currentMso + 1, resultColumn, currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                                }
                                else
                                {
                                    _cells.GetRange(1, currentRow - currentMso + 1, resultColumn, currentRow).Style = StyleConstants.Success;
                                }
                                currentMso = 1;
                            }
                            else
                                currentMso++;

                            currentRow++;
                        }

                        if (replySheet.mapName != "MSM907A") return;
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (_eFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (replySheet.functionKeys.Contains("XMIT-Confirm"))
                        {
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            _cells.GetCell(3, 5).Value = replySheet.message;
                            arrayFields = new ArrayScreenNameValue();
                            requestSheet.screenFields = arrayFields.ToArray();
                        }

                        if (_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetRange(1, currentRow - currentMso, resultColumn, currentRow - 1).Style = StyleConstants.Error;
                            _cells.GetCell(resultColumn, currentRow - 1).Value = replySheet.message;
                        }
                        else
                        {
                            _cells.GetRange(1, currentRow - currentMso, resultColumn, currentRow - 1).Style = StyleConstants.Success;
                            _cells.GetCell(2, 5).Value = replySheet.message ?? _cells.GetCell(2, 5).Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, currentRow).Style = StyleConstants.Error;
                _cells.GetCell(resultColumn, currentRow).Value = "ERROR:  " + ex.Message;
                //                ErrorLogger.LogError("RibbonEllipse.cs:LoadSheet()", ex.Message, EFunctions.DebugErrors);
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
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
    }
}

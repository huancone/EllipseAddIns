using System;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMSO905ExcelAddIn.Properties;

namespace EllipseMSO905ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int MaxRows = 1000;
        private static readonly EllipseFunctions EFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private int _resultColumn = 9;
        private readonly string _sheetName01 = "MSO905";
        private static int _tittleRow = 11;
        private int _currentRow = _tittleRow + 1;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
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

            excelSheet.Name = _sheetName01;

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

            _cells.GetCell(1, _tittleRow).Value = "Account Code";
            _cells.GetCell(2, _tittleRow).Value = "W/Order Or Project";
            _cells.GetCell(3, _tittleRow).Value = "W/P";
            _cells.GetCell(4, _tittleRow).Value = "Journal Description";
            _cells.GetCell(5, _tittleRow).Value = "Amount (+/-) Pesos";
            _cells.GetCell(6, _tittleRow).Value = "Document Ref";
            _cells.GetCell(7, _tittleRow).Value = "Foreign";
            _cells.GetCell(8, _tittleRow).Value = "Dolars";
            _cells.GetRange(1, _tittleRow, _resultColumn - 1, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell(_resultColumn, _tittleRow).Value = "Result";
            _cells.GetCell(_resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnValidate_Click(object sender, RibbonControlEventArgs e)
        {
            ValidarDatos();
        }

        private void ValidarDatos()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.GetCell(_resultColumn, _currentRow).ClearContents();
            _cells.GetCell(_resultColumn, _currentRow).Style = _cells.GetStyle(StyleConstants.Normal);

            _currentRow = _tittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, _currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, _currentRow).Select();

                    //Valida Centro de Costo
                    var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, _currentRow).Value);
                    var projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, _currentRow).Value) ?? "";
                    var projectInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, _currentRow).Value) ?? "";

                    var accountCode = new AccountCode("ICOR", account);
                    _cells.GetCell(_resultColumn, _currentRow).Value += accountCode.Error;

                    if (accountCode.Error == null && accountCode.ActiveStatus == "A")
                    {
                        _cells.GetRange(1, _currentRow, 3, _currentRow).Style =
                            _cells.GetStyle(StyleConstants.Success);
                    }

                    if (accountCode.Error != null)
                    {
                        _cells.GetCell(1, _currentRow).Style = StyleConstants.Error;
                    }

                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(_resultColumn, _currentRow).Value += " Numero de Proyecto Requerido";
                        _cells.GetCell(2, _currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        _cells.GetRange(1, _currentRow, 3, _currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    if (accountCode.WorkOrderEntryInd == "M" && (projectNo == "" || projectInd == "P"))
                    {
                        _cells.GetCell(_resultColumn, _currentRow).Value += " Numero de Orden Requerido";
                        _cells.GetRange(1, _currentRow, 3, _currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    //valido si se necesita Subledger
                    if (accountCode.SubLedgerInd == "M" && !projectNo.Contains(";"))
                    {
                        _cells.GetCell(_resultColumn, _currentRow).Value += " Subledger Requerido";
                        _cells.GetCell(6, _currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }

                    var nit = _cells.GetEmptyIfNull(_cells.GetCell(4, _currentRow).Value);

                    if (nit.Contains("#") & nit.Contains("@")) continue;

                    var startIndex = nit.IndexOf("#", 1, StringComparison.Ordinal);
                    var endIndex = nit.IndexOf("@", startIndex+1, StringComparison.Ordinal);
                    nit = nit.Substring(startIndex + 1, endIndex - startIndex - 1);

                    var ellipseNit = new EllipseNit(nit);

                    _cells.GetCell(4, _currentRow).Style = (ellipseNit.Nit == null) ? _cells.GetStyle(StyleConstants.Error) : _cells.GetStyle(StyleConstants.Success);
                    _cells.GetCell(_resultColumn + 1, _currentRow).Value += ellipseNit.Error;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, _currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, _currentRow).Value = "ERROR:  " + ex.Message;
                }
                finally
                {
                    _currentRow++;
                }
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                    LoadSheet();
            }
            else
                MessageBox.Show(Resources.RibbonEllipse_btnLoad_Click_Invalid_Format);
        }

        private void LoadSheet()
        {
            try
            {
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(_resultColumn, _tittleRow + 1, _resultColumn, MaxRows).Clear();

                _currentRow = _tittleRow + 1;
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var proxySheet = new screen.ScreenService();
                var requestSheet = new screen.ScreenSubmitRequestDTO();

                proxySheet.Url = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

                var opSheet = new screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings
                };
                _cells.GetCell(7, _currentRow).Select();

                const string option = "3";
                var fullPeriod = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value);
                const string foreignInd = "Y";
                var journalNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value);
                var journalType = _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value);
                var journalDesc = _cells.GetEmptyIfNull(_cells.GetCell(2, 7).Value);
                var accrualJournal = _cells.GetEmptyIfNull(_cells.GetCell(2, 8).Value);

                EFunctions.RevertOperation(opSheet, proxySheet);
                var replySheet = proxySheet.executeScreen(opSheet, "MSO905");

                if (EFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell(_resultColumn, _currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, _currentRow).Value = replySheet.message;
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

                    while (EFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, _currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, _currentRow).Value = replySheet.message;
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

                        while (_cells.GetEmptyIfNull(_cells.GetCell(1, _currentRow).Value) != "")
                        {


                            _cells.GetCell(1, _currentRow).Select();
                            var accountCode = _cells.GetEmptyIfNull(_cells.GetCell(1, _currentRow).Value);
                            var workOrderProject = _cells.GetEmptyIfNull(_cells.GetCell(2, _currentRow).Value);
                            var workOrderProjectInd = _cells.GetEmptyIfNull(_cells.GetCell(3, _currentRow).Value);
                            var journalDescItem = _cells.GetEmptyIfNull(_cells.GetCell(4, _currentRow).Value);
                            var memoAmount = _cells.GetEmptyIfNull(_cells.GetCell(5, _currentRow).Value);
                            var documentReference = _cells.GetEmptyIfNull(_cells.GetCell(6, _currentRow).Value);
                            var foreingCurrency = _cells.GetEmptyIfNull(_cells.GetCell(7, _currentRow).Value);
                            var tranAmount = _cells.GetEmptyIfNull(_cells.GetCell(8, _currentRow).Value);

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

                                while (EFunctions.CheckReplyWarning(replySheet))
                                    replySheet = proxySheet.submit(opSheet, requestSheet);

                                while (replySheet.functionKeys.Contains("XMIT-Confirm"))
                                {
                                    replySheet = proxySheet.submit(opSheet, requestSheet);
                                    _cells.GetCell(3, 5).Value = replySheet.message;
                                    arrayFields = new ArrayScreenNameValue();
                                    requestSheet.screenFields = arrayFields.ToArray();
                                }

                                if (EFunctions.CheckReplyError(replySheet))
                                {
                                    _cells.GetRange(1, _currentRow - currentMso + 1, _resultColumn, _currentRow).Style = StyleConstants.Error;
                                    _cells.GetCell(_resultColumn, _currentRow).Value = replySheet.message;
                                }
                                else
                                {
                                    _cells.GetRange(1, _currentRow - currentMso + 1, _resultColumn, _currentRow).Style = StyleConstants.Success;
                                }
                                currentMso = 1;
                            }
                            else
                                currentMso++;

                            _currentRow++;
                        }

                        if (replySheet.mapName != "MSM907A") return;
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (EFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (replySheet.functionKeys.Contains("XMIT-Confirm"))
                        {
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            _cells.GetCell(3, 5).Value = replySheet.message;
                            arrayFields = new ArrayScreenNameValue();
                            requestSheet.screenFields = arrayFields.ToArray();
                        }

                        if (EFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetRange(1, _currentRow - currentMso, _resultColumn, _currentRow - 1).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, _currentRow - 1).Value = replySheet.message;
                        }
                        else
                        {
                            _cells.GetRange(1, _currentRow - currentMso, _resultColumn, _currentRow - 1).Style = StyleConstants.Success;
                            _cells.GetCell(2, 5).Value = replySheet.message ?? _cells.GetCell(2, 5).Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _cells.GetCell(1, _currentRow).Style = StyleConstants.Error;
                _cells.GetCell(_resultColumn, _currentRow).Value = "ERROR:  " + ex.Message;
                //                ErrorLogger.LogError("RibbonEllipse.cs:LoadSheet()", ex.Message, EFunctions.DebugErrors);
            }

        }

        public class EllipseNit
        {
            public string Nit { get; set; }
            public string Error { get; set; }

            public EllipseNit(string nit)
            {
                try
                {
                    if (string.IsNullOrEmpty(nit))
                    {
                        Error = "Nit Invalido";
                        return;
                    }

                    var ellipseNitQuery = Queries.GetSupplierNit(nit, EFunctions.dbReference, EFunctions.dbLink);
                    var drEllipseNit = EFunctions.GetQueryResult(ellipseNitQuery);
                    if (drEllipseNit != null && !drEllipseNit.IsClosed && drEllipseNit.HasRows)
                    {
                        drEllipseNit.Read();
                        Nit = drEllipseNit["NIT"].ToString();
                        Error = null;
                    }
                    else
                    {
                        Nit = null;
                        Error = "Nit no registrado";
                    }
                    if (drEllipseNit != null) drEllipseNit.Close();
                }
                catch (Exception error)
                {
                    Error = error.Message;
                }
            }
        }

        public class AccountCode
        {
            public AccountCode(string districtCode, string accountCode)
            {
                try
                {
                    if (string.IsNullOrEmpty(districtCode) || string.IsNullOrEmpty(accountCode))
                    {
                        Error = "AccoundeCode Invalida";
                        return;
                    }

                    if (accountCode.Contains(";"))
                    {
                        Mnemonic = accountCode.Contains(";")
                        ? accountCode.Substring(accountCode.IndexOf(";", StringComparison.Ordinal) + 1, accountCode.Length - accountCode.IndexOf(";", StringComparison.Ordinal) - 1).Replace("=", "")
                        : "";
                        accountCode = accountCode.Contains(";")
                            ? accountCode.Substring(0, accountCode.IndexOf(";", StringComparison.Ordinal))
                            : accountCode;



                        var mnemonicQuery = Queries.GetSupplierMnemonic(Mnemonic, EFunctions.dbReference, EFunctions.dbLink);
                        var drMnemonic = EFunctions.GetQueryResult(mnemonicQuery);
                        if (drMnemonic != null && !drMnemonic.IsClosed && drMnemonic.HasRows)
                        {
                            drMnemonic.Read();
                            if (Convert.ToInt32(drMnemonic["CANTIDAD"].ToString()) == 1)
                                Account = accountCode + ";" + drMnemonic["GL_COLLOQ_CD"];
                            else
                                Error = " No se puede determinar la cuenta";
                        }
                        else
                            Error = " Mnenomico No Valido";
                        if (drMnemonic != null) drMnemonic.Close();
                    }

                    var sqlQuery = Queries.GetAccountCodeInfo(districtCode, accountCode, EFunctions.dbReference,
                        EFunctions.dbLink);

                    var drAccountCode = EFunctions.GetQueryResult(sqlQuery);

                    if (drAccountCode != null && !drAccountCode.IsClosed && drAccountCode.HasRows)
                    {
                        while (drAccountCode.Read())
                        {
                            ActiveStatus = drAccountCode["ACTIVE_STATUS"].ToString();
                            ProjectEntriInd = drAccountCode["PROJ_ENTRY_IND"].ToString();
                            WorkOrderEntryInd = drAccountCode["WO_ENTRY_IND"].ToString();
                            SubLedgerInd = drAccountCode["SUBLEDGER_IND"].ToString();
                            Error = (drAccountCode["ACTIVE_STATUS"].ToString() == "I") ? ", AccountCode Inactivo" : Error;
                        }
                    }
                    else
                    {
                        Error = " Centro de Costos No Valido";
                    }
                    if (drAccountCode != null) drAccountCode.Close();
                }
                catch (Exception error)
                {
                    Error = error.Message;
                }
            }

            public string Error { get; set; }
            public string ActiveStatus { get; set; }
            public string Account { get; set; }
            public string ProjectEntriInd { get; set; }
            public string WorkOrderEntryInd { get; set; }
            public string SubLedgerInd { get; set; }
            public string Mnemonic { get; set; }
        }

        public static class Queries
        {
            public static string GetEmployeeName(string employeeId, string dbReference, string dbLink)
            {
                var sqlQuery = " " +
                               "SELECT DISTINCT " +
                               "  EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE " +
                               "FROM " +
                               "  " + dbReference + ".MSF870" + dbLink + " POS " +
                               "INNER JOIN " + dbReference + ".MSF878" + dbLink + " EMPOS " +
                               "ON" +
                               "  EMPOS.POSITION_ID = POS.POSITION_ID " +
                               "AND " +
                               "  (" +
                               "    EMPOS.POS_STOP_DATE > TO_CHAR ( SYSDATE, 'YYYYMMDD' ) " +
                               "  OR EMPOS.POS_STOP_DATE = '00000000' " +
                               "  ) " +
                               "INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP " +
                               "ON " +
                               "  EMPOS.EMPLOYEE_ID = EMP.EMPLOYEE_ID " +
                               "WHERE " +
                               "EMPOS.EMPLOYEE_ID = '" + employeeId + "' ";
                return sqlQuery;
            }

            public static string GetTransactionInfo(string districtCode, string numTransaction, string dbReference, string dbLink)
            {
                var processDate = numTransaction.Substring(0, 8);
                var transNo = numTransaction.Substring(8, 11);
                var userNo = numTransaction.Substring(19, 4);
                var recType = numTransaction.Substring(23, 1);
                var sqlQuery = " " +
                               "SELECT " +
                               "  TR.FULL_PERIOD, " +
                               "  TR.ACCOUNT_CODE, " +
                               "  DECODE ( TRIM ( TR.PROJECT_NO ), NULL, DECODE ( TRIM ( TR.WORK_ORDER ), NULL, ' ', TR.WORK_ORDER ), TR.PROJECT_NO ) PROJECT_NO, " +
                               "  DECODE ( TRIM ( TR.PROJECT_NO ), NULL, DECODE ( TRIM ( TR.WORK_ORDER ), NULL, ' ', 'W' ), 'P' ) IND," +
                               "  TR.TRAN_AMOUNT, " +
                               "  TR.TRAN_AMOUNT_S " +
                               "FROM " +
                               "  " + dbReference + ".MSF900" + dbLink + " TR " +
                               "WHERE " +
                               "  TR.DSTRCT_CODE = '" + districtCode + "' " +
                               "AND TR.PROCESS_DATE = '" + processDate + "' " +
                               "AND TR.TRANS_NO = '" + transNo + "' " +
                               "AND TR.USERNO = '" + userNo + "' " +
                               "AND TR.REC900_TYPE = '" + recType + "' ";

                return sqlQuery;
            }

            public static string GetAccountCodeInfo(string districtCode, string accountCode, string dbReference, string dbLink)
            {
                var sqlQuery = " " +
                               "SELECT " +
                               "  CC.ACTIVE_STATUS, " +
                               "  CC.ACCOUNT_CODE, " +
                               "  CC.PROJ_ENTRY_IND, " +
                               "  CC.WO_ENTRY_IND, " +
                               "  CC.SUBLEDGER_IND " +
                               "FROM " +
                               "  " + dbReference + ".MSF966" + dbLink + " CC " +
                               "WHERE " +
                               "  CC.DSTRCT_CODE = '" + districtCode + "' " +
                               "AND CC.ACCOUNT_CODE = '" + accountCode + "'";
                return sqlQuery;
            }

            public static string GetSupplierName(string districtCode, string supplierId, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
                               "SELECT " +
                               "  SUP.SUPPLIER_NO, " +
                               "  SUP.SUPPLIER_NAME " +
                               "FROM " +
                               "  " + dbReference + ".MSF200 SUP" + dbLink + " " +
                               "INNER JOIN " + dbReference + ".MSF203" + dbLink + " SD " +
                               "ON " +
                               "  SD.SUPPLIER_NO = SUP.SUPPLIER_NO " +
                               "WHERE " +
                               "  SUP.SUPPLIER_NO = '" + supplierId + "' " +
                               "  AND SD.DSTRCT_CODE = '" + districtCode + "'";
                return sqlQuery;
            }

            public static string GetContractNameDesc(string document, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
                               "SELECT " +
                               "  CONTRACT_DESC " +
                               "FROM " +
                               "  " + dbReference + ".MSF384" + dbLink + " " +
                               "WHERE " +
                               "  CONTRACT_NO = '" + document + "'";
                return sqlQuery;
            }

            public static string GetPurchaseOrder(string document, string supplierNo, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
                               "SELECT " +
                               "  PO_NO " +
                               "FROM " +
                               "  " + dbReference + ".MSF220" + dbLink + " " +
                               "WHERE " +
                               "  PO_NO = '" + document + "' " +
                               "AND SUPPLIER_NO = '" + supplierNo + "'";
                return sqlQuery;
            }

            public static string GetSupplierMnemonic(string mnemonic, string dbReference, string dbLink)
            {
                var sqlQuery = "" +
                               "SELECT " +
                               "  GL_COLLOQ_CD, " +
                               "  COUNT ( * ) OVER ( ) CANTIDAD " +
                               "FROM " +
                               "  " + dbReference + ".MSF922" + dbLink + " " +
                               "WHERE " +
                               "  GL_COLLOQ_TY = '7' " +
                               "AND COLLOQ_NAME = '" + mnemonic + "'";
                return sqlQuery;
            }

            public static string GetJournalInfo(string districtCode, string journal, string dbReference, string dbLink)
            {
                var sqlQuery = " " +
                    "SELECT " +
                    "  TR.DSTRCT_CODE, " +
                    "  TR.TRAN_GROUP_KEY, " +
                    "  ( TR.PROCESS_DATE || TR.TRANSACTION_NO || TR.USERNO || TR.REC900_TYPE ) NUMTXT, " +
                    "  TR.FULL_PERIOD, " +
                    "  TR.REC900_TYPE, " +
                    "  TR.TRAN_TYPE, " +
                    "  TR.POSTED_STATUS, " +
                    "  TR.MANJNL_VCHR, " +
                    "  TR.ACCOUNTANT, " +
                    "  TR.ACCOUNT_CODE, " +
                    "  TR.PROJECT_NO, " +
                    "  TR.TRAN_AMOUNT, " +
                    "  TR.TRAN_AMOUNT_S, " +
                    "  TR.CURRENCY_TYPE, " +
                    "  TR.CREATION_DATE, " +
                    "  TR.CREATION_TIME, " +
                    "  TR.CREATION_USER, " +
                    "  TR.MIMS_SL_KEY, " +
                    "  TR.JOURNAL_DESC, " +
                    "  TR.DOCUMENT_REF, " +
                    "  TR.AUTO_JNL_FLG, " +
                    "  TR.JOURNAL_TYPE " +
                    "FROM " +
                    "  " + dbReference + ".MSF900" + dbLink + " TR " +
                    "INNER JOIN " + dbReference + ".MSFX90 X90" + dbLink + " " +
                    "ON " +
                    "  TR.TRAN_GROUP_KEY = X90.DSTRCT_CODE || X90.PROCESS_DATE || X90.TRANSACTION_NO || X90.USERNO || X90.REC900_TYPE " +
                    "WHERE " +
                    "  X90.DSTRCT_CODE = '" + districtCode + "' " +
                    "AND X90.JOURNAL_NO = '" + journal + "' " +
                    "ORDER BY " +
                    "  1, " +
                    "  2, " +
                    "  3 ";

                return sqlQuery;
            }

            public static string GetSupplierNit(string nit, string dbReference, string dbLink)
            {
                var sqlQuery = " " +
                               "SELECT " +
                               " DISTINCT NIT " +
                               "FROM " +
                               "  ( " +
                               "    SELECT " +
                               "      trim(TAX_FILE_NO) NIT " +
                               "    FROM " +
                               "      " + dbReference + ".MSF203" + dbLink + " " +
                               "    WHERE " +
                               "      TRIM (TAX_FILE_NO) = '" + nit + "' " +
                               "    UNION " +
                               "    SELECT " +
                               "      TRIM(GOVT_ID_NO) NIT " +
                               "    FROM " +
                               "      " + dbReference + ".MSF503" + dbLink + " " +
                               "    WHERE " +
                               "      TRIM (GOVT_ID_NO) = '" + nit + "' " +
                               "    UNION " +
                               "    SELECT " +
                               "      TRIM(TABLE_CODE) NIT " +
                               "    FROM " +
                               "      " + dbReference + ".MSF010" + dbLink + " " +
                               "    WHERE " +
                               "      TABLE_TYPE          = '+NIT' " +
                               "    AND TRIM (TABLE_CODE) = '" + nit + "' " +
                               "  )  ";

                return sqlQuery;
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

    }
}


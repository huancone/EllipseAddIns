using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseMSQ901ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int MaxRows = 1000;
        private EllipseFunctions _eFunctions;
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private int _resultColumn = 35;
        private const string SheetNameSupplier = "MSQ901-ConsultaSupplier";
        private const string SheetNameJournal = "MSQ901-ConsultaJournal";
        private const string SheetNameCustomer = "MSQ901-ConsultaCustomer";
        //private string _sheetName01;
        private int _tittleRow = 8;
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            var settings = new SharedClassLibrary.Ellipse.Settings();
            _eFunctions = new EllipseFunctions();
            //_frmAuth = new FormAuthenticate();
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
        private void btnConsultar_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                if (_thread != null && _thread.IsAlive) return;
                switch (_excelApp.ActiveWorkbook.ActiveSheet.Name)
                {
                    case SheetNameSupplier:
                    {
                        _thread = new Thread(InvoiceSupplierHeaderChange);
                        _thread.SetApartmentState(ApartmentState.STA);
                        _thread.Start();
                            break;
                    }
                    case SheetNameJournal:
                    {
                        _thread = new Thread(JournalHeaderRangeChange);
                        _thread.SetApartmentState(ApartmentState.STA);
                        _thread.Start();
                            break;
                    }
                    case SheetNameCustomer:
                    {
                        _thread = new Thread(CustomerInvoiceHeaderRangeChange);
                        _thread.SetApartmentState(ApartmentState.STA);
                        _thread.Start();
                            break;
                    }
                    default:
                    {
                        _thread = null;
                        break;
                    }
                }
                if (_thread == null)
                    return;
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:Consultar()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnJournal_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoJournal();
        }

        private void btnFormatoSupplierInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoSupplierInvoice();
        }

        private void FormatoSupplierInvoice()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();


                Microsoft.Office.Tools.Excel.Worksheet workSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var sheetName = "MSQ901-ConsultaSupplier";
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                #region Instructions

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "CONSULTA POR SUPPLIER E INVOICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
                _cells.MergeCells(2, 1, 7, 2);

                #endregion

                #region Datos
                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "SUPPLIER";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 6).Value = "INVOICE";
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);



                _cells.GetCell(1, _tittleRow).Value = "DSTRCT_CODE";
                _cells.GetCell(2, _tittleRow).Value = "TRAN_GROUP_KEY";
                _cells.GetCell(3, _tittleRow).Value = "NUMTXT";
                _cells.GetCell(4, _tittleRow).Value = "FULL_PERIOD";
                _cells.GetCell(5, _tittleRow).Value = "CREATION_USER";
                _cells.GetCell(6, _tittleRow).Value = "REC900_TYPE";
                _cells.GetCell(7, _tittleRow).Value = "TRAN_TYPE";
                _cells.GetCell(8, _tittleRow).Value = "POSTED_STATUS";
                _cells.GetCell(9, _tittleRow).Value = "SUPPLIER";
                _cells.GetCell(10, _tittleRow).Value = "EXT_INV_NO";
                _cells.GetCell(11, _tittleRow).Value = "ACCOUNT_CODE";
                _cells.GetCell(12, _tittleRow).Value = "PROJECT_NO";
                _cells.GetCell(13, _tittleRow).Value = "TRAN_AMOUNT";
                _cells.GetCell(14, _tittleRow).Value = "TRAN_AMOUNT_S";
                _cells.GetCell(15, _tittleRow).Value = "CURRENCY_TYPE";
                _cells.GetCell(16, _tittleRow).Value = "CREATION_DATE";
                _cells.GetCell(17, _tittleRow).Value = "CREATION_TIME";
                _cells.GetCell(18, _tittleRow).Value = "CREATION_USER";
                _cells.GetCell(19, _tittleRow).Value = "MIMS_SL_KEY";
                _cells.GetCell(20, _tittleRow).Value = "CONTRATO";
                _cells.GetCell(21, _tittleRow).Value = "ACTA";
                _cells.GetCell(22, _tittleRow).Value = "PO_NO";
                _cells.GetCell(23, _tittleRow).Value = "PO_ITEM";
                _cells.GetCell(24, _tittleRow).Value = "PORTION_NO";
                _cells.GetCell(25, _tittleRow).Value = "INV_ITEM_NO";
                _cells.GetCell(26, _tittleRow).Value = "INV_ITEM_DESC";
                _cells.GetCell(27, _tittleRow).Value = "RECEIVED_BY";
                _cells.GetCell(28, _tittleRow).Value = "AUTH_BY";
                _cells.GetCell(29, _tittleRow).Value = "ATAX_CODE";
                _cells.GetCell(30, _tittleRow).Value = "ATAX_RATE_9";
                _cells.GetCell(31, _tittleRow).Value = "OVERRIDE_SW";
                _cells.GetCell(32, _tittleRow).Value = "BRANCH_CODE";
                _cells.GetCell(33, _tittleRow).Value = "BANK_ACCT_NO";
                _cells.GetCell(34, _tittleRow).Value = "CHEQUE_RUN_NO";
                _cells.GetCell(35, _tittleRow).Value = "CHEQUE_NO";

                _cells.GetRange(1, _tittleRow, _resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                #endregion

                var invoiceSupplierHeaderRange = workSheet.Controls.AddNamedRange(workSheet.Range["B4", "B6"], "invoiceSupplierHeaderRange");
                invoiceSupplierHeaderRange.Change += invoiceSupplierHeaderRange_Change;


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

                _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatoSupplierInvoice()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
}

        private void FormatoJournal()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();


                Microsoft.Office.Tools.Excel.Worksheet workSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var sheetName = "MSQ901-ConsultaJournal";
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                #region Instructions

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "CONSULTA POR JOURNAL";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
                _cells.MergeCells(2, 1, 7, 2);

                #endregion

                #region Datos
                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "JOURNAL";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(1, _tittleRow).Value = "DSTRCT_CODE";
                _cells.GetCell(2, _tittleRow).Value = "TRAN_GROUP_KEY";
                _cells.GetCell(3, _tittleRow).Value = "NUMTXT";
                _cells.GetCell(4, _tittleRow).Value = "FULL_PERIOD";
                _cells.GetCell(5, _tittleRow).Value = "REC900_TYPE";
                _cells.GetCell(6, _tittleRow).Value = "TRAN_TYPE";
                _cells.GetCell(7, _tittleRow).Value = "POSTED_STATUS";
                _cells.GetCell(8, _tittleRow).Value = "MANJNL_VCHR";
                _cells.GetCell(9, _tittleRow).Value = "ACCOUNTANT";
                _cells.GetCell(10, _tittleRow).Value = "ACCOUNT_CODE";
                _cells.GetCell(11, _tittleRow).Value = "PROJECT_NO";
                _cells.GetCell(12, _tittleRow).Value = "TRAN_AMOUNT";
                _cells.GetCell(13, _tittleRow).Value = "TRAN_AMOUNT_S";
                _cells.GetCell(14, _tittleRow).Value = "CURRENCY_TYPE";
                _cells.GetCell(15, _tittleRow).Value = "CREATION_DATE";
                _cells.GetCell(16, _tittleRow).Value = "CREATION_TIME";
                _cells.GetCell(17, _tittleRow).Value = "CREATION_USER";
                _cells.GetCell(18, _tittleRow).Value = "MIMS_SL_KEY";
                _cells.GetCell(19, _tittleRow).Value = "JOURNAL_DESC";
                _cells.GetCell(20, _tittleRow).Value = "DOCUMENT_REF";
                _cells.GetCell(21, _tittleRow).Value = "AUTO_JNL_FLG";
                _cells.GetCell(22, _tittleRow).Value = "JOURNAL_TYPE";

                _cells.GetRange(1, _tittleRow, _resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                #endregion

                var journalHeaderRange = workSheet.Controls.AddNamedRange(workSheet.Range["B4", "B5"], "journalHeaderRange");
                journalHeaderRange.Change += journalHeaderRange_Change;


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

                _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatoJournal()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void FormatoCustomerInvoice()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();


                Microsoft.Office.Tools.Excel.Worksheet workSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var sheetName = "MSQ901-ConsultaCustomer";
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _resultColumn = 24;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "CONSULTA POR CUSTOMER E INVOICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
                _cells.MergeCells(2, 1, 7, 2);

                #region Datos
                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "CUSTOMER";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 6).Value = "INVOICE";
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(1, _tittleRow).Value = "DSTRCT_CODE";
                _cells.GetCell(2, _tittleRow).Value = "TRAN_GROUP_KEY";
                _cells.GetCell(3, _tittleRow).Value = "NUMTXT";
                _cells.GetCell(4, _tittleRow).Value = "FULL_PERIOD";
                _cells.GetCell(5, _tittleRow).Value = "REC900_TYPE";
                _cells.GetCell(6, _tittleRow).Value = "TRAN_TYPE";
                _cells.GetCell(7, _tittleRow).Value = "POSTED_STATUS";
                _cells.GetCell(8, _tittleRow).Value = "CUST_NO";
                _cells.GetCell(9, _tittleRow).Value = "AR_INV_NO";
                _cells.GetCell(10, _tittleRow).Value = "ACCOUNT_CODE";
                _cells.GetCell(11, _tittleRow).Value = "PROJECT_NO";
                _cells.GetCell(12, _tittleRow).Value = "TRAN_AMOUNT";
                _cells.GetCell(13, _tittleRow).Value = "TRAN_AMOUNT_S";
                _cells.GetCell(14, _tittleRow).Value = "CURRENCY_TYPE";
                _cells.GetCell(15, _tittleRow).Value = "CREATION_DATE";
                _cells.GetCell(16, _tittleRow).Value = "CREATION_TIME";
                _cells.GetCell(17, _tittleRow).Value = "CREATION_USER";
                _cells.GetCell(18, _tittleRow).Value = "MIMS_SL_KEY";
                _cells.GetCell(19, _tittleRow).Value = "INV_DATE";
                _cells.GetCell(20, _tittleRow).Value = "ITEM_NO";
                _cells.GetCell(21, _tittleRow).Value = "REVENUE_CODE";
                _cells.GetCell(22, _tittleRow).Value = "CUST_REF";
                _cells.GetCell(23, _tittleRow).Value = "SALES_PERSON_ID";

                _cells.GetRange(1, _tittleRow, _resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                #endregion

                var customerInvoiceHeaderRange = workSheet.Controls.AddNamedRange(workSheet.Range["B4", "B6"], "customerInvoiceHeaderRange");
                customerInvoiceHeaderRange.Change += customerInvoiceHeaderRange_Change;


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

                _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatoCustomerInvoice()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void invoiceSupplierHeaderRange_Change(Range target)
        {
            InvoiceSupplierHeaderChange();
        }

        private void journalHeaderRange_Change(Range target)
        {
            JournalHeaderRangeChange();
        }

        private void JournalHeaderRangeChange()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);


                _cells.GetRange(2, 4, 2, 5).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 5).NumberFormat = "@";
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).ClearContents();
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);

                var districtNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value).ToUpper();
                var journal = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value).ToUpper();


                if (string.IsNullOrEmpty(districtNo) || string.IsNullOrEmpty(journal)) return;

                var sqlQuery = Queries.GetJournalInfo(districtNo, journal, _eFunctions.DbReference, _eFunctions.DbLink);

                var drinfo = _eFunctions.GetQueryResult(sqlQuery);

                if (drinfo != null && !drinfo.IsClosed)
                {
                    var currentRow = _tittleRow + 1;
                    while (drinfo.Read())
                    {
                        _cells.GetCell(1, currentRow).Value = drinfo["DSTRCT_CODE"].ToString();
                        _cells.GetCell(2, currentRow).Value = drinfo["TRAN_GROUP_KEY"].ToString();
                        _cells.GetCell(3, currentRow).Value = drinfo["NUMTXT"].ToString();
                        _cells.GetCell(4, currentRow).Value = drinfo["FULL_PERIOD"].ToString();
                        _cells.GetCell(5, currentRow).Value = drinfo["REC900_TYPE"].ToString();
                        _cells.GetCell(6, currentRow).Value = drinfo["TRAN_TYPE"].ToString();
                        _cells.GetCell(7, currentRow).Value = drinfo["POSTED_STATUS"].ToString();
                        _cells.GetCell(8, currentRow).Value = drinfo["MANJNL_VCHR"].ToString();
                        _cells.GetCell(9, currentRow).Value = drinfo["ACCOUNTANT"].ToString();
                        _cells.GetCell(10, currentRow).Value = drinfo["ACCOUNT_CODE"].ToString();
                        _cells.GetCell(11, currentRow).Value = drinfo["PROJECT_NO"].ToString();
                        _cells.GetCell(12, currentRow).Value = drinfo["TRAN_AMOUNT"].ToString();
                        _cells.GetCell(13, currentRow).Value = drinfo["TRAN_AMOUNT_S"].ToString();
                        _cells.GetCell(14, currentRow).Value = drinfo["CURRENCY_TYPE"].ToString();
                        _cells.GetCell(15, currentRow).Value = drinfo["CREATION_DATE"].ToString();
                        _cells.GetCell(16, currentRow).Value = drinfo["CREATION_TIME"].ToString();
                        _cells.GetCell(17, currentRow).Value = drinfo["CREATION_USER"].ToString();
                        _cells.GetCell(18, currentRow).Value = drinfo["MIMS_SL_KEY"].ToString();
                        _cells.GetCell(19, currentRow).Value = drinfo["JOURNAL_DESC"].ToString();
                        _cells.GetCell(20, currentRow).Value = drinfo["DOCUMENT_REF"].ToString();
                        _cells.GetCell(21, currentRow).Value = drinfo["AUTO_JNL_FLG"].ToString();
                        _cells.GetCell(22, currentRow).Value = drinfo["JOURNAL_TYPE"].ToString();
                        currentRow++;
                    }

                    _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                    _cells.GetRange(12, _tittleRow + 1, 13, MaxRows).NumberFormat = "$#,###.00;[Red]-$#,###.00";
                    _cells.GetRange(2, 4, 2, 5).NumberFormat = "@";
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();
                }
                else
                {
                    _cells.GetRange(2, 4, 2, 5).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                    _cells.GetRange(2, 4, 2, 5).NumberFormat = "@";
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:JournalHeaderRangeChange()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void InvoiceSupplierHeaderChange()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).ClearContents();
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);

                var districtNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value).ToUpper();
                var supplier = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value).ToUpper();
                var invoice = _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value).ToUpper();

                if (string.IsNullOrEmpty(districtNo) || string.IsNullOrEmpty(supplier) || string.IsNullOrEmpty(invoice)) return;

                var sqlQuery = Queries.GetSupplierInvoiceInfo(districtNo, supplier, invoice, _eFunctions.DbReference, _eFunctions.DbLink);

                var drinfo = _eFunctions.GetQueryResult(sqlQuery);

                if (drinfo != null && !drinfo.IsClosed)
                {
                    var currentRow = _tittleRow + 1;
                    while (drinfo.Read())
                    {
                        _cells.GetCell(1, currentRow).Value = drinfo["DSTRCT_CODE"].ToString();
                        _cells.GetCell(2, currentRow).Value = drinfo["TRAN_GROUP_KEY"].ToString();
                        _cells.GetCell(3, currentRow).Value = drinfo["NUMTXT"].ToString();
                        _cells.GetCell(4, currentRow).Value = drinfo["FULL_PERIOD"].ToString();
                        _cells.GetCell(5, currentRow).Value = drinfo["CREATION_USER"].ToString();
                        _cells.GetCell(6, currentRow).Value = drinfo["REC900_TYPE"].ToString();
                        _cells.GetCell(7, currentRow).Value = drinfo["TRAN_TYPE"].ToString();
                        _cells.GetCell(8, currentRow).Value = drinfo["POSTED_STATUS"].ToString();
                        _cells.GetCell(9, currentRow).Value = drinfo["SUPPLIER"].ToString();
                        _cells.GetCell(10, currentRow).Value = drinfo["EXT_INV_NO"].ToString();
                        _cells.GetCell(11, currentRow).Value = drinfo["ACCOUNT_CODE"].ToString();
                        _cells.GetCell(12, currentRow).Value = drinfo["PROJECT_NO"].ToString();
                        _cells.GetCell(13, currentRow).Value = drinfo["TRAN_AMOUNT"].ToString();
                        _cells.GetCell(14, currentRow).Value = drinfo["TRAN_AMOUNT_S"].ToString();
                        _cells.GetCell(15, currentRow).Value = drinfo["CURRENCY_TYPE"].ToString();
                        _cells.GetCell(16, currentRow).Value = drinfo["CREATION_DATE"].ToString();
                        _cells.GetCell(17, currentRow).Value = drinfo["CREATION_TIME"].ToString();
                        _cells.GetCell(18, currentRow).Value = drinfo["CREATION_USER"].ToString();
                        _cells.GetCell(19, currentRow).Value = drinfo["MIMS_SL_KEY"].ToString();
                        _cells.GetCell(20, currentRow).Value = drinfo["CONTRATO"].ToString();
                        _cells.GetCell(21, currentRow).Value = drinfo["ACTA"].ToString();
                        _cells.GetCell(22, currentRow).Value = drinfo["PO_NO"].ToString();
                        _cells.GetCell(23, currentRow).Value = drinfo["PO_ITEM"].ToString();
                        _cells.GetCell(24, currentRow).Value = drinfo["PORTION_NO"].ToString();
                        _cells.GetCell(25, currentRow).Value = drinfo["INV_ITEM_NO"].ToString();
                        _cells.GetCell(26, currentRow).Value = drinfo["INV_ITEM_DESC"].ToString();
                        _cells.GetCell(27, currentRow).Value = drinfo["RECEIVED_BY"].ToString();
                        _cells.GetCell(28, currentRow).Value = drinfo["AUTH_BY"].ToString();
                        _cells.GetCell(29, currentRow).Value = drinfo["ATAX_CODE"].ToString();
                        _cells.GetCell(30, currentRow).Value = drinfo["ATAX_RATE_9"].ToString();
                        _cells.GetCell(31, currentRow).Value = drinfo["OVERRIDE_SW"].ToString();
                        _cells.GetCell(32, currentRow).Value = drinfo["BRANCH_CODE"].ToString();
                        _cells.GetCell(33, currentRow).Value = drinfo["BANK_ACCT_NO"].ToString();
                        _cells.GetCell(34, currentRow).Value = drinfo["CHEQUE_RUN_NO"].ToString();
                        _cells.GetCell(35, currentRow).Value = drinfo["CHEQUE_NO"].ToString();
                        currentRow++;
                    }
                    _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                    _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
                    _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                    _cells.GetRange(13, _tittleRow + 1, 14, MaxRows).NumberFormat = "$#,###.00;[Red]-$#,###.00";
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();
                }
                else
                {
                    _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                    _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:InvoiceSupplierHeaderChange()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCustomerInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoCustomerInvoice();
        }

        private void customerInvoiceHeaderRange_Change(Range target)
        {
            CustomerInvoiceHeaderRangeChange();
        }

        private void CustomerInvoiceHeaderRangeChange()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.ActiveWorkbook;
                Worksheet excelSheet = excelBook.ActiveSheet;


                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).ClearContents();
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                _cells.GetRange(12, _tittleRow + 1, 13, MaxRows).NumberFormat = "$#,###.00;[Red]-$#,###.00";

                _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";


                var districtNo = _cells.GetEmptyIfNull(_cells.GetCell(2, 4).Value).ToUpper();
                var customer = _cells.GetEmptyIfNull(_cells.GetCell(2, 5).Value).ToUpper();
                var invoice = _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value).ToUpper();

                if (string.IsNullOrEmpty(districtNo) || string.IsNullOrEmpty(customer) || string.IsNullOrEmpty(invoice)) return;

                var sqlQuery = Queries.GetCustomerInvoiceInfo(districtNo, customer, invoice, _eFunctions.DbReference, _eFunctions.DbLink);

                var drinfo = _eFunctions.GetQueryResult(sqlQuery);

                if (drinfo != null && !drinfo.IsClosed)
                {
                    var currentRow = _tittleRow + 1;
                    while (drinfo.Read())
                    {
                        _cells.GetCell(1, currentRow).Value = drinfo["DSTRCT_CODE"].ToString();
                        _cells.GetCell(2, currentRow).Value = drinfo["TRAN_GROUP_KEY"].ToString();
                        _cells.GetCell(3, currentRow).Value = drinfo["NUMTXT"].ToString();
                        _cells.GetCell(4, currentRow).Value = drinfo["FULL_PERIOD"].ToString();
                        _cells.GetCell(5, currentRow).Value = drinfo["REC900_TYPE"].ToString();
                        _cells.GetCell(6, currentRow).Value = drinfo["TRAN_TYPE"].ToString();
                        _cells.GetCell(7, currentRow).Value = drinfo["POSTED_STATUS"].ToString();
                        _cells.GetCell(8, currentRow).Value = drinfo["CUST_NO"].ToString();
                        _cells.GetCell(9, currentRow).Value = drinfo["AR_INV_NO"].ToString();
                        _cells.GetCell(10, currentRow).Value = drinfo["ACCOUNT_CODE"].ToString();
                        _cells.GetCell(11, currentRow).Value = drinfo["PROJECT_NO"].ToString();
                        _cells.GetCell(12, currentRow).Value = drinfo["TRAN_AMOUNT"].ToString();
                        _cells.GetCell(13, currentRow).Value = drinfo["TRAN_AMOUNT_S"].ToString();
                        _cells.GetCell(14, currentRow).Value = drinfo["CURRENCY_TYPE"].ToString();
                        _cells.GetCell(15, currentRow).Value = drinfo["CREATION_DATE"].ToString();
                        _cells.GetCell(16, currentRow).Value = drinfo["CREATION_TIME"].ToString();
                        _cells.GetCell(17, currentRow).Value = drinfo["CREATION_USER"].ToString();
                        _cells.GetCell(18, currentRow).Value = drinfo["MIMS_SL_KEY"].ToString();
                        _cells.GetCell(19, currentRow).Value = drinfo["INV_DATE"].ToString();
                        _cells.GetCell(20, currentRow).Value = drinfo["ITEM_NO"].ToString();
                        _cells.GetCell(21, currentRow).Value = drinfo["REVENUE_CODE"].ToString();
                        _cells.GetCell(22, currentRow).Value = drinfo["CUST_REF"].ToString();
                        _cells.GetCell(23, currentRow).Value = drinfo["SALES_PERSON_ID"].ToString();
                        currentRow++;
                    }
                    excelSheet.Cells.Columns.AutoFit();
                    excelSheet.Cells.Rows.AutoFit();
                }
                else
                {
                    _cells.GetRange(2, 4, 2, 6).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";
                    _cells.GetRange(2, 4, 2, 6).NumberFormat = "@";
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CustomerInvoiceHeaderRangeChange()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
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

    internal static class Queries
    {
        public static string GetSupplierInvoiceInfo(string districtCode, string supplierNo, string invoiceNo, string dbReference, string dbLink)
        {
            var sqlQuery = " " +
                "SELECT " +
                "  TT.DSTRCT_CODE, " +
                "  TT.TRAN_GROUP_KEY, " +
                "  ( TT.PROCESS_DATE || TT.TRANSACTION_NO || TT.USERNO || TT.REC900_TYPE ) NUMTXT, " +
                "  TT.REC900_TYPE, " +
                "  TT.TRAN_TYPE, " +
                "  TT.POSTED_STATUS, " +
                "  DECODE ( TT.REC900_TYPE, 'C', MAX ( TT.PMT_SUPPLIER ) OVER ( PARTITION BY TRAN_GROUP_KEY ), TT.SUPPLIER_NO ) SUPPLIER, " +
                "  DECODE ( TT.REC900_TYPE, 'C', MAX ( TT.EXT_INV_NO ) OVER ( PARTITION BY TRAN_GROUP_KEY ), TT.EXT_INV_NO ) EXT_INV_NO, " +
                "  TT.ACCOUNT_CODE, " +
                "  TT.PROJECT_NO, " +
                "  TT.FULL_PERIOD, " +
                "  TT.TRAN_AMOUNT, " +
                "  TT.TRAN_AMOUNT_S, " +
                "  TT.CURRENCY_TYPE, " +
                "  TT.CREATION_DATE, " +
                "  TT.CREATION_TIME, " +
                "  TT.CREATION_USER, " +
                "  TT.MIMS_SL_KEY, " +
                "  DECODE ( TRIM ( TT.CONTRACT_NO ), NULL, DECODE ( SUBSTR ( TT.INV_ITEM_DESC, 1, 3 ), 'CNT', SUBSTR ( TT.INV_ITEM_DESC, 5, INSTR ( TT.INV_ITEM_DESC, ' ' ) - 5 ), NULL ), TT.CONTRACT_NO ) CONTRATO, " +
                "  DECODE ( SUBSTR ( TT.INV_ITEM_DESC, 1, 3 ), 'CNT', SUBSTR ( TT.INV_ITEM_DESC, INSTR ( TT.INV_ITEM_DESC, ' ' ) + 1, INSTR ( TT.INV_ITEM_DESC, ' ', 2 ) - 2 ), NULL ) ACTA, " +
                "  TT.PO_NO, " +
                "  TT.PO_ITEM, " +
                "  TT.PORTION_NO, " +
                "  TT.INV_ITEM_NO, " +
                "  TT.INV_ITEM_DESC, " +
                "  TT.RECEIVED_BY, " +
                "  TT.AUTH_BY, " +
                "  TT.ATAX_CODE, " +
                "  TT.ATAX_RATE_9, " +
                "  TT.OVERRIDE_SW, " +
                "  TT.BRANCH_CODE, " +
                "  TT.BANK_ACCT_NO, " +
                "  TT.CHEQUE_RUN_NO, " +
                "  TT.CHEQUE_NO " +
                "FROM " +
                "  " + dbReference + ".MSF900" + dbLink + " TT " +
                "INNER JOIN " + dbReference + ".MSFX9B" + dbLink + " X9B " +
                "ON " +
                "  X9B.DSTRCT_CODE       = TT.DSTRCT_CODE " +
                "  AND X9B.PROCESS_DATE   = TT.PROCESS_DATE " +
                "  AND X9B.TRANSACTION_NO = TT.TRANSACTION_NO " +
                "  AND X9B.USERNO         = TT.USERNO " +
                "  AND X9B.REC900_TYPE    = TT.REC900_TYPE " +
                "INNER JOIN " + dbReference + ".MSF260" + dbLink + " INV " +
                "ON " +
                "  INV.INV_NO = X9B.INV_NO " +
                "  AND INV.SUPPLIER_NO = X9B.SUPPLIER_NO " +
                "WHERE " +
                "  X9B.DSTRCT_CODE = '" + districtCode + "' " +
                "AND X9B.SUPPLIER_NO = '" + supplierNo + "' " +
                "AND INV.EXT_INV_NO = '" + invoiceNo + "' " +
                "ORDER BY " +
                "  1, " +
                "  2, " +
                "  3";

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
                "  X90.DSTRCT_CODE       = TR.DSTRCT_CODE " +
                "  AND X90.PROCESS_DATE   = TR.PROCESS_DATE " +
                "  AND X90.TRANSACTION_NO = TR.TRANSACTION_NO " +
                "  AND X90.USERNO         = TR.USERNO " +
                "  AND X90.REC900_TYPE    = TR.REC900_TYPE " +
                "WHERE " +
                "  X90.DSTRCT_CODE = '" + districtCode + "' " +
                "AND X90.JOURNAL_NO = '" + journal + "' " +
                "ORDER BY " +
                "  1, " +
                "  2, " +
                "  3 ";

            return sqlQuery;
        }

        public static string GetCustomerInvoiceInfo(string districtCode, string customer, string invoice, string dbReference, string dbLink)
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
                "  TR.CUST_NO, " +
                "  DECODE ( TR.REC900_TYPE, 'Z', TRIM ( SUBSTR ( TR.CONSOLIDATE_Z, 2, 10 ) ), TR.AR_INV_NO ) AR_INV_NO, " +
                "  TR.ACCOUNT_CODE, " +
                "  TR.PROJECT_NO, " +
                "  TR.TRAN_AMOUNT, " +
                "  TR.TRAN_AMOUNT_S, " +
                "  TR.CURRENCY_TYPE, " +
                "  TR.CREATION_DATE, " +
                "  TR.CREATION_TIME, " +
                "  TR.CREATION_USER, " +
                "  TR.MIMS_SL_KEY, " +
                "  TR.INV_DATE, " +
                "  TR.ITEM_NO, " +
                "  TR.REVENUE_CODE, " +
                "  TR.CUST_REF, " +
                "  TR.SALES_PERSON_ID " +
                "FROM " +
                "  " + dbReference + ".MSF900" + dbLink + " TR " +
                "INNER JOIN " + dbReference + ".MSFX93" + dbLink + " X93 " +
                "ON " +
                "  X93.DSTRCT_CODE       = TR.DSTRCT_CODE " +
                "  AND X93.PROCESS_DATE   = TR.PROCESS_DATE " +
                "  AND X93.TRANSACTION_NO = TR.TRANSACTION_NO " +
                "  AND X93.USERNO         = TR.USERNO " +
                "  AND X93.REC900_TYPE    = TR.REC900_TYPE " +
                "WHERE " +
                "  X93.CUST_NO = '" + customer + "' " +
                "AND X93.AR_INV_NO = '" + invoice + "' " +
                "AND X93.DSTRCT_CODE = '" + districtCode + "' " +
                "ORDER BY " +
                "  1, " +
                "  2, " +
                "  3 ";

            return sqlQuery;
        }
    }
}

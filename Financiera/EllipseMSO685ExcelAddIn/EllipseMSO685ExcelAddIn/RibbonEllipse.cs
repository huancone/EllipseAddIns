using System;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseMSO685ExcelAddIn
{
    public partial class RibbonEllipse
    {

        private const int TittleRow = 5;
        private static int _resultColumn = 13;
        public static EllipseFunctions EFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private string _sheetName01;
        
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormatSubAssetsDep_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSubAssetsDep();
        }

        private void FormatSubAssetsDep()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;
            _sheetName01 = "MSO685 Opcion 3";
            excelSheet.Name = _sheetName01;
            _resultColumn = 19;

            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                _cells.GetRange(1, TittleRow, _resultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "MSO685 Opcion 3 Maintain Sub-Asset Depreciation Details";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";

            _cells.GetCell(1, TittleRow).Value = "Asset Reference *";
            _cells.GetCell(2, TittleRow).Value = "Sub Asset Number *";
            _cells.GetCell(3, TittleRow).Value = "Book Type *";
            _cells.GetCell(4, TittleRow).Value = "Depreciation Method *";
            _cells.GetCell(5, TittleRow).Value = "Depreciation Rate";
            _cells.GetCell(6, TittleRow).Value = "Manual Period Depn";
            _cells.GetCell(7, TittleRow).Value = "Until Period";
            _cells.GetCell(8, TittleRow).Value = "Accelerated Depn Rate";
            _cells.GetCell(9, TittleRow).Value = "Until Period";
            _cells.GetCell(10, TittleRow).Value = "Rate Table";
            _cells.GetCell(11, TittleRow).Value = "Recovery Period";
            _cells.GetCell(12, TittleRow).Value = "Dividend Statistic";
            _cells.GetCell(13, TittleRow).Value = "Divisor Statistic";
            _cells.GetCell(14, TittleRow).Value = "Estimated Life (months)";
            _cells.GetCell(15, TittleRow).Value = "Useful Life Group Code";
            _cells.GetCell(16, TittleRow).Value = "Est Retirement Value - Local";
            _cells.GetCell(17, TittleRow).Value = "Foreign Currency Cost";
            _cells.GetCell(18, TittleRow).Value = "Foreign Currency Type";
            _cells.GetCell(_resultColumn, TittleRow).Value = "Message";

            _cells.GetRange(1, TittleRow, _resultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnAccion3_Click(object sender, RibbonControlEventArgs e)
        {
            var excelBook = _excelApp.ActiveWorkbook;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            Worksheet excelSheet = excelBook.ActiveSheet;
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            Accion3();
        }

        private void Accion3()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();
            proxySheet.Url = EFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    EFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO685");
                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {


                        if (replySheet.mapName != "MSM685A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "3");
                        arrayFields.Add("ASSET_REF1I", _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
                        arrayFields.Add("SUB_ASSET_NO1I",
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                        arrayFields.Add("BOOK_OR_TAX1I",
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value));
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (EFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                        if (EFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            if (replySheet.mapName != "MSM685C") return;
                            arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("DEPR_METHOD3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value));
                            arrayFields.Add("DEPR_RATE3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value));
                            arrayFields.Add("MAN_PER_DEPR3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                            arrayFields.Add("FIN_MAN_PER3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value));
                            arrayFields.Add("ACCEL_DEPR_RT3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                            arrayFields.Add("FIN_ACCEL_PER3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));
                            arrayFields.Add("RATE_TABLE3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value));
                            arrayFields.Add("RECOV_PERIOD3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value));
                            arrayFields.Add("DIVIDEND_STAT3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value));
                            arrayFields.Add("DIVISOR_STAT3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value));
                            arrayFields.Add("EST_MM_LIFE3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value));
                            arrayFields.Add("LIFE_GRP_CODE3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value));
                            arrayFields.Add("EST_DISPOS_VAL3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value));
                            arrayFields.Add("FOR_CURR_AMT3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value));
                            arrayFields.Add("FOREIGN_CURR3I",
                                _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value));

                            requestSheet.screenFields = arrayFields.ToArray();
                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);
                            while (EFunctions.CheckReplyWarning(replySheet))
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (replySheet.message.Contains("Confirm"))
                                replySheet = proxySheet.submit(opSheet, requestSheet);
                            if (EFunctions.CheckReplyError(replySheet))
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Success;
                                _cells.GetCell(_resultColumn, currentRow).Value = "Procesado";
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
    }
}

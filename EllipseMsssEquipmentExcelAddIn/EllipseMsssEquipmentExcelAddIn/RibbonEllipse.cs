using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseMsssEquipmentExcelAddIn.Properties;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Settings = SharedClassLibrary.Ellipse.Settings;

namespace EllipseMsssEquipmentExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const int TittleRow = 6;
        private const int ResultColumn = 23;
        private const int MaxRows = 10000;
        private static readonly string SheetName01 = Resources.RibbonEllipse_SheetName01;

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

            //settings.SetDefaultCustomSettingValue("FlagEstDuration", "Y");
            //settings.SetDefaultCustomSettingValue("ValidateTaskPlanDates", "Y");
            //settings.SetDefaultCustomSettingValue("IgnoreClosedStatus", "N");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //var flagEstDur = MyUtilities.IsTrue(settings.GetCustomSettingValue("FlagEstDuration"));
            //var valdTaskPlanDates = MyUtilities.IsTrue(settings.GetCustomSettingValue("ValidateTaskPlanDates"));
            //var ignoreCldStat = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreClosedStatus"));
            //
            //cbFlagEstDuration.Checked = flagEstDur;
            //cbValidateTaskPlanDates.Checked = valdTaskPlanDates;
            //cbIgnoreClosedStatus.Checked = ignoreCldStat;
            //
            settings.SaveCustomSettings();
        }
        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
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

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("A3").Value = "MSSS SERVICE";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A3", "G3");

                _cells.GetCell("A4").Value = "equipmentGrpId";


                _cells.GetCell(1, TittleRow).Value = "Action";
                _cells.GetCell(2, TittleRow).Value = "equipmentGrpId";
                _cells.GetCell(3, TittleRow).Value = "equipmentGrpIdDescription";
                _cells.GetCell(4, TittleRow).Value = "compCode";
                _cells.GetCell(5, TittleRow).Value = "compcodeDescription";
                _cells.GetCell(6, TittleRow).Value = "compModCode";
                _cells.GetCell(7, TittleRow).Value = "failureMode";
                _cells.GetCell(8, TittleRow).Value = "failureModeDescription";
                _cells.GetCell(9, TittleRow).Value = "failureCode";
                _cells.GetCell(10, TittleRow).Value = "failureCodeDescription";
                _cells.GetCell(11, TittleRow).Value = "functionCode";
                _cells.GetCell(12, TittleRow).Value = "functionCodeDescription";
                _cells.GetCell(13, TittleRow).Value = "consequence";
                _cells.GetCell(14, TittleRow).Value = "consequenceDescription";
                _cells.GetCell(15, TittleRow).Value = "effect";
                _cells.GetCell(16, TittleRow).Value = "strategy";
                _cells.GetCell(17, TittleRow).Value = "strategyDescription";
                _cells.GetCell(18, TittleRow).Value = "agreedAction";
                _cells.GetCell(19, TittleRow).Value = "failureClass";
                _cells.GetCell(20, TittleRow).Value = "failureClassDescription";
                _cells.GetCell(21, TittleRow).Value = "functionClass";
                _cells.GetCell(22, TittleRow).Value = "functionClassDescription";
                _cells.GetCell(23, TittleRow).Value = "Result";

                _cells.GetCell(1, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(4, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(6, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(9, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(11, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(12, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(13, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(14, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(15, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(16, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(17, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(18, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(19, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(20, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(21, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(22, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(23, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                Microsoft.Office.Tools.Excel.Worksheet workSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var equipmentGrpId = workSheet.Controls.AddNamedRange(workSheet.Range["B4"], "equipmentGrpId");
                equipmentGrpId.Change += equipmentGrpIdRange_Change;

                var optionList = new List<string>
                {
                    "Create", 
                    "Delete", 
                    "Modify"
                };
                _cells.SetValidationList(_cells.GetRange(1, TittleRow + 1, 1, 200000), optionList);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();

                _cells.SetCursorDefault();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, Resources.RibbonEllipse_FormatSheet_Error);
            }
        }

        private void equipmentGrpIdRange_Change(Excel.Range target)
        {
            EquipmentGrpIdRangeChange(target);
        }

        private void EquipmentGrpIdRangeChange(Excel.Range target)
        {
            _excelApp = Globals.ThisAddIn.Application;
            var excelBook = _excelApp.ActiveWorkbook;
            Excel.Worksheet excelSheet = excelBook.ActiveSheet;

            var equipmentGrpId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(target.Column, target.Row).Value).ToUpper();

            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Clear();

            if (string.IsNullOrEmpty(equipmentGrpId)) return;
            var sqlQuery = Queries.GetMsssInfo(equipmentGrpId, _eFunctions.DbReference, _eFunctions.DbLink);
            var drMsss = _eFunctions.GetQueryResult(sqlQuery);

            if (drMsss == null || drMsss.IsClosed) return;

            var currentRow = TittleRow + 1;
            while (drMsss.Read())
            {
                _cells.GetCell(2, currentRow).Value = drMsss["EQUIP_GRP_ID"].ToString();
                _cells.GetCell(4, currentRow).Value = drMsss["COMP_CODE"].ToString();
                _cells.GetCell(5, currentRow).Value = drMsss["COMP_CODE_DESC"].ToString();
                _cells.GetCell(6, currentRow).Value = drMsss["COMP_MOD_CODE"].ToString();
                _cells.GetCell(7, currentRow).Value = drMsss["FAILURE_MODE"].ToString();
                _cells.GetCell(8, currentRow).Value = drMsss["FAILURE_MODE_DESC"].ToString();
                _cells.GetCell(9, currentRow).Value = drMsss["FAILURE_CODE"].ToString();
                _cells.GetCell(10, currentRow).Value = drMsss["FAILURE_CODE_DESC"].ToString();
                _cells.GetCell(11, currentRow).Value = drMsss["FUNCTION_CODE"].ToString();
                _cells.GetCell(12, currentRow).Value = drMsss["FUNCTION_CODE_DESC"].ToString();
                _cells.GetCell(13, currentRow).Value = drMsss["CONSEQUENCE"].ToString();
                _cells.GetCell(14, currentRow).Value = drMsss["CONSEQUENCE_DESC"].ToString();
                _cells.GetCell(15, currentRow).Value = drMsss["EFFECT"].ToString();
                _cells.GetCell(16, currentRow).Value = drMsss["STRATEGY"].ToString();
                _cells.GetCell(17, currentRow).Value = drMsss["STRATEGY_DESC"].ToString();
                _cells.GetCell(18, currentRow).Value = drMsss["AGREED_ACTION"].ToString();
                _cells.GetCell(19, currentRow).Value = drMsss["FAILURE_CLASS"].ToString();
                _cells.GetCell(20, currentRow).Value = drMsss["FAILURE_CLASS_DESC"].ToString();
                _cells.GetCell(21, currentRow).Value = drMsss["FUNCTION_CLASS"].ToString();
                _cells.GetCell(22, currentRow).Value = drMsss["FUNCTION_CLASS_DESC"].ToString();
                currentRow++;
            }
            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                LoadSheet();
            }
            else
                MessageBox.Show(Resources.Loadsheet_error);
        }

        private void LoadSheet()
        {
            var msssService = new MSSSService.MSSSService();
            var msssOpContext = new MSSSService.OperationContext();

            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            msssService.Url = urlService + "/MSSSService";
            msssOpContext.district = _frmAuth.EllipseDsct;
            msssOpContext.position = _frmAuth.EllipsePost;
            msssOpContext.maxInstances = 100;
            msssOpContext.returnWarnings = Debugger.DebugWarnings;


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var currentRow = TittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    string action = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    var msssItem = new MssItemDto
                    {
                        EquipmentGrpId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        CompCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        CompCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        CompModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        FailureMode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        FailureModeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        FailureCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        FailureCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        FunctionCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        FunctionCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value),
                        Consequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                        ConsequenceDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        Effect = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                        Strategy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                        StrategyDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value),
                        AgreedAction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value),
                        FailureClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value),
                        FailureClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, currentRow).Value),
                        FunctionClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, currentRow).Value),
                        FunctionClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, currentRow).Value)
                    };


                    switch (action)
                    {
                        case "Create":
                            msssService.create(msssOpContext, msssItem.ToCreateRequestDto());
                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetCell(ResultColumn, currentRow).Value = "Creado";
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                                _cells.GetStyle(StyleConstants.Success);

                            break;

                        case "Delete":
                            msssService.delete(msssOpContext, msssItem.ToDeleteRequestDto());

                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetCell(ResultColumn, currentRow).Value = "Borrado";
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                                _cells.GetStyle(StyleConstants.Success);

                            break;
                    }
                }
                catch (Exception error)
                {
                    _cells.GetCell(ResultColumn, currentRow).Select();
                    _cells.GetCell(ResultColumn, currentRow).Value = error.Message;
                    _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                        _cells.GetStyle(StyleConstants.Error);
                }
                finally
                {
                    _cells.GetCell(ResultColumn, currentRow).Select();
                    currentRow++;
                }
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    internal static class Queries
    {
        public static string GetMsssInfo(string equipmentGrpId, string dbReference, string dbLink)
        {
            var query = "" +
                           "SELECT " +
                           "MSSS.EQUIP_GRP_ID, " +
                           "MSSS.COMP_CODE, " +
                           "CO.TABLE_DESC COMP_CODE_DESC, " +
                           "MSSS.COMP_MOD_CODE, " +
                           "MO.TABLE_DESC COMP_MOD_CODE_DESC, " +
                           "MSSS.FAILURE_MODE, " +
                           "W0.TABLE_DESC FAILURE_MODE_DESC, " +
                           "MSSS.FAILURE_CODE, " +
                           "W1.TABLE_DESC FAILURE_CODE_DESC, " +
                           "MSSS.FUNCTION_CODE, " +
                           "W2.TABLE_DESC FUNCTION_CODE_DESC, " +
                           "MSSS.CONSEQUENCE, " +
                           "W3.TABLE_DESC CONSEQUENCE_DESC, " +
                           "MSSS.EFFECT, " +
                           "MSSS.STRATEGY, " +
                           "STRT.TABLE_DESC STRATEGY_DESC, " +
                           "MSSS.AGREED_ACTION, " +
                           "MSSS.FAILURE_CLASS, " +
                           "FLCL.TABLE_DESC FAILURE_CLASS_DESC, " +
                           "MSSS.FUNCTION_CLASS, " +
                           "FNCL.TABLE_DESC FUNCTION_CLASS_DESC " +
                           "FROM " +
                           "  " + dbReference + ".MSF6A1" + dbLink + " MSSS " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " CO " +
                           "ON " +
                           "  CO.TABLE_CODE = MSSS.COMP_CODE " +
                           "AND CO.TABLE_TYPE = 'CO' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " MO " +
                           "ON " +
                           "  MO.TABLE_CODE = MSSS.COMP_MOD_CODE " +
                           "AND MO.TABLE_TYPE = 'MO' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " W0 " +
                           "ON " +
                           "  W0.TABLE_CODE = MSSS.FAILURE_MODE " +
                           "AND W0.TABLE_TYPE = 'W0' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " W1 " +
                           "ON " +
                           "  W1.TABLE_CODE = MSSS.FAILURE_CODE " +
                           "AND W1.TABLE_TYPE = 'W1' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " W2 " +
                           "ON " +
                           "  W2.TABLE_CODE = MSSS.FUNCTION_CODE " +
                           "AND W2.TABLE_TYPE = 'W2' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " W3 " +
                           "ON " +
                           "  W3.TABLE_CODE = MSSS.CONSEQUENCE " +
                           "AND W3.TABLE_TYPE = 'W3' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " STRT " +
                           "ON " +
                           "  STRT.TABLE_CODE = MSSS.STRATEGY " +
                           "AND STRT.TABLE_TYPE = 'STRT' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " FLCL " +
                           "ON " +
                           "  FLCL.TABLE_CODE = MSSS.FAILURE_CLASS " +
                           "AND FLCL.TABLE_TYPE = 'FLCL' " +
                           "LEFT JOIN " + dbReference + ".MSF010" + dbLink + " FNCL " +
                           "ON " +
                           "  FNCL.TABLE_CODE = MSSS.FUNCTION_CLASS " +
                           "AND FNCL.TABLE_TYPE = 'FNCL' " +
                           "WHERE " +
                           "  MSSS.EQUIP_GRP_ID = '" + equipmentGrpId + "'";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
    }
}

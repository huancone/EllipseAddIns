using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Services.Ellipse;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseMsssEquipmentExcelAddIn.Properties;

namespace EllipseMsssEquipmentExcelAddIn
{
    public partial class RibbonEllipse
    {
        private static readonly string SheetName01 = Resources.RibbonEllipse_SheetName01;
        private const int TittleRow = 6;
        private const int ResultColumn = 23;
        private const int MaxRows = 10000;
        private static readonly EllipseFunctions EFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Excel.Application _excelApp;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var environmentList = Environments.GetEnvironmentList();
            foreach (var item in environmentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpEnvironment.Items.Add(drpItem);
            }

            drpEnvironment.SelectedItem.Label = Resources.RibbonEllipse_RibbonEllipse_Load_DefaultEnvironment;
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
                var excelBook = _excelApp.Workbooks.Add();
                Excel.Worksheet excelSheet = excelBook.ActiveSheet;

                Microsoft.Office.Tools.Excel.Worksheet workSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);


                excelSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "MSSS SERVICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "G2");

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

                var equipmentGrpId = workSheet.Controls.AddNamedRange(workSheet.Range["B4"], "equipmentGrpId");
                equipmentGrpId.Change += equipmentGrpIdRange_Change;

                var optionList = new List<string>
                {
                    "Create", 
                    "Delete", 
                    "Modify"
                };
                _cells.SetValidationList(_cells.GetRange(1, TittleRow + 1, 1, 200000), optionList);

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
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
            var sqlQuery = Queries.GetMsssInfo(equipmentGrpId, EFunctions.dbReference, EFunctions.dbLink);
            var drMsss = EFunctions.GetQueryResult(sqlQuery);

            if (drMsss == null || drMsss.IsClosed || !drMsss.HasRows) return;

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
            var msssProxy = new MSSSService.MSSSService();
            var msssOp = new MSSSService.OperationContext();

            msssProxy.Url = EFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/MSSSService";
            msssOp.district = _frmAuth.EllipseDsct;
            msssOp.position = _frmAuth.EllipsePost;
            msssOp.maxInstances = 100;
            msssOp.returnWarnings = Debugger.DebugWarnings;


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var currentRow = TittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                string action = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                switch (action)
                {
                    case "Create":
                        try
                        {
                            var msssItem = new MSSSService.MSSSServiceCreateRequestDTO
                                            {
                                                equipmentGrpId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                                                compCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                                                compModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                                                failureMode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                                                failureCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                                                functionCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                                                consequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                                                effect = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                                                strategy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                                                agreedAction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value),
                                                failureClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value),
                                                functionClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, currentRow).Value)
                                            };
                            msssProxy.create(msssOp, msssItem);
                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetCell(ResultColumn, currentRow).Value = "Creado";
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                        catch (Exception error)
                        {
                            _cells.GetCell(ResultColumn, currentRow).Value = error.Message;
                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        }
                        break;

                    case "Delete":
                        try
                        {
                            var urlEnvironment = EFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                            EFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnvironment);
                            var responseDto = EFunctions.InitiatePostConnection();

                            if (responseDto.GotErrorMessages()) return;
                            var equipmentGrpId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
                            var compCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                            var compcodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                            var compModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                            var failureMode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                            var failureModeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                            var failureCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value);
                            var failureCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value);
                            var functionCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value);
                            var functionCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value);
                            var consequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value);
                            var consequenceDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                            var effect = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);
                            var strategy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
                            var strategyDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
                            var agreedAction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value);
                            var failureClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value);
                            var failureClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, currentRow).Value);
                            var functionClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, currentRow).Value);
                            var functionClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, currentRow).Value);


                            var requestXml = "";
                            requestXml = requestXml + "<interaction>";
                            requestXml = requestXml + "	<actions>";
                            requestXml = requestXml + "		<action>";
                            requestXml = requestXml + "			<name>service</name>";
                            requestXml = requestXml + "			<data>";
                            requestXml = requestXml + "				<name>com.mincom.enterpriseservice.ellipse.msss.MSSSService</name>";
                            requestXml = requestXml + "				<operation>delete</operation>";
                            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
                            requestXml = requestXml + "				<dto uuid=\"" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
                            requestXml = requestXml + "					<consequenceDescription>" + consequenceDescription + "</consequenceDescription>";
                            requestXml = requestXml + "					<functionClass>" + functionClass + "</functionClass>";
                            requestXml = requestXml + "					<compcodeDescription>" + compcodeDescription + "</compcodeDescription>";
                            requestXml = requestXml + "					<strategy>" + strategy + "</strategy>";
                            requestXml = requestXml + "					<equipmentGrpId>" + equipmentGrpId + "</equipmentGrpId>";
                            requestXml = requestXml + "					<failureClass>" + failureClass + "</failureClass>";
                            requestXml = requestXml + "					<failureCodeDescription>" + failureCodeDescription + "</failureCodeDescription>";
                            requestXml = requestXml + "					<strategyDescription>" + strategyDescription + "</strategyDescription>";
                            requestXml = requestXml + "					<effect>" + effect + "</effect>";
                            requestXml = requestXml + "					<functionCodeDescription>" + functionCodeDescription + "</functionCodeDescription>";
                            requestXml = requestXml + "					<functionCode>" + functionCode + "</functionCode>";
                            requestXml = requestXml + "					<equipmentGrpIdDescription>BANDA TRANSPORTADORA BC402</equipmentGrpIdDescription>";
                            requestXml = requestXml + "					<failureCode>" + failureCode + "</failureCode>";
                            requestXml = requestXml + "					<failureMode>" + failureMode + "</failureMode>";
                            requestXml = requestXml + "					<consequence>" + consequence + "</consequence>";
                            requestXml = requestXml + "					<failureClassDescription>" + failureClassDescription + "</failureClassDescription>";
                            requestXml = requestXml + "					<failureModeDescription>" + failureModeDescription + "</failureModeDescription>";
                            requestXml = requestXml + "					<functionClassDescription>" + functionClassDescription + "</functionClassDescription>";
                            requestXml = requestXml + "					<agreedAction>" + agreedAction + "</agreedAction>";
                            requestXml = requestXml + "					<compCode>" + compCode + "</compCode>";
                            requestXml = requestXml + "					<compModCode>" + compModCode + "</compModCode>";
                            requestXml = requestXml + "				</dto>";
                            requestXml = requestXml + "			</data>";
                            requestXml = requestXml + "			<id>" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "</id>";
                            requestXml = requestXml + "		</action>";
                            requestXml = requestXml + "	</actions>";
                            requestXml = requestXml + "	<chains/>";
                            requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                            requestXml = requestXml + "	<application>mse6a1</application>";
                            requestXml = requestXml + "	<applicationPage>read</applicationPage>";
                            requestXml = requestXml + "	<transaction>true</transaction>";
                            requestXml = requestXml + "</interaction>";

                            responseDto = EFunctions.ExecutePostRequest(requestXml);

                            var errorMessage = responseDto.Errors.Aggregate("",
                                (current, msg) => current + (msg.Field + " " + msg.Text));
                            if (errorMessage.Equals(""))
                            {
                                _cells.GetCell(ResultColumn, currentRow).Select();
                                _cells.GetCell(ResultColumn, currentRow).Value = "Borrado";
                                _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                                    _cells.GetStyle(StyleConstants.Success);
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn, currentRow).Select();
                                _cells.GetCell(ResultColumn, currentRow).Value = errorMessage;
                                _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                                    _cells.GetStyle(StyleConstants.Error);
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
                        }
                        break;
                }
                currentRow++;
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
        //        private void LoadSheet()
        //        {
        //            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
        //            {
        //                var urlEnvironment = EFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
        //                EFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnvironment);
        //                var responseDto = EFunctions.InitiatePostConnection();
        //
        //                if (responseDto.GotErrorMessages()) return;
        //
        //                var currentRow = TittleRow + 1;
        //
        //                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
        //                {
        //                    try
        //                    {
        //
        //                        var equipmentGrpId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
        //                        var compCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
        //                        var compcodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
        //                        var compModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
        //                        var failureMode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
        //                        var failureModeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
        //                        var failureCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value);
        //                        var failureCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value);
        //                        var functionCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value);
        //                        var functionCodeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value);
        //                        var consequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value);
        //                        var consequenceDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
        //                        var effect = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);
        //                        var strategy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
        //                        var strategyDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
        //                        var agreedAction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value);
        //                        var failureClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value);
        //                        var failureClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, currentRow).Value);
        //                        var functionClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, currentRow).Value);
        //                        var functionClassDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, currentRow).Value);
        //
        //
        //                        var requestXml = "";
        //                        requestXml = requestXml + "<interaction>";
        //                        requestXml = requestXml + "	<actions>";
        //                        requestXml = requestXml + "		<action>";
        //                        requestXml = requestXml + "			<name>service</name>";
        //                        requestXml = requestXml + "			<data>";
        //                        requestXml = requestXml + "				<name>com.mincom.enterpriseservice.ellipse.msss.MSSSService</name>";
        //                        requestXml = requestXml + "				<operation>delete</operation>";
        //                        requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
        //                        requestXml = requestXml + "				<dto uuid=\"" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
        //                        requestXml = requestXml + "					<consequenceDescription>" + consequenceDescription + "</consequenceDescription>";
        //                        requestXml = requestXml + "					<functionClass>" + functionClass + "</functionClass>";
        //                        requestXml = requestXml + "					<compcodeDescription>" + compcodeDescription + "</compcodeDescription>";
        //                        requestXml = requestXml + "					<strategy>" + strategy + "</strategy>";
        //                        requestXml = requestXml + "					<equipmentGrpId>" + equipmentGrpId + "</equipmentGrpId>";
        //                        requestXml = requestXml + "					<failureClass>" + failureClass + "</failureClass>";
        //                        requestXml = requestXml + "					<failureCodeDescription>" + failureCodeDescription + "</failureCodeDescription>";
        //                        requestXml = requestXml + "					<strategyDescription>" + strategyDescription + "</strategyDescription>";
        //                        requestXml = requestXml + "					<effect>" + effect + "</effect>";
        //                        requestXml = requestXml + "					<functionCodeDescription>" + functionCodeDescription + "</functionCodeDescription>";
        //                        requestXml = requestXml + "					<functionCode>" + functionCode + "</functionCode>";
        //                        requestXml = requestXml + "					<equipmentGrpIdDescription>BANDA TRANSPORTADORA BC402</equipmentGrpIdDescription>";
        //                        requestXml = requestXml + "					<failureCode>" + failureCode + "</failureCode>";
        //                        requestXml = requestXml + "					<failureMode>" + failureMode + "</failureMode>";
        //                        requestXml = requestXml + "					<consequence>" + consequence + "</consequence>";
        //                        requestXml = requestXml + "					<failureClassDescription>" + failureClassDescription + "</failureClassDescription>";
        //                        requestXml = requestXml + "					<failureModeDescription>" + failureModeDescription + "</failureModeDescription>";
        //                        requestXml = requestXml + "					<functionClassDescription>" + functionClassDescription + "</functionClassDescription>";
        //                        requestXml = requestXml + "					<agreedAction>" + agreedAction + "</agreedAction>";
        //                        requestXml = requestXml + "					<compCode>" + compCode + "</compCode>";
        //                        requestXml = requestXml + "					<compModCode>" + compModCode + "</compModCode>";
        //                        requestXml = requestXml + "				</dto>";
        //                        requestXml = requestXml + "			</data>";
        //                        requestXml = requestXml + "			<id>" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "</id>";
        //                        requestXml = requestXml + "		</action>";
        //                        requestXml = requestXml + "	</actions>";
        //                        requestXml = requestXml + "	<chains/>";
        //                        requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
        //                        requestXml = requestXml + "	<application>mse6a1</application>";
        //                        requestXml = requestXml + "	<applicationPage>read</applicationPage>";
        //                        requestXml = requestXml + "	<transaction>true</transaction>";
        //                        requestXml = requestXml + "</interaction>";
        //                        responseDto = EFunctions.ExecutePostRequest(requestXml);
        //
        //                        var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
        //                        if (errorMessage.Equals(""))
        //                        {
        //                            _cells.GetCell(ResultColumn, currentRow).Value = errorMessage;
        //                            _cells.GetCell(ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
        //                        }
        //                        else
        //                        {
        //                            _cells.GetCell(ResultColumn, currentRow).Value = errorMessage;
        //                            _cells.GetCell(ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
        //                        }
        //                    }
        //                    catch (Exception error)
        //                    {
        //                        _cells.GetCell(ResultColumn, currentRow).Value = error.Message;
        //                        _cells.GetCell(ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
        //                    }
        //                    finally
        //                    {
        //                        currentRow++;
        //                    }
        //                }
        //
        //
        //            }
        //            else
        //                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        //        }
    }

    public static class Queries
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

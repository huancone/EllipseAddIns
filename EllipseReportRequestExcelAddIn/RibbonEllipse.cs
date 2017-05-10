using System;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseReportRequestExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        EllipseFunctions _eFunctions = new EllipseFunctions();
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Application _excelApp;

        private const string SheetName01 = "ReportRequest";
        private const int TitleRow01 = 4;
        private const int ResultColumn01 = 19;
        private const string TableName01 = "ReportRequestTable";
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
            FormatSheetHeaderData();
        }

        private void btnExecuteRequest_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
            {
                LoadRequest();
            }
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para la solicitud");
        }


        public void FormatSheetHeaderData()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "REQUEST MSO080 - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, TitleRow01, 7, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, TitleRow01).Value = "PROGRAMA";
                _cells.GetCell(2, TitleRow01).Value = "EJECUTAR (Y/N)";
                _cells.GetCell(3, TitleRow01).Value = "IMPRESORA";
                _cells.GetCell(4, TitleRow01).Value = "MEDIO";
                _cells.GetCell(5, TitleRow01).Value = "PARAM 1";
                _cells.GetCell(6, TitleRow01).Value = "PARAM 2";
                _cells.GetCell(7, TitleRow01).Value = "PARAM 3";
                _cells.GetCell(8, TitleRow01).Value = "PARAM 4";
                _cells.GetCell(9, TitleRow01).Value = "PARAM 5";
                _cells.GetCell(10, TitleRow01).Value = "PARAM 6";
                _cells.GetCell(11, TitleRow01).Value = "PARAM 7";
                _cells.GetCell(12, TitleRow01).Value = "PARAM 8";
                _cells.GetCell(13, TitleRow01).Value = "PARAM 9";
                _cells.GetCell(14, TitleRow01).Value = "PARAM 10";
                _cells.GetCell(15, TitleRow01).Value = "PARAM 11";
                _cells.GetCell(16, TitleRow01).Value = "PARAM 12";
                _cells.GetCell(17, TitleRow01).Value = "PARAM 13";
                _cells.GetCell(18, TitleRow01).Value = "PARAM 14";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);




                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void LoadRequest()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(drpEnviroment.SelectedItem.Label))
                    throw new Exception("Seleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                //ScreenService Opción en reemplazo de los servicios
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                var proxySheet = new Screen.ScreenService
                {
                    Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService"
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = TitleRow01 + 1;
                while ("" + _cells.GetCell(1, i).Value != "")
                {
                    try
                    {
                        _eFunctions.RevertOperation(opSheet, proxySheet);


                        var program = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value2);
                        var execute = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2);
                        var printer = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value2);
                        var medium = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value2);
                        var param1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value2);
                        var param2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value2);
                        var param3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value2);
                        var param4 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value2);
                        var param5 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value2);
                        var param6 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value2);
                        var param7 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value2);
                        var param8 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value2);
                        var param9 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value2);
                        var param10 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value2);
                        var param11 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value2);
                        var param12 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value2);
                        var param13 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value2);
                        var param14 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, i).Value2);


                        //ejecutamos el programa
                        var reply = proxySheet.executeScreen(opSheet, "MSO080");
                        //Validamos el ingreso
                        if (reply.mapName != "MSM080A") continue;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("RESTART1I", program);
                        var request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);

                        if (reply.mapName != "MSM080A") continue;
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("SKLITEM1I", "1");

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);
                        //Validamos el ingreso
                        if (reply.mapName != "MSM080B") continue;

                        //se adicionan los valores a los campos
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("SUBMIT_FLG2I", execute);
                        arrayFields.Add("PRINTER_NAME2I", printer);
                        arrayFields.Add("MEDIUM2I", medium);

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);

                        if (reply.mapName != "MSM080C") continue;

                        //se adicionan los valores a los campos
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("PARM3I1", param1);
                        arrayFields.Add("PARM3I2", param2);
                        arrayFields.Add("PARM3I3", param3);
                        arrayFields.Add("PARM3I4", param4);
                        arrayFields.Add("PARM3I5", param5);
                        arrayFields.Add("PARM3I6", param6);
                        arrayFields.Add("PARM3I7", param7);
                        arrayFields.Add("PARM3I8", param8);
                        arrayFields.Add("PARM3I9", param9);
                        arrayFields.Add("PARM3I10", param10);
                        arrayFields.Add("PARM3I11", param11);
                        arrayFields.Add("PARM3I12", param12);
                        arrayFields.Add("PARM3I13", param13);
                        arrayFields.Add("PARM3I14", param14);


                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);

                        _cells.GetCell(ResultColumn01, i).Value2 = reply.message;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn01, i).Value2 = "ERROR: " + ex.Message;
                    }
                    finally
                    {
                        i++;
                        _cells.GetCell(ResultColumn01, i).Select();
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:LoadRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}

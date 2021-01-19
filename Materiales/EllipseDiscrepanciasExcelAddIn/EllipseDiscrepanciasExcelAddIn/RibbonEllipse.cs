using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using EllipseDiscrepanciasExcelAddIn.CountTaskService;
using EllipseDiscrepanciasExcelAddIn.DiscrepancyTaskService;
using EllipseDiscrepanciasExcelAddIn.DiscrepancyLogService;
using EllipseDiscrepanciasExcelAddIn.AdjustDiscrepantHoldingService;
using SharedClassLibrary;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using System.Threading;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;

namespace EllipseDiscrepanciasExcelAddIn
{
    public partial class RibbonEllipse
    {

        ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;
        string _sheetName01 = "MSE1SF-MSE1SX";
        string _sheetName02 = "MSE1TD";
        Worksheet _worksheet;
        Worksheet _worksheet2;
        string _colHeader = "E";
        string _colFinal = "F";
        string _colOcultar = "G1";
        int _rowCabezera = 10;
        int _rowInicial = 11;
        int _maxRow = 10000;
        string _colHeader2 = "G";
        string _colFinal2 = "H";
        string _colOcultar2 = "I1";
        int _rowCabezera2 = 9;
        int _rowInicial2 = 10;
        string _custodianId;
        string _taskId;
        string _sourceTaskId;
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

                MessageBox.Show(ex.Message, Resources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
        
        public void SetSheetHeaderData()
        {
            try
            {
                this._excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();

                while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                    _excelApp.ActiveWorkbook.Worksheets.Add();                

                if (_cells == null)

                    _cells = new ExcelStyleCells(this._excelApp);

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = _sheetName01;

                _cells.GetCell(_colFinal + "1").Value = "OBLIGATORIO";
                _cells.GetCell(_colFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal + "2").Value = "OPCIONAL";
                _cells.GetCell(_colFinal + "2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(_colFinal + "3").Value = "INFORMATIVO";
                _cells.GetCell(_colFinal + "3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(_colFinal + "4").Value = "RESULTADO INCORRECTO";
                _cells.GetCell(_colFinal + "4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell(_colFinal + "5").Value = "RESULTADO CORRECTO";
                _cells.GetCell(_colFinal + "5").Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(_colOcultar, "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A" + (_rowCabezera-3)).Value = "ASIGNADO A";
                _cells.GetCell("A" + (_rowCabezera-3)).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B" + (_rowCabezera-3)).AddComment("ASIGNADO A");

                _cells.GetCell("A" + (_rowCabezera - 2)).Value = "DISTRITO";
                _cells.GetCell("A" + (_rowCabezera - 2)).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("B" + (_rowCabezera - 2)).Value = "ICOR";
                _cells.GetCell("B" + (_rowCabezera - 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                _cells.GetCell("A" + _rowCabezera).Value = "STOCKCODE";
                _cells.GetCell("A" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("B" + _rowCabezera).Value = "BODEGA";
                _cells.GetCell("B" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
              //  _cells.GetCell("B" + RowCabezera).AddComment("NI para pedido normal o CR para devolucion");

                _cells.GetCell("C" + _rowCabezera).Value = "CANTIDAD CONTADA";
                _cells.GetCell("C" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
              //  _cells.GetCell("C" + RowCabezera).AddComment("Formato YYYYMMDD");

                _cells.GetCell("D" + _rowCabezera).Value = "RAZON DISCREPANCIA";
                _cells.GetCell("D" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
             //   _cells.GetCell("D" + RowCabezera).AddComment("USUARIO");

                _cells.GetCell("E" + _rowCabezera).Value = "COMENTARIO DISCREPANCIA";
                _cells.GetCell("E" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("E" + _rowCabezera).AddComment("Los primeros 8 caracteres deben ser el documento OnBase. Ej:MD123456. Comentario de la Discrepancia");

                _cells.GetCell("F" + _rowCabezera).Value = "RESULTADO";
                _cells.GetCell("F" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MSE1SF-MSE1SX - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", _colHeader + "5");
                _cells.GetRange("B1", _colHeader + "5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", _colHeader + "5").Borders.Weight = "2";
                /*_cells.merge_cells("C6", "L11");
                _cells.getRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.getRange("C6", "L11").Borders.Weight = "2";
                
                */
                _cells.MergeCells("A" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4));
                _cells.GetRange("A" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4)).Borders.Weight = "2";
                _cells.GetRange("A" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4)).EntireColumn.AutoFit();

                _cells.MergeCells("C" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 2));
                _cells.GetRange("B" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 2)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 2)).Borders.Weight = "2";
                _cells.GetRange("B" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 2)).EntireColumn.AutoFit();

                _cells.MergeCells("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1));
                _cells.GetRange("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1)).Borders.Weight = "2";
                _cells.GetRange("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1)).EntireColumn.AutoFit();

                _cells.GetCell("A" + _rowInicial).Select();

                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = _sheetName02;

                _cells.GetCell(_colFinal2 + "1").Value = "OBLIGATORIO";
                _cells.GetCell(_colFinal2 + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal2 + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal2 + "2").Value = "OPCIONAL";
                _cells.GetCell(_colFinal2 + "2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(_colFinal2 + "3").Value = "INFORMATIVO";
                _cells.GetCell(_colFinal2 + "3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(_colFinal2 + "4").Value = "RESULTADO INCORRECTO";
                _cells.GetCell(_colFinal2 + "4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell(_colFinal2 + "5").Value = "RESULTADO CORRECTO";
                _cells.GetCell(_colFinal2 + "5").Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(_colOcultar2, "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A" + (_rowCabezera2 - 2)).Value = "DISTRITO";
                _cells.GetCell("A" + (_rowCabezera2 - 2)).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("B" + (_rowCabezera2 - 2)).Value = "ICOR";
                _cells.GetCell("B" + (_rowCabezera2 - 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                _cells.GetCell("A" + _rowCabezera2).Value = "STOCKCODE";
                _cells.GetCell("A" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("B" + _rowCabezera2).Value = "BODEGA";
                _cells.GetCell("B" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                //  _cells.GetCell("B" + RowCabezera).AddComment("NI para pedido normal o CR para devolucion");

                _cells.GetCell("C" + _rowCabezera2).Value = "ACCION";
                _cells.GetCell("C" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                //  _cells.GetCell("C" + RowCabezera).AddComment("Formato YYYYMMDD");

                _cells.GetCell("D" + _rowCabezera2).Value = "WRITE OFF";
                _cells.GetCell("D" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                //   _cells.GetCell("D" + RowCabezera).AddComment("USUARIO");

                _cells.GetCell("E" + _rowCabezera2).Value = "BRING ON";
                _cells.GetCell("E" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("F" + _rowCabezera2).Value = "CODIGO RESOLUCION";
                _cells.GetCell("F" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("G" + _rowCabezera2).Value = "DOCUMENTO ONBASE";
                _cells.GetCell("G" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("G" + _rowCabezera2).AddComment("Solo se permiten 8 Caracteres. Ej: MD123456");

                _cells.GetCell("H" + _rowCabezera2).Value = "RESULTADO";
                _cells.GetCell("H" + _rowCabezera2).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MSE1TD - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", _colHeader2 + "5");
                _cells.GetRange("B1", _colHeader2 + "5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", _colHeader2 + "5").Borders.Weight = "2";
                /*_cells.merge_cells("C6", "L11");
                _cells.getRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.getRange("C6", "L11").Borders.Weight = "2";
                
                */

                _cells.MergeCells("A" + (_rowCabezera2 - 3), _colFinal2 + (_rowCabezera2 - 3));
                _cells.GetRange("A" + (_rowCabezera2 - 3), _colFinal2 + (_rowCabezera2 - 3)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A" + (_rowCabezera2 - 3), _colFinal2 + (_rowCabezera2 - 3)).Borders.Weight = "2";
                _cells.GetRange("A" + (_rowCabezera2 - 3), _colFinal2 + (_rowCabezera2 - 3)).EntireColumn.AutoFit();

                _cells.MergeCells("C" + (_rowCabezera2 - 2), _colFinal2 + (_rowCabezera2 - 2));
                _cells.GetRange("C" + (_rowCabezera2 - 2), _colFinal2 + (_rowCabezera2 - 2)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("C" + (_rowCabezera2 - 2), _colFinal2 + (_rowCabezera2 - 2)).Borders.Weight = "2";
                _cells.GetRange("C" + (_rowCabezera2 - 2), _colFinal2 + (_rowCabezera2 - 2)).EntireColumn.AutoFit();

                _cells.MergeCells("A" + (_rowCabezera2 - 1), _colFinal2 + (_rowCabezera2 - 1));
                _cells.GetRange("A" + (_rowCabezera2 - 1), _colFinal2 + (_rowCabezera2 - 1)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A" + (_rowCabezera2 - 1), _colFinal2 + (_rowCabezera2 - 1)).Borders.Weight = "2";
                _cells.GetRange("A" + (_rowCabezera2 - 1), _colFinal2 + (_rowCabezera2 - 1)).EntireColumn.AutoFit();

                _cells.GetCell("A" + _rowInicial2).Select();

                _excelApp.ActiveWorkbook.Sheets[1].Select();
            }
            catch (Exception ex)
            {
                //ErrorLogger.LogError("ExcelStyle_cells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, this.debugErrors);
                MessageBox.Show("Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
        }

        public void AutoAjuste(Excel.Range target)
        {
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        private void Limpiar()
        {

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                _cells.GetCell("A" + _rowInicial + ":" + _colHeader + _maxRow).ClearContents();
                _cells.GetCell(_colFinal + _rowInicial + ":" + _colFinal + _maxRow).Clear();
            }

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName02)
            {
                _cells.GetCell("A" + _rowInicial2 + ":" + _colHeader2 + _maxRow).ClearContents();
                _cells.GetCell(_colFinal2 + _rowInicial2 + ":" + _colFinal2 + _maxRow).Clear();
            }

        }

        private void LimpiarResultado()
        {

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                _cells.GetCell(_colFinal + _rowInicial + ":" + _colFinal + _maxRow).ClearContents();
                _cells.GetCell(_colFinal + _rowInicial + ":" + _colFinal + _maxRow).Clear();
            }

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName02)
            {
                _cells.GetCell(_colFinal2 + _rowInicial2 + ":" + _colFinal2 + _maxRow).ClearContents();
                _cells.GetCell(_colFinal2 + _rowInicial2 + ":" + _colFinal2 + _maxRow).Clear();
            }

        }

        private void Centrar()
        {
            var bodegas = GetBodegas();
            var razones = GetRazones();

            _cells.GetCell("A" + _rowInicial + ":" + _colHeader + _maxRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            _cells.GetCell("A" + _rowInicial + ":A" + _maxRow).NumberFormat = "@";
            _cells.GetCell("B" + _rowInicial + ":B" + _maxRow).NumberFormat = "@";
            _cells.GetCell("C" + _rowInicial + ":C" + _maxRow).NumberFormat = "@";
            _cells.GetCell("D" + _rowInicial + ":D" + _maxRow).NumberFormat = "@";
            _cells.GetCell("E" + _rowInicial + ":E" + _maxRow).NumberFormat = "@";
            _cells.GetCell("F" + _rowInicial + ":F" + _maxRow).NumberFormat = "@";
            _cells.SetValidationList(_cells.GetCell("B" + _rowInicial + ":B" + _maxRow), bodegas);
            _cells.SetValidationList(_cells.GetCell("D" + _rowInicial + ":D" + _maxRow), razones);

        }

        private void Centrar2()
        {

            _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);

            var bodegas = GetBodegas();
            var resoluciones = GetResoluciones();
            var acciones = GetAcciones();

            _cells.GetCell("A" + _rowInicial2 + ":" + _colHeader + _maxRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            _cells.GetCell("A" + _rowInicial2 + ":A" + _maxRow).NumberFormat = "@";
            _cells.GetCell("B" + _rowInicial2 + ":B" + _maxRow).NumberFormat = "@";
            _cells.GetCell("C" + _rowInicial2 + ":C" + _maxRow).NumberFormat = "@";
            _cells.GetCell("D" + _rowInicial2 + ":D" + _maxRow).NumberFormat = "@";
            _cells.GetCell("E" + _rowInicial2 + ":E" + _maxRow).NumberFormat = "@";
            _cells.GetCell("F" + _rowInicial2 + ":F" + _maxRow).NumberFormat = "@";
            _cells.GetCell("G" + _rowInicial2 + ":G" + _maxRow).NumberFormat = "@";
            _cells.GetCell("H" + _rowInicial2 + ":H" + _maxRow).NumberFormat = "@";
            _cells.GetCell("G" + _rowInicial2 + ":G" + _maxRow).Validation.Add(Excel.XlDVType.xlValidateTextLength, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlLess,9,Type.Missing);
            _cells.SetValidationList(_cells.GetCell("B" + _rowInicial2 + ":B" + _maxRow), bodegas);
            _cells.SetValidationList(_cells.GetCell("C" + _rowInicial2 + ":C" + _maxRow), acciones);
            _cells.SetValidationList(_cells.GetCell("F" + _rowInicial2 + ":F" + _maxRow), resoluciones);

        }

        public List<string> GetAcciones()
        {

            var getAcciones = new List<string>();

            getAcciones.Add("Adjust");
            getAcciones.Add("Resolution");

            return getAcciones;
        }

        public List<string> GetBodegas()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "SELECT TRIM(SUBSTR(TABLE_CODE,5,LENGTH(TABLE_CODE))) AS BODEGA FROM ELLIPSE.MSF010 WHERE TABLE_TYPE = 'WH' AND SUBSTR(TABLE_CODE,1,4) = 'ICOR' ORDER BY 1";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getBodegas = new List<string>();

            while (odr.Read())
            {
                getBodegas.Add("" + odr["BODEGA"]);
            }
            return getBodegas;
        }

        public List<string> GetRazones()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "SELECT TRIM(TABLE_CODE) AS RAZON FROM ELLIPSE.MSF010 WHERE TABLE_TYPE = 'DRRS'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getRazones = new List<string>();

            while (odr.Read())
            {
                getRazones.Add("" + odr["RAZON"]);
            }
            return getRazones;
        }

        public List<string> GetResoluciones()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "SELECT TRIM(TABLE_CODE) AS RESOLUCION FROM ELLIPSE.MSF010 WHERE TABLE_TYPE = 'DRRC'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getResoluciones = new List<string>();

            while (odr.Read())
            {
                getResoluciones.Add("" + odr["RESOLUCION"]);
            }
            return getResoluciones;
        }

        private void bProcesar_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
                {
                    LimpiarResultado();

                    if (_frmAuth.ShowDialog() == DialogResult.OK)
                    //if (true)
                    {
                        //  _frmAuth.EllipseDsct = "ICOR";
                        //  _frmAuth.EllipsePost = "ADMIN";
                        //  _frmAuth.EllipseUser = "ljuvinao";
                        //  _frmAuth.EllipsePswd = "";
                        //  _cells.GetCell("A1").Value = "Conectado";

                        var proxySheet = new CountTaskService.CountTaskService();
                        var opSheet = new CountTaskService.OperationContext();

                        var currentRow = _rowInicial;

                        string assigned = "" + _cells.GetCell("B7").Value;
                        string districtCode = "" + _cells.GetCell("B8").Value;
                        string stockCode = "" + _cells.GetCell("A" + currentRow).Value;
                        string whouse = "" + _cells.GetCell("B" + currentRow).Value;
                        string countQuantity = "" + _cells.GetCell("C" + currentRow).Value;
                        string codeReason = "" + _cells.GetCell("D" + currentRow).Value;
                        string comments = "" + _cells.GetCell("E" + currentRow).Value;


                        while (!string.IsNullOrEmpty(stockCode))
                        {
                            try
                            {
                                var requestParamsSheet = new CountTaskDTO();


                                proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CountTaskService";

                                opSheet.district = _frmAuth.EllipseDsct;
                                opSheet.position = _frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                requestParamsSheet.assignedTo = assigned;
                                requestParamsSheet.stockCode = stockCode;
                                requestParamsSheet.districtCode = districtCode;
                                requestParamsSheet.warehouseId = whouse;

                                var replySheet = proxySheet.create(opSheet, requestParamsSheet);

                                if (replySheet.errors.Count() > 0)
                                {
                                    _cells.GetCell(_colFinal + currentRow).Value = replySheet.errors[0].messageText;
                                    _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(_colFinal + currentRow).Select();
                                }
                                else
                                {

                                    var proxySheet2 = new CountTaskService.CountTaskService();
                                    var opSheet2 = new CountTaskService.OperationContext();

                                    var requestParamsSheet2 = new CountTaskDTO[1];

                                    proxySheet2.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CountTaskService";

                                    opSheet2.district = _frmAuth.EllipseDsct;
                                    opSheet2.position = _frmAuth.EllipsePost;
                                    opSheet2.maxInstances = 100;
                                    opSheet2.returnWarnings = Debugger.DebugWarnings;

                                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                                    requestParamsSheet2[0] = new CountTaskDTO();
                                    requestParamsSheet2[0].stockCode = replySheet.countTaskDTO.stockCode;
                                    requestParamsSheet2[0].custodianId = replySheet.countTaskDTO.custodianId;
                                    requestParamsSheet2[0].taskId = replySheet.countTaskDTO.taskId;
                                    requestParamsSheet2[0].countQty = Convert.ToDecimal(countQuantity);
                                    requestParamsSheet2[0].countQtySpecified = true;

                                    var replySheet2 = proxySheet2.multipleUpdate(opSheet2, requestParamsSheet2);


                                    if (replySheet2[0].errors.Count() > 0)
                                    {
                                        _cells.GetCell(_colFinal + currentRow).Value = replySheet2[0].errors[0].messageText;
                                        _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                        _cells.GetCell(_colFinal + currentRow).Select();
                                    }
                                    else
                                    {

                                        var proxySheet3 = new CountTaskService.CountTaskService();
                                        var opSheet3 = new CountTaskService.OperationContext();

                                        var requestParamsSheet3 = new CountTaskDTO[1];

                                        proxySheet3.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CountTaskService";

                                        opSheet3.district = _frmAuth.EllipseDsct;
                                        opSheet3.position = _frmAuth.EllipsePost;
                                        opSheet3.maxInstances = 100;
                                        opSheet3.returnWarnings = Debugger.DebugWarnings;

                                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                        requestParamsSheet3[0] = new CountTaskDTO();

                                        requestParamsSheet3[0].stockCode = replySheet.countTaskDTO.stockCode;
                                        requestParamsSheet3[0].custodianId = replySheet.countTaskDTO.custodianId;
                                        requestParamsSheet3[0].taskId = replySheet.countTaskDTO.taskId;



                                        var replySheet3 = proxySheet3.multipleReconcile(opSheet3, requestParamsSheet3);

                                        if (replySheet3[0].errors.Count() > 0)
                                        {

                                            _cells.GetCell(_colFinal + currentRow).Value = replySheet3[0].errors[0].messageText;
                                            _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                            _cells.GetCell(_colFinal + currentRow).Select();

                                        }
                                        else
                                        {

                                            if (replySheet3[0].countTaskDTO.countErrorStatus == "SH")
                                            {

                                                var proxySheet4 = new CountTaskService.CountTaskService();
                                                var opSheet4 = new CountTaskService.OperationContext();

                                                var requestParamsSheet4 = new CountTaskDTO[1];

                                                proxySheet4.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/CountTaskService";

                                                opSheet4.district = _frmAuth.EllipseDsct;
                                                opSheet4.position = _frmAuth.EllipsePost;
                                                opSheet4.maxInstances = 100;
                                                opSheet4.returnWarnings = Debugger.DebugWarnings;

                                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                                requestParamsSheet4[0] = new CountTaskDTO();

                                                requestParamsSheet4[0].stockCode = replySheet.countTaskDTO.stockCode;
                                                requestParamsSheet4[0].custodianId = replySheet.countTaskDTO.custodianId;
                                                requestParamsSheet4[0].taskId = replySheet.countTaskDTO.taskId;
                                                requestParamsSheet4[0].discrepancyReason = codeReason;
                                                requestParamsSheet4[0].discrepancyComment = comments;

                                                var replySheet4 = proxySheet4.multipleRaiseDiscrepancy(opSheet4, requestParamsSheet4);

                                                if (replySheet3[0].errors.Count() > 0)
                                                {
                                                    _cells.GetCell(_colFinal + currentRow).Value = replySheet3[0].errors[0].messageText;
                                                    _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                                    _cells.GetCell(_colFinal + currentRow).Select();
                                                }
                                                else
                                                {
                                                    _cells.GetCell(_colFinal + currentRow).Value = replySheet3[0].informationalMessages[0].messageText;
                                                    _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleResult);
                                                    _cells.GetCell(_colFinal + currentRow).Select();

                                                }


                                            }
                                            else
                                            {
                                                _cells.GetCell(_colFinal + currentRow).Value = replySheet3[0].informationalMessages[0].messageText;
                                                _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleResult);
                                                _cells.GetCell(_colFinal + currentRow).Select();
                                            }
                                        }

                                    }

                                }

                            }
                            catch (Exception ex)
                            {
        
                                    _cells.GetCell(_colFinal + currentRow).Value = ex.Message;
                                    _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(_colFinal + currentRow).Select();


                            }
                            finally
                            {
                                currentRow++;
                                stockCode = "" + _cells.GetCell("A" + currentRow).Value;
                                whouse = "" + _cells.GetCell("B" + currentRow).Value;
                                countQuantity = "" + _cells.GetCell("C" + currentRow).Value;
                                codeReason = "" + _cells.GetCell("D" + currentRow).Value;
                                comments = "" + _cells.GetCell("E" + currentRow).Value;
                            }

                        }

                        MessageBox.Show("Proceso Finalizado Correctamente");

                    }

                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                //ebugger.LogError("RibbonEllipse:CreateWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            
        }

        public bool IsNumeric(object expression)
        {

            bool isNum;

            double retNum;

            isNum = double.TryParse(Convert.ToString(expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);

            return isNum;

        }

        private void bProcesarMSE1TD_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName02)
                {

                    LimpiarResultado();

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                    if (_frmAuth.ShowDialog() == DialogResult.OK)
                    //if (true)
                    {
                        //  _frmAuth.EllipseDsct = "ICOR";
                        //  _frmAuth.EllipsePost = "ADMIN";
                        //  _frmAuth.EllipseUser = "ljuvinao";
                        //  _frmAuth.EllipsePswd = "";
                        //  _cells.GetCell("A1").Value = "Conectado";

                        var proxySheet = new DiscrepancyTaskService.DiscrepancyTaskService();
                        var opSheet = new DiscrepancyTaskService.OperationContext();

                        var currentRow = _rowInicial2;

                        string districtCode = "" + _cells.GetCell("B7").Value;
                        string stockCode = "" + _cells.GetCell("A" + currentRow).Value;
                        string whouse = "" + _cells.GetCell("B" + currentRow).Value;
                        string action = "" + _cells.GetCell("C" + currentRow).Value;
                        string writeOff = "" + _cells.GetCell("D" + currentRow).Value;
                        string bringOn = "" + _cells.GetCell("E" + currentRow).Value;
                        string resolution = "" + _cells.GetCell("F" + currentRow).Value;
                        string docOnbase = "" + _cells.GetCell("G" + currentRow).Value;


                        while (!string.IsNullOrEmpty(stockCode))
                        {
                            try
                            {
                                var requestParamsSheet = new DiscrepancyTaskSearchParam();
                                var replySheet = new DiscrepancyTaskServiceResult[100];

                                proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/DiscrepancyTaskService";

                                opSheet.district = _frmAuth.EllipseDsct;
                                opSheet.position = _frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                requestParamsSheet.districtCode = districtCode;
                                requestParamsSheet.warehouseId = whouse;
                                requestParamsSheet.stockCode = stockCode;

                                replySheet = proxySheet.search(opSheet, requestParamsSheet, null);

                                if (replySheet[0].errors.Count() > 0)
                                {
                                    _cells.GetCell(_colFinal2 + currentRow).Value = replySheet[0].errors[0].messageText;
                                    _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(_colFinal2 + currentRow).Select();
                                }
                                else
                                {

                                    var numDisc = replySheet.Count();
                                    var iter = 0;

                                    if (numDisc > 1 && replySheet[0].errors.Count() == 0)
                                    {

                                        while (iter < numDisc)
                                        {

                                            var proxySheet2 = new DiscrepancyLogService.DiscrepancyLogService();
                                            var opSheet2 = new DiscrepancyLogService.OperationContext();

                                            var requestParamsSheet2 = new DiscrepancyLogSearchParam();

                                            proxySheet2.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/DiscrepancyLogService";

                                            opSheet2.district = _frmAuth.EllipseDsct;
                                            opSheet2.position = _frmAuth.EllipsePost;
                                            opSheet2.maxInstances = 100;
                                            opSheet2.returnWarnings = Debugger.DebugWarnings;

                                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                            requestParamsSheet2.custodianId = replySheet[iter].discrepancyTaskDTO.custodianId;
                                            requestParamsSheet2.taskId = replySheet[iter].discrepancyTaskDTO.taskId;

                                            var replySheet2 = proxySheet2.search(opSheet2, requestParamsSheet2, null);

                                            var comentario = replySheet2[0].discrepancyLogDTO.comment.Trim();

                                            comentario = comentario.Substring(0, 8);

                                            if (comentario == docOnbase)
                                            {
                                                _custodianId = replySheet[iter].discrepancyTaskDTO.custodianId;
                                                _taskId = replySheet[iter].discrepancyTaskDTO.taskId;
                                                _sourceTaskId = replySheet[iter].discrepancyTaskDTO.sourceTaskId;

                                                iter = numDisc + 1;
                                            }

                                            iter++;

                                        }

                                    }
                                    else
                                    {
                                        _custodianId = replySheet[0].discrepancyTaskDTO.custodianId;
                                        _taskId = replySheet[0].discrepancyTaskDTO.taskId;
                                        _sourceTaskId = replySheet[0].discrepancyTaskDTO.sourceTaskId;                                        
                                    }

                                    if (action == "Adjust" && !string.IsNullOrEmpty(_custodianId) && !string.IsNullOrEmpty(_taskId) && !string.IsNullOrEmpty(_sourceTaskId))
                                    {

                                        var proxySheet3 = new AdjustDiscrepantHoldingService.AdjustDiscrepantHoldingService();
                                        var opSheet3 = new AdjustDiscrepantHoldingService.OperationContext();

                                        var searchParams = new AdjustDiscrepantHoldingItemSearchParam();

                                        proxySheet3.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/AdjustDiscrepantHoldingService";

                                        opSheet3.district = _frmAuth.EllipseDsct;
                                        opSheet3.position = _frmAuth.EllipsePost;
                                        opSheet3.maxInstances = 100;
                                        opSheet3.returnWarnings = Debugger.DebugWarnings;

                                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                        searchParams.custodianId = _custodianId;
                                        searchParams.taskId = _taskId;
                                        searchParams.stockCode = stockCode;

                                        var replySheet3 = proxySheet3.searchItems(opSheet3, searchParams, null);

                                        var l = replySheet3.GetLength(0);

                                        if (l > 0)
                                        {                                        

                                        var proxySheet4 = new AdjustDiscrepantHoldingService.AdjustDiscrepantHoldingService();
                                        var opSheet4 = new AdjustDiscrepantHoldingService.OperationContext();

                                        var requestParamsSheet4 = new AdjustDiscrepantHoldingItemDTO[1];

                                        proxySheet4.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/AdjustDiscrepantHoldingService";

                                        opSheet4.district = _frmAuth.EllipseDsct;
                                        opSheet4.position = _frmAuth.EllipsePost;
                                        opSheet4.maxInstances = 100;
                                        opSheet4.returnWarnings = Debugger.DebugWarnings;

                                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                        requestParamsSheet4[0] = new AdjustDiscrepantHoldingItemDTO();

                                        requestParamsSheet4[0].custodianId = _custodianId;
                                        requestParamsSheet4[0].fixedStatus = replySheet3[0].adjustdiscrepantholdingitemdto.fixedStatus;
                                        requestParamsSheet4[0].holdingId = replySheet3[0].adjustdiscrepantholdingitemdto.holdingId;
                                        requestParamsSheet4[0].holdingType = replySheet3[0].adjustdiscrepantholdingitemdto.holdingType;

                                        if (!string.IsNullOrEmpty(writeOff))
                                        {
                                            requestParamsSheet4[0].writeOff = Convert.ToDecimal(writeOff);
                                            requestParamsSheet4[0].writeOffSpecified = true;
                                        }

                                        if (!string.IsNullOrEmpty(bringOn))
                                        {
                                            requestParamsSheet4[0].bringOn = Convert.ToDecimal(bringOn);
                                            requestParamsSheet4[0].bringOnSpecified = true;
                                        }

                                        requestParamsSheet4[0].taskId = _taskId;
                                        requestParamsSheet4[0].stockCode = stockCode;

                                        var replySheet4 = proxySheet4.multipleUpdateItem(opSheet4, requestParamsSheet4);

                                        if (replySheet4[0].errors.Count() > 0)
                                        {
                                            _cells.GetCell(_colFinal2 + currentRow).Value = replySheet4[0].errors[0].messageText;
                                            _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                            _cells.GetCell(_colFinal2 + currentRow).Select();
                                        }
                                        else
                                        {
                                            var proxySheet5 = new AdjustDiscrepantHoldingService.AdjustDiscrepantHoldingService();
                                            var opSheet5 = new AdjustDiscrepantHoldingService.OperationContext();

                                            var requestParamsSheet5 = new AdjustDiscrepantHoldingDTO();

                                            proxySheet5.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/AdjustDiscrepantHoldingService";

                                            opSheet5.district = _frmAuth.EllipseDsct;
                                            opSheet5.position = _frmAuth.EllipsePost;
                                            opSheet5.maxInstances = 100;
                                            opSheet5.returnWarnings = Debugger.DebugWarnings;

                                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                            requestParamsSheet5.districtCode = districtCode;
                                            requestParamsSheet5.stockCode = stockCode;
                                            requestParamsSheet5.custodianId = _custodianId;
                                            requestParamsSheet5.taskId = _taskId;
                                            requestParamsSheet5.resolutionCode = resolution;
                                            requestParamsSheet5.warehouseId = whouse;

                                            var replySheet5 = proxySheet5.finalise(opSheet5, requestParamsSheet5);

                                            if (replySheet5.errors.Count() > 0)
                                            {
                                                _cells.GetCell(_colFinal2 + currentRow).Value = replySheet5.errors[0].messageText;
                                                _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                                _cells.GetCell(_colFinal2 + currentRow).Select();
                                            }
                                            else
                                            {
                                                _cells.GetCell(_colFinal2 + currentRow).Value = replySheet5.informationalMessages[0].messageText;
                                                _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleResult);
                                                _cells.GetCell(_colFinal2 + currentRow).Select();
                                            }

                                        }

                                        }
                                        else
                                        {
                                            var proxySheet4 = new AdjustDiscrepantHoldingService.AdjustDiscrepantHoldingService();
                                            var opSheet4 = new AdjustDiscrepantHoldingService.OperationContext();

                                            var requestParamsSheet4 = new AdjustDiscrepantHoldingItemDTO[1];

                                            proxySheet4.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/AdjustDiscrepantHoldingService";

                                            opSheet4.district = _frmAuth.EllipseDsct;
                                            opSheet4.position = _frmAuth.EllipsePost;
                                            opSheet4.maxInstances = 100;
                                            opSheet4.returnWarnings = Debugger.DebugWarnings;

                                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                            requestParamsSheet4[0] = new AdjustDiscrepantHoldingItemDTO();

                                            requestParamsSheet4[0].custodianId = _custodianId;

                                            if (!string.IsNullOrEmpty(writeOff))
                                            {
                                                requestParamsSheet4[0].writeOff = Convert.ToDecimal(writeOff);
                                                requestParamsSheet4[0].writeOffSpecified = true;
                                            }

                                            if (!string.IsNullOrEmpty(bringOn))
                                            {
                                                requestParamsSheet4[0].bringOn = Convert.ToDecimal(bringOn);
                                                requestParamsSheet4[0].bringOnSpecified = true;
                                            }

                                            requestParamsSheet4[0].taskId = _taskId;
                                            requestParamsSheet4[0].stockCode = stockCode;
                                            requestParamsSheet4[0].stockOwnershipIndicator = "O";

                                            var replySheet4 = proxySheet4.multipleCreateItem(opSheet4, requestParamsSheet4);

                                            if (replySheet4[0].errors.Count() > 0)
                                            {
                                                _cells.GetCell(_colFinal2 + currentRow).Value = replySheet4[0].errors[0].messageText;
                                                _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                                _cells.GetCell(_colFinal2 + currentRow).Select();
                                            }
                                            else
                                            {
                                                var proxySheet5 = new AdjustDiscrepantHoldingService.AdjustDiscrepantHoldingService();
                                                var opSheet5 = new AdjustDiscrepantHoldingService.OperationContext();

                                                var requestParamsSheet5 = new AdjustDiscrepantHoldingDTO();

                                                proxySheet5.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/AdjustDiscrepantHoldingService";

                                                opSheet5.district = _frmAuth.EllipseDsct;
                                                opSheet5.position = _frmAuth.EllipsePost;
                                                opSheet5.maxInstances = 100;
                                                opSheet5.returnWarnings = Debugger.DebugWarnings;

                                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                                requestParamsSheet5.districtCode = districtCode;
                                                requestParamsSheet5.stockCode = stockCode;
                                                requestParamsSheet5.custodianId = _custodianId;
                                                requestParamsSheet5.taskId = _taskId;
                                                requestParamsSheet5.resolutionCode = resolution;
                                                requestParamsSheet5.warehouseId = whouse;

                                                var replySheet5 = proxySheet5.finalise(opSheet5, requestParamsSheet5);

                                                if (replySheet5.errors.Count() > 0)
                                                {
                                                    _cells.GetCell(_colFinal2 + currentRow).Value = replySheet5.errors[0].messageText;
                                                    _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                                    _cells.GetCell(_colFinal2 + currentRow).Select();
                                                }
                                                else
                                                {
                                                    _cells.GetCell(_colFinal2 + currentRow).Value = replySheet5.informationalMessages[0].messageText;
                                                    _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleResult);
                                                    _cells.GetCell(_colFinal2 + currentRow).Select();
                                                }

                                            }
                                        }

                                    }

                                    if (action == "Resolution" && !string.IsNullOrEmpty(_custodianId) && !string.IsNullOrEmpty(_taskId) && !string.IsNullOrEmpty(_sourceTaskId))
                                    {

                                        var proxySheet6 = new DiscrepancyTaskService.DiscrepancyTaskService();
                                        var opSheet6 = new DiscrepancyTaskService.OperationContext();

                                        var requestParamsSheet6 = new DiscrepancyTaskDTO();

                                        proxySheet6.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/DiscrepancyTaskService";

                                        opSheet6.district = _frmAuth.EllipseDsct;
                                        opSheet6.position = _frmAuth.EllipsePost;
                                        opSheet6.maxInstances = 100;
                                        opSheet6.returnWarnings = Debugger.DebugWarnings;

                                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                        requestParamsSheet6.districtCode = districtCode;
                                        requestParamsSheet6.stockCode = stockCode;
                                        requestParamsSheet6.custodianId = _custodianId;
                                        requestParamsSheet6.sourceTaskId = _sourceTaskId;
                                        requestParamsSheet6.taskId = _taskId;
                                        requestParamsSheet6.resolutionCode = resolution;

                                        var replySheet6 = proxySheet6.resolve(opSheet6, requestParamsSheet6);

                                        if (replySheet6.errors.Count() > 0)
                                        {
                                            _cells.GetCell(_colFinal2 + currentRow).Value = replySheet6.errors[0].messageText;
                                            _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                            _cells.GetCell(_colFinal2 + currentRow).Select();
                                        }
                                        else
                                        {
                                            _cells.GetCell(_colFinal2 + currentRow).Value = replySheet6.informationalMessages[0].messageText;
                                            _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleResult);
                                            _cells.GetCell(_colFinal2 + currentRow).Select();
                                        }

                                    }

                                    if ((string.IsNullOrEmpty(action)))
                                    {
                                        _cells.GetCell(_colFinal2 + currentRow).Value = "Debe Seleccionar Alguna Accion (Adjust o Resolution)";
                                        _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                        _cells.GetCell(_colFinal2 + currentRow).Select();

                                    }

                                    if (string.IsNullOrEmpty(_custodianId))
                                    {
                                        _cells.GetCell(_colFinal2 + currentRow).Value = "Documento OnBase No Encontrado";
                                        _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                        _cells.GetCell(_colFinal2 + currentRow).Select();
                                    }

                                }
                                
                                
                            }
                            catch (Exception ex)
                            {

                                if(ex.Message == "Index was outside the bounds of the array.")
                                {
                                    _cells.GetCell(_colFinal2 + currentRow).Value = "StockCode no tiene Discrepancias";
                                    _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(_colFinal2 + currentRow).Select();


                                }
                                else
                                {
                                    _cells.GetCell(_colFinal2 + currentRow).Value = ex.Message;
                                    _cells.GetCell(_colFinal2 + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(_colFinal2 + currentRow).Select();


                                }

                            }
                            finally
                            {
                                currentRow++;
                                stockCode = "" + _cells.GetCell("A" + currentRow).Value;
                                whouse = "" + _cells.GetCell("B" + currentRow).Value;
                                action = "" + _cells.GetCell("C" + currentRow).Value;
                                writeOff = "" + _cells.GetCell("D" + currentRow).Value;
                                bringOn = "" + _cells.GetCell("E" + currentRow).Value;
                                resolution = "" + _cells.GetCell("F" + currentRow).Value;
                                docOnbase = "" + _cells.GetCell("G" + currentRow).Value;
                            }
                        }
                        MessageBox.Show("Proceso Finalizado Correctamente");
                    }
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                //Debugger.LogError("RibbonEllipse:CreateWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void bLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            Limpiar();
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            SetSheetHeaderData();
            Centrar();
            Centrar2();
            _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);

            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            Microsoft.Office.Tools.Excel.NamedRange groupRange;
            var groupCells = _worksheet.Range["A" + _rowInicial + ":" + _colFinal + _maxRow];
            groupRange = _worksheet.Controls.AddNamedRange(groupCells, "GroupRange");

            groupRange.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(AutoAjuste);

            _worksheet2 = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[2]);

            Microsoft.Office.Tools.Excel.NamedRange groupRange2;
            var groupCells2 = _worksheet2.Range["A" + _rowInicial2 + ":" + _colFinal2 + _maxRow];
            groupRange2 = _worksheet2.Controls.AddNamedRange(groupCells2, "GroupRange2");

            groupRange2.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(AutoAjuste);
        }
    }
}

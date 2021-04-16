using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Connections;
using EllipseSAO900AddIn.Properties;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService;

namespace EllipseSAO900AddIn
{
    public partial class RibbonEllipse
    {
        private const int MaxRows = 1000;
        private EllipseFunctions _eFunctions;
        private ExcelStyleCells _cells;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;
        private const int ResultColumn = 12;
        private const int TitleRow01 = 9;
        private string _sheetName01; //Variable según el formato
        
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
        private void btnFormatoReclasificaciones_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoReclasificaciones();
        }

        private void btnFormatoModificaciones_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoModificaciones();
        }

        private void btnFormatoCausaciones_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoCausaciones();
        }

        private void btnFormatoDistribuciones_Click(object sender, RibbonControlEventArgs e)
        {
            FormatoDistribuciones();
        }

        private void btnValidar_Click(object sender, RibbonControlEventArgs e)
        {
            ValidarDatos();
        }

        private void btnExportar_Click(object sender, RibbonControlEventArgs e)
        {
            ExportarDatos();
        }

        private void FormatoReclasificaciones()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                Microsoft.Office.Tools.Excel.Worksheet workSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _sheetName01 = "SAO900-Reclasificaciones";
                excelSheet.Name = _sheetName01;
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "RECLASIFICACIONES MESES ANTERIORES";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(2, 1, 7, 2);

                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "AUTORIZADOR";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 5, 7, 5);
                _cells.GetCell(1, 6).Value = "RAZON DEL CAMBIO";
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 7).Value = "USUARIO";
                _cells.GetCell(1, 7).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 7, 7, 7);

                _cells.GetCell(1, titleRow).Value = "NUM_TRANSACCION";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "PERIODO";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = "CCOSTOS / API";
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, titleRow).Value = "PROJ/WO";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = "IND";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, titleRow).Value = "CCOSTOS_DESTINO";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, titleRow).Value = "EQUIPO";
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "PROJ/WO_DESTINO";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "IND_DESTINO";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, titleRow).Value = "DOLARES";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(11, titleRow).Value = "PESOS";
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange1 = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "employeeRange1");
                employeeRange1.Change += changesEmployeeRange_Change;

                var employeeRange2 = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "employeeRange2");
                employeeRange2.Change += changesEmployeeRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_FormatoReclasificaciones_Error);
            }
        }

        private void FormatoModificaciones()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                Microsoft.Office.Tools.Excel.Worksheet workSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _sheetName01 = "SAO900-Modificaciones";
                var resultColumn = ResultColumn;
                var titleRow = TitleRow01;
                excelSheet.Name = _sheetName01;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "MODIFICACIONES MES CORRIENTE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(2, 1, 7, 2);

                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "AUTORIZADOR";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 5, 7, 5);
                _cells.GetCell(1, 6).Value = "RAZON DEL CAMBIO";
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 7).Value = "USUARIO";
                _cells.GetCell(1, 7).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 7, 7, 7);

                _cells.GetCell(1, titleRow).Value = "NUM_TRANSACCION";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "PERIODO";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = "CCOSTOS / API";
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, titleRow).Value = "PROJ/WO";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = "IND";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, titleRow).Value = "CCOSTOS_DESTINO";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, titleRow).Value = "EQUIPO";
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, titleRow).Value = "PROJ/WO_DESTINO";
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, titleRow).Value = "IND_DESTINO";
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, titleRow).Value = "DOLARES";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(11, titleRow).Value = "PESOS";
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange1 = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "employeeRange1");
                employeeRange1.Change += changesEmployeeRange_Change;

                var employeeRange2 = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "employeeRange2");
                employeeRange2.Change += changesEmployeeRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_FormatoReclasificaciones_Error);
            }
        }

        private void FormatoCausaciones()
        {
            var resultColumn = 6;
            var titleRow = 13;
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                Microsoft.Office.Tools.Excel.Worksheet workSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _sheetName01 = "SAO900-Causaciones";
                excelSheet.Name = _sheetName01;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "Causaciones";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(2, 1, 7, 2);

                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "SUPPLIER";
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 5, 7, 5);
                _cells.GetCell(1, 6).Value = "TIPO DE DOCUMENTO";
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 7).Value = "NUM DE DOCUMENTO";
                _cells.GetCell(1, 7).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 8).Value = "FECHA SOLICITUD";
                _cells.GetCell(1, 8).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 9).Value = "MONEDA: (Currency)";
                _cells.GetCell(1, 9).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 10).Value = "VALOR";
                _cells.GetCell(1, 10).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 11).Value = "SOLICITADO";
                _cells.GetCell(1, 11).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells(3, 11, 7, 11);

                _cells.GetRange(2, 2, 2, 11).NumberFormat = "@";

                _cells.GetCell(1, titleRow).Value = "CCOSTOS/DETALLE";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "PROYECTO / WO";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = "P/W";
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, titleRow).Value = "EQUIPO";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = "VALOR (PES o USD)";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange = workSheet.Controls.AddNamedRange(workSheet.Range["B11"], "employeeRange");
                employeeRange.Change += changesEmployeeRange_Change;

                var documentRange = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "documentRange");
                documentRange.Change += changesdocumentRangeRange_Change;

                var supplierRange = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "supplierRange");
                supplierRange.Change += changessupplierRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_FormatoCausaciones_Error);
            }
        }

        private void FormatoDistribuciones()
        {
            var resultColumn = 7;
            var titleRow = 12;
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                Microsoft.Office.Tools.Excel.Worksheet workSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                _sheetName01 = "SAO900-Distribuciones";
                excelSheet.Name = _sheetName01;

                _cells.GetCell(1, 1).Value = "CERREJÓN";
                _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(1, 1, 1, 2);
                _cells.GetCell("B1").Value = "Distribuciones";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells(2, 1, 7, 2);

                _cells.GetCell(1, 4).Value = "DISTRITO";
                _cells.GetCell(1, 4).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 5).Value = "AUTORIZADOR";
                _cells.MergeCells(3, 5, 7, 5);
                _cells.GetCell(1, 5).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 6).Value = "USUARIO";
                _cells.MergeCells(3, 6, 7, 6);
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 7).Value = "Num Transaccion";
                _cells.GetCell(1, 7).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 8).Value = "Vr Dolraes";
                _cells.GetCell(1, 8).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 9).Value = "Vr pesos";
                _cells.GetCell(1, 9).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, 10).Value = "XCentro de Costos";
                _cells.GetCell(1, 10).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetRange(2, 2, 2, 10).NumberFormat = "@";

                _cells.GetCell(1, titleRow).Value = "CCOSTOS";
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = "PROJ/WO";
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = "IND";
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, titleRow).Value = "DOLARES";
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = "PESOS";
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, titleRow).Value = "RAZON DEL CAMBIO ";
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "employeeRange");
                employeeRange.Change += changesEmployeeRange_Change;

                var employeeRange2 = workSheet.Controls.AddNamedRange(workSheet.Range["B6"], "employeeRange2");
                employeeRange2.Change += changesEmployeeRange_Change;

                var transactionRange = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "transactionRange");
                transactionRange.Change += changesTransactionRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_FormatoDistribuciones_Error);
            }
        }

        private void changessupplierRange_Change(Range target)
        {
            GetSupplierName(target);
        }

        private void changesdocumentRangeRange_Change(Range target)
        {
            GetDocument(target);
        }

        private void changesEmployeeRange_Change(Range target)
        {
            GetEmployeeName(target);
        }

        private void changesTransactionRange_Change(Range target)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper();
            var transactionNumber = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(target.Column, target.Row).Value);

            var transactionNo = new Transaction(_eFunctions, districtCode, transactionNumber);

            _cells.GetCell(target.Column + 1, target.Row).Value = transactionNo.Error;
            _cells.GetCell(target.Column, target.Row).Style =
                _cells.GetStyle(transactionNo.Error == null ? StyleConstants.Success : StyleConstants.Error);

            _cells.GetCell(2, 8).Value = Convert.ToDecimal(transactionNo.TranAmount);
            _cells.GetCell(2, 9).Value = Convert.ToDecimal(transactionNo.TranAmountS);
            _cells.GetCell(2, 10).Value = transactionNo.Account;
        }

        private void GetDocument(Range target)
        {
            var document = (_cells.GetEmptyIfNull(_cells.GetCell(target.Column, target.Row).Value)).ToUpper();
            string documentType = (_cells.GetEmptyIfNull(_cells.GetCell(target.Column, target.Row - 1).Value)).ToUpper();
            var supplierNo = (_cells.GetEmptyIfNull(_cells.GetCell(target.Column, target.Row - 2).Value)).ToUpper();

            if (string.IsNullOrEmpty(documentType))
            {
                _cells.GetCell(target.Column, target.Row - 1).ClearComments();
                _cells.GetCell(target.Column, target.Row - 1).AddComment("Digite el tipo de documento");
                _cells.GetCell(target.Column, target.Row - 1).Style = _cells.GetStyle(StyleConstants.Error);
                return;
            }

            if (string.IsNullOrEmpty(document)) return;

            string sqlQuery;

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            switch (documentType)
            {
                case "CO":
                    sqlQuery = Queries.GetContractNameDesc(document, _eFunctions.DbReference, _eFunctions.DbLink);
                    var drContract = _eFunctions.GetQueryResult(sqlQuery);
                    if (drContract != null && !drContract.IsClosed)
                    {
                        while (drContract.Read())
                        {
                            _cells.GetCell(target.Column + 1, target.Row).Value = drContract["CONTRACT_DESC"].ToString();
                            _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Success);
                            _cells.GetCell(target.Column, target.Row - 1).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                    }
                    else
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = "El contrato No Existe";
                        _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    break;
                case "PO":
                    sqlQuery = Queries.GetPurchaseOrder(document, supplierNo, _eFunctions.DbReference, _eFunctions.DbLink);
                    var drPurchaseOrder = _eFunctions.GetQueryResult(sqlQuery);
                    if (drPurchaseOrder != null && !drPurchaseOrder.IsClosed)
                    {
                        while (drPurchaseOrder.Read())
                        {
                            _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Success);
                            _cells.GetCell(target.Column, target.Row - 1).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                    }
                    else
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = "Purchase Order no Existe";
                        _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    break;
                case "IN":
                    //
                    break;
            }
        }

        private void GetSupplierName(Range target)
        {
            try
            {
                var supplierId = (_cells.GetEmptyIfNull(_cells.GetCell(target.Column, target.Row).Value)).ToUpper();
                var districtCode =
                    (_cells.GetEmptyIfNull(_cells.GetCell(target.Column, target.Row - 1).Value)).ToUpper();

                _cells.GetCell(target.Column, target.Row - 1).ClearFormats();

                if (string.IsNullOrEmpty(districtCode))
                {
                    _cells.GetCell(target.Column, target.Row - 1).Value += "Digite el Distrito";
                    _cells.GetCell(target.Column, target.Row - 1).Style = _cells.GetStyle(StyleConstants.Error);
                    return;
                }

                if (string.IsNullOrEmpty(supplierId)) return;

                var sqlQuery = Queries.GetSupplierName(districtCode, supplierId, _eFunctions.DbReference,
                    _eFunctions.DbLink);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var drSupplierName = _eFunctions.GetQueryResult(sqlQuery);

                if (drSupplierName != null && !drSupplierName.IsClosed && drSupplierName.HasRows)
                {
                    while (drSupplierName.Read())
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = drSupplierName["SUPPLIER_NAME"].ToString();
                        _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Success);
                        _cells.GetCell(target.Column, target.Row - 1).Style = _cells.GetStyle(StyleConstants.Success);
                    }
                }
                else
                {
                    _cells.GetCell(target.Column + 1, target.Row).Value =
                        SaoResources.RibbonEllipse_GetSupplierName_DoesntExist;
                    _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
                }
            }
            catch (Exception error)
            {
                _cells.GetCell(target.Column + 1, target.Row).Value = error.Message;
                _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
            }
        }

        private void GetEmployeeName(Range target)
        {
            try
            {
                var employeeId =
                    (_cells.GetNullOrTrimmedValue(_cells.GetCell(target.Column, target.Row).Value)).ToUpper();

                if (string.IsNullOrEmpty(employeeId)) return;

                var sqlQuery = Queries.GetEmployeeName(employeeId, _eFunctions.DbReference, _eFunctions.DbLink);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var drEmployeeName = _eFunctions.GetQueryResult(sqlQuery);

                if (drEmployeeName != null && !drEmployeeName.IsClosed && drEmployeeName.HasRows)
                {
                    while (drEmployeeName.Read())
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = drEmployeeName["NOMBRE"].ToString();
                    }
                }
                else
                {
                    _cells.GetCell(target.Column + 1, target.Row).Value =
                        SaoResources.RibbonEllipse_GetEmployeeName_DoesntExist;
                    _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
                }
            }
            catch (Exception error)
            {
                _cells.GetCell(target.Column + 1, target.Row).Value = error.Message;
                _cells.GetCell(target.Column, target.Row).Style = _cells.GetStyle(StyleConstants.Error);
            }
        }

        private void ValidarDatos()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.ActiveWorkbook;
                Worksheet excelSheet = excelBook.ActiveSheet;

                var resultColumn = ResultColumn;
                var titleRow = 9;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.GetRange(resultColumn, titleRow + 1, resultColumn, MaxRows).ClearContents();
                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).ClearComments();
                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, titleRow + 1, resultColumn, MaxRows).NumberFormat = "@";

                switch (excelSheet.Name)
                {
                    case "SAO900-Reclasificaciones":
                        ValidarReclasificaciones();
                        break;
                    case "SAO900-Modificaciones":
                        ValidarReclasificaciones();
                        break;
                    case "SAO900-Causaciones":
                        ValidarCausaciones();
                        break;
                    case "SAO900-Distribuciones":
                        ValidarDistribuciones();
                        break;
                }
                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ValidarDatos_Error);
            }
        }

        private void ExportarDatos()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var excelBook = _excelApp.ActiveWorkbook;
                Worksheet excelSheet = excelBook.ActiveSheet;

                switch (excelSheet.Name)
                {
                    case "SAO900-Reclasificaciones":
                        ExportarReclasificaciones();
                        break;
                    case "SAO900-Modificaciones":
                        ExportarModificaciones();
                        break;
                    case "SAO900-Causaciones":
                        ExportarCausaciones();
                        break;
                    case "SAO900-Distribuciones":
                        ExportarDistribuciones();
                        break;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ExportarDatos_Error);
            }
        }

        private void ExportarDistribuciones()
        {
            try
            {
                var distribuciones = new List<Distribuciones>();
                var titleRow = TitleRow01;
                var currentRow = titleRow + 1;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {
                    var itemDistribuciones = new Distribuciones
                    {
                        Action = "A",
                        Autorizador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value),
                        Distrito = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper(),
                        NumeroTransaccion = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value),
                        Centro = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 10).Value),
                        Dolares = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        Pesos = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        CentroDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        Equipo = "",
                        ProyectoOrdenDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        IndicadorDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        Razon = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value)
                    };

                    distribuciones.Add(itemDistribuciones);
                    currentRow++;
                }

                var outputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = '\'',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };

                var cc = new CsvContext();

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "D_" + _cells.GetEmptyIfNull(_cells.GetCell(2, 6).Value) + "_MMDD_CC.csv",
                    Filter = @"Archivos CSV|*.csv",
                    InitialDirectory = @"C:\\",
                    Title = @"Programa de Lectura"
                };

                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;

                cc.Write(distribuciones, saveFileDialog.FileName, outputFileDescription);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ExportarDistribuciones_Error);
            }
        }

        private void ExportarCausaciones()
        {
            try
            {
                var causaciones = new List<Causaciones>();
                var titleRow = TitleRow01;
                var currentRow = titleRow + 1;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {
                    var itemCausaciones = new Causaciones
                    {
                        Action = "A",
                        Item = (currentRow - titleRow).ToString(),
                        Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value),
                        TipoDocumento = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value),
                        NumeroDocumento = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value),
                        FechaSolicitud = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 8).Value),
                        Moneda = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 9).Value),
                        ValorTotal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 10).Value),
                        SolicitadorPor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 11).Value),
                        Distrito = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper(),
                        Centro = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        Equipo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        ProyectoOrden = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Ind = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        Valor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value)
                    };

                    causaciones.Add(itemCausaciones);
                    currentRow++;
                }

                var outputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = '\'',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };

                var cc = new CsvContext();

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "C_" + _cells.GetEmptyIfNull(_cells.GetCell(2, 11).Value) + "_MMDD_CC.csv",
                    Filter = @"Archivos CSV|*.csv",
                    InitialDirectory = @"C:\\",
                    Title = @"Programa de Lectura"
                };

                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;

                cc.Write(causaciones, saveFileDialog.FileName, outputFileDescription);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ExportarCausaciones_Error);
            }
        }

        private void ExportarModificaciones()
        {
            try
            {
                var reclasificaciones = new List<Reclasificaciones>();
                var titleRow = TitleRow01;
                var currentRow = titleRow + 1;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {
                    var itemReclasificacion = new Reclasificaciones
                    {
                        Action = "A",
                        Autorizador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value),
                        Distrito = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper(),
                        NumTransaccion = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        Centro = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        ProyectoOrden = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        Indicador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Dolares = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        Pesos = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        CentroDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),

                        Equipo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        ProyectoOrdenDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        //                        Equipo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        //                        ProyectoOrdenDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        IndicadorDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        RazonCambio = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value)
                    };

                    reclasificaciones.Add(itemReclasificacion);
                    currentRow++;
                }

                var outputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = '\'',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };

                var cc = new CsvContext();

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "M_" + _cells.GetEmptyIfNull(_cells.GetCell(2, 7).Value) + "_MMDD_CC.csv",
                    Filter = @"Archivos CSV|*.csv",
                    InitialDirectory = @"C:\\",
                    Title = @"Programa de Lectura"
                };

                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;

                cc.Write(reclasificaciones, saveFileDialog.FileName, outputFileDescription);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ExportarReclasificaciones_Error);
            }
        }

        private void ExportarReclasificaciones()
        {
            try
            {
                var reclasificaciones = new List<Reclasificaciones>();
                var titleRow = TitleRow01;
                var currentRow = titleRow + 1;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {
                    var itemReclasificacion = new Reclasificaciones
                    {
                        Action = "A",
                        Autorizador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value),
                        Distrito = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper(),
                        NumTransaccion = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        Centro = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        ProyectoOrden = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        Indicador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Dolares = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        Pesos = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        CentroDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        Equipo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        ProyectoOrdenDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        IndicadorDestino = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        RazonCambio = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value)
                    };

                    reclasificaciones.Add(itemReclasificacion);
                    currentRow++;
                }

                var outputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = '\'',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };

                var cc = new CsvContext();

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "R_" + _cells.GetEmptyIfNull(_cells.GetCell(2, 7).Value) + "_MMDD_CC.csv",
                    Filter = @"Archivos CSV|*.csv",
                    InitialDirectory = @"C:\\",
                    Title = @"Programa de Lectura"
                };

                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;

                cc.Write(reclasificaciones, saveFileDialog.FileName, outputFileDescription);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ExportarReclasificaciones_Error);
            }
        }

        private void ValidarDistribuciones()
        {
            var titleRow = TitleRow01;
            var resultColumn = ResultColumn;
            var currentRow = titleRow + 1;
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            _cells.GetCell(2, 4).ClearComments();
            if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value) == null)
            {
                _cells.GetCell(2, 4).AddComment("Digite el Distrito");
                _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Error);
                return;
            }
            _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Success);
            _cells.GetCell(2, 5).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);
            _cells.GetCell(2, 6).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);
            _cells.GetCell(2, 7).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);
            _cells.GetCell(2, 8).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 8).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);
            _cells.GetCell(2, 9).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 9).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);
            _cells.GetCell(2, 10).Style =
                _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 10).Value) == null
                    ? StyleConstants.Error
                    : StyleConstants.Success);

            _cells.GetCell(2, 8).Value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 8).Value) != null
                ? Convert.ToDecimal(_cells.GetCell(2, 8).Value)
                : _cells.GetCell(2, 4).Value;
            _cells.GetCell(2, 9).Value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 9).Value) != null
                ? Convert.ToDecimal(_cells.GetCell(2, 9).Value)
                : _cells.GetCell(2, 5).Value;
            _cells.GetRange(2, 8, 2, 9).Style = "Currency";
            _cells.GetRange(2, 8, 2, 9).NumberFormat = "$#,###.00";


            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                var district = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper();
                var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                var projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) ?? "";
                var projectInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) ?? "";

                _cells.GetCell(4, currentRow).Style =
                    _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value) == null
                        ? StyleConstants.Error
                        : StyleConstants.Success);
                _cells.GetCell(5, currentRow).Style =
                    _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value) == null
                        ? StyleConstants.Error
                        : StyleConstants.Success);
                _cells.GetCell(6, currentRow).Style =
                    _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value) == null
                        ? StyleConstants.Error
                        : StyleConstants.Success);

                var accountCode = new AccountCode(_eFunctions, district, account);
                _cells.GetCell(resultColumn, currentRow).Value += accountCode.Error;
                _cells.GetCell(1, currentRow).Style =
                    _cells.GetStyle((accountCode.Error == null && accountCode.ActiveStatus == "A")
                        ? StyleConstants.Success
                        : StyleConstants.Error);

                if (accountCode.Error == null)
                {
                    //Valido si el proyecto es mandatorio
                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
                        _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    //                    else if (accountCode.ProjectEntriInd == "O" && (projectNo != "" || projectInd != ""))
                    //                    {
                    //                        _cells.GetCell(_resultColumn, currentRow).Value += " Proyecto No Requerido";
                    //                        _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    //                    }
                    else
                    {
                        //valido si la orden es mandatoria
                        if (accountCode.WorkOrderEntryInd == "M" && (projectNo == "" || projectInd == "P"))
                        {
                            _cells.GetCell(resultColumn, currentRow).Value += " Numero de Orden Requerido";
                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        }
                        //                        else if (accountCode.WorkOrderEntryInd == "O" && (projectNo != "" || projectInd != ""))
                        //                        {
                        //                            _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden No Requerido";
                        //                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        //                        }
                        else
                        {
                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                    }
                }

                //valido si se necesita Subledger
                if (accountCode.SubLedgerInd == "M" && !projectNo.Contains(";"))
                {
                    _cells.GetCell(resultColumn, currentRow).Value += " Subledger Requerido";
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }

                currentRow++;
            }

            _cells.GetRange(4, titleRow + 1, 5, MaxRows).Style = "Currency";
            _cells.GetRange(4, titleRow + 1, 5, MaxRows).NumberFormat = "$#,###.00";
        }

        private void ValidarCausaciones()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var titleRow = TitleRow01;
            var resultColumn = ResultColumn;

            var currentRow = titleRow + 1;

            _cells.GetCell(2, 4).ClearComments();
            if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value) == null)
            {
                _cells.GetCell(2, 4).AddComment("Digite el Distrito");
                _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Error);
                return;
            }

            _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Success);
            _cells.GetCell(2, 5).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 6).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 7).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 8).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 8).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 9).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 9).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 10).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 10).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 11).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 11).Value) == null? StyleConstants.Error: StyleConstants.Success);
            _cells.GetCell(2, 10).Value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 10).Value) != null? Convert.ToDecimal(_cells.GetCell(2, 10).Value): _cells.GetCell(2, 10).Value;
            _cells.GetCell(2, 10).Style = "Currency";
            _cells.GetCell(2, 10).NumberFormat = "$#,###.00";
            
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                var district = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper();
                var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                var projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) ?? "";
                var projectInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) ?? "";

                _cells.GetCell(5, currentRow).Style =
                    _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value) == null? StyleConstants.Error: StyleConstants.Success);

                var accountCode = new AccountCode(_eFunctions, district, account);
                _cells.GetCell(resultColumn, currentRow).Value += accountCode.Error;
                _cells.GetCell(1, currentRow).Style =_cells.GetStyle((accountCode.Error != null && accountCode.ActiveStatus != "A")? StyleConstants.Error: StyleConstants.Success);

                if (accountCode.Error == null)
                {
                    //Valido si el proyecto es mandatorio
                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
                        _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    //                    else if (accountCode.ProjectEntriInd == "O" && (projectNo != "" || projectInd != ""))
                    //                    {
                    //                        _cells.GetCell(_resultColumn, currentRow).Value += " Proyecto No Requerido";
                    //                        _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    //                    }
                    else
                    {
                        //valido si la orden es mandatoria
                        if (accountCode.WorkOrderEntryInd == "M" && (projectNo == "" || projectInd == "P"))
                        {
                            _cells.GetCell(resultColumn, currentRow).Value += " Numero de Orden Requerido";
                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        }
                        //                        else if (accountCode.WorkOrderEntryInd == "O" && (projectNo != "" || projectInd != ""))
                        //                        {
                        //                            _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden No Requerido";
                        //                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        //                        }
                        else
                        {
                            _cells.GetRange(2, currentRow, 3, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                    }
                }
                //valido si se necesita Subledger
                if (accountCode.SubLedgerInd == "M" && !projectNo.Contains(";"))
                {
                    _cells.GetCell(resultColumn, currentRow).Value += " Subledger Requerido";
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }

                currentRow++;
            }
            _cells.GetRange(5, titleRow + 1, 5, MaxRows).Style = "Currency";
            _cells.GetRange(5, titleRow + 1, 5, MaxRows).NumberFormat = "$#,###.00";
        }

        private void ValidarReclasificaciones()
        {
            var titleRow = TitleRow01;
            var resultColumn = ResultColumn;
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            try
            {
                const int transactionColumn = 1;
                ValidarTransaccion904(transactionColumn);
                
                var currentRow = titleRow + 1;
                _cells.GetCell(2, 4).ClearComments();
                if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value) == null)
                {
                    _cells.GetCell(2, 4).AddComment("Digite el Distrito");
                    _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Error);
                    return;
                }

                _cells.GetCell(2, 4).Style = _cells.GetStyle(StyleConstants.Success);
                _cells.GetCell(2, 5).Style = _cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value) == null ? StyleConstants.Error : StyleConstants.Success);
                _cells.GetCell(2, 6).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value) == null? StyleConstants.Error: StyleConstants.Success);
                _cells.GetCell(2, 7).Style =_cells.GetStyle(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value) == null? StyleConstants.Error: StyleConstants.Success);

                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {

                    _cells.GetCell(1, currentRow).Select();

                    //valida la transaccion contable

                    var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper();
                    var transaction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                    var projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value) ?? "";
                    var projectInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value) ?? "";

                    var transactionNo = new Transaction(_eFunctions, districtCode, transaction);
                    _cells.GetCell(resultColumn, currentRow).Value += " " + transactionNo.Error;
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(transactionNo.Error == null ? StyleConstants.Success : StyleConstants.Error);

                    _cells.GetCell(2, currentRow).Value = transactionNo.FullPeriod;
                    _cells.GetCell(3, currentRow).Value = transactionNo.Account;
                    _cells.GetCell(4, currentRow).Value = transactionNo.ProjectNo;
                    _cells.GetCell(5, currentRow).Value = transactionNo.Ind;
                    _cells.GetCell(10, currentRow).Value = Convert.ToDecimal(transactionNo.TranAmount);
                    _cells.GetCell(11, currentRow).Value = Convert.ToDecimal(transactionNo.TranAmountS);

                    //Valido el Centro de Costo Destino
                    var accountCode = new AccountCode(_eFunctions, districtCode, account);
                    _cells.GetCell(resultColumn, currentRow).Value += accountCode.Error;
                    _cells.GetCell(6, currentRow).Style =
                        _cells.GetStyle(accountCode.Error != null && accountCode.ActiveStatus != "A"
                            ? StyleConstants.Error
                            : StyleConstants.Success);

                    //Valido si el proyecto es mandatorio
                    if (accountCode.Error == null)
                    {
                        if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                        {
                            _cells.GetCell(resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
                            _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                            _cells.GetRange(8, currentRow, 9, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        }
                        //                        else if (accountCode.ProjectEntriInd == "O" && (projectNo != "" || projectInd != ""))
                        //                        {
                        //                            _cells.GetCell(_resultColumn, currentRow).Value += " Proyecto No Requerido";
                        //                            _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        //                            _cells.GetRange(8, currentRow, 9, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                        //                        }
                        else
                        {
                            //valido si la orden es mandatoria
                            if (accountCode.WorkOrderEntryInd == "M" && (projectNo == "" || projectInd == "P"))
                            {
                                _cells.GetCell(resultColumn, currentRow).Value += " Numero de Orden Requerido";
                                _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                                _cells.GetRange(8, currentRow, 9, currentRow).Style =
                                    _cells.GetStyle(StyleConstants.Error);
                            }
                            //                            else if (accountCode.WorkOrderEntryInd == "O" && (projectNo != "" || projectInd != ""))
                            //                            {
                            //                                _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden No Requerido";
                            //                                _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                            //                                _cells.GetRange(8, currentRow, 9, currentRow).Style =
                            //                                    _cells.GetStyle(StyleConstants.Error);
                            //                            }
                            else
                            {
                                _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                                _cells.GetRange(8, currentRow, 9, currentRow).Style =
                                    _cells.GetStyle(StyleConstants.Success);
                            }
                        }
                    }

                    //valido si se necesita Subledger
                    if (accountCode.SubLedgerInd == "M" && !projectNo.Contains(";"))
                    {
                        _cells.GetCell(resultColumn, currentRow).Value += " Subledger Requerido";
                        _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    currentRow++;
                }

                _cells.GetRange(10, titleRow + 1, 11, MaxRows).Style = "Currency";
                _cells.GetRange(10, titleRow + 1, 11, MaxRows).NumberFormat = "$#,###.00";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, SaoResources.RibbonEllipse_ValidarReclasificaciones_Error);
            }
        }

        private void ValidarTransaccion904(int transactionColumn)
        {
            var titleRow = TitleRow01;
            var resultColumn = ResultColumn;
            var currentRow = titleRow + 1;

            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

            var opSheet = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(transactionColumn, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    //valida la transaccion contable

                    var districtCode = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct).ToUpper();
                    var transaction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(transactionColumn, currentRow).Value);


                    _eFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO904");

                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM904A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("DSTRCT_CODE1I", districtCode);
                        arrayFields.Add("TRAN_ID1I", transaction);
                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                    }
                    if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, currentRow).Value = replySheet.message;
                    }

                    _eFunctions.RevertOperation(opSheet, proxySheet);
                }
                catch (Exception error)
                {
                    _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, currentRow).Value = error.Message;
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
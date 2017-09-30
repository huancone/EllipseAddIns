using System;
using System.Collections.Generic;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseSAO900AddIn.Properties;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseSAO900AddIn
{
    public partial class RibbonEllipse
    {
        private const int MaxRows = 1000;
        private static readonly EllipseFunctions EFunctions = new EllipseFunctions();
        private static ExcelStyleCells _cells;
        private static Application _excelApp;
        private int _resultColumn = 12;
        private string _sheetName01;
        private static int _tittleRow = 9;
        private static readonly FormAuthenticate FrmAuth = new FormAuthenticate();

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

                _cells.GetCell(1, _tittleRow).Value = "NUM_TRANSACCION";
                _cells.GetCell(1, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, _tittleRow).Value = "PERIODO";
                _cells.GetCell(2, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, _tittleRow).Value = "CCOSTOS / API";
                _cells.GetCell(3, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, _tittleRow).Value = "PROJ/WO";
                _cells.GetCell(4, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, _tittleRow).Value = "IND";
                _cells.GetCell(5, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, _tittleRow).Value = "CCOSTOS_DESTINO";
                _cells.GetCell(6, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, _tittleRow).Value = "EQUIPO";
                _cells.GetCell(7, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, _tittleRow).Value = "PROJ/WO_DESTINO";
                _cells.GetCell(8, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, _tittleRow).Value = "IND_DESTINO";
                _cells.GetCell(9, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, _tittleRow).Value = "DOLARES";
                _cells.GetCell(10, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(11, _tittleRow).Value = "PESOS";
                _cells.GetCell(11, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_resultColumn, _tittleRow).Value = "Resultado";
                _cells.GetCell(_resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange1 = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "employeeRange1");
                employeeRange1.Change += changesEmployeeRange_Change;

                var employeeRange2 = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "employeeRange2");
                employeeRange2.Change += changesEmployeeRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, Resources.RibbonEllipse_FormatoReclasificaciones_Error);
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

                _cells.GetCell(1, _tittleRow).Value = "NUM_TRANSACCION";
                _cells.GetCell(1, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, _tittleRow).Value = "PERIODO";
                _cells.GetCell(2, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, _tittleRow).Value = "CCOSTOS / API";
                _cells.GetCell(3, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, _tittleRow).Value = "PROJ/WO";
                _cells.GetCell(4, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, _tittleRow).Value = "IND";
                _cells.GetCell(5, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, _tittleRow).Value = "CCOSTOS_DESTINO";
                _cells.GetCell(6, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(7, _tittleRow).Value = "EQUIPO";
                _cells.GetCell(7, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, _tittleRow).Value = "PROJ/WO_DESTINO";
                _cells.GetCell(8, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, _tittleRow).Value = "IND_DESTINO";
                _cells.GetCell(9, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, _tittleRow).Value = "DOLARES";
                _cells.GetCell(10, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(11, _tittleRow).Value = "PESOS";
                _cells.GetCell(11, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_resultColumn, _tittleRow).Value = "Resultado";
                _cells.GetCell(_resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

                var employeeRange1 = workSheet.Controls.AddNamedRange(workSheet.Range["B5"], "employeeRange1");
                employeeRange1.Change += changesEmployeeRange_Change;

                var employeeRange2 = workSheet.Controls.AddNamedRange(workSheet.Range["B7"], "employeeRange2");
                employeeRange2.Change += changesEmployeeRange_Change;

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, Resources.RibbonEllipse_FormatoReclasificaciones_Error);
            }
        }

        private void FormatoCausaciones()
        {
            _resultColumn = 6;
            _tittleRow = 13;
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

                _cells.GetCell(1, _tittleRow).Value = "CCOSTOS/DETALLE";
                _cells.GetCell(1, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, _tittleRow).Value = "PROYECTO / WO";
                _cells.GetCell(2, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, _tittleRow).Value = "P/W";
                _cells.GetCell(3, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, _tittleRow).Value = "EQUIPO";
                _cells.GetCell(4, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, _tittleRow).Value = "VALOR (PES o USD)";
                _cells.GetCell(5, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_resultColumn, _tittleRow).Value = "Resultado";
                _cells.GetCell(_resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_FormatoCausaciones_Error);
            }
        }

        private void FormatoDistribuciones()
        {
            _resultColumn = 7;
            _tittleRow = 12;
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

                _cells.GetCell(1, _tittleRow).Value = "CCOSTOS";
                _cells.GetCell(1, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, _tittleRow).Value = "PROJ/WO";
                _cells.GetCell(2, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, _tittleRow).Value = "IND";
                _cells.GetCell(3, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, _tittleRow).Value = "DOLARES";
                _cells.GetCell(4, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, _tittleRow).Value = "PESOS";
                _cells.GetCell(5, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, _tittleRow).Value = "RAZON DEL CAMBIO ";
                _cells.GetCell(6, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_resultColumn, _tittleRow).Value = "Resultado";
                _cells.GetCell(_resultColumn, _tittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_FormatoDistribuciones_Error);
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
            var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value).ToUpper();
            var transactionNumber = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(target.Column, target.Row).Value);

            var transactionNo = new Transaction(districtCode, transactionNumber);

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

            EFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            switch (documentType)
            {
                case "CO":
                    sqlQuery = Queries.GetContractNameDesc(document, EFunctions.dbReference, EFunctions.dbLink);
                    var drContract = EFunctions.GetQueryResult(sqlQuery);
                    if (drContract != null && !drContract.IsClosed && drContract.HasRows)
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
                    sqlQuery = Queries.GetPurchaseOrder(document, supplierNo, EFunctions.dbReference, EFunctions.dbLink);
                    var drPurchaseOrder = EFunctions.GetQueryResult(sqlQuery);
                    if (drPurchaseOrder != null && !drPurchaseOrder.IsClosed && drPurchaseOrder.HasRows)
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

                var sqlQuery = Queries.GetSupplierName(districtCode, supplierId, EFunctions.dbReference,
                    EFunctions.dbLink);

                EFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var drSupplierName = EFunctions.GetQueryResult(sqlQuery);

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
                        Resources.RibbonEllipse_GetSupplierName_DoesntExist;
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

                var sqlQuery = Queries.GetEmployeeName(employeeId, EFunctions.dbReference, EFunctions.dbLink);

                EFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var drEmployeeName = EFunctions.GetQueryResult(sqlQuery);

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
                        Resources.RibbonEllipse_GetEmployeeName_DoesntExist;
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

                EFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                _cells.GetRange(_resultColumn, _tittleRow + 1, _resultColumn, MaxRows).ClearContents();
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).ClearComments();
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, _tittleRow + 1, _resultColumn, MaxRows).NumberFormat = "@";

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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ValidarDatos_Error);
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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ExportarDatos_Error);
            }
        }

        private void ExportarDistribuciones()
        {
            try
            {
                var distribuciones = new List<Distribuciones>();
                var currentRow = _tittleRow + 1;
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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ExportarDistribuciones_Error);
            }
        }

        private void ExportarCausaciones()
        {
            try
            {
                var causaciones = new List<Causaciones>();
                var currentRow = _tittleRow + 1;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
                {
                    var itemCausaciones = new Causaciones
                    {
                        Action = "A",
                        Item = (currentRow - _tittleRow).ToString(),
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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ExportarCausaciones_Error);
            }
        }

        private void ExportarModificaciones()
        {
            try
            {
                var reclasificaciones = new List<Reclasificaciones>();
                var currentRow = _tittleRow + 1;
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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ExportarReclasificaciones_Error);
            }
        }

        private void ExportarReclasificaciones()
        {
            try
            {
                var reclasificaciones = new List<Reclasificaciones>();
                var currentRow = _tittleRow + 1;
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
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ExportarReclasificaciones_Error);
            }
        }

        private void ValidarDistribuciones()
        {
            var currentRow = _tittleRow + 1;

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

                var accountCode = new AccountCode(district, account);
                _cells.GetCell(_resultColumn, currentRow).Value += accountCode.Error;
                _cells.GetCell(1, currentRow).Style =
                    _cells.GetStyle((accountCode.Error == null && accountCode.ActiveStatus == "A")
                        ? StyleConstants.Success
                        : StyleConstants.Error);

                if (accountCode.Error == null)
                {
                    //Valido si el proyecto es mandatorio
                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
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
                            _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden Requerido";
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
                    _cells.GetCell(_resultColumn, currentRow).Value += " Subledger Requerido";
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }

                currentRow++;
            }

            _cells.GetRange(4, _tittleRow + 1, 5, MaxRows).Style = "Currency";
            _cells.GetRange(4, _tittleRow + 1, 5, MaxRows).NumberFormat = "$#,###.00";
        }

        private void ValidarCausaciones()
        {
            var currentRow = _tittleRow + 1;

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

                var accountCode = new AccountCode(district, account);
                _cells.GetCell(_resultColumn, currentRow).Value += accountCode.Error;
                _cells.GetCell(1, currentRow).Style =_cells.GetStyle((accountCode.Error != null && accountCode.ActiveStatus != "A")? StyleConstants.Error: StyleConstants.Success);

                if (accountCode.Error == null)
                {
                    //Valido si el proyecto es mandatorio
                    if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
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
                            _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden Requerido";
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
                    _cells.GetCell(_resultColumn, currentRow).Value += " Subledger Requerido";
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                }

                currentRow++;
            }
            _cells.GetRange(5, _tittleRow + 1, 5, MaxRows).Style = "Currency";
            _cells.GetRange(5, _tittleRow + 1, 5, MaxRows).NumberFormat = "$#,###.00";
        }

        private void ValidarReclasificaciones()
        {
            try
            {
                const int transactionColumn = 1;
                ValidarTransaccion904(transactionColumn);
                
                var currentRow = _tittleRow + 1;
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

                    var transactionNo = new Transaction(districtCode, transaction);
                    _cells.GetCell(_resultColumn, currentRow).Value += " " + transactionNo.Error;
                    _cells.GetCell(1, currentRow).Style = _cells.GetStyle(transactionNo.Error == null ? StyleConstants.Success : StyleConstants.Error);

                    _cells.GetCell(2, currentRow).Value = transactionNo.FullPeriod;
                    _cells.GetCell(3, currentRow).Value = transactionNo.Account;
                    _cells.GetCell(4, currentRow).Value = transactionNo.ProjectNo;
                    _cells.GetCell(5, currentRow).Value = transactionNo.Ind;
                    _cells.GetCell(10, currentRow).Value = Convert.ToDecimal(transactionNo.TranAmount);
                    _cells.GetCell(11, currentRow).Value = Convert.ToDecimal(transactionNo.TranAmountS);

                    //Valido el Centro de Costo Destino
                    var accountCode = new AccountCode(districtCode, account);
                    _cells.GetCell(_resultColumn, currentRow).Value += accountCode.Error;
                    _cells.GetCell(6, currentRow).Style =
                        _cells.GetStyle(accountCode.Error != null && accountCode.ActiveStatus != "A"
                            ? StyleConstants.Error
                            : StyleConstants.Success);

                    //Valido si el proyecto es mandatorio
                    if (accountCode.Error == null)
                    {
                        if (accountCode.ProjectEntriInd == "M" && (projectNo == "" || projectInd == "W"))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Proyecto Requerido";
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
                                _cells.GetCell(_resultColumn, currentRow).Value += " Numero de Orden Requerido";
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
                        _cells.GetCell(_resultColumn, currentRow).Value += " Subledger Requerido";
                        _cells.GetCell(6, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    currentRow++;
                }

                _cells.GetRange(10, _tittleRow + 1, 11, MaxRows).Style = "Currency";
                _cells.GetRange(10, _tittleRow + 1, 11, MaxRows).NumberFormat = "$#,###.00";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, Resources.RibbonEllipse_ValidarReclasificaciones_Error);
            }
        }

        private void ValidarTransaccion904(int transactionColumn)
        {
            var currentRow = _tittleRow + 1;

            FrmAuth.StartPosition = FormStartPosition.CenterScreen;
            FrmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            if (FrmAuth.ShowDialog() != DialogResult.OK) return;
            ClientConversation.authenticate(FrmAuth.EllipseUser, FrmAuth.EllipsePswd);

            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();

            proxySheet.Url = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

            var opSheet = new screen.OperationContext
            {
                district = FrmAuth.EllipseDsct,
                position = FrmAuth.EllipsePost,
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

                    var districtCode = _cells.GetNullIfTrimmedEmpty(FrmAuth.EllipseDsct).ToUpper();
                    var transaction = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(transactionColumn, currentRow).Value);


                    EFunctions.RevertOperation(opSheet, proxySheet);
                    var replySheet = proxySheet.executeScreen(opSheet, "MSO904");

                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
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
                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }

                    EFunctions.RevertOperation(opSheet, proxySheet);
                }
                catch (Exception error)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, currentRow).Value = error.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        /// <summary>
        ///     Crea un objeto con los datos de la transaccion contable, se contruye con la informacion del distrito y el
        ///     transaction group key, que busca esta llave en la tabla msf900
        /// </summary>
        private class Transaction
        {

            public Transaction(string districtCode, string transactionNo)
            {
                try
                {
                    if (string.IsNullOrEmpty(districtCode) || string.IsNullOrEmpty(transactionNo))
                    {
                        Error = "Transaccion Invalida";
                        return;
                    }
                    var sqlQuery = Queries.GetTransactionInfo(districtCode, transactionNo, EFunctions.dbReference,
                        EFunctions.dbLink);

                    var drTransactionNo = EFunctions.GetQueryResult(sqlQuery);

                    if (drTransactionNo != null && !drTransactionNo.IsClosed && drTransactionNo.HasRows)
                    {
                        while (drTransactionNo.Read())
                        {
                            FullPeriod = drTransactionNo["FULL_PERIOD"].ToString();
                            Account = drTransactionNo["ACCOUNT_CODE"].ToString();
                            ProjectNo = drTransactionNo["PROJECT_NO"].ToString();
                            Ind = drTransactionNo["IND"].ToString();
                            TranAmount = drTransactionNo["TRAN_AMOUNT"].ToString();
                            TranAmountS = drTransactionNo["TRAN_AMOUNT_S"].ToString();
                        }
                    }
                    else
                    {
                        Error = "La Transaccion no Existe";
                    }
                }
                catch (Exception error)
                {
                    Error = error.Message;
                }
            }

            public string FullPeriod { get; private set; }
            public string Account { get; private set; }
            public string ProjectNo { get; private set; }
            public string Ind { get; private set; }
            public string TranAmount { get; private set; }
            public string TranAmountS { get; private set; }
            public string Error { get; private set; }
        }

        /// <summary>
        ///     Crea un objeto con la informacion de la combinacion centro/detalle verificando el estado y las variables de control
        ///     de proyecto, orden y subledger.
        /// </summary>
        private class AccountCode
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

                    accountCode = accountCode.Contains(";")
                        ? accountCode.Substring(0, accountCode.IndexOf(";", StringComparison.Ordinal))
                        : accountCode;

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
                            Error = (drAccountCode["ACTIVE_STATUS"].ToString() == "I") ? "AccountCode Inactivo" : null;
                        }



                    }
                    else
                    {
                        Error = "Centro de Costos Destino No Valido";
                    }
                }
                catch (Exception error)
                {
                    Error = error.Message;
                }
            }

            public string Error { get; private set; }
            public string ActiveStatus { get; private set; }
            public string ProjectEntriInd { get; private set; }
            public string WorkOrderEntryInd { get; private set; }
            public string SubLedgerInd { get; private set; }
        }

        /// <summary>
        ///     Consultas SQL a las bases de datos de Ellipse 8
        /// </summary>
        private static class Queries
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

            public static string GetTransactionInfo(string districtCode, string numTransaction, string dbReference,
                string dbLink)
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

            public static string GetAccountCodeInfo(string districtCode, string accountCode, string dbReference,
                string dbLink)
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

            public static string GetSupplierName(string districtCode, string supplierId, string dbReference,
                string dbLink)
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
        }

        /// <summary>
        ///     Clase usada para exportar el documento csv
        /// </summary>
        private class Reclasificaciones
        {
            [CsvColumn(Name = "ACTION", FieldIndex = 1)]
            public string Action { get; set; }

            [CsvColumn(Name = "AUTORIZADOR", FieldIndex = 2)]
            public string Autorizador { get; set; }

            [CsvColumn(Name = "DISTRITO", FieldIndex = 3)]
            public string Distrito { get; set; }

            [CsvColumn(Name = "NUM_TRANSACCION", FieldIndex = 4)]
            public string NumTransaccion { get; set; }

            [CsvColumn(Name = "CCOSTOS", FieldIndex = 5)]
            public string Centro { get; set; }

            [CsvColumn(Name = "PROJ/WO", FieldIndex = 6)]
            public string ProyectoOrden { get; set; }

            [CsvColumn(Name = "IND", FieldIndex = 7)]
            public string Indicador { get; set; }

            [CsvColumn(Name = "DOLARES", FieldIndex = 8)]
            public string Dolares { get; set; }

            [CsvColumn(Name = "PESOS", FieldIndex = 9)]
            public string Pesos { get; set; }

            [CsvColumn(Name = "CCOSTOS_DESTINO", FieldIndex = 10)]
            public string CentroDestino { get; set; }

            [CsvColumn(Name = "EQUIPO", FieldIndex = 11)]
            public string Equipo { get; set; }

            [CsvColumn(Name = "PROJ/WO_DESTINO", FieldIndex = 12)]
            public string ProyectoOrdenDestino { get; set; }

            [CsvColumn(Name = "IND_DESTINO", FieldIndex = 13)]
            public string IndicadorDestino { get; set; }

            [CsvColumn(Name = "RAZON DEL CAMBIO", FieldIndex = 14)]
            public string RazonCambio { get; set; }
        }

        /// <summary>
        ///     Clase usada para exportar el documento csv
        /// </summary>
        private class Modificaciones
        {
            [CsvColumn(Name = "ACTION", FieldIndex = 1)]
            public string Action { get; set; }

            [CsvColumn(Name = "AUTORIZADOR", FieldIndex = 2)]
            public string Autorizador { get; set; }

            [CsvColumn(Name = "DISTRITO", FieldIndex = 3)]
            public string Distrito { get; set; }

            [CsvColumn(Name = "NUM_TRANSACCION", FieldIndex = 4)]
            public string NumTransaccion { get; set; }

            [CsvColumn(Name = "CCOSTOS", FieldIndex = 5)]
            public string Centro { get; set; }

            [CsvColumn(Name = "PROJ/WO", FieldIndex = 6)]
            public string ProyectoOrden { get; set; }

            [CsvColumn(Name = "IND", FieldIndex = 7)]
            public string Indicador { get; set; }

            [CsvColumn(Name = "DOLARES", FieldIndex = 8)]
            public string Dolares { get; set; }

            [CsvColumn(Name = "PESOS", FieldIndex = 9)]
            public string Pesos { get; set; }

            [CsvColumn(Name = "CCOSTOS_DESTINO", FieldIndex = 10)]
            public string CentroDestino { get; set; }

            [CsvColumn(Name = "EQUIPO", FieldIndex = 11)]
            public string Equipo { get; set; }

            [CsvColumn(Name = "PROJ/WO_DESTINO", FieldIndex = 12)]
            public string ProyectoOrdenDestino { get; set; }

            [CsvColumn(Name = "IND_DESTINO", FieldIndex = 13)]
            public string IndicadorDestino { get; set; }

            [CsvColumn(Name = "RAZON DEL CAMBIO", FieldIndex = 14)]
            public string RazonCambio { get; set; }
        }

        /// <summary>
        ///     Clase usada para exportar el documento csv
        /// </summary>
        private class Causaciones
        {
            [CsvColumn(Name = "ACCION", FieldIndex = 1)]
            public string Action { get; set; }

            [CsvColumn(Name = "ITEM", FieldIndex = 2)]
            public string Item { get; set; }

            [CsvColumn(Name = "SUPPLIER", FieldIndex = 3)]
            public string Supplier { get; set; }

            [CsvColumn(Name = "TIPO_DE_DOC", FieldIndex = 4)]
            public string TipoDocumento { get; set; }

            [CsvColumn(Name = "NUM_DE_DOC", FieldIndex = 5)]
            public string NumeroDocumento { get; set; }

            [CsvColumn(Name = "FECHA_SOLCT", FieldIndex = 6)]
            public string FechaSolicitud { get; set; }

            [CsvColumn(Name = "MONEDA", FieldIndex = 7)]
            public string Moneda { get; set; }

            [CsvColumn(Name = "VALOR_TOTAL", FieldIndex = 8)]
            public string ValorTotal { get; set; }

            [CsvColumn(Name = "SOLICITADO_POR", FieldIndex = 9)]
            public string SolicitadorPor { get; set; }

            [CsvColumn(Name = "DISTRITO", FieldIndex = 10)]
            public string Distrito { get; set; }

            [CsvColumn(Name = "C_COSTOS_DETALLE", FieldIndex = 11)]
            public string Centro { get; set; }

            [CsvColumn(Name = "EQUIPO", FieldIndex = 12)]
            public string Equipo { get; set; }

            [CsvColumn(Name = "PROYECTO_WO", FieldIndex = 13)]
            public string ProyectoOrden { get; set; }

            [CsvColumn(Name = "P_W", FieldIndex = 14)]
            public string Ind { get; set; }

            [CsvColumn(Name = "VALOR_PES_o_USD", FieldIndex = 15)]
            public string Valor { get; set; }
        }

        /// <summary>
        ///     Clase usada para exportar el documento csv
        /// </summary>
        private class Distribuciones
        {
            [CsvColumn(Name = "ACTION", FieldIndex = 1)]
            public string Action { get; set; }

            [CsvColumn(Name = "AUTORIZADOR", FieldIndex = 2)]
            public string Autorizador { get; set; }

            [CsvColumn(Name = "DISTRITO", FieldIndex = 3)]
            public string Distrito { get; set; }

            [CsvColumn(Name = "NUM_TRANSACCION", FieldIndex = 4)]
            public string NumeroTransaccion { get; set; }

            [CsvColumn(Name = "CCOSTOS", FieldIndex = 5)]
            public string Centro { get; set; }

            [CsvColumn(Name = "PROJ/WO", FieldIndex = 6)]
            public string ProyectoOrden { get; set; }

            [CsvColumn(Name = "IND", FieldIndex = 7)]
            public string Indicador { get; set; }

            [CsvColumn(Name = "DOLARES", FieldIndex = 8)]
            public string Dolares { get; set; }

            [CsvColumn(Name = "PESOS", FieldIndex = 9)]
            public string Pesos { get; set; }

            [CsvColumn(Name = "CCOSTOS_DESTINO", FieldIndex = 10)]
            public string CentroDestino { get; set; }

            [CsvColumn(Name = "EQUIPO", FieldIndex = 11)]
            public string Equipo { get; set; }

            [CsvColumn(Name = "PROJ/WO_DESTINO", FieldIndex = 12)]
            public string ProyectoOrdenDestino { get; set; }

            [CsvColumn(Name = "IND_DESTINO", FieldIndex = 13)]
            public string IndicadorDestino { get; set; }

            [CsvColumn(Name = "RAZON DEL CAMBIO", FieldIndex = 14)]
            public string Razon { get; set; }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
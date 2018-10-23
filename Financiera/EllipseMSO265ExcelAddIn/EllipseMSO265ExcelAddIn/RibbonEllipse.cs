using System.Threading;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary.Utilities;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using Invoice = EllipseMSO265ExcelAddIn.Invoice265.Invoice;
using InvoiceItem = EllipseMSO265ExcelAddIn.Invoice265.InvoiceItem;
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable LocalizableElement
// ReSharper disable UnusedMember.Local
// ReSharper disable UseNullPropagation
// ReSharper disable LoopCanBeConvertedToQuery
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace EllipseMSO265ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int TitleRow01 = 5;
        private const int TitleRow02 = 5;
        private const int ResultColumn01C = 26;
        private const int ResultColumn01N = 19;
        private const int ResultColumn01X = 21;
        private const int ResultColumn02 = 4;
        private static EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private const string SheetName01C = "MSO265 Cesantias";
        private const string SheetName01N = "MSO265 Nomina";
        private const string SheetName01X = "MSO265 NonOrderInvoice";
        private const string SheetName02 = "Comentarios";
        private const string TableName02 = "TablaComentarios";
        private const string TableName01C = "TablaCesantias";
        private const string TableName01N = "TablaNomina";
        private const string TableName01X = "TablaNonOrderInvoice";

        private const string ValidationSheetName = "ListaImpuestos";
        private Thread _thread;

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

        private void btnNomina_Click(object sender, RibbonControlEventArgs e)
        {
            FormatNomina();
        }

        private void btnCalculateTaxes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01N ||  _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01X)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(CalculateTaxes);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CalculateTaxes()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnFormatNonOrderInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            FormatNonOrderInvoice();
        }

        private void FormatNonOrderInvoice()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _excelApp.Workbooks.Add();
            while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                _excelApp.ActiveWorkbook.Worksheets.Add();

            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01X;
            _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
            #region Titulo

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Non Order Invoice Entry";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Encabezados


            _cells.GetRange(1, TitleRow01, ResultColumn01X - 1, TitleRow01).Style = StyleConstants.TitleRequired;
            _cells.GetCell(ResultColumn01X, TitleRow01).Style = StyleConstants.TitleResult;

            _cells.GetCell(1, TitleRow01).Value = "District";
            _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(2, TitleRow01).Value = "Supplier";
            _cells.GetCell(3, TitleRow01).Value = "Mnemonic";
            _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(4, TitleRow01).Value = "Invoice No.";
            _cells.GetCell(5, TitleRow01).Value = "Invoice AMT";
            _cells.GetCell(6, TitleRow01).Value = "Add Tax Value";
            _cells.GetCell(6, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(7, TitleRow01).Value = "Group Tax";
            _cells.GetCell(7, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(8, TitleRow01).Value = "Add Tax Codes";
            _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(9, TitleRow01).Value = "Currency";
            _cells.GetCell(9, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(10, TitleRow01).Value = "Invoice date (YYYYMMDD)";
            _cells.GetCell(11, TitleRow01).Value = "Invoice received (YYYYMMDD)";
            _cells.GetCell(12, TitleRow01).Value = "Due date (YYYYMMDD)";
            _cells.GetCell(13, TitleRow01).Value = "Bank branch";
            _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(14, TitleRow01).Value = "Account No.";
            _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(15, TitleRow01).Value = "Description";
            _cells.GetCell(16, TitleRow01).Value = "Value";
            _cells.GetCell(17, TitleRow01).Value = "Auth by";
            _cells.GetCell(18, TitleRow01).Value = "Account";
            _cells.GetCell(19, TitleRow01).Value = "Wo/Proj No.";
            _cells.GetCell(19, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(20, TitleRow01).Value = "W/P";
            _cells.GetCell(20, TitleRow01).Style = StyleConstants.TitleOptional;
            _cells.GetCell(ResultColumn01X, TitleRow01).Value = "Result";

            _cells.GetCell(8, TitleRow01).AddComment("Puede agregar diversos códigos separando con punto y coma. Ejemplo: US10; HOP1");

            _cells.GetCell(1, TitleRow01 + 1).Value = "ICOR";
            _cells.GetCell(9, TitleRow01 + 1).Value = "PES";

            var taxGroupCodeList = Invoice265.GetTaxGroupCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(7, TitleRow01 + 1), taxGroupCodeList, ValidationSheetName, 1, false);

            var taxCodeList = Invoice265.GetTaxCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(8, TitleRow01 + 1), taxCodeList, ValidationSheetName, 2, false);

            var workProjectIndicatorList = new List<string> {"W - WorkOrder", "P - Project"};
            _cells.SetValidationList(_cells.GetCell(20, TitleRow01 + 1), workProjectIndicatorList, ValidationSheetName, 3, false);

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01X, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.GetCell(5, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.GetCell(16, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01X, TitleRow01 + 1), TableName01X);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            #endregion
            #region Comentarios
            _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("A1", "B2");
            _cells.GetCell("C1").Value = "COMENTARIOS";
            _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("C1", "J2");

            _cells.GetCell("K1").Value = "OBLIGATORIO";
            _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell("K2").Value = "OPCIONAL";
            _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
            _cells.GetCell("K3").Value = "INFORMATIVO";
            _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.GetCell("A3").Value = "DISTRITO";
            _cells.GetCell("B3").Value = "ICOR";
            _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

            _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;
            _cells.GetCell(1, TitleRow02).Value = "Supplier";
            _cells.GetCell(2, TitleRow02).Value = "Referencia";
            _cells.GetCell(3, TitleRow02).Value = "Comentario";
            _cells.GetCell(3, TitleRow02 + 1).WrapText = true;
            _cells.GetCell(3, TitleRow02 + 1).ColumnWidth = 60;

            _cells.GetCell(ResultColumn02, TitleRow02 + 1).ColumnWidth = 36;
            _cells.GetCell(ResultColumn02, TitleRow02).Value = "Resultado";
            _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;
            _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            #endregion

            _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }

        private void FormatNomina()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _excelApp.Workbooks.Add();
            while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                _excelApp.ActiveWorkbook.Worksheets.Add();

            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01N;
            _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
            #region Titulo

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pago de Nomina";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Encabezados



            _cells.GetCell(1, TitleRow01).Value = "Codigo Banco";
            _cells.GetCell(2, TitleRow01).Value = "Cuenta Banco";
            _cells.GetCell(3, TitleRow01).Value = "Analista";
            _cells.GetCell(4, TitleRow01).Value = "Supplier";
            _cells.GetCell(5, TitleRow01).Value = "Cedula";
            _cells.GetCell(6, TitleRow01).Value = "Moneda";
            _cells.GetCell(7, TitleRow01).Value = "NumFactura";
            _cells.GetCell(8, TitleRow01).Value = "Fecha Factura";
            _cells.GetCell(9, TitleRow01).Value = "Fecha Pago";
            _cells.GetCell(10, TitleRow01).Value = "Valor Total";
            _cells.GetCell(11, TitleRow01).Value = "Descripción";
            _cells.GetCell(12, TitleRow01).Value = "REF";
            _cells.GetCell(13, TitleRow01).Value = "Valor Item";
            _cells.GetCell(14, TitleRow01).Value = "Cuenta";
            _cells.GetCell(15, TitleRow01).Value = "Posicion Aprobador";
            _cells.GetCell(16, TitleRow01).Value = "Valor Impuesto";
            _cells.GetCell(17, TitleRow01).Value = "Grupo Impuesto";
            _cells.GetCell(18, TitleRow01).Value = "Adicional Impuesto";
            _cells.GetCell(18, TitleRow01).AddComment("Puede agregar diversos códigos separando con punto y coma. Ejemplo: US10; HOP1");

            var taxGroupCodeList = Invoice265.GetTaxGroupCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(17, TitleRow01 + 1), taxGroupCodeList, ValidationSheetName, 1, false);

            var taxCodeList = Invoice265.GetTaxCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(18, TitleRow01 + 1), taxCodeList, ValidationSheetName, 2, false);

            _cells.GetCell(ResultColumn01N, TitleRow01).Value = "Result";

            _cells.GetRange(1, TitleRow01, ResultColumn01N - 1, TitleRow01).Style = StyleConstants.TitleRequired;
            _cells.GetCell(ResultColumn01N, TitleRow01).Style = StyleConstants.TitleResult;

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01N, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.GetRange(10, TitleRow01 + 1, 10, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.GetRange(13, TitleRow01 + 1, 13, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.GetRange(16, TitleRow01 + 1, 16, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01N, TitleRow01 + 1), TableName01N);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            ImportFileNomina();

            #endregion
            #region Comentarios
            _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("A1", "B2");
            _cells.GetCell("C1").Value = "COMENTARIOS";
            _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("C1", "J2");

            _cells.GetCell("K1").Value = "OBLIGATORIO";
            _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell("K2").Value = "OPCIONAL";
            _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
            _cells.GetCell("K3").Value = "INFORMATIVO";
            _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.GetCell("A3").Value = "DISTRITO";
            _cells.GetCell("B3").Value = "ICOR";
            _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

            _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;
            _cells.GetCell(1, TitleRow02).Value = "Supplier";
            _cells.GetCell(2, TitleRow02).Value = "Referencia";
            _cells.GetCell(3, TitleRow02).Value = "Comentario";
            _cells.GetCell(3, TitleRow02 + 1).WrapText = true;
            _cells.GetCell(3, TitleRow02 + 1).ColumnWidth = 60;

            _cells.GetCell(ResultColumn02, TitleRow02 + 1).ColumnWidth = 36;
            _cells.GetCell(ResultColumn02, TitleRow02).Value = "Resultado";
            _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;
            _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            #endregion

            _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }

        private void ImportFileNomina()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01N) return;

            _cells.ClearTableRange(TableName01N);

            var openFileDialog2 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"ChequMim nominaquincenalfebrero.csv",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\Data\Loaders\Parametros"
            };

            if (openFileDialog2.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog2.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var nominaParameters = cc.Read<Invoice265.NominaParameters>(filePath, inputFileDescription);

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01N, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.GetRange(10, TitleRow01 + 1, 10, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.GetRange(13, TitleRow01 + 1, 13, TitleRow01 + 1).NumberFormat = "$ #,##0.00";

            var currentRow = TitleRow01 + 1;
            foreach (var c in nominaParameters)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Value = c.BranchCode;
                    _cells.GetCell(2, currentRow).Value = c.BankAccount;
                    _cells.GetCell(3, currentRow).Value = c.Accountant;
                    _cells.GetCell(4, currentRow).Value = c.SupplierNo;
                    _cells.GetCell(5, currentRow).Value = c.SupplierMnemonic;
                    _cells.GetCell(6, currentRow).Value = c.Currency;
                    _cells.GetCell(7, currentRow).Value = c.InvoiceNo;
                    _cells.GetCell(8, currentRow).Value = c.InvoiceDate;
                    _cells.GetCell(9, currentRow).Value = c.DueDate;
                    _cells.GetCell(10, currentRow).Value = c.InvoiceAmount;
                    _cells.GetCell(11, currentRow).Value = c.Description;
                    _cells.GetCell(12, currentRow).Value = c.Ref;
                    _cells.GetCell(13, currentRow).Value = c.ItemValue;
                    _cells.GetCell(14, currentRow).Value = c.Account;
                    _cells.GetCell(15, currentRow).Value = c.AuthorizedBy;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ImportFileNomina()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    currentRow++;
                }
            }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        private void btnFormatCesantias_Click(object sender, RibbonControlEventArgs e)
        {
            FormatCesantias();
        }

        private void FormatCesantias()
        {
            _excelApp = Globals.ThisAddIn.Application;
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _excelApp.Workbooks.Add();
            while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                _excelApp.ActiveWorkbook.Worksheets.Add();
            _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01C;
            _cells.SetCursorWait();

            #region Instructions

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pago de Cesantias";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Datos

            _cells.GetCell(1, TitleRow01).Value = "Cedula";
            _cells.GetCell(2, TitleRow01).Value = "Nombre";
            _cells.GetCell(3, TitleRow01).Value = "Referencia";
            _cells.GetCell(4, TitleRow01).Value = "Descripcion";
            _cells.GetCell(5, TitleRow01).Value = "Fecha Factura";
            _cells.GetCell(6, TitleRow01).Value = "Fecha Pago";
            _cells.GetCell(7, TitleRow01).Value = "Cuenta";
            _cells.GetCell(8, TitleRow01).Value = "Moneda";
            _cells.GetCell(9, TitleRow01).Value = "Valor Item";
            _cells.GetCell(10, TitleRow01).Value = "Valor Total";
            _cells.GetCell(11, TitleRow01).Value = "Posicion Aprobador";
            _cells.GetCell(12, TitleRow01).Value = "Codigo Banco";
            _cells.GetCell(13, TitleRow01).Value = "Cuenta Banco";
            _cells.GetCell(14, TitleRow01).Value = "Banco";
            _cells.GetCell(15, TitleRow01).Value = "Sucursal Banco";
            _cells.GetCell(16, TitleRow01).Value = "Analista";
            _cells.GetCell(17, TitleRow01).Value = "Supplier";
            _cells.GetCell(18, TitleRow01).Value = "Sucursal Banco Ellipse";
            _cells.GetCell(19, TitleRow01).Value = "Cuenta Banco Ellipse";
            _cells.GetCell(20, TitleRow01).Value = "ST Adress";
            _cells.GetCell(21, TitleRow01).Value = "ST Business";
            _cells.GetCell(22, TitleRow01).Value = "ST Status";
            _cells.GetCell(23, TitleRow01).Value = "Valor Impuesto";
            _cells.GetCell(24, TitleRow01).Value = "Grupo Impuesto";
            _cells.GetCell(25, TitleRow01).Value = "Impuesto Adicional";
            _cells.GetCell(25, TitleRow01).AddComment("Puede agregar diversos códigos separando con punto y coma. Ejemplo: US10; HOP1");
            _cells.GetCell(ResultColumn01C, TitleRow01).Value = "Result";

            _cells.GetRange(1, TitleRow01, ResultColumn01C - 1, TitleRow01).Style = StyleConstants.TitleRequired;
            _cells.GetCell(ResultColumn01C, TitleRow01).Style = StyleConstants.TitleResult;


            var taxGroupCodeList = Invoice265.GetTaxGroupCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(24, TitleRow01 + 1), taxGroupCodeList, ValidationSheetName, 1, false);

            var taxCodeList = Invoice265.GetTaxCodeList(_eFunctions).Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(25, TitleRow01 + 1), taxCodeList, ValidationSheetName, 2, false);

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01C, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.GetRange(9, TitleRow01 + 1, 10, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.GetCell(23, TitleRow01 + 1).NumberFormat = "$ #,##0.00";
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01C, TitleRow01 + 1), TableName01C);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            ImportFileCesantias();
            _cells.SetCursorDefault();
            #endregion

            #region Comentarios
            _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("A1", "B2");
            _cells.GetCell("C1").Value = "COMENTARIOS";
            _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells("C1", "J2");

            _cells.GetCell("K1").Value = "OBLIGATORIO";
            _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
            _cells.GetCell("K2").Value = "OPCIONAL";
            _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
            _cells.GetCell("K3").Value = "INFORMATIVO";
            _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

            _cells.GetCell("A3").Value = "DISTRITO";
            _cells.GetCell("B3").Value = "ICOR";
            _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
            _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

            _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;
            _cells.GetCell(1, TitleRow02).Value = "Supplier";
            _cells.GetCell(2, TitleRow02).Value = "Referencia";
            _cells.GetCell(3, TitleRow02).Value = "Comentario";
            _cells.GetCell(3, TitleRow02 + 1).WrapText = true;
            _cells.GetCell(3, TitleRow02 + 1).ColumnWidth = 60;

            _cells.GetCell(ResultColumn02, TitleRow02 + 1).ColumnWidth = 36;
            _cells.GetCell(ResultColumn02, TitleRow02).Value = "Resultado";
            _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;
            _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
            _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            #endregion

            _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }

        private void ReviewInternalText()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);

            var i = TitleRow02 + 1;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            string districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = EllipseStdTextClassLibrary.StdText.GetCustomOpContext(districtCode, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    //GENERAL
                    var supplier = "" + _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var extendedInvoice = "" + _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);

                    var sqlQuery = "SELECT INV_NO FROM ELLIPSE.MSF260 WHERE SUPPLIER_NO = '" + supplier + "' AND EXT_INV_NO = '" + extendedInvoice + "'";
                    var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                    if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
                    {
                        _eFunctions.CloseConnection();
                        throw new Exception("No se ha encontrado una combinación válida para el supplier y la referencia ingresada");
                    }

                    dataReader.Read();

                    var invoice = dataReader["INV_NO"].ToString();

                    var stdTextId = "II" + districtCode + supplier + invoice;
                    var internalText = EllipseStdTextClassLibrary.StdText.GetText(urlService, opContext, stdTextId);


                    _cells.GetCell(3, i).Value = internalText;
                    _cells.GetCell(ResultColumn02, i).Value = "CONSULTADO SystemInvoice: " + invoice;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn02 - 1, i).Value = "";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewInternalText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                }
                finally
                {
                    _cells.GetCell(3, i).WrapText = true;
                    _cells.GetCell(3, i).ColumnWidth = 60;
                    _cells.GetCell(ResultColumn02, i).ColumnWidth = 36;

                    _cells.GetCell(1, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void UpdateInternalText()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);

            var i = TitleRow02 + 1;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            string districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = EllipseStdTextClassLibrary.StdText.GetCustomOpContext(districtCode, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    //GENERAL
                    var supplier = "" + _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var extendedInvoice = "" + _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var newInternalText = "" + _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    var sqlQuery = "SELECT INV_NO FROM ELLIPSE.MSF260 WHERE SUPPLIER_NO = '" + supplier + "' AND EXT_INV_NO = '" + extendedInvoice + "'";
                    var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                    if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
                        throw new Exception("No se ha encontrado una combinación válida para el supplier y la referencia ingresada");

                    dataReader.Read();

                    var invoice = dataReader["INV_NO"].ToString();

                    var stdTextId = "II" + districtCode + supplier + invoice;
                    EllipseStdTextClassLibrary.StdText.SetText(urlService, opContext, stdTextId, newInternalText);
                    
                    _cells.GetCell(ResultColumn02, i).Value = "ACTUALIZADO SystemInvoice: " + invoice;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateInternalText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                }
                finally
                {
                    _cells.GetCell(3, i).WrapText = true;
                    _cells.GetCell(3, i).ColumnWidth = 60;
                    _cells.GetCell(ResultColumn02, i).ColumnWidth = 36;

                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ImportFileCesantias()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName01C);

            var openFileDialog1 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"cesantias.csv",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\Data\Loaders\Parametros"
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog1.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var cesantiasParameters = cc.Read<Invoice265.CesantiasParameters>(filePath, inputFileDescription);

            _cells.GetRange(1, TitleRow01 + 1, ResultColumn01C, TitleRow01 + 1).NumberFormat = "@";
            _cells.GetRange(9, TitleRow01 + 1, 10, TitleRow01 + 1).NumberFormat = "$ #,##0.00";

            var currentRow = TitleRow01 + 1;
            foreach (var c in cesantiasParameters)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Value = c.SupplierMnemonic;
                    _cells.GetCell(2, currentRow).Value = c.SupplierName;
                    _cells.GetCell(3, currentRow).Value = c.Reference;
                    _cells.GetCell(4, currentRow).Value = c.Description;
                    _cells.GetCell(5, currentRow).Value = c.InvoiceDate;
                    _cells.GetCell(6, currentRow).Value = c.DueDate;
                    _cells.GetCell(7, currentRow).Value = c.Account;
                    _cells.GetCell(8, currentRow).Value = c.Currency;
                    _cells.GetCell(9, currentRow).Value = c.ItemValue;
                    _cells.GetCell(10, currentRow).Value = c.InvoiceAmount;
                    _cells.GetCell(11, currentRow).Value = c.AuthorizedBy;
                    _cells.GetCell(12, currentRow).Value = c.BranchCode;
                    _cells.GetCell(13, currentRow).Value = c.BankAccount;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ImportFileCesantias()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    currentRow++;
                }
            }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();

            ValidateCesantias();

            _cells.SetCursorDefault();
        }

        private void btnValidate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ValidateCesantias);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ValidateCesantias()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ValidateCesantias()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var currentRow = TitleRow01 + 1;

            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();


            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                _cells.GetRange(12, currentRow, 13, currentRow).Style = StyleConstants.Normal;
                _cells.GetRange(18, currentRow, 19, currentRow).Style = StyleConstants.Normal;

                _cells.GetRange(1, currentRow, ResultColumn01C, currentRow).NumberFormat = "@";
                _cells.GetRange(9, currentRow, 10, currentRow).NumberFormat = "$ #,##0.00";
                try
                {
                    var supplierTaxFileId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    var supplierInfo = new Invoice265.Supplier(_eFunctions, null, supplierTaxFileId);
                    _cells.GetCell(14, currentRow).Select();

                    _cells.GetCell(17, currentRow).Value = supplierInfo.SupplierNo;
                    _cells.GetCell(18, currentRow).Value = supplierInfo.AccountName.Substring(2, 4);
                    _cells.GetCell(19, currentRow).Value = supplierInfo.AccountNo;
                    _cells.GetCell(20, currentRow).Value = supplierInfo.StAdress;
                    _cells.GetCell(21, currentRow).Value = supplierInfo.StBusiness;
                    _cells.GetCell(22, currentRow).Value = supplierInfo.Status;

                    var bankCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value);

                    if (supplierInfo.AccountName.Substring(2, 4) == bankCode)
                    {
                        _cells.GetCell(12, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(18, currentRow).Style = StyleConstants.Success;
                    }
                    else
                    {
                        _cells.GetCell(12, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(18, currentRow).Style = StyleConstants.Error;
                    }

                    var bankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value);

                    if (supplierInfo.AccountNo == bankAccount)
                    {
                        _cells.GetCell(13, currentRow).Style = StyleConstants.Success;
                        _cells.GetCell(19, currentRow).Style = StyleConstants.Success;
                    }
                    else
                    {
                        _cells.GetCell(13, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(19, currentRow).Style = StyleConstants.Error;
                    }
                    _cells.GetCell(ResultColumn01C, currentRow).Value = "Proceso Exitoso";
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ValidateCesantias()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    _cells.GetCell(ResultColumn01C, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
            _cells.SetCursorDefault();
        }

        private void btnReloadParameters_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ImportFileCesantias);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReloadParameters()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01N || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01X)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C)
                        _thread = new Thread(LoadCesantiasPost);
                    else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01N)
                        _thread = new Thread(LoadNominaPost);
                    else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01X)
                        _thread = new Thread(LoadNonInvoiceOrder);
                    else
                        throw new Exception("@La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:Load()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void CalculateTaxes()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var currentRow = TitleRow01 + 1;

            //selección de acción hoja
            int resultColumn = 0;
            int itemIndexValue = 0;
            int taxIndexValue = 0;
            int groupTaxIndexValue = 0;
            int additionalTaxIndexValue = 0;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01N)
            {
                resultColumn = ResultColumn01N;
                itemIndexValue = 13;
                taxIndexValue = 16;
                groupTaxIndexValue = 17;
                additionalTaxIndexValue = 18;
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01C)
            {
                resultColumn = ResultColumn01C;
                itemIndexValue = 9;
                taxIndexValue = 23;
                groupTaxIndexValue = 24;
                additionalTaxIndexValue = 25;
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01X)
            {
                    resultColumn = ResultColumn01X;
                    itemIndexValue = 16;
                    taxIndexValue = 6;
                    groupTaxIndexValue = 7;
                    additionalTaxIndexValue = 8;
            }
            // -Fin selección de acción hoja


            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    _cells.GetCell(1, resultColumn).Style = StyleConstants.Normal;
                    var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                    var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);

                    var valorItem = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal);
                    var valorImpuesto = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal);

                    var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));
                    var additionalTaxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                    var taxItemList = Invoice265.InvoiceActions.GetItemTaxList(_eFunctions, groupTaxCode, additionalTaxCodeList);
                    decimal calculatedTaxValue = Invoice265.InvoiceActions.GetCalculatedItemTaxValue(valorItem, taxItemList);

                    //Comparo los valores de impuesto ingresados (si existe)
                    if (string.IsNullOrWhiteSpace(valorImpuestoString) || valorImpuesto == calculatedTaxValue)
                    {
                        _cells.GetCell(taxIndexValue, currentRow).Value = calculatedTaxValue;
                        _cells.GetCell(resultColumn, currentRow).Value = "Impuesto Calculado: " + calculatedTaxValue;
                        _cells.GetCell(resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                    }
                    else if (valorImpuesto != calculatedTaxValue)
                    {
                        _cells.GetCell(taxIndexValue, currentRow).Value = calculatedTaxValue;
                        _cells.GetCell(resultColumn, currentRow).Value = "Impuesto Calculado: " + calculatedTaxValue + ". Valor anterior: " + valorImpuesto;
                        _cells.GetCell(resultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Warning);
                    }
                    _cells.GetCell(resultColumn, currentRow).Select();
                    //
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:CalculateTaxes()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    _cells.GetCell(resultColumn, currentRow).Select();
                    _cells.GetCell(resultColumn, currentRow).Value = ex.Message;
                    _cells.GetCell(resultColumn, currentRow).Style = StyleConstants.Error;
                }
                finally
                {
                    currentRow++;
                }
            }
            _eFunctions.CloseConnection();
            _cells.SetCursorDefault();
        }

        private void LoadNonInvoiceOrder()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var currentRow = TitleRow01 + 1;
            var startRow = TitleRow01 + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var invoice = new Invoice
                    {
                        District = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        SupplierMnemonic = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        InvoiceDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        InvoiceReceivedDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value),
                        BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                        BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        TaxAmount = 0
                    };

                    var invoiceItemList = new List<InvoiceItem>();
                    const int itemIndexValue = 16;
                    const int taxIndexValue = 6;
                    const int groupTaxIndexValue = 7;
                    const int additionalTaxIndexValue = 8;

                    Invoice nextInvoice;
                    do
                    {
                        var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                        var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);

                        var invoiceItem = new InvoiceItem
                        {
                            Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                            ItemValue = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal),
                            TaxValue = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal),
                            AuthorizedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value),
                            Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value),
                            WorkOrderProjectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value),
                            WorkOrderProjectIndicator = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, currentRow).Value)
                        };

                        var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));
                        var additionalTaxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                        invoiceItem.TaxList = Invoice265.InvoiceActions.GetItemTaxList(_eFunctions, groupTaxCode, additionalTaxCodeList);
                        decimal calculatedTaxValue = Invoice265.InvoiceActions.GetCalculatedItemTaxValue(invoiceItem.ItemValue, invoiceItem.TaxList);

                        if (invoiceItem.TaxValue != 0 && invoiceItem.TaxValue != calculatedTaxValue)
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        if (invoiceItem.TaxValue == 0 && calculatedTaxValue != 0)
                        {
                            invoiceItem.TaxValue = calculatedTaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Value = invoiceItem.TaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        }

                        invoice.TaxAmount += invoiceItem.TaxValue;
                        //Requerido para ajuste manual de impuestos
                        invoiceItem.FirstTaxAdjustment = Invoice265.InvoiceActions.GetItemTaxAdjustment(invoiceItem.ItemValue, calculatedTaxValue, invoiceItem.TaxValue, invoiceItem.TaxList);
                        //
                        invoiceItemList.Add(invoiceItem);

                        //Siguiente Invoice para comparación
                        nextInvoice = new Invoice
                        {
                            District = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow + 1).Value),
                            SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow + 1).Value),
                            SupplierMnemonic = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow + 1).Value),
                            InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow + 1).Value),
                            InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow + 1).Value),
                            Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow + 1).Value),
                            InvoiceDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow + 1).Value),
                            InvoiceReceivedDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow + 1).Value),
                            DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow + 1).Value),
                            BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow + 1).Value),
                            BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow + 1).Value),
                        };
                        //
                        if(invoice.Equals(nextInvoice))
                            currentRow++;
                    } while (invoice.Equals(nextInvoice));

                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");
                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);
                    Invoice265.InvoiceActions.LoadNonInvoice(_eFunctions, invoice, invoiceItemList);

                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01X, i).Select();
                        _cells.GetCell(ResultColumn01X, i).Value = "Creado";
                        _cells.GetRange(1, i, ResultColumn01X, i).Style = _cells.GetStyle(StyleConstants.Success);

                        //Escribe el branch code y bank account si está en blanco
                        if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value) != null && _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value) != null) continue;
                        _cells.GetCell(13, i).Value = invoice.BankBranchCode;
                        _cells.GetCell(14, i).Value = invoice.BankAccountNo;
                    }
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:LoadNonInvoiceOrder()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01X, i).Select();
                        _cells.GetCell(ResultColumn01X, i).Value = ex.Message;
                        _cells.GetCell(ResultColumn01X, i).Style = StyleConstants.Error;
                        _cells.GetRange(1, i, ResultColumn01X, i).Style = StyleConstants.Error;
                    }
                }
                finally
                {
                    currentRow++;
                    startRow = currentRow;
                }
            }
            _cells.SetCursorDefault();
        }
        
        private void LoadNominaPost()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var currentRow = TitleRow01 + 1;
            var startRow = TitleRow01 + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var fechaFactura = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value), "MMddyy", CultureInfo.InvariantCulture);

                    var fechaPago = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value), "MMddyy", CultureInfo.InvariantCulture);

                    var invoice = new Invoice
                    {
                        BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        SupplierMnemonic = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        InvoiceDate = fechaFactura.ToString("yyyyMMdd"),
                        DueDate = fechaPago.ToString("yyyyMMdd"),
                        InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),

                        Ref = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value),
                        TaxAmount = 0
                    };

                    var invoiceItemList = new List<InvoiceItem>();
                    const int itemIndexValue = 13;
                    const int taxIndexValue = 16;
                    const int groupTaxIndexValue = 17;
                    const int additionalTaxIndexValue = 18;

                    Invoice nextInvoice;
                    do
                    {
                        var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                        var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);

                        var invoiceItem = new InvoiceItem
                        {
                            Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                            ItemValue = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal),
                            TaxValue = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal),
                            AuthorizedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                            Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        };

                        var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));
                        var additionalTaxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                        invoiceItem.TaxList = Invoice265.InvoiceActions.GetItemTaxList(_eFunctions, groupTaxCode, additionalTaxCodeList);
                        decimal calculatedTaxValue = Invoice265.InvoiceActions.GetCalculatedItemTaxValue(invoiceItem.ItemValue, invoiceItem.TaxList);

                        if (invoiceItem.TaxValue != 0 && invoiceItem.TaxValue != calculatedTaxValue)
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        if (invoiceItem.TaxValue == 0 && calculatedTaxValue != 0)
                        {
                            invoiceItem.TaxValue = calculatedTaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Value = invoiceItem.TaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        }

                        invoice.TaxAmount += invoiceItem.TaxValue;
                        //Requerido para ajuste manual de impuestos
                        invoiceItem.FirstTaxAdjustment = Invoice265.InvoiceActions.GetItemTaxAdjustment(invoiceItem.ItemValue, calculatedTaxValue, invoiceItem.TaxValue, invoiceItem.TaxList);
                        //
                        invoiceItemList.Add(invoiceItem);

                        //Siguiente Invoice para comparación

                        var nextFechaFactura = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow + 1).Value), "MMddyy", CultureInfo.InvariantCulture);
                        var nextFechaPago = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow + 1).Value), "MMddyy", CultureInfo.InvariantCulture);

                        nextInvoice = new Invoice
                        {
                            BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow + 1).Value),
                            BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow + 1).Value),
                            Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow + 1).Value),
                            SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow + 1).Value),
                            SupplierMnemonic = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow + 1).Value),
                            Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow + 1).Value),
                            InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow + 1).Value),
                            InvoiceDate = nextFechaFactura.ToString("yyyyMMdd"),
                            DueDate = nextFechaPago.ToString("yyyyMMdd"),
                            InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow + 1).Value),

                            Ref = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow + 1).Value),
                        };
                        //
                        if (invoice.Equals(nextInvoice))
                            currentRow++;
                    } while (invoice.Equals(nextInvoice));

                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");
                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);
                    Invoice265.InvoiceActions.LoadNonInvoice(_eFunctions, invoice, invoiceItemList);

                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01N, i).Select();
                        _cells.GetCell(ResultColumn01N, i).Value = "Creado";
                        _cells.GetRange(1, i, ResultColumn01N, i).Style = _cells.GetStyle(StyleConstants.Success);

                        //Escribe el branch code y bank account si está en blanco
                        if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null && _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null) continue;
                        _cells.GetCell(1, i).Value = invoice.BankBranchCode;
                        _cells.GetCell(2, i).Value = invoice.BankAccountNo;
                    }
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:LoadNominaPost()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01N, i).Select();
                        _cells.GetCell(ResultColumn01N, i).Value = ex.Message;
                        _cells.GetCell(ResultColumn01N, i).Style = StyleConstants.Error;
                        _cells.GetRange(1, i, ResultColumn01N, i).Style = StyleConstants.Error;
                    }
                }
                finally
                {
                    currentRow++;
                    startRow = currentRow;
                }
            }
            _cells.SetCursorDefault();
        }

        private void LoadCesantiasPost()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var currentRow = TitleRow01 + 1;
            var startRow = TitleRow01 + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var invoice = new Invoice
                    {
                        SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value),
                        Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                        InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        InvoiceDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                        TaxAmount = 0
                    };

                    var invoiceItemList = new List<InvoiceItem>();
                    const int itemIndexValue = 9;
                    const int taxIndexValue = 23;
                    const int groupTaxIndexValue = 24;
                    const int additionalTaxIndexValue = 25;

                    Invoice nextInvoice;
                    do
                    {
                        var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                        var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);

                        var invoiceItem = new InvoiceItem
                        {
                            Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                            ItemValue = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal),
                            TaxValue = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal),
                            AuthorizedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                            Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        };

                        var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));
                        var additionalTaxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                        invoiceItem.TaxList = Invoice265.InvoiceActions.GetItemTaxList(_eFunctions, groupTaxCode, additionalTaxCodeList);
                        decimal calculatedTaxValue = Invoice265.InvoiceActions.GetCalculatedItemTaxValue(invoiceItem.ItemValue, invoiceItem.TaxList);

                        if (invoiceItem.TaxValue != 0 && invoiceItem.TaxValue != calculatedTaxValue)
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        if (invoiceItem.TaxValue == 0 && calculatedTaxValue != 0)
                        {
                            invoiceItem.TaxValue = calculatedTaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Value = invoiceItem.TaxValue;
                            _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                        }

                        invoice.TaxAmount += invoiceItem.TaxValue;
                        //Requerido para ajuste manual de impuestos
                        invoiceItem.FirstTaxAdjustment = Invoice265.InvoiceActions.GetItemTaxAdjustment(invoiceItem.ItemValue, calculatedTaxValue, invoiceItem.TaxValue, invoiceItem.TaxList);
                        //
                        invoiceItemList.Add(invoiceItem);

                        //Siguiente Invoice para comparación
                        nextInvoice = new Invoice
                        {
                            SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow + 1).Value),
                            Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow + 1).Value),
                            InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow + 1).Value),
                            InvoiceDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow + 1).Value),
                            DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow + 1).Value),
                            Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow + 1).Value),
                            InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow + 1).Value),
                            BankBranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow + 1).Value),
                            BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow + 1).Value),
                        };
                        //
                        if (invoice.Equals(nextInvoice))
                            currentRow++;
                    } while (invoice.Equals(nextInvoice));

                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");
                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);
                    Invoice265.InvoiceActions.LoadNonInvoice(_eFunctions, invoice, invoiceItemList);

                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01C, i).Select();
                        _cells.GetCell(ResultColumn01C, i).Value = "Creado";
                        _cells.GetRange(1, i, ResultColumn01C, i).Style = _cells.GetStyle(StyleConstants.Success);

                        //Escribe el branch code y bank account si está en blanco
                        if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value) != null && _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value) != null) continue;
                        _cells.GetCell(14, i).Value = invoice.BankBranchCode;
                        _cells.GetCell(15, i).Value = invoice.BankAccountNo;
                    }
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:LoadCesantiasPost()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    for (int i = startRow; i <= currentRow; i++)
                    {
                        _cells.GetCell(ResultColumn01C, i).Select();
                        _cells.GetCell(ResultColumn01C, i).Value = ex.Message;
                        _cells.GetCell(ResultColumn01C, i).Style = StyleConstants.Error;
                        _cells.GetRange(1, i, ResultColumn01C, i).Style = StyleConstants.Error;
                    }
                }
                finally
                {
                    currentRow++;
                    startRow = currentRow;
                }
            }
            _cells.SetCursorDefault();
        }
        
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnReviewInternalComments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReviewInternalText);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewInternalText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateInternalComments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateInternalText);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewInternalText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
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

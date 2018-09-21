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
using Util = System.Web.Services.Ellipse.Post.Util;
using EllipseCommonsClassLibrary.Utilities;
using Screen = EllipseCommonsClassLibrary.ScreenService;
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



            _cells.GetCell(1, TitleRow01).Value = "District";
            _cells.GetCell(2, TitleRow01).Value = "Supplier";
            _cells.GetCell(3, TitleRow01).Value = "Mnemonic";
            _cells.GetCell(4, TitleRow01).Value = "Invoice No.";
            _cells.GetCell(5, TitleRow01).Value = "Invoice AMT";
            _cells.GetCell(6, TitleRow01).Value = "Add Tax Value";
            _cells.GetCell(7, TitleRow01).Value = "Group Tax";
            _cells.GetCell(8, TitleRow01).Value = "Add Tax Codes";
            _cells.GetCell(9, TitleRow01).Value = "Currency";
            _cells.GetCell(10, TitleRow01).Value = "Invoice date (YYYYMMDD)";
            _cells.GetCell(11, TitleRow01).Value = "Invoice received (YYYYMMDD)";
            _cells.GetCell(12, TitleRow01).Value = "Due date (YYYYMMDD)";
            _cells.GetCell(13, TitleRow01).Value = "Bank branch";
            _cells.GetCell(14, TitleRow01).Value = "Account No.";
            _cells.GetCell(15, TitleRow01).Value = "Description";
            _cells.GetCell(16, TitleRow01).Value = "Value";
            _cells.GetCell(17, TitleRow01).Value = "Auth by";
            _cells.GetCell(18, TitleRow01).Value = "Account";
            _cells.GetCell(19, TitleRow01).Value = "Wo/Proj No.";
            _cells.GetCell(20, TitleRow01).Value = "W/P";

            _cells.GetCell(8, TitleRow01).AddComment("Puede agregar diversos códigos separando con punto y coma. Ejemplo: US10; HOP1");

            var taxGroupCodeList = GetTaxGroupCodeList().Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(7, TitleRow01 + 1), taxGroupCodeList, ValidationSheetName, 1, false);

            var taxCodeList = GetTaxCodeList().Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(8, TitleRow01 + 1), taxCodeList, ValidationSheetName, 2, false);

            var workProjectIndicatorList = new List<string> {"W - WorkOrder", "P - Project"};
            _cells.SetValidationList(_cells.GetCell(8, TitleRow01 + 1), workProjectIndicatorList, ValidationSheetName, 3, false);

            _cells.GetCell(ResultColumn01X, TitleRow01).Value = "Result";

            _cells.GetRange(1, TitleRow01, ResultColumn01X - 1, TitleRow01).Style = StyleConstants.TitleRequired;
            _cells.GetCell(ResultColumn01X, TitleRow01).Style = StyleConstants.TitleResult;

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

            var taxGroupCodeList = GetTaxGroupCodeList().Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
            _cells.SetValidationList(_cells.GetCell(17, TitleRow01 + 1), taxGroupCodeList, ValidationSheetName, 1, false);

            var taxCodeList = GetTaxCodeList().Select(item => item.TaxCode + " - " + item.TaxDescription).ToList();
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

            var nominaParameters = cc.Read<NominaParameters>(filePath, inputFileDescription);

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
                catch (Exception error)
                {
                    MessageBox.Show("Error: " + error.Message);
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
                    Debugger.LogError("RibbonEllipse.cs:ReviewInternalText()", ex.Message);
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
                    Debugger.LogError("RibbonEllipse.cs:ReviewInternalText()", ex.Message);
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

            var cesantiasParameters = cc.Read<CesantiasParameters>(filePath, inputFileDescription);

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
                catch (Exception error)
                {
                    MessageBox.Show("Error: " + error.Message);
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
            var currentRow = TitleRow01 + 1;
            var supplierInfo = new SupplierInfo();

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
                    var supplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    supplierInfo = new SupplierInfo(supplierNo, drpEnviroment.SelectedItem.Label);
                    _cells.GetCell(14, currentRow).Select();

                    _cells.GetCell(17, currentRow).Value = supplierInfo.SupplierNo;
                    _cells.GetCell(18, currentRow).Value = supplierInfo.AccountName.Substring(2, 4);
                    _cells.GetCell(19, currentRow).Value = supplierInfo.AccountNo;
                    _cells.GetCell(20, currentRow).Value = supplierInfo.StAdress;
                    _cells.GetCell(21, currentRow).Value = supplierInfo.StBusiness;
                    _cells.GetCell(22, currentRow).Value = supplierInfo.Status;

                    _cells.GetCell(12, currentRow).Style =
                        _cells.GetCell(18, currentRow).Style =
                            _cells.GetStyle(supplierInfo.AccountName.Substring(2, 4) == _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value)
                                ? StyleConstants.Success
                                : StyleConstants.Error);


                    _cells.GetCell(13, currentRow).Style =
                        _cells.GetCell(19, currentRow).Style =
                            _cells.GetStyle(supplierInfo.AccountNo == _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value)
                                ? StyleConstants.Success
                                : StyleConstants.Error);
                    _cells.GetCell(ResultColumn01C, currentRow).Value = supplierInfo.Error;
                }
                catch (Exception ex)
                {
                    supplierInfo.Error = ex.Message;
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
                    var valorItem = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal);
                    var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);
                    var valorImpuesto = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal);

                    var listTaxes = new List<string>();
                    var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));

                    var groupTaxCodeList = GetTaxCodeList(groupTaxCode);
                    if (groupTaxCodeList != null && groupTaxCodeList.Count > 0)
                        foreach (var taxItem in groupTaxCodeList)
                            listTaxes.Add(taxItem.TaxCode);

                    //additional taxes
                    var taxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                    if (!string.IsNullOrWhiteSpace(taxCodeList) && taxCodeList.Contains(";"))
                    {
                        var splitArray = taxCodeList.Split(';');
                        foreach (var item in splitArray)
                            listTaxes.Add(item);
                    }
                    else if (!string.IsNullOrWhiteSpace(taxCodeList))
                    {
                        taxCodeList = MyUtilities.GetCodeKey(taxCodeList);
                        listTaxes.Add(taxCodeList);
                    }
                    //
                    if (listTaxes.Count != listTaxes.Distinct().Count())
                        throw new Exception("Impuesto Duplicado");

                    var listTaxItems = GetTaxCodeList(listTaxes);

                    decimal calculatedTaxValue = 0;
                    foreach (var tax in listTaxItems)
                    {

                        decimal taxValueItem = valorItem * (tax.TaxRatePerc / 100);
                        if (MyUtilities.IsTrue(tax.Deduct))
                            taxValueItem = taxValueItem * -1;

                        calculatedTaxValue += taxValueItem;
                    }
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
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var nominaInfo = new Nomina
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
                        BranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                        BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value),
                        ItemValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                        AuthorizedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value),
                        Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, currentRow).Value),
                        WorkOrderProjectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, currentRow).Value),
                        WorkOrderProjectIndicator = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, currentRow).Value),


                    };

                    const int itemIndexValue = 16;
                    const int taxIndexValue = 6;
                    const int groupTaxIndexValue = 7;
                    const int additionalTaxIndexValue = 8;

                    var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                    var valorItem = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal);
                    var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);
                    var valorImpuesto = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal);

                    var listTaxes = new List<string>();
                    var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));

                    var groupTaxCodeList = GetTaxCodeList(groupTaxCode);
                    if (groupTaxCodeList != null && groupTaxCodeList.Count > 0)
                        foreach (var taxItem in groupTaxCodeList)
                            listTaxes.Add(taxItem.TaxCode);

                    //additional taxes
                    var taxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                    if (!string.IsNullOrWhiteSpace(taxCodeList) && taxCodeList.Contains(";"))
                    {
                        var splitArray = taxCodeList.Split(';');
                        foreach (var item in splitArray)
                            listTaxes.Add(item);
                    }
                    else if (!string.IsNullOrWhiteSpace(taxCodeList))
                    {
                        taxCodeList = MyUtilities.GetCodeKey(taxCodeList);
                        listTaxes.Add(taxCodeList);
                    }
                    //
                    if (listTaxes.Count != listTaxes.Distinct().Count())
                        throw new Exception("Impuesto Duplicado");

                    var listTaxItems = GetTaxCodeList(listTaxes);

                    decimal calculatedTaxValue = 0;
                    foreach (var tax in listTaxItems)
                    {

                        decimal taxValueItem = valorItem * (tax.TaxRatePerc / 100);
                        if (MyUtilities.IsTrue(tax.Deduct))
                            taxValueItem = taxValueItem * -1;

                        calculatedTaxValue += taxValueItem;
                    }

                    if (valorImpuesto != 0 && valorImpuesto != calculatedTaxValue)
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    if (valorImpuesto == 0 && calculatedTaxValue != 0)
                    {
                        valorImpuesto = calculatedTaxValue;
                        _cells.GetCell(taxIndexValue, currentRow).Value = valorImpuesto;
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    }
                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");

                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);

                    var responseDto = _eFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Abrimos la pantalla
                    var requestXml = "<interaction>" +
                                     "   <actions>" +
                                     "       <action>" +
                                     "           <name>initialScreen</name>" +
                                     "           <data>" +
                                     "               <screenName>MSO265</screenName>" +
                                     "           </data>" +
                                     "           <id>" + Util.GetNewOperationId() + "</id>" +
                                     "           </action>" +
                                     "   </actions>" +
                                     "   <connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                     "   <application>ServiceInteraction</application>" +
                                     "   <applicationPage>unknown</applicationPage>" +
                                     "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Ingresamos la información principal
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("No se ha podido ingresar al programa MSO265");
                    requestXml = "<interaction>                                                     ";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<inputs>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DSTRCT_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>SUPPLIER_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.SupplierNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (string.IsNullOrWhiteSpace(nominaInfo.SupplierNo) && !string.IsNullOrWhiteSpace(nominaInfo.SupplierMnemonic))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>MNEMONIC1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.SupplierMnemonic + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.InvoiceNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "                 <screenField>";
                    requestXml = requestXml + "                 	<name>INV_AMT1I</name>";
                    requestXml = requestXml + "                 	<value>" + nominaInfo.InvoiceAmount + "</value>";
                    requestXml = requestXml + "                 </screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>CURRENCY_TYPE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Currency + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>HANDLE_CDE1I</name>";
                    requestXml = requestXml + "						<value>PN</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.InvoiceDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_RCPT_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.InvoiceReceivedDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DUE_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.DueDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BRANCH_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.BranchCode + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BANK_ACCT_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.BankAccountNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_ITEM_DESC1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Description + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I1</name>";
                    requestXml = requestXml + "					    <value>" + nominaInfo.ItemValue + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I1</name>";
                    requestXml = requestXml + "					   	<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>AUTH_BY1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.AuthorizedBy + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNT1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Account + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>WORK_ORDER1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.WorkOrderProjectNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>WORK_PROJ_IND1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.WorkOrderProjectIndicator + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (valorImpuesto > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ADD_TAX_AMOUNT1I</name>";
                        requestXml = requestXml + "						<value>" + valorImpuesto + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACTION1I1</name>";
                        requestXml = requestXml + "						<value>T</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    requestXml = requestXml + "				</inputs>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>                                                       ";
                    requestXml = requestXml + "	<chains/>                                                        ";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                    requestXml = requestXml + "</interaction>                                                    ";

                    requestXml = requestXml.Replace("&", "&amp;");
                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    
                    //Pantalla de información del proveedor a la que ingresa internamente por el MNEMONIC / Cedula
                    if (string.IsNullOrWhiteSpace(nominaInfo.SupplierNo) && !string.IsNullOrWhiteSpace(nominaInfo.SupplierMnemonic))
                    {
                        if (!responseDto.ResponseString.Contains("MSM202A"))
                            throw new Exception("Se ha producido un error al intentar validar la información del Supplier");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        requestXml = requestXml + "					<screenField> ";
                        requestXml = requestXml + "						<name>SUP_MNEMONIC1I</name> ";
                        requestXml = requestXml + "						<value>" + nominaInfo.SupplierMnemonic + "</value> ";
                        requestXml = requestXml + "					</screenField> ";
                        requestXml = requestXml + "					<screenField> ";
                        requestXml = requestXml + "						<name>SUP_STATUS_IND1I</name> ";
                        requestXml = requestXml + "						<value>A</value> ";
                        requestXml = requestXml + "					</screenField> ";
                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM202A</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    }
                    // - supplier selection

                    //Pantalla de Impuestos
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        var taxIndex = 1;
                        foreach (var tax in listTaxes)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value>" + tax + "</value> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }
                        while (taxIndex <= 12)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value/> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }

                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);
                        
                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                        //confirmación impuestos
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);
                        
                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    }
                    //

                    //Pantalla de confirmación inicial
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Pantalla de confirmación final
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    _cells.GetCell(ResultColumn01X, currentRow).Select();
                    _cells.GetCell(ResultColumn01X, currentRow).Value = "Creado";
                    _cells.GetRange(1, currentRow, ResultColumn01X, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01X, currentRow).Select();
                    _cells.GetCell(ResultColumn01X, currentRow).Value = ex.Message;
                    _cells.GetCell(ResultColumn01X, currentRow).Style = StyleConstants.Error;
                    _cells.GetRange(1, currentRow, ResultColumn01X, currentRow).Style = StyleConstants.Error;
                }
                finally
                {
                    currentRow++;
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
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var fechaFactura = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value), "MMddyy", CultureInfo.InvariantCulture);

                    var fechaPago = DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value), "MMddyy", CultureInfo.InvariantCulture);

                    var nominaInfo = new Nomina
                    {
                        BranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        BankAccountNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        SupplierMnemonic = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Currency = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        InvoiceNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        InvoiceDate = fechaFactura.ToString("yyyyMMdd"),
                        DueDate = fechaPago.ToString("yyyyMMdd"),
                        InvoiceAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        Ref = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value),
                        ItemValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                        Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        AuthorizedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)
                    };

                    const int itemIndexValue = 13;
                    const int taxIndexValue = 16;
                    const int groupTaxIndexValue = 17;
                    const int additionalTaxIndexValue = 18;

                    var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                    var valorItem = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal);
                    var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);
                    var valorImpuesto = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal);

                    var listTaxes = new List<string>();
                    var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));

                    var groupTaxCodeList = GetTaxCodeList(groupTaxCode);
                    if (groupTaxCodeList != null && groupTaxCodeList.Count > 0)
                        foreach (var taxItem in groupTaxCodeList)
                            listTaxes.Add(taxItem.TaxCode);

                    //additional taxes
                    var taxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                    if (!string.IsNullOrWhiteSpace(taxCodeList) && taxCodeList.Contains(";"))
                    {
                        var splitArray = taxCodeList.Split(';');
                        foreach (var item in splitArray)
                            listTaxes.Add(item);
                    }
                    else if (!string.IsNullOrWhiteSpace(taxCodeList))
                    {
                        taxCodeList = MyUtilities.GetCodeKey(taxCodeList);
                        listTaxes.Add(taxCodeList);
                    }
                    //
                    if (listTaxes.Count != listTaxes.Distinct().Count())
                        throw new Exception("Impuesto Duplicado");

                    var listTaxItems = GetTaxCodeList(listTaxes);

                    decimal calculatedTaxValue = 0;
                    foreach (var tax in listTaxItems)
                    {

                        decimal taxValueItem = valorItem * (tax.TaxRatePerc / 100);
                        if (MyUtilities.IsTrue(tax.Deduct))
                            taxValueItem = taxValueItem * -1;

                        calculatedTaxValue += taxValueItem;
                    }

                    if (valorImpuesto != 0 && valorImpuesto != calculatedTaxValue)
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    if (valorImpuesto == 0 && calculatedTaxValue != 0)
                    {
                        valorImpuesto = calculatedTaxValue;
                        _cells.GetCell(taxIndexValue, currentRow).Value = valorImpuesto;
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    }
                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");

                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);

                    var responseDto = _eFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Abrimos la pantalla
                    var requestXml = "<interaction>" +
                                     "   <actions>" +
                                     "       <action>" +
                                     "           <name>initialScreen</name>" +
                                     "           <data>" +
                                     "               <screenName>MSO265</screenName>" +
                                     "           </data>" +
                                     "           <id>" + Util.GetNewOperationId() + "</id>" +
                                     "           </action>" +
                                     "   </actions>" +
                                     "   <connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                     "   <application>ServiceInteraction</application>" +
                                     "   <applicationPage>unknown</applicationPage>" +
                                     "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Ingresamos la información principal
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("No se ha podido ingresar al programa MSO265");
                    requestXml = "<interaction>                                                     ";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<inputs>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DSTRCT_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>SUPPLIER_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.SupplierNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (string.IsNullOrWhiteSpace(nominaInfo.SupplierNo) && !string.IsNullOrWhiteSpace(nominaInfo.SupplierMnemonic))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>MNEMONIC1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.SupplierMnemonic + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.InvoiceNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "                 <screenField>";
                    requestXml = requestXml + "                 	<name>INV_AMT1I</name>";
                    requestXml = requestXml + "                 	<value>" + nominaInfo.InvoiceAmount + "</value>";
                    requestXml = requestXml + "                 </screenField>";
                    requestXml = requestXml + "                 <screenField>";
                    requestXml = requestXml + "                 	<name>ACCOUNTANT1I</name>";
                    requestXml = requestXml + "                 	<value>" + nominaInfo.Accountant + "</value>";
                    requestXml = requestXml + "                 </screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>CURRENCY_TYPE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Currency + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>HANDLE_CDE1I</name>";
                    requestXml = requestXml + "						<value>PN</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.InvoiceDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DUE_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.DueDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BRANCH_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.BranchCode + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BANK_ACCT_NO1I</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.BankAccountNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_ITEM_DESC1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Description + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I1</name>";
                    requestXml = requestXml + "					    <value>" + nominaInfo.ItemValue + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I1</name>";
                    requestXml = requestXml + "					   	<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>AUTH_BY1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.AuthorizedBy + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNT1I1</name>";
                    requestXml = requestXml + "						<value>" + nominaInfo.Account + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (valorImpuesto > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ADD_TAX_AMOUNT1I</name>";
                        requestXml = requestXml + "						<value>" + valorImpuesto + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACTION1I1</name>";
                        requestXml = requestXml + "						<value>T</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    requestXml = requestXml + "				</inputs>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>                                                       ";
                    requestXml = requestXml + "	<chains/>                                                        ";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                    requestXml = requestXml + "</interaction>                                                    ";

                    requestXml = requestXml.Replace("&", "&amp;");
                    responseDto = _eFunctions.ExecutePostRequest(requestXml);

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Pantalla de información del proveedor a la que ingresa internamente por el MNEMONIC / Cedula
                    if (string.IsNullOrWhiteSpace(nominaInfo.SupplierNo) && !string.IsNullOrWhiteSpace(nominaInfo.SupplierMnemonic))
                    {
                        if (!responseDto.ResponseString.Contains("MSM202A"))
                            throw new Exception("Se ha producido un error al intentar validar la información del Supplier");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        requestXml = requestXml + "					<screenField> ";
                        requestXml = requestXml + "						<name>SUP_MNEMONIC1I</name> ";
                        requestXml = requestXml + "						<value>" + nominaInfo.SupplierMnemonic + "</value> ";
                        requestXml = requestXml + "					</screenField> ";
                        requestXml = requestXml + "					<screenField> ";
                        requestXml = requestXml + "						<name>SUP_STATUS_IND1I</name> ";
                        requestXml = requestXml + "						<value>A</value> ";
                        requestXml = requestXml + "					</screenField> ";
                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM202A</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    }
                    // - supplier selection

                    //Pantalla de Impuestos
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        var taxIndex = 1;
                        foreach (var tax in listTaxes)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value>" + tax + "</value> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }
                        while (taxIndex <= 12)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value/> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }

                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                        //confirmación impuestos
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    }
                    //

                    //Pantalla de confirmación inicial
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Pantalla de confirmación final
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    _cells.GetCell(ResultColumn01X, currentRow).Select();
                    _cells.GetCell(ResultColumn01X, currentRow).Value = "Creado";
                    _cells.GetRange(1, currentRow, ResultColumn01X, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01X, currentRow).Select();
                    _cells.GetCell(ResultColumn01X, currentRow).Value = ex.Message;
                    _cells.GetCell(ResultColumn01X, currentRow).Style = StyleConstants.Error;
                    _cells.GetRange(1, currentRow, ResultColumn01X, currentRow).Style = StyleConstants.Error;
                }
                finally
                {
                    currentRow++;
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
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var supplierInfo = new SupplierInfo
                    {
                        SupplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value),
                        Accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value),
                        InvNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        InvDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        CurrencyType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value),
                        InvAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        InvItemDesc = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        InvItemValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value),
                        AuthBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        Account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        BranchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        BankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)
                    };

                    const int itemIndexValue = 9;
                    const int taxIndexValue = 23;
                    const int groupTaxIndexValue = 24;
                    const int additionalTaxIndexValue = 25;

                    var valorItemString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(itemIndexValue, currentRow).Value);
                    var valorItem = !string.IsNullOrWhiteSpace(valorItemString) ? Convert.ToDecimal(valorItemString) : default(decimal);
                    var valorImpuestoString = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(taxIndexValue, currentRow).Value);
                    var valorImpuesto = !string.IsNullOrWhiteSpace(valorImpuestoString) ? Convert.ToDecimal(valorImpuestoString) : default(decimal);

                    var listTaxes = new List<string>();
                    var groupTaxCode = "" + MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(groupTaxIndexValue, currentRow).Value));

                    var groupTaxCodeList = GetTaxCodeList(groupTaxCode);
                    if (groupTaxCodeList != null && groupTaxCodeList.Count > 0)
                        foreach (var taxItem in groupTaxCodeList)
                            listTaxes.Add(taxItem.TaxCode);

                    //additional taxes
                    var taxCodeList = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(additionalTaxIndexValue, currentRow).Value);

                    if (!string.IsNullOrWhiteSpace(taxCodeList) && taxCodeList.Contains(";"))
                    {
                        var splitArray = taxCodeList.Split(';');
                        foreach (var item in splitArray)
                            listTaxes.Add(item);
                    }
                    else if (!string.IsNullOrWhiteSpace(taxCodeList))
                    {
                        taxCodeList = MyUtilities.GetCodeKey(taxCodeList);
                        listTaxes.Add(taxCodeList);
                    }
                    //
                    if (listTaxes.Count != listTaxes.Distinct().Count())
                        throw new Exception("Impuesto Duplicado");

                    var listTaxItems = GetTaxCodeList(listTaxes);

                    decimal calculatedTaxValue = 0;
                    foreach (var tax in listTaxItems)
                    {

                        decimal taxValueItem = valorItem * (tax.TaxRatePerc / 100);
                        if (MyUtilities.IsTrue(tax.Deduct))
                            taxValueItem = taxValueItem * -1;

                        calculatedTaxValue += taxValueItem;
                    }

                    if (valorImpuesto != 0 && valorImpuesto != calculatedTaxValue)
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    if (valorImpuesto == 0 && calculatedTaxValue != 0)
                    {
                        valorImpuesto = calculatedTaxValue;
                        _cells.GetCell(taxIndexValue, currentRow).Value = valorImpuesto;
                        _cells.GetCell(taxIndexValue, currentRow).Style = StyleConstants.Warning;
                    }
                    var urlEnviroment = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");

                    _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);

                    var responseDto = _eFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Abrimos la pantalla
                    var requestXml = "<interaction>" +
                                     "   <actions>" +
                                     "       <action>" +
                                     "           <name>initialScreen</name>" +
                                     "           <data>" +
                                     "               <screenName>MSO265</screenName>" +
                                     "           </data>" +
                                     "           <id>" + Util.GetNewOperationId() + "</id>" +
                                     "           </action>" +
                                     "   </actions>" +
                                     "   <connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                     "   <application>ServiceInteraction</application>" +
                                     "   <applicationPage>unknown</applicationPage>" +
                                     "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Ingresamos la información principal
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("No se ha podido ingresar al programa MSO265");
                    requestXml = "<interaction>                                                     ";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<inputs>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DSTRCT_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>SUPPLIER_NO1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.SupplierNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_NO1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.InvNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "                 <screenField>";
                    requestXml = requestXml + "                 	<name>INV_AMT1I</name>";
                    requestXml = requestXml + "                 	<value>" + supplierInfo.InvAmount + "</value>";
                    requestXml = requestXml + "                 </screenField>";
                    requestXml = requestXml + "                 <screenField>";
                    requestXml = requestXml + "                 	<name>ACCOUNTANT1I</name>";
                    requestXml = requestXml + "                 	<value>" + supplierInfo.Accountant + "</value>";
                    requestXml = requestXml + "                 </screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>CURRENCY_TYPE1I</name>";
                    requestXml = requestXml + "						<value>PES</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>HANDLE_CDE1I</name>";
                    requestXml = requestXml + "						<value>PN</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.InvDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>DUE_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.DueDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BRANCH_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.BranchCode + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>BANK_ACCT_NO1I</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.BankAccount + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_ITEM_DESC1I1</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.InvItemDesc + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I1</name>";
                    requestXml = requestXml + "					    <value>" + supplierInfo.InvItemValue + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I1</name>";
                    requestXml = requestXml + "					   	<value>" + _frmAuth.EllipseDsct + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>AUTH_BY1I1</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.AuthBy + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNT1I1</name>";
                    requestXml = requestXml + "						<value>" + supplierInfo.Account + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (valorImpuesto > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ADD_TAX_AMOUNT1I</name>";
                        requestXml = requestXml + "						<value>" + valorImpuesto + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACTION1I1</name>";
                        requestXml = requestXml + "						<value>T</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    requestXml = requestXml + "				</inputs>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>                                                       ";
                    requestXml = requestXml + "	<chains/>                                                        ";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                    requestXml = requestXml + "</interaction>                                                    ";

                    requestXml = requestXml.Replace("&", "&amp;");
                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Pantalla de Impuestos
                    if (listTaxes != null && listTaxes.Count > 0)
                    {
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        var taxIndex = 1;
                        foreach (var tax in listTaxes)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value>" + tax + "</value> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }
                        while (taxIndex <= 12)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value/> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }

                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);
                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                        //confirmación impuestos
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = _eFunctions.ExecutePostRequest(requestXml);
                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                    }
                    //

                    //Pantalla de confirmación inicial
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    //Pantalla de confirmación final
                    if (!responseDto.ResponseString.Contains("MSM265A"))
                        throw new Exception("Se ha producido un error al intentar completar el proceso");
                    requestXml = "<interaction>";
                    requestXml = requestXml + "	<actions>";
                    requestXml = requestXml + "		<action>";
                    requestXml = requestXml + "			<name>submitScreen</name>";
                    requestXml = requestXml + "			<data>";
                    requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                    requestXml = requestXml + "			</data>";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                    requestXml = requestXml + "		</action>";
                    requestXml = requestXml + "	</actions>";
                    requestXml = requestXml + "	<connectionId>" + _eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                    requestXml = requestXml + "</interaction>";

                    responseDto = _eFunctions.ExecutePostRequest(requestXml);
                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                    _cells.GetCell(ResultColumn01N, currentRow).Select();
                    _cells.GetCell(ResultColumn01N, currentRow).Value = "Creado";
                    _cells.GetRange(1, currentRow, ResultColumn01N, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01N, currentRow).Select();
                    _cells.GetCell(ResultColumn01N, currentRow).Value = ex.Message;
                    _cells.GetCell(ResultColumn01N, currentRow).Style = StyleConstants.Error;
                    _cells.GetRange(1, currentRow, ResultColumn01N, currentRow).Style = StyleConstants.Error;
                }
                finally
                {
                    currentRow++;
                }
            }
            _cells.SetCursorDefault();
        }

        private class Nomina
        {
            public string BranchCode { get; set; }
            public string BankAccountNo { get; set; }
            public string Accountant { get; set; }
            public string SupplierNo { get; set; }
            public string SupplierMnemonic { get; set; }
            public string Currency { get; set; }
            public string InvoiceNo { get; set; }
            public string InvoiceAmount { get; set; }
            public string Description { get; set; }
            public string ItemValue { get; set; }
            public string Account { get; set; }
            public string AuthorizedBy { get; set; }
            public string District { get; set; }
            public string InvoiceDate { get; set; }
            public string InvoiceReceivedDate { get; set; }
            public string DueDate { get; set; }
            public string WorkOrderProjectNo { get; set; }
            public string WorkOrderProjectIndicator { get; set; }

            public string Ref { get; set; }
        }

        public class SupplierInfo
        {
            public SupplierInfo()
            {
            }

            public SupplierInfo(string supplier, string enviroment)
            {
                var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", supplier, _eFunctions.dbReference,
                    _eFunctions.dbLink);
                _eFunctions.SetDBSettings(enviroment);

                var drSupplierInfo = _eFunctions.GetQueryResult(sqlQuery);

                if (!drSupplierInfo.Read())
                {
                    Error = "No existen datos";
                }
                else
                {
                    var cant = Convert.ToInt16(drSupplierInfo["CANTIDAD_REGISTROS"].ToString());
                    if (cant > 1)
                    {
                        Error = "Más de un Supplier Activo";
                    }
                    else
                    {
                        SupplierNo = drSupplierInfo["SUPPLIER_NO"].ToString();
                        TaxFileNo = drSupplierInfo["TAX_FILE_NO"].ToString();
                        StAdress = drSupplierInfo["ST_ADRESS"].ToString();
                        StBusiness = drSupplierInfo["ST_BUSINESS"].ToString();
                        SupplierName = drSupplierInfo["SUPPLIER_NAME"].ToString();
                        CurrencyType = drSupplierInfo["CURRENCY_TYPE"].ToString();
                        AccountName = drSupplierInfo["BANK_ACCT_NAME"].ToString();
                        AccountNo = drSupplierInfo["BANK_ACCT_NO"].ToString();
                        StAdress = drSupplierInfo["ST_ADRESS"].ToString();
                        StBusiness = drSupplierInfo["ST_BUSINESS"].ToString();
                        Status = drSupplierInfo["SUP_STATUS"].ToString();
                        Error = "Success";
                    }

                }
            }

            public string SupplierNo { get; set; }
            public string InvNo { get; set; }
            public string InvAmount { get; set; }
            public string Accountant { get; set; }
            public string CurrencyType { get; set; }
            public string InvDate { get; set; }
            public string DueDate { get; set; }
            public string BranchCode { get; set; }
            public string BankAccount { get; set; }
            public string InvItemDesc { get; set; }
            public string InvItemValue { get; set; }
            public string AuthBy { get; set; }
            public string Account { get; set; }
            public string TaxFileNo { get; set; }
            public string StAdress { get; set; }
            public string StBusiness { get; set; }
            public string SupplierName { get; set; }
            public string AccountName { get; set; }
            public string AccountNo { get; set; }
            public string Error { get; set; }
            public string Status { get; set; }
        }

        public class CesantiasParameters
        {
            [CsvColumn(FieldIndex = 1)]
            public string SupplierMnemonic { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string SupplierName{ get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string Reference { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string Description { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string InvoiceDate { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string DueDate { get; set; }

            [CsvColumn(FieldIndex = 7)]
            public string Account { get; set; }

            [CsvColumn(FieldIndex = 8)]
            public string Currency { get; set; }

            [CsvColumn(FieldIndex = 9)]
            public string ItemValue { get; set; }

            [CsvColumn(FieldIndex = 10)]
            public string InvoiceAmount { get; set; }

            [CsvColumn(FieldIndex = 11)]
            public string AuthorizedBy { get; set; }

            [CsvColumn(FieldIndex = 12)]
            public string BranchCode { get; set; }

            [CsvColumn(FieldIndex = 13)]
            public string BankAccount { get; set; }
        }

        public static class Queries
        {
            public static string GetSupplierInvoiceInfo(string districtCode, string supplierNo, string dbReference,
                string dbLink)
            {
                var sqlQuery = "SELECT " +
                               "   TRIM(A.SUPPLIER_NO) SUPPLIER_NO, " +
                               "   TRIM(B.TAX_FILE_NO) TAX_FILE_NO, " +
                               "   TRIM(A.SUP_STATUS) ST_ADRESS, " +
                               "   TRIM(B.SUP_STATUS) ST_BUSINESS, " +
                               "   TRIM(A.SUPPLIER_NAME) SUPPLIER_NAME, " +
                               "   TRIM(A.CURRENCY_TYPE) CURRENCY_TYPE, " +
                               "   TRIM(B.BANK_ACCT_NAME) BANK_ACCT_NAME, " +
                               "   TRIM(B.BANK_ACCT_NO) BANK_ACCT_NO, " +
                               "   B.SUP_STATUS, " +
                               "   COUNT(B.SUPPLIER_NO) OVER(PARTITION BY B.TAX_FILE_NO) CANTIDAD_REGISTROS " +
                               " FROM " +
                               "   ELLIPSE.MSF200 A " +
                               " INNER JOIN ELLIPSE.MSF203 B " +
                               " ON " +
                               "   A.SUPPLIER_NO = B.SUPPLIER_NO " +
                               " AND B.DSTRCT_CODE = '" + districtCode + "' " +
                               " AND B.TAX_FILE_NO = '" + supplierNo + "' " +
                               " AND B.SUP_STATUS <> 9";
                return sqlQuery;
            }
        }

        public class NominaParameters
        {
            [CsvColumn(FieldIndex = 1)]
            public string BranchCode { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string BankAccount { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string Accountant { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string SupplierNo { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string SupplierMnemonic { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string Currency { get; set; }

            [CsvColumn(FieldIndex = 7)]
            public string InvoiceNo { get; set; }

            [CsvColumn(FieldIndex = 8)]
            public string InvoiceDate { get; set; }

            [CsvColumn(FieldIndex = 9)]
            public string DueDate { get; set; }

            [CsvColumn(FieldIndex = 10)]
            public string InvoiceAmount { get; set; }

            [CsvColumn(FieldIndex = 11)]
            public string Description { get; set; }

            [CsvColumn(FieldIndex = 12)]
            public string Ref { get; set; }

            [CsvColumn(FieldIndex = 13)]
            public string ItemValue { get; set; }

            [CsvColumn(FieldIndex = 14)]
            public string Account { get; set; }

            [CsvColumn(FieldIndex = 15)]
            public string AuthorizedBy { get; set; }

            [CsvColumn(FieldIndex = 16)]
            public string Value01 { get; set; }

            [CsvColumn(FieldIndex = 17)]
            public string Value02 { get; set; }

            [CsvColumn(FieldIndex = 18)]
            public string Value03 { get; set; }

            [CsvColumn(FieldIndex = 19)]
            public string Value04 { get; set; }

            [CsvColumn(FieldIndex = 20)]
            public string Value05 { get; set; }

            [CsvColumn(FieldIndex = 21)]
            public string Value06 { get; set; }

            [CsvColumn(FieldIndex = 22)]
            public string Value07 { get; set; }

            [CsvColumn(FieldIndex = 23)]
            public string Value08 { get; set; }

            [CsvColumn(FieldIndex = 24)]
            public string Value09 { get; set; }

            [CsvColumn(FieldIndex = 25)]
            public string Value10 { get; set; }
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
        
        public List<TaxCodeItem> GetTaxCodeList()
        {
            return GetTaxCodeList(null, null);
        }
        public List<TaxCodeItem> GetTaxCodeList(List<string> taxCodeParamList)
        {
            if (taxCodeParamList == null || !taxCodeParamList.Any())
                return new List<TaxCodeItem>();
            return GetTaxCodeList(taxCodeParamList, null);
        }
        public List<TaxCodeItem> GetTaxCodeList(string taxGroupCode)
        {
            return string.IsNullOrWhiteSpace(taxGroupCode) ? null : GetTaxCodeList(null, taxGroupCode);
        }
        private static List<TaxCodeItem> GetTaxCodeList(List<string> taxCodesParamList, string taxGroupCode)
        {
            var taxList = new List<TaxCodeItem>();

            var paramTaxes = "";
            if (taxCodesParamList != null && taxCodesParamList.Any())
                paramTaxes = " AND TXC.ATAX_CODE IN (" + MyUtilities.GetListInSeparator(taxCodesParamList, ",", "'") + ")";


            var paramGroupIndicator = "";
            paramGroupIndicator = " AND (TRIM(GRP_LEVEL_IND) IS NULL OR TRIM(GRP_LEVEL_IND) = 'N')";

            var conditionalGroup = "";
            var paramGroupCode = "";
            if (!string.IsNullOrWhiteSpace(taxGroupCode))
            {
                conditionalGroup = " JOIN ELLIPSE.MSF014 TXG ON TXG.REL_ATAX_CODE = TXC.ATAX_CODE";
                paramGroupCode = " AND TXG.ATAX_CODE = '" + taxGroupCode + "'";
            }
            var sqlQuery = "SELECT TC.TABLE_CODE, TC.TABLE_DESC, TXC.ATAX_CODE, TXC.DESCRIPTION, TXC.TAX_REF, TXC.ATAX_RATE_9, TXC.DEFAULTED_IND, TXC.DEDUCT_SW" +
                           " FROM ELLIPSE.MSF010 TC JOIN ELLIPSE.MSF013 TXC ON TC.TABLE_CODE = TXC.ATAX_CODE" + conditionalGroup +
                           " WHERE TC.TABLE_TYPE = '+ADD' " +
                           paramGroupIndicator +
                           paramTaxes +
                           paramGroupCode;

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");
            var dataReader = _eFunctions.GetQueryResult(sqlQuery);

            if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
            {
                _eFunctions.CloseConnection();
                return taxList;
            }

            while (dataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage);//!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            _eFunctions.CloseConnection();
            return taxList;
        }
        private static List<TaxCodeItem> GetTaxGroupCodeList(List<string> taxGroupCodeParamList = null)
        {
            var taxList = new List<TaxCodeItem>();

            var paramTaxes = "";
            if (taxGroupCodeParamList != null && taxGroupCodeParamList.Count > 0)
                paramTaxes = " AND TXC.ATAX_CODE IN (" + MyUtilities.GetListInSeparator(taxGroupCodeParamList, ",", "'") + ")";

            var paramGroupIndicator = " AND TRIM(GRP_LEVEL_IND) = 'Y'";

            var sqlQuery = "SELECT TXC.ATAX_CODE, TXC.DESCRIPTION, TXC.TAX_REF, TXC.ATAX_RATE_9, TXC.DEFAULTED_IND, TXC.DEDUCT_SW" +
                           " FROM ELLIPSE.MSF013 TXC " +
                           " WHERE " +
                           paramGroupIndicator +
                           paramTaxes;

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");
            var dataReader = _eFunctions.GetQueryResult(sqlQuery);

            if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
            {
                _eFunctions.CloseConnection();
                return taxList;
            }

            while (dataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage);//!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            _eFunctions.CloseConnection();
            return taxList;
        }

    }

    public class TaxCodeItem
    {
        public string TaxCode;
        public string TaxDescription;
        public string TaxReference;
        public decimal TaxRatePerc;
        public string DefaultToInvoiceItem;
        public string Deduct;
    }
}

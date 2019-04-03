using System;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using EllipseNonOrderInvoiceEntryExcelAddIn.Properties;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseNonOrderInvoiceEntryExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "NonOrder Invoice Entry";
        private Application _excelApp;
        private ExcelStyleCells _cells;
        private const int TittleRow = 7;
        private const int ResultColumn = 18;
        private const int MaxRows = 5000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();


        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            Debugger.DebugErrors = false;
            Debugger.DebugQueries = false;
            Debugger.DebugWarnings = false;

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
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;
            excelSheet.Name = SheetName01;
            _cells = new ExcelStyleCells(_excelApp);

            _cells.GetCell("A1").Value = "CERREJÓN";
            _cells.GetCell("B1").Value = "MSO265 - NonOrder Invoice Entry";

            _cells.GetRange("A1", "B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.GetRange("B1", "D1").Merge();

            _cells.GetCell(1, TittleRow).Value = "District";
            _cells.GetCell(2, TittleRow).Value = "Supplier";
            _cells.GetCell(3, TittleRow).Value = "Mnemonic";
            _cells.GetCell(4, TittleRow).Value = "Invoice No.";
            _cells.GetCell(5, TittleRow).Value = "Invoice AMT";
            _cells.GetCell(6, TittleRow).Value = "Add Tax";
            _cells.GetCell(7, TittleRow).Value = "Currency";
            _cells.GetCell(8, TittleRow).Value = "Invoice date (YYYYMMDD)";
            _cells.GetCell(9, TittleRow).Value = "Invoice received (YYYYMMDD)";
            _cells.GetCell(10, TittleRow).Value = "Due date (YYYYMMDD)";
            _cells.GetCell(11, TittleRow).Value = "Bank branch";
            _cells.GetCell(12, TittleRow).Value = "Account No.";
            _cells.GetCell(13, TittleRow).Value = "Description";
            _cells.GetCell(14, TittleRow).Value = "Value";
            _cells.GetCell(15, TittleRow).Value = "Auth by";
            _cells.GetCell(16, TittleRow).Value = "Account";
            _cells.GetCell(17, TittleRow).Value = "Wo/Proj No.";
            _cells.GetCell(17, TittleRow).Value = "W/P";
            _cells.GetCell(ResultColumn, TittleRow).Value = "Result";
            _cells.GetRange(1, TittleRow, ResultColumn, MaxRows).NumberFormat = "@";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnLoadFile_Click(object sender, RibbonControlEventArgs e)
        {
            LoadFile();
        }

        private void LoadFile()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;

            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearContents();
            _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";

            var openFileDialog = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"Seleccione un archivo de Texto",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\\"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var invoice = cc.Read<Invoice>(filePath, inputFileDescription);

            var currentRow = TittleRow + 1;
            foreach (var inv in invoice)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Value = inv.District;
                    _cells.GetCell(2, currentRow).Value = inv.Supplier;
                    _cells.GetCell(3, currentRow).Value = inv.Mnemonic;
                    _cells.GetCell(4, currentRow).Value = inv.InvoiceNo;
                    _cells.GetCell(5, currentRow).Value = inv.InvoiceAMT;
                    _cells.GetCell(6, currentRow).Value = inv.AddTax;
                    _cells.GetCell(7, currentRow).Value = inv.Currency ;
                    _cells.GetCell(8, currentRow).Value = inv.Invoicedate ;
                    _cells.GetCell(9, currentRow).Value = inv.Invoicereceived;
                    _cells.GetCell(10, currentRow).Value = inv.Duedate ;
                    _cells.GetCell(11, currentRow).Value = inv.Bankbranch ;
                    _cells.GetCell(12, currentRow).Value = inv.AccountNo ;
                    _cells.GetCell(13, currentRow).Value = inv.Description;
                    _cells.GetCell(14, currentRow).Value = inv.Value;
                    _cells.GetCell(15, currentRow).Value = inv.Authby;
                    _cells.GetCell(16, currentRow).Value = inv.WoProjNo;
                    _cells.GetCell(17, currentRow).Value = inv.Wp;
                }
                catch (Exception error)
                {
                    _cells.GetCell(ResultColumn, currentRow).Value = "Error: " + error.Message;
                }
                finally { currentRow++; }
            }

        }

        private class Invoice
        {
            [CsvColumn(FieldIndex = 1)]
            public string District { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string Supplier { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string Mnemonic { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string InvoiceNo { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string InvoiceAMT { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string AddTax { get; set; }

            [CsvColumn(FieldIndex = 7)]
            public string Currency { get; set; }

            [CsvColumn(FieldIndex = 8)]
            public string Invoicedate { get; set; }

            [CsvColumn(FieldIndex = 9)]
            public string Invoicereceived { get; set; }

            [CsvColumn(FieldIndex = 10)]
            public string Duedate { get; set; }

            [CsvColumn(FieldIndex = 11)]
            public string Bankbranch { get; set; }

            [CsvColumn(FieldIndex = 12)]
            public string AccountNo { get; set; }

            [CsvColumn(FieldIndex = 13)]
            public string Description { get; set; }

            [CsvColumn(FieldIndex = 14)]
            public string Value { get; set; }
            
            [CsvColumn(FieldIndex = 15)]
            public string Authby { get; set; }
            
            [CsvColumn(FieldIndex = 16)]
            public string WoProjNo { get; set; }
            
            [CsvColumn(FieldIndex = 17)]
            public string Wp { get; set; }
        }
    }
}

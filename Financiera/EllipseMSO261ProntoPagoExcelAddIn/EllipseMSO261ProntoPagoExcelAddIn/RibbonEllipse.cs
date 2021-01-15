using System;
using System.Globalization;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using EllipseMSO261ProntoPagoExcelAddIn.Properties;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Math = System.Math;
using screen = SharedClassLibrary.Ellipse.ScreenService;


namespace EllipseMSO261ProntoPagoExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int TittleRow = 8;
        private const int ResultColumn = 16;
        private static EllipseFunctions _eFunctions;
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private FormAuthenticate _frmAuth;
        private const string SheetName01 = "MSO261 - Pronto Pago";

        ListObject _excelSheetItems;

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
        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;

            excelSheet.Name = SheetName01;

            #region Instructions

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pronto Pagos";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Datos

            _cells.GetCell(1, 4).Value = "Porcentaje de Descuento";
            _cells.GetCell(1, 5).Value = "Dias";
            _cells.GetCell(1, 6).Value = "Branch Code";
            _cells.GetCell(1, 7).Value = "Bank Account";
            _cells.GetRange(1, 4, 1, 7).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            _cells.GetCell(1, TittleRow).Value = "Supplier";
            _cells.GetCell(2, TittleRow).Value = "Factura";
            _cells.GetCell(3, TittleRow).Value = "Fecha pago solicitada";
            _cells.GetCell(4, TittleRow).Value = "Fecha pago original";
            _cells.GetCell(5, TittleRow).Value = "PMT Status";
            _cells.GetCell(6, TittleRow).Value = "Proveedor";
            _cells.GetCell(7, TittleRow).Value = "Codigo Banco Original";
            _cells.GetCell(8, TittleRow).Value = "ST";
            _cells.GetCell(9, TittleRow).Value = "Vr total factura";
            _cells.GetCell(10, TittleRow).Value = "Vr Base de  Descuento";
            _cells.GetCell(11, TittleRow).Value = "Diferencia";
            _cells.GetCell(12, TittleRow).Value = "Descuento calculado";
            _cells.GetCell(13, TittleRow).Value = "Vr descuento aplicado	";
            _cells.GetCell(14, TittleRow).Value = "Fecha de pago modificada";
            _cells.GetCell(15, TittleRow).Value = "Banco de Pago Modificado";
            _cells.GetCell(ResultColumn, TittleRow).Value = "Result";

            _cells.GetRange(1, TittleRow, ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetCell(2, 4).NumberFormat = "0.00%";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            ImportFile();

            #endregion
        }

        private void btnReloadParameters_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFile();
        }

        private void ImportFile()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != SheetName01) return;

            _cells.GetRange(2, 4, 2, 7).ClearContents();

            var openFileDialog1 = new OpenFileDialog
            {
                Filter = @"Archivos CSV|*.csv",
                FileName = @"Parametros Modify Invoice.csv",
                Title = @"Programa de Lectura",
                InitialDirectory = @"C:\Data\Loaders\Parametros"
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            var filePath = openFileDialog1.FileName;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true,
                EnforceCsvColumnAttribute = true
            };

            var cc = new CsvContext();

            var invoiceParameters = cc.Read<InvoiceParameters>(filePath, inputFileDescription);

            foreach (var p in invoiceParameters)
            {
                try
                {
                    _cells.GetCell(2, 4).Value = p.Percentage;
                    _cells.GetCell(2, 5).Value = p.Days;
                    _cells.GetCell(2, 6).Value = p.Branchcode;
                    _cells.GetCell(2, 7).Value = p.Bankaccount;
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error: " + error.Message);
                }
            }

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetCell(2, 4).NumberFormat = "0.00%";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            MessageBox.Show(PpResources.RibbonEllipse_ImportFile_Parametros_Cargados);
        }


        private void btnGetInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            GetInvoice();
        }

        private void GetInvoice()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            if (excelSheet.Name != SheetName01) return;

            if (drpEnvironment.Label == null || drpEnvironment.Label.Equals("")) return;

            _cells.GetRange(4, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();
            _cells.GetRange(4, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearFormats();

            var currentRow = TittleRow + 1;

            while ((_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value)) != null || (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value)) != null)
            {
                try
                {
                    var supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
                    var factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                    var supplierInfo = SupplierInvoiceInfo.GetSupplierInvoiceInfo(supplier, factura, _eFunctions);

                    if(supplierInfo == null)
                        throw new Exception("Supplier Invoice Not Found");
                    _cells.GetCell(12, currentRow).Select();
//                    _cells.GetCell(1, currentRow).Value = supplierInfo.Supplier;
//                    _cells.GetCell(2, currentRow).Value = supplierInfo.Factura;
                    _cells.GetCell(3, currentRow).Value = supplierInfo.Fechapagosolicitada;
                    _cells.GetCell(4, currentRow).Value = supplierInfo.Fechapagooriginal;
                    _cells.GetCell(5, currentRow).Value = supplierInfo.PmtStatus;
                    _cells.GetCell(6, currentRow).Value = supplierInfo.Proveedor;
                    _cells.GetCell(7, currentRow).Value = supplierInfo.CodigoBancoOriginal;
                    _cells.GetCell(8, currentRow).Value = supplierInfo.St;
                    _cells.GetCell(9, currentRow).Value = supplierInfo.Vrtotalfactura;
                    _cells.GetCell(10, currentRow).Value = supplierInfo.VrBasedeDescuento;
                    _cells.GetCell(11, currentRow).Value = supplierInfo.Diferencia;
                    _cells.GetCell(12, currentRow).Value = supplierInfo.Descuentocalculado;
                    _cells.GetCell(ResultColumn, currentRow).Value = "CONSULTADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetCell(2, 4).NumberFormat = "0.00%";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            _eFunctions.CloseConnection();
            MessageBox.Show(PpResources.RibbonEllipse_GetInvoice_Facturas_Consultadas__Favor_Verificar);
        }


        private void btnCalculateDiscount_Click(object sender, RibbonControlEventArgs e)
        {
            CalculateDiscount();
        }

        private void CalculateDiscount()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            var supplierInfo = new SupplierInvoiceInfo();

            if (excelSheet.Name != SheetName01) return;

            if (drpEnvironment.SelectedItem.Label == null || drpEnvironment.SelectedItem.Label.Equals("")) return;

            var invoiceParameters = new InvoiceParameters
            {
                Percentage = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value)), 
                Days = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value)), 
                Branchcode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value), 
                Bankaccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value)
            };

            _cells.GetRange(11, TittleRow + 1, 12, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();

            var currentRow = TittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null & _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value) != null & _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value) != null & Math.Abs(invoiceParameters.Percentage) > 0 & Math.Abs(invoiceParameters.Days) > 0)
            {

                supplierInfo.Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                supplierInfo.Factura = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
                supplierInfo.Fechapagosolicitada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                supplierInfo.Fechapagooriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                supplierInfo.PmtStatus = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                supplierInfo.Proveedor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                supplierInfo.CodigoBancoOriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                supplierInfo.St = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                supplierInfo.Vrtotalfactura = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));
                supplierInfo.VrBasedeDescuento = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value));
                supplierInfo.Diferencia = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value));
                supplierInfo.Descuentocalculado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value));
                supplierInfo.Vrdescuentoaplicado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value));
                supplierInfo.Fechadepagomodificada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                supplierInfo.BancodePagoModificado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);

                try
                {
                    var fechapagooriginal = DateTime.ParseExact(supplierInfo.Fechapagooriginal, "yyyyMMdd", CultureInfo.InvariantCulture);
                    var fechapagosolicitada = DateTime.ParseExact(supplierInfo.Fechapagosolicitada, "yyyyMMdd", CultureInfo.InvariantCulture);
                    supplierInfo.Diferencia = (fechapagooriginal - fechapagosolicitada).TotalDays;

                    supplierInfo.Descuentocalculado = (supplierInfo.Vrtotalfactura - supplierInfo.VrBasedeDescuento) < 0.000 ? Math.Round(supplierInfo.Diferencia * supplierInfo.Vrtotalfactura * invoiceParameters.Percentage / invoiceParameters.Days) : Math.Round(supplierInfo.Diferencia * supplierInfo.VrBasedeDescuento * invoiceParameters.Percentage / invoiceParameters.Days);
                    _cells.GetCell(11, currentRow).Value = supplierInfo.Diferencia;
                    _cells.GetCell(12, currentRow).Value = supplierInfo.Descuentocalculado;
                    _cells.GetRange(11, currentRow, 12, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
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

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetCell(2, 4).NumberFormat = "0.00%";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            MessageBox.Show(PpResources.RibbonEllipse_CalculateDiscount_Descuentos_Calculados);
            

        }

        private void btnModifyInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                    ModifyInvoice();
            }
            else
                MessageBox.Show(PpResources.RibbonEllipse_btnLoad_Click_Invalid_Format);

        }

        private void ModifyInvoice()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var supplierInfo = new SupplierInvoiceInfo();

            var currentRow = TittleRow + 1;

            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();

            _cells.GetRange(ResultColumn, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();
            _cells.GetRange(ResultColumn, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).Style = _cells.GetStyle(StyleConstants.Normal);

            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            _cells.GetCell(1, currentRow).Select();

            var invoiceParameters = new InvoiceParameters
            {
                Percentage = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value)),
                Days = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value)),
                Branchcode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value),
                Bankaccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value)
            };

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    supplierInfo.Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    supplierInfo.Factura = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
                    supplierInfo.Fechapagosolicitada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                    supplierInfo.Fechapagooriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                    supplierInfo.PmtStatus = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                    supplierInfo.Proveedor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                    supplierInfo.CodigoBancoOriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                    supplierInfo.St = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                    supplierInfo.Vrtotalfactura = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));
                    supplierInfo.VrBasedeDescuento = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value));
                    supplierInfo.Diferencia = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value));
                    supplierInfo.Descuentocalculado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value));
                    supplierInfo.Vrdescuentoaplicado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value));
                    supplierInfo.Fechadepagomodificada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                    supplierInfo.BancodePagoModificado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);

                    _eFunctions.RevertOperation(opSheet, proxySheet);

                    var replySheet = proxySheet.executeScreen(opSheet, "MSO261");

                    if (_eFunctions.CheckReplyError(replySheet))
                        throw new Exception(replySheet.message);
                    
                    if (replySheet.mapName != "MSM261A") return;
                    var arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("OPTION1I", "1");
                    arrayFields.Add("DSTRCT_CODE1I", "ICOR");
                    arrayFields.Add("SUPPLIER_NO1I", supplierInfo.Supplier);
                    arrayFields.Add("INV_NO1I", supplierInfo.Factura);
                    requestSheet.screenFields = arrayFields.ToArray();

                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                    while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (_eFunctions.CheckReplyError(replySheet))
                        throw new Exception(replySheet.message);

                    if (!replySheet.mapName.Equals("MSM261B", StringComparison.InvariantCultureIgnoreCase))
                        throw new Exception("NO SE PUDO ACCEDER AL PROGRAMA MSM261B");

                    arrayFields = new ArrayScreenNameValue();

                    arrayFields.Add("BRANCH_CODE2I", invoiceParameters.Branchcode);
                    arrayFields.Add("BANK_ACCT_NO2I", invoiceParameters.Bankaccount);
                    arrayFields.Add("SD_AMOUNT2I", supplierInfo.Descuentocalculado.ToString(CultureInfo.InvariantCulture));
                    arrayFields.Add("SD_DATE2I", supplierInfo.Fechapagosolicitada);
                    requestSheet.screenFields = arrayFields.ToArray();

                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                    while (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (_eFunctions.CheckReplyError(replySheet))
                        throw new Exception(replySheet.message);

                    if (!replySheet.message.Contains("X2:3730 - MODIFICATIONS MADE TO INVOICE"))
                        throw new Exception(replySheet.message);

                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn, currentRow).Value = "Success";

                    _cells.GetCell(13, currentRow).Value = supplierInfo.Vrdescuentoaplicado = supplierInfo.Descuentocalculado;
                    _cells.GetCell(14, currentRow).Value = supplierInfo.Fechadepagomodificada = supplierInfo.Fechapagosolicitada;
                    _cells.GetCell(15, currentRow).Value = supplierInfo.BancodePagoModificado = invoiceParameters.Branchcode + " - " + invoiceParameters.Bankaccount;

                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn, currentRow).Value = "MODIFICADO";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }

            MessageBox.Show(PpResources.RibbonEllipse_ModifyInvoice_Facturas_Modificadas);
        }


        private void btnVerifyInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            VerifyInvoice();
        }

        private void VerifyInvoice()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (excelSheet.Name != SheetName01) return;

            if (drpEnvironment.Label == null || drpEnvironment.Label.Equals("")) return;

            _cells.GetRange(13, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();

            var currentRow = TittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) != null)
            {
                try
                {
                    var supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
                    var factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                    var supplierInfo = SupplierInvoiceInfo.GetSupplierInvoiceInfo(supplier, factura, _eFunctions);

                    if (supplierInfo == null)
                        throw new Exception("Supplier Invoice Not Found");
                    _cells.GetCell(14, currentRow).Select();
                    _cells.GetCell(13, currentRow).Value = supplierInfo.Vrdescuentoaplicado;
                    _cells.GetCell(14, currentRow).Value = supplierInfo.Fechadepagomodificada;
                    _cells.GetCell(15, currentRow).Value = supplierInfo.BancodePagoModificado;
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn, currentRow).Value = "VERIFICADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
            _eFunctions.CloseConnection();
            MessageBox.Show(PpResources.RibbonEllipse_VerifyInvoice_Datos_de_Facturas_Consuladas__Favor_Verificar);
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
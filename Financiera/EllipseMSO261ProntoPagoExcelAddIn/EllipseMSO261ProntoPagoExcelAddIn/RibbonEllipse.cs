using System;
using System.Globalization;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseMSO261ProntoPagoExcelAddIn.Properties;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;


namespace EllipseMSO261ProntoPagoExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int TittleRow = 8;
        private const int ResultColumn = 16;
        public static EllipseFunctions EFunctions = new EllipseFunctions();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private string _sheetName01 = "Pronto Pagos";
        ListObject _excelSheetItems;
        public InvoiceParameters Parameters;
        public SupplierInvoiceInfo SupplierInfo;

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


            _sheetName01 = "MSO261 - Pronto Pago";
            excelSheet.Name = _sheetName01;

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

            if (excelSheet.Name != _sheetName01) return;

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
                    Parameters = new InvoiceParameters();
                    Parameters = p;

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

            MessageBox.Show(Resources.RibbonEllipse_ImportFile_Parametros_Cargados);


        }


        private void btnGetInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            GetInvoice();
        }

        private void GetInvoice()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            SupplierInfo = new SupplierInvoiceInfo();

            if (excelSheet.Name != _sheetName01) return;

            if (drpEnviroment.Label == null || drpEnviroment.Label.Equals("")) return;

            _cells.GetRange(4, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();
            _cells.GetRange(4, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearFormats();

            var currentRow = TittleRow + 1;

            var supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
            var factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));

            while (supplier != null & factura != null)
            {
                try
                {
                    SupplierInfo = new SupplierInvoiceInfo(supplier, factura, drpEnviroment.SelectedItem.Label);
                    _cells.GetCell(12, currentRow).Select();
//                    _cells.GetCell(1, currentRow).Value = supplierInfo.Supplier;
//                    _cells.GetCell(2, currentRow).Value = supplierInfo.Factura;
                    _cells.GetCell(3, currentRow).Value = SupplierInfo.Fechapagosolicitada;
                    _cells.GetCell(4, currentRow).Value = SupplierInfo.Fechapagooriginal;
                    _cells.GetCell(5, currentRow).Value = SupplierInfo.PmtStatus;
                    _cells.GetCell(6, currentRow).Value = SupplierInfo.Proveedor;
                    _cells.GetCell(7, currentRow).Value = SupplierInfo.CodigoBancoOriginal;
                    _cells.GetCell(8, currentRow).Value = SupplierInfo.St;
                    _cells.GetCell(9, currentRow).Value = SupplierInfo.Vrtotalfactura;
                    _cells.GetCell(10, currentRow).Value = SupplierInfo.VrBasedeDescuento;
                    _cells.GetCell(11, currentRow).Value = SupplierInfo.Diferencia;
                    _cells.GetCell(12, currentRow).Value = SupplierInfo.Descuentocalculado;
                    _cells.GetCell(ResultColumn, currentRow).Value = SupplierInfo.Error;

                }
                catch (Exception ex)
                {
                    if (SupplierInfo != null) SupplierInfo.Error = ex.Message;
                }
                finally
                {
                    currentRow++;
                    supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
                    factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                    SupplierInfo = new SupplierInvoiceInfo(supplier, factura, drpEnviroment.SelectedItem.Label);
                }
            }

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";
            _cells.GetCell(2, 4).NumberFormat = "0.00%";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            MessageBox.Show(Resources.RibbonEllipse_GetInvoice_Facturas_Consultadas__Favor_Verificar);
        }


        private void btnCalculateDiscount_Click(object sender, RibbonControlEventArgs e)
        {
            CalculateDiscount();
        }

        private void CalculateDiscount()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            SupplierInfo = new SupplierInvoiceInfo();

            if (excelSheet.Name != _sheetName01) return;

            if (drpEnviroment.SelectedItem.Label == null || drpEnviroment.SelectedItem.Label.Equals("")) return;

            Parameters.Percentage = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 4).Value));
            Parameters.Days = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 5).Value));
            Parameters.Branchcode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 6).Value);
            Parameters.Bankaccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, 7).Value);

            _cells.GetRange(11, TittleRow + 1, 12, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();

            var currentRow = TittleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null & _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value) != null & _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value) != null & Math.Abs(Parameters.Percentage) > 0 & Math.Abs(Parameters.Days) > 0)
            {

                SupplierInfo.Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                SupplierInfo.Factura = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
                SupplierInfo.Fechapagosolicitada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                SupplierInfo.Fechapagooriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                SupplierInfo.PmtStatus = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                SupplierInfo.Proveedor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                SupplierInfo.CodigoBancoOriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                SupplierInfo.St = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                SupplierInfo.Vrtotalfactura = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));
                SupplierInfo.VrBasedeDescuento = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value));
                SupplierInfo.Diferencia = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value));
                SupplierInfo.Descuentocalculado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value));
                SupplierInfo.Vrdescuentoaplicado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value));
                SupplierInfo.Fechadepagomodificada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                SupplierInfo.BancodePagoModificado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);

                try
                {
                    var fechapagooriginal = DateTime.ParseExact(SupplierInfo.Fechapagooriginal, "yyyyMMdd", CultureInfo.InvariantCulture);
                    var fechapagosolicitada = DateTime.ParseExact(SupplierInfo.Fechapagosolicitada, "yyyyMMdd", CultureInfo.InvariantCulture);
                    SupplierInfo.Diferencia = (fechapagooriginal - fechapagosolicitada).TotalDays;

                    SupplierInfo.Descuentocalculado = (SupplierInfo.Vrtotalfactura - SupplierInfo.VrBasedeDescuento) < 0.000 ? Math.Round(SupplierInfo.Diferencia * SupplierInfo.Vrtotalfactura * Parameters.Percentage / Parameters.Days) : Math.Round(SupplierInfo.Diferencia * SupplierInfo.VrBasedeDescuento * Parameters.Percentage / Parameters.Days);
                    _cells.GetCell(11, currentRow).Value = SupplierInfo.Diferencia;
                    _cells.GetCell(12, currentRow).Value = SupplierInfo.Descuentocalculado;
                    _cells.GetRange(11, currentRow, 12, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                }
                catch (Exception ex)
                {
                    SupplierInfo.Error = ex.Message;
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

            MessageBox.Show(Resources.RibbonEllipse_CalculateDiscount_Descuentos_Calculados);
            

        }

        private void btnModifyInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                    ModifyInvoice();
            }
            else
                MessageBox.Show(Resources.RibbonEllipse_btnLoad_Click_Invalid_Format);

        }

        private void ModifyInvoice()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            SupplierInfo = new SupplierInvoiceInfo();

            var currentRow = TittleRow + 1;

            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();

            _cells.GetRange(ResultColumn, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();
            _cells.GetRange(ResultColumn, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).Style = _cells.GetStyle(StyleConstants.Normal);

            proxySheet.Url = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";

            var opSheet = new screen.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            _cells.GetCell(1, currentRow).Select();

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();
                    SupplierInfo.Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    SupplierInfo.Factura = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value);
                    SupplierInfo.Fechapagosolicitada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                    SupplierInfo.Fechapagooriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                    SupplierInfo.PmtStatus = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                    SupplierInfo.Proveedor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                    SupplierInfo.CodigoBancoOriginal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                    SupplierInfo.St = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value));
                    SupplierInfo.Vrtotalfactura = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value));
                    SupplierInfo.VrBasedeDescuento = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value));
                    SupplierInfo.Diferencia = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value));
                    SupplierInfo.Descuentocalculado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value));
                    SupplierInfo.Vrdescuentoaplicado = Convert.ToDouble(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value));
                    SupplierInfo.Fechadepagomodificada = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                    SupplierInfo.BancodePagoModificado = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);

                    EFunctions.RevertOperation(opSheet, proxySheet);

                    var replySheet = proxySheet.executeScreen(opSheet, "MSO261");

                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM261A") return;
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("OPTION1I", "1");
                        arrayFields.Add("DSTRCT_CODE1I", "ICOR");
                        arrayFields.Add("SUPPLIER_NO1I", SupplierInfo.Supplier);
                        arrayFields.Add("INV_NO1I", SupplierInfo.Factura);
                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (EFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (EFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                        }
                        else if (replySheet.mapName == "MSM261B")
                        {
                            arrayFields = new ArrayScreenNameValue();

                            arrayFields.Add("BRANCH_CODE2I", Parameters.Branchcode);
                            arrayFields.Add("BANK_ACCT_NO2I", Parameters.Bankaccount);
                            arrayFields.Add("SD_AMOUNT2I", SupplierInfo.Descuentocalculado.ToString(CultureInfo.InvariantCulture));
                            arrayFields.Add("SD_DATE2I", SupplierInfo.Fechapagosolicitada);
                            requestSheet.screenFields = arrayFields.ToArray();

                            requestSheet.screenKey = "1";
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                            while (EFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys.Contains("XMIT-Confirm"))
                                replySheet = proxySheet.submit(opSheet, requestSheet);

                            if (EFunctions.CheckReplyError(replySheet) & !replySheet.message.Contains("X2:3730 - MODIFICATIONS MADE TO INVOICE"))
                            {
                                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                                _cells.GetCell(ResultColumn, currentRow).Value = "Success";

                                _cells.GetCell(13, currentRow).Value = SupplierInfo.Vrdescuentoaplicado = SupplierInfo.Descuentocalculado;
                                _cells.GetCell(14, currentRow).Value = SupplierInfo.Fechadepagomodificada = SupplierInfo.Fechapagosolicitada;
                                _cells.GetCell(15, currentRow).Value = SupplierInfo.BancodePagoModificado = Parameters.Branchcode + " - " + Parameters.Bankaccount;
                            }
                        }
                    }
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

            MessageBox.Show(Resources.RibbonEllipse_ModifyInvoice_Facturas_Modificadas);
        }


        private void btnVerifyInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            VerifyInvoice();
        }

        private void VerifyInvoice()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            SupplierInfo = new SupplierInvoiceInfo();

            if (excelSheet.Name != _sheetName01) return;

            if (drpEnviroment.Label == null || drpEnviroment.Label.Equals("")) return;

            _cells.GetRange(13, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).ClearContents();

            var currentRow = TittleRow + 1;

            var supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
            var factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));

            while (supplier != null & factura != null)
            {
                try
                {
                    SupplierInfo = new SupplierInvoiceInfo(supplier, factura, drpEnviroment.SelectedItem.Label);
                    _cells.GetCell(14, currentRow).Select();

                    _cells.GetCell(13, currentRow).Value = SupplierInfo.Vrdescuentoaplicado;
                    _cells.GetCell(14, currentRow).Value = SupplierInfo.Fechadepagomodificada;
                    _cells.GetCell(15, currentRow).Value = SupplierInfo.BancodePagoModificado;

                }
                catch (Exception ex)
                {
                    if (SupplierInfo != null) SupplierInfo.Error = ex.Message;
                }
                finally
                {
                    currentRow++;
                    supplier = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value));
                    factura = (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value));
                    SupplierInfo = new SupplierInvoiceInfo(supplier, factura, drpEnviroment.SelectedItem.Label);
                }
            }
            MessageBox.Show(Resources.RibbonEllipse_VerifyInvoice_Datos_de_Facturas_Consuladas__Favor_Verificar);
        }

        public class SupplierInvoiceInfo
        {

            public SupplierInvoiceInfo()
            {

            }

            public SupplierInvoiceInfo(string supplier, string factura, string enviroment)
            {
                var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", supplier, factura, EFunctions.dbReference, EFunctions.dbLink);
                EFunctions.SetDBSettings(enviroment);

                var drSupplierInvoiceInfo = EFunctions.GetQueryResult(sqlQuery);

                if (!drSupplierInvoiceInfo.Read())
                {
                    Error = "No existen datos";
                }
                else
                {
                    Supplier = drSupplierInvoiceInfo["SUPPLIER_NO"].ToString();
                    Factura = drSupplierInvoiceInfo["EXT_INV_NO"].ToString();
                    Fechapagosolicitada = null;
                    Fechapagooriginal = drSupplierInvoiceInfo["DUE_DATE"].ToString();
                    PmtStatus = drSupplierInvoiceInfo["PMT_STATUS"].ToString();
                    Proveedor = drSupplierInvoiceInfo["NOM_SUPPLIER"].ToString();
                    CodigoBancoOriginal = drSupplierInvoiceInfo["ORIG_BANK"].ToString();
                    St = Convert.ToDouble(drSupplierInvoiceInfo["NO_OF_DAYS_PAY"].ToString());
                    Vrtotalfactura = Convert.ToDouble(drSupplierInvoiceInfo["VLR_FACTURA"].ToString());
                    VrBasedeDescuento = Convert.ToDouble(drSupplierInvoiceInfo["VRBASE"].ToString());
                    Diferencia = Convert.ToDouble(drSupplierInvoiceInfo["DIREFENCIA"].ToString());
                    Descuentocalculado = Convert.ToDouble(drSupplierInvoiceInfo["VR_OTROS_DESCTS"].ToString());
                    Vrdescuentoaplicado = Convert.ToDouble(drSupplierInvoiceInfo["VR_OTROS_DESCTS"].ToString());
                    Fechadepagomodificada = drSupplierInvoiceInfo["FEC_MOD_PAGO"].ToString();
                    BancodePagoModificado = drSupplierInvoiceInfo["ORIG_BANK"].ToString();
                    Error = "Success";
                }
            }

            public string Supplier { get; set; }
            public string Factura { get; set; }
            public string Fechapagosolicitada { get; set; }
            public string Fechapagooriginal { get; set; }
            public string PmtStatus { get; set; }
            public string Proveedor { get; set; }
            public string CodigoBancoOriginal { get; set; }
            public double St { get; set; }
            public double Vrtotalfactura { get; set; }
            public double VrBasedeDescuento { get; set; }
            public double Diferencia { get; set; }
            public double Descuentocalculado { get; set; }
            public double Vrdescuentoaplicado { get; set; }
            public string Fechadepagomodificada { get; set; }
            public string BancodePagoModificado { get; set; }
            public string Error { get; set; }
        }

        public class InvoiceParameters
        {
            [CsvColumn(FieldIndex = 1)]
            public double X { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public double Y { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public double Z { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string Sitio { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string Tajo { get; set; }


            [CsvColumn(FieldIndex = 4)]
            public string TipoSitio { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string SiNombretio { get; set; }

        }

        public static class Queries
        {
            public static string GetSupplierInvoiceInfo(string districtCode, string supplierNo, string invoiceNo, string dbReference, string dbLink)
            {
                var sqlQuery = "SELECT " +
                               "  INV.SUPPLIER_NO, " +
                               "  INV.EXT_INV_NO, " +
                               "  INV.SD_DATE FEC_MOD_PAGO, " +
                               "  INV.DUE_DATE, " +
                               "  INV.PMT_STATUS, " +
                               "  SUP.SUPPLIER_NAME NOM_SUPPLIER, " +
                               "  TRIM ( INV.BRANCH_CODE ) || '-' || TRIM ( INV.BANK_ACCT_NO ) ORIG_BANK, " +
                               "  SBI.NO_OF_DAYS_PAY, " +
                               "  DECODE ( INV.CURRENCY_TYPE, 'USD ', DECODE ( INV.LOC_INV_AMD, '0', INV.LOC_INV_ORIG, INV.LOC_INV_AMD ), DECODE ( INV.FOR_INV_AMD, '0', INV.FOR_INV_ORIG, INV.FOR_INV_AMD ) ) VLR_FACTURA, " +
                               "  DECODE ( SUBSTR ( INVOICE_LINE_ITEM.INV_ITEM_DESC, 1, 4 ), 'CNT:', INVOICE_LINE_ITEM.FOR_VAL_INVD, DECODE ( INV.CURRENCY_TYPE, 'USD ', DECODE ( INV.LOC_INV_AMD, '0', INV.LOC_INV_ORIG, INV.LOC_INV_AMD ), DECODE ( INV.FOR_INV_AMD, '0', INV.FOR_INV_ORIG, INV.FOR_INV_AMD ) ) - DECODE ( INV.CURRENCY_TYPE, 'PES ', NVL ( 0, INV.ATAX_AMT_FOR ), 0 ) ) VRBASE, " +
                               "  SBI.NO_OF_DAYS_PAY - ( TO_DATE ( DECODE ( TRIM ( INV.SD_DATE ), NULL, INV.INV_RCPT_DATE, INV.SD_DATE ), 'YYYYMMDD' ) - TO_DATE ( INV.INV_RCPT_DATE, 'YYYYMMDD' ) ) DIREFENCIA, " +
                               "  INV.SD_AMOUNT VR_OTROS_DESCTS " +
                               "FROM " +
                               "  ELLIPSE.MSF260 INV " +
                               "LEFT JOIN ELLIPSE.MSF203 SBI " +
                               "ON " +
                               "  INV.DSTRCT_CODE = SBI.DSTRCT_CODE " +
                               "AND INV.SUPPLIER_NO = SBI.SUPPLIER_NO " +
                               "LEFT JOIN ELLIPSE.MSF200 SUP " +
                               "ON " +
                               "  SBI.SUPPLIER_NO = SUP.SUPPLIER_NO " +
                               "INNER JOIN ELLIPSE.MSF26A INVOICE_LINE_ITEM " +
                               "ON " +
                               "  INVOICE_LINE_ITEM.SUPPLIER_NO = INV.SUPPLIER_NO " +
                               "AND INVOICE_LINE_ITEM.INV_NO = INV.INV_NO " +
                               "AND INVOICE_LINE_ITEM.DSTRCT_CODE = INV.DSTRCT_CODE " +
                               "AND INVOICE_LINE_ITEM.INV_ITEM_NO = '001' " +
                                "WHERE " +
                                "   INV.DSTRCT_CODE = '" + districtCode + "' " +
                                "AND INV.SUPPLIER_NO = '" + supplierNo + "' " +
                                "AND INV.INV_NO = '" + invoiceNo + "' ";

                return sqlQuery;
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
using System;
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

namespace EllipseMSO265ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int TittleRow = 5;
        private static int _resultColumn = 23;
        public static EllipseFunctions EFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;
        private string _sheetName01 = "MSO265 Cesantias";

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

        private void FormatNomina()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;


            _sheetName01 = "MSO265 Nomina";
            excelSheet.Name = _sheetName01;
            _resultColumn = 16;

            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                _cells.GetRange(1, TittleRow, _resultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            #region Titulo

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pago de Nomina";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Encabezados

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";

            _cells.GetCell(1, TittleRow).Value = "Codigo Banco";
            _cells.GetCell(2, TittleRow).Value = "Cuenta Banco";
            _cells.GetCell(3, TittleRow).Value = "Analista";
            _cells.GetCell(4, TittleRow).Value = "Supplier";
            _cells.GetCell(5, TittleRow).Value = "Cedula";
            _cells.GetCell(6, TittleRow).Value = "Moneda";
            _cells.GetCell(7, TittleRow).Value = "NumFcatura";
            _cells.GetCell(8, TittleRow).Value = "Fecha Factura";
            _cells.GetCell(9, TittleRow).Value = "Fecha Pago";
            _cells.GetCell(10, TittleRow).Value = "Valor Total";
            _cells.GetCell(11, TittleRow).Value = "DESCRIPCION";
            _cells.GetCell(12, TittleRow).Value = "REF";
            _cells.GetCell(13, TittleRow).Value = "Valor Item";
            _cells.GetCell(14, TittleRow).Value = "Cuenta";
            _cells.GetCell(15, TittleRow).Value = "Posicion Aprobador";
            _cells.GetCell(_resultColumn, TittleRow).Value = "Result";

            _cells.GetRange(1, TittleRow, _resultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

            _cells.GetRange(10, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";
            _cells.GetRange(13, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            ImportFileNomina();

            #endregion
        }

        private void ImportFileNomina()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != _sheetName01) return;

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).Delete();

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
            var currentRow = TittleRow + 1;
            foreach (var c in nominaParameters)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Value = c.CodigoBanco;
                    _cells.GetCell(2, currentRow).Value = c.CuentaBanco;
                    _cells.GetCell(3, currentRow).Value = c.Analista;
                    _cells.GetCell(4, currentRow).Value = c.Supplier;
                    _cells.GetCell(5, currentRow).Value = c.Cedula;
                    _cells.GetCell(6, currentRow).Value = c.Moneda;
                    _cells.GetCell(7, currentRow).Value = c.NumFcatura;
                    _cells.GetCell(8, currentRow).Value = c.FechaFactura;
                    _cells.GetCell(9, currentRow).Value = c.FechaPago;
                    _cells.GetCell(10, currentRow).Value = c.ValorTotal;
                    _cells.GetCell(11, currentRow).Value = c.Descripcion;
                    _cells.GetCell(12, currentRow).Value = c.Ref;
                    _cells.GetCell(13, currentRow).Value = c.ValorItem;
                    _cells.GetCell(14, currentRow).Value = c.Cuenta;
                    _cells.GetCell(15, currentRow).Value = c.PosicionAprobador;
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

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";
            _cells.GetRange(10, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";
            _cells.GetRange(13, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();
        }

        private void btnFormatCesantias_Click(object sender, RibbonControlEventArgs e)
        {
            FormatCesantias();
        }

        private void FormatCesantias()
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;


            _sheetName01 = "MSO265 Cesantias";
            excelSheet.Name = _sheetName01;


            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                _cells.GetRange(1, TittleRow, _resultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            #region Instructions

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pago de Cesantias";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Datos

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";

            _cells.GetCell(1, TittleRow).Value = "Cedula";
            _cells.GetCell(2, TittleRow).Value = "Nombre";
            _cells.GetCell(3, TittleRow).Value = "Referencia";
            _cells.GetCell(4, TittleRow).Value = "Descripcion";
            _cells.GetCell(5, TittleRow).Value = "Fecha Factura";
            _cells.GetCell(6, TittleRow).Value = "Fecha Pago";
            _cells.GetCell(7, TittleRow).Value = "Cuenta";
            _cells.GetCell(8, TittleRow).Value = "Moneda";
            _cells.GetCell(9, TittleRow).Value = "Valor Item";
            _cells.GetCell(10, TittleRow).Value = "Valor Total";
            _cells.GetCell(11, TittleRow).Value = "Posicion Aprobador";
            _cells.GetCell(12, TittleRow).Value = "Codigo Banco";
            _cells.GetCell(13, TittleRow).Value = "Cuenta Banco";
            _cells.GetCell(14, TittleRow).Value = "Banco";
            _cells.GetCell(15, TittleRow).Value = "Sucursal Banco";
            _cells.GetCell(16, TittleRow).Value = "Analista";
            _cells.GetCell(17, TittleRow).Value = "Supplier";
            _cells.GetCell(18, TittleRow).Value = "Sucursal Banco Ellipse";
            _cells.GetCell(19, TittleRow).Value = "Cuenta Banco Ellipse";
            _cells.GetCell(20, TittleRow).Value = "ST Adress";
            _cells.GetCell(21, TittleRow).Value = "ST Business";
            _cells.GetCell(22, TittleRow).Value = "ST Status";
            _cells.GetCell(_resultColumn, TittleRow).Value = "Result";

            _cells.GetRange(1, TittleRow, _resultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            ImportFileCesantias();

            #endregion
        }

        private void ImportFileCesantias()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;

            if (excelSheet.Name != _sheetName01) return;

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).Delete();

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
            var currentRow = TittleRow + 1;
            foreach (var c in cesantiasParameters)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Value = c.Cedula;
                    _cells.GetCell(2, currentRow).Value = c.Nombre;
                    _cells.GetCell(3, currentRow).Value = c.Referencia;
                    _cells.GetCell(4, currentRow).Value = c.Descripcion;
                    _cells.GetCell(5, currentRow).Value = c.FechaFactura;
                    _cells.GetCell(6, currentRow).Value = c.FechaPago;
                    _cells.GetCell(7, currentRow).Value = c.Cuenta;
                    _cells.GetCell(8, currentRow).Value = c.Moneda;
                    _cells.GetCell(9, currentRow).Value = c.ValorItem;
                    _cells.GetCell(10, currentRow).Value = c.ValorTotal;
                    _cells.GetCell(11, currentRow).Value = c.PosicionAprobador;
                    _cells.GetCell(12, currentRow).Value = c.CodigoBanco;
                    _cells.GetCell(13, currentRow).Value = c.CuentaBanco;
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

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat =
                "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            ValidateCesantias();
        }

        private void btnValidate_Click(object sender, RibbonControlEventArgs e)
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            switch (excelSheet.Name)
            {
                case "MSO265 Cesantias":
                    ValidateCesantias();
                    break;
            }
        }

        private void ValidateCesantias()
        {
            var currentRow = TittleRow + 1;
            var supplierInfo = new SupplierInfo();

            _cells.GetRange(12, TittleRow + 1, 13, _excelSheetItems.ListRows.Count + TittleRow).Style = _cells.GetStyle(StyleConstants.Normal);
            _cells.GetRange(18, TittleRow + 1, 19, _excelSheetItems.ListRows.Count + TittleRow).Style = _cells.GetStyle(StyleConstants.Normal);

            _cells.GetRange(1, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    var supplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    supplierInfo = new SupplierInfo(supplierNo, drpEnviroment.SelectedItem.Label);
                    _cells.GetCell(14, currentRow).Select();

                    _cells.GetCell(17, currentRow).Value = supplierInfo.SupplierNo;
                    _cells.GetCell(18, currentRow).Value = supplierInfo.AccountName.Substring(2, 4) ?? "";
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
                    _cells.GetCell(_resultColumn, currentRow).Value = supplierInfo.Error;
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
        }

        private void btnReloadParameters_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFileCesantias();
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            var excelBook = _excelApp.ActiveWorkbook;
            Worksheet excelSheet = excelBook.ActiveSheet;
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            switch (excelSheet.Name)
            {
                case "MSO265 Cesantias":
                    //LoadCesantias();
                    LoadCesantiasPost();
                    break;
                case "MSO265 Nomina":
                    LoadNomina();
                    break;
            }
        }

        private void LoadNomina()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var fechaFactura =
                        DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value), "MMddyy",
                            CultureInfo.InvariantCulture);
                    var fechaPago =
                        DateTime.ParseExact(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value), "MMddyy",
                            CultureInfo.InvariantCulture);

                    var nominaInfo = new Nomina
                    {
                        CodigoBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value),
                        CuentaBanco = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value),
                        Analista = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value),
                        Supplier = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value),
                        Cedula = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value),
                        Moneda = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value),
                        NumFcatura = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value),
                        FechaFactura = fechaFactura.ToString("yyyyMMdd"),
                        FechaPago = fechaPago.ToString("yyyyMMdd"),
                        ValorTotal = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value),
                        Descripcion = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value),
                        Ref = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value),
                        ValorItem = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value),
                        Cuenta = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value),
                        PosicionAprobador = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)
                    };


                    var urlEnviroment = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");
                    EFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost,
                        _frmAuth.EllipseDsct, urlEnviroment);
                    var responseDto = EFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages()) return;

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
                                     "   <connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                     "   <application>ServiceInteraction</application>" +
                                     "   <applicationPage>unknown</applicationPage>" +
                                     "</interaction>";

                    responseDto = EFunctions.ExecutePostRequest(requestXml);

                    var errorMessage = responseDto.Errors.Aggregate("",
                        (current, msg) => current + (msg.Field + " " + msg.Text));
                    if (errorMessage.Equals(""))
                    {
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
                        requestXml = requestXml + "						<name>MNEMONIC1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.Cedula + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>INV_NO1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.NumFcatura + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "                 <screenField>";
                        requestXml = requestXml + "                 	<name>INV_AMT1I</name>";
                        requestXml = requestXml + "                 	<value>" + nominaInfo.ValorTotal + "</value>";
                        requestXml = requestXml + "                 </screenField>";
                        requestXml = requestXml + "                 <screenField>";
                        requestXml = requestXml + "                 	<name>ACCOUNTANT1I</name>";
                        requestXml = requestXml + "                 	<value>" + nominaInfo.Analista + "</value>";
                        requestXml = requestXml + "                      </screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>CURRENCY_TYPE1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.Moneda + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>HANDLE_CDE1I</name>";
                        requestXml = requestXml + "						<value>PN</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>INV_DATE1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.FechaFactura + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>DUE_DATE1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.FechaPago + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>BRANCH_CODE1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.CodigoBanco + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>BANK_ACCT_NO1I</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.CuentaBanco + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>INV_ITEM_DESC1I1</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.Descripcion + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I1</name>";
                        requestXml = requestXml + "					    <value>" + nominaInfo.ValorItem + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I1</name>";
                        requestXml = requestXml + "					   	<value>" + _frmAuth.EllipseDsct + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>AUTH_BY1I1</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.PosicionAprobador + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACCOUNT1I1</name>";
                        requestXml = requestXml + "						<value>" + nominaInfo.Cuenta + "</value>";
                        requestXml = requestXml + "					</screenField>";
                        requestXml = requestXml + "				</inputs>";
                        requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                        requestXml = requestXml + "			</data>";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                        requestXml = requestXml + "		</action>";
                        requestXml = requestXml + "	</actions>                                                       ";
                        requestXml = requestXml + "	<chains/>                                                        ";
                        requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId +
                                     "</connectionId>";
                        requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                        requestXml = requestXml + "</interaction>                                                    ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = EFunctions.ExecutePostRequest(requestXml);
                        errorMessage = responseDto.Errors.Aggregate("",
                            (current, msg) => current + (msg.Field + " " + msg.Text));

                        if (errorMessage.Equals(""))
                        {
                            if (responseDto.ResponseString.Contains("MSM202A"))
                            {
                                requestXml = requestXml + "<interaction> ";
                                requestXml = requestXml + "	<actions> ";
                                requestXml = requestXml + "		<action> ";
                                requestXml = requestXml + "			<name>submitScreen</name> ";
                                requestXml = requestXml + "			<data> ";
                                requestXml = requestXml + "				<inputs> ";
                                requestXml = requestXml + "					<screenField> ";
                                requestXml = requestXml + "						<name>SUP_MNEMONIC1I</name> ";
                                requestXml = requestXml + "						<value>" + nominaInfo.Cedula + "</value> ";
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
                                requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId +
                                             "</connectionId> ";
                                requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                                requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                                requestXml = requestXml + "</interaction> ";

                                requestXml = requestXml.Replace("&", "&amp;");
                                responseDto = EFunctions.ExecutePostRequest(requestXml);
                                errorMessage = responseDto.Errors.Aggregate("",
                                    (current, msg) => current + (msg.Field + " " + msg.Text));

                                if (errorMessage.Equals(""))
                                {
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
                                    requestXml = requestXml + "	<connectionId>" +
                                                 EFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                                    requestXml = requestXml + "	<application>ServiceInteraction</application>";
                                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                                    requestXml = requestXml + "</interaction>";

                                    responseDto = EFunctions.ExecutePostRequest(requestXml);
                                    errorMessage = responseDto.Errors.Aggregate("",
                                        (current, msg) => current + (msg.Field + " " + msg.Text));

                                    if (errorMessage.Equals(""))
                                    {
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
                                        requestXml = requestXml + "	<connectionId>" +
                                                     EFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                                        requestXml = requestXml + "	<application>ServiceInteraction</application>";
                                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                                        requestXml = requestXml + "</interaction>";

                                        responseDto = EFunctions.ExecutePostRequest(requestXml);
                                        errorMessage = responseDto.Errors.Aggregate("",
                                            (current, msg) => current + (msg.Field + " " + msg.Text));

                                        if (errorMessage.Equals(""))
                                        {
                                            _cells.GetCell(_resultColumn, currentRow).Select();
                                            _cells.GetCell(_resultColumn, currentRow).Value = "Creado";
                                            _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                                _cells.GetStyle(StyleConstants.Success);
                                        }
                                        else
                                        {
                                            _cells.GetCell(_resultColumn, currentRow).Select();
                                            _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                                            _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                                _cells.GetStyle(StyleConstants.Error);
                                        }
                                    }
                                    else
                                    {
                                        _cells.GetCell(_resultColumn, currentRow).Select();
                                        _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                                        _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                            _cells.GetStyle(StyleConstants.Error);
                                    }
                                }
                                else
                                {
                                    _cells.GetCell(_resultColumn, currentRow).Select();
                                    _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                                    _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                        _cells.GetStyle(StyleConstants.Error);
                                }
                            }
                        }
                        else
                        {
                            _cells.GetCell(_resultColumn, currentRow).Select();
                            _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                            _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                _cells.GetStyle(StyleConstants.Error);
                        }
                    }

                    else
                    {
                        _cells.GetCell(_resultColumn, currentRow).Select();
                        _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                        _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                            _cells.GetStyle(StyleConstants.Error);
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void LoadCesantiasPost()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TittleRow + 1;
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

                    var urlEnviroment = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, "POST");
                    EFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost,
                        _frmAuth.EllipseDsct, urlEnviroment);
                    var responseDto = EFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages()) return;

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
                                     "   <connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                     "   <application>ServiceInteraction</application>" +
                                     "   <applicationPage>unknown</applicationPage>" +
                                     "</interaction>";

                    responseDto = EFunctions.ExecutePostRequest(requestXml);

                    var errorMessage = responseDto.Errors.Aggregate("",
                        (current, msg) => current + (msg.Field + " " + msg.Text));

                    if (errorMessage.Equals(""))
                    {
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
                        requestXml = requestXml + "                      </screenField>";
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
                        requestXml = requestXml + "				</inputs>";
                        requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                        requestXml = requestXml + "			</data>";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                        requestXml = requestXml + "		</action>";
                        requestXml = requestXml + "	</actions>                                                       ";
                        requestXml = requestXml + "	<chains/>                                                        ";
                        requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId +
                                     "</connectionId>";
                        requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                        requestXml = requestXml + "</interaction>                                                    ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = EFunctions.ExecutePostRequest(requestXml);
                        errorMessage = responseDto.Errors.Aggregate("",
                            (current, msg) => current + (msg.Field + " " + msg.Text));

                        if (errorMessage.Equals(""))
                        {
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
                            requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId +
                                         "</connectionId>";
                            requestXml = requestXml + "	<application>ServiceInteraction</application>";
                            requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                            requestXml = requestXml + "</interaction>";

                            responseDto = EFunctions.ExecutePostRequest(requestXml);
                            errorMessage = responseDto.Errors.Aggregate("",
                                (current, msg) => current + (msg.Field + " " + msg.Text));

                            if (errorMessage.Equals(""))
                            {
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
                                requestXml = requestXml + "	<connectionId>" + EFunctions.PostServiceProxy.ConnectionId +
                                             "</connectionId>";
                                requestXml = requestXml + "	<application>ServiceInteraction</application>";
                                requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                                requestXml = requestXml + "</interaction>";

                                responseDto = EFunctions.ExecutePostRequest(requestXml);
                                errorMessage = responseDto.Errors.Aggregate("",
                                    (current, msg) => current + (msg.Field + " " + msg.Text));

                                if (errorMessage.Equals(""))
                                {
                                    _cells.GetCell(_resultColumn, currentRow).Select();
                                    _cells.GetCell(_resultColumn, currentRow).Value = "Creado";
                                    _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                        _cells.GetStyle(StyleConstants.Success);
                                }
                                else
                                {
                                    _cells.GetCell(_resultColumn, currentRow).Select();
                                    _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                                    _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                        _cells.GetStyle(StyleConstants.Error);
                                }
                            }
                            else
                            {
                                _cells.GetCell(_resultColumn, currentRow).Select();
                                _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                                _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                    _cells.GetStyle(StyleConstants.Error);
                            }
                        }
                        else
                        {
                            _cells.GetCell(_resultColumn, currentRow).Select();
                            _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                            _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                                _cells.GetStyle(StyleConstants.Error);
                        }
                    }
                    else
                    {
                        _cells.GetCell(_resultColumn, currentRow).Select();
                        _cells.GetCell(_resultColumn, currentRow).Value = errorMessage;
                        _cells.GetRange(1, currentRow, _resultColumn, currentRow).Style =
                            _cells.GetStyle(StyleConstants.Error);
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        private void LoadCesantias()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TittleRow + 1;

            var proxySheet = new screen.ScreenService();
            var requestSheet = new screen.ScreenSubmitRequestDTO();

            _cells.GetRange(_resultColumn, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow)
                .ClearContents();
            _cells.GetRange(_resultColumn, TittleRow + 1, _resultColumn, _excelSheetItems.ListRows.Count + TittleRow)
                .Style = _cells.GetStyle(StyleConstants.Normal);

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


                    EFunctions.RevertOperation(opSheet, proxySheet);

                    var replySheet = proxySheet.executeScreen(opSheet, "MSO265");

                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {
                        if (replySheet.mapName != "MSM265A") return;
                        var arrayFields = new ArrayScreenNameValue();

                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDsct);
                        arrayFields.Add("SUPPLIER_NO1I", supplierInfo.SupplierNo);
                        arrayFields.Add("ACCOUNTANT1I", supplierInfo.Accountant);
                        arrayFields.Add("INV_NO1I", supplierInfo.InvNo);
                        arrayFields.Add("INV_DATE1I", supplierInfo.InvDate);
                        arrayFields.Add("DUE_DATE1I", supplierInfo.DueDate);
                        arrayFields.Add("CURRENCY_TYPE1I", supplierInfo.CurrencyType);
                        arrayFields.Add("HANDLE_CDE1I", "PN");
                        arrayFields.Add("INV_AMT1I", supplierInfo.InvAmount);
                        arrayFields.Add("INV_ITEM_DESC1I1", supplierInfo.InvItemDesc);
                        arrayFields.Add("INV_ITEM_VALUE1I1", supplierInfo.InvItemValue);
                        arrayFields.Add("AUTH_BY1I1", supplierInfo.AuthBy);
                        arrayFields.Add("ACCOUNT1I1", supplierInfo.Account);
                        arrayFields.Add("BRANCH_CODE1I", supplierInfo.BranchCode);
                        arrayFields.Add("BANK_ACCT_NO1I", supplierInfo.BankAccount);
                        arrayFields.Add("ACCT_DSTRCT1I1", _frmAuth.EllipseDsct);

                        requestSheet.screenFields = arrayFields.ToArray();

                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        while (EFunctions.CheckReplyWarning(replySheet))
                            replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet.message.Contains("Confirm"))
                            replySheet = proxySheet.submit(opSheet, requestSheet);


                        if (EFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(_resultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Success;
                            _cells.GetCell(_resultColumn, currentRow).Value = "Success";
                        }
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(_resultColumn, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(_resultColumn, currentRow).Value = ex.Message;
                }
                finally
                {
                    currentRow++;
                }
            }
        }

        public class Nomina
        {
            public string CodigoBanco { get; set; }
            public string CuentaBanco { get; set; }
            public string Analista { get; set; }
            public string Supplier { get; set; }
            public string Cedula { get; set; }
            public string Moneda { get; set; }
            public string NumFcatura { get; set; }
            public string FechaFactura { get; set; }
            public string FechaPago { get; set; }
            public string ValorTotal { get; set; }
            public string Descripcion { get; set; }
            public string Ref { get; set; }
            public string ValorItem { get; set; }
            public string Cuenta { get; set; }
            public string PosicionAprobador { get; set; }
        }

        public class SupplierInfo
        {
            public SupplierInfo()
            {
            }

            public SupplierInfo(string supplier, string enviroment)
            {
                var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", supplier, EFunctions.dbReference,
                    EFunctions.dbLink);
                EFunctions.SetDBSettings(enviroment);

                var drSupplierInfo = EFunctions.GetQueryResult(sqlQuery);

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
            public string Cedula { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string Nombre { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string Referencia { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string Descripcion { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string FechaFactura { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string FechaPago { get; set; }

            [CsvColumn(FieldIndex = 7)]
            public string Cuenta { get; set; }

            [CsvColumn(FieldIndex = 8)]
            public string Moneda { get; set; }

            [CsvColumn(FieldIndex = 9)]
            public string ValorItem { get; set; }

            [CsvColumn(FieldIndex = 10)]
            public string ValorTotal { get; set; }

            [CsvColumn(FieldIndex = 11)]
            public string PosicionAprobador { get; set; }

            [CsvColumn(FieldIndex = 12)]
            public string CodigoBanco { get; set; }

            [CsvColumn(FieldIndex = 13)]
            public string CuentaBanco { get; set; }
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
            public string CodigoBanco { get; set; }

            [CsvColumn(FieldIndex = 2)]
            public string CuentaBanco { get; set; }

            [CsvColumn(FieldIndex = 3)]
            public string Analista { get; set; }

            [CsvColumn(FieldIndex = 4)]
            public string Supplier { get; set; }

            [CsvColumn(FieldIndex = 5)]
            public string Cedula { get; set; }

            [CsvColumn(FieldIndex = 6)]
            public string Moneda { get; set; }

            [CsvColumn(FieldIndex = 7)]
            public string NumFcatura { get; set; }

            [CsvColumn(FieldIndex = 8)]
            public string FechaFactura { get; set; }

            [CsvColumn(FieldIndex = 9)]
            public string FechaPago { get; set; }

            [CsvColumn(FieldIndex = 10)]
            public string ValorTotal { get; set; }

            [CsvColumn(FieldIndex = 11)]
            public string Descripcion { get; set; }

            [CsvColumn(FieldIndex = 12)]
            public string Ref { get; set; }

            [CsvColumn(FieldIndex = 13)]
            public string ValorItem { get; set; }

            [CsvColumn(FieldIndex = 14)]
            public string Cuenta { get; set; }

            [CsvColumn(FieldIndex = 15)]
            public string PosicionAprobador { get; set; }

            [CsvColumn(FieldIndex = 16)]
            public string value01 { get; set; }

            [CsvColumn(FieldIndex = 17)]
            public string value02 { get; set; }

            [CsvColumn(FieldIndex = 18)]
            public string value03 { get; set; }

            [CsvColumn(FieldIndex = 19)]
            public string value04 { get; set; }

            [CsvColumn(FieldIndex = 20)]
            public string value05 { get; set; }

            [CsvColumn(FieldIndex = 21)]
            public string value06 { get; set; }

            [CsvColumn(FieldIndex = 22)]
            public string value07 { get; set; }

            [CsvColumn(FieldIndex = 23)]
            public string value08 { get; set; }

            [CsvColumn(FieldIndex = 24)]
            public string value09 { get; set; }

            [CsvColumn(FieldIndex = 25)]
            public string value10 { get; set; }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
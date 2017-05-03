using System;
using System.Data;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using LINQtoCSV;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseMSO265ExcelAddIn
{
    public partial class RibbonEllipse
    {

        private const int TittleRow = 5;
        private const int ResultColumn = 22;
        public static EllipseFunctions EFunctions = new EllipseFunctions();
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private string _sheetName01 = "MSO265 Cesantias";
        ListObject _excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            EFunctions.DebugQueries = false;
            EFunctions.DebugErrors = false;
            EFunctions.DebugWarnings = false;
            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void btnFormatCesantias_Click(object sender, RibbonControlEventArgs e)
        {
            FormatCesantias();
        }

        private void FormatCesantias()
        {
            _excelApp = Globals.ThisAddIn.Application;
            var excelBook = _excelApp.Workbooks.Add();
            Worksheet excelSheet = excelBook.ActiveSheet;


            _sheetName01 = "MSO265 Cesantias";
            excelSheet.Name = _sheetName01;


            _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            #region Instructions

            _cells.GetCell(1, 1).Value = "CERREJÓN";
            _cells.GetCell(1, 1).Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            _cells.MergeCells(1, 1, 1, 2);
            _cells.GetCell("B1").Value = "Registro y Verificacion de Pago de Cesantias";
            _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderSize17);
            _cells.MergeCells(2, 1, 7, 2);

            #endregion

            #region Datos

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";

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
            _cells.GetCell(18, TittleRow).Value = "Sucursal Banco";
            _cells.GetCell(19, TittleRow).Value = "Cuenta Banco";
            _cells.GetCell(20, TittleRow).Value = "ST Adress";
            _cells.GetCell(21, TittleRow).Value = "ST Business";
            _cells.GetCell(ResultColumn, TittleRow).Value = "Result";

            _cells.GetRange(1, TittleRow, ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);


            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

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

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).Delete();

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
                FirstLineHasColumnNames = true,
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
                    MessageBox.Show(string.Format("Error: {0}", error.Message));
                }
                finally
                {
                    currentRow++;
                }
            }

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

            excelSheet.Cells.Columns.AutoFit();
            excelSheet.Cells.Rows.AutoFit();

            //ValidateCesantias();
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

            _cells.GetRange(1, TittleRow + 1, ResultColumn, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "@";
            _cells.GetRange(9, TittleRow + 1, 10, _excelSheetItems.ListRows.Count + TittleRow).NumberFormat = "$ #,##0.00";

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    var supplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value);
                    supplierInfo = new SupplierInfo(supplierNo, drpEnviroment.SelectedItem.Label);
                    _cells.GetCell(14, currentRow).Select();

                    _cells.GetCell(17, currentRow).Value = supplierInfo.SupplierNo;
                    _cells.GetCell(18, currentRow).Value = supplierInfo.AccountName;
                    _cells.GetCell(19, currentRow).Value = supplierInfo.AccountNo;
                    _cells.GetCell(20, currentRow).Value = supplierInfo.StAdress;
                    _cells.GetCell(21, currentRow).Value = supplierInfo.StBusiness;

                    _cells.GetCell(12, currentRow).Style = _cells.GetCell(18, currentRow).Style = _cells.GetStyle(supplierInfo.AccountName.Substring(2, 4) == _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, currentRow).Value) ? StyleConstants.Success : StyleConstants.Error);


                    _cells.GetCell(13, currentRow).Style = _cells.GetCell(19, currentRow).Style = _cells.GetStyle(supplierInfo.AccountNo == _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, currentRow).Value) ? StyleConstants.Success : StyleConstants.Error);
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
            LoadCesantias();
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
            }
        }

        private void LoadCesantiasPost()
        {

//            public void SetPostService(string ellipseUser, string ellipsePswd, string ellipsePost, string ellipseDsct, string urlEnviroment)
//            {
//                urlEnviroment = urlEnviroment.Replace("/ews/services", "") + "/ria-Ellipse-8.4.23.2_59/bind?app=";
//                urlEnviroment = urlEnviroment.Replace("http://ews", "http://ellipse");
//                PostServiceProxy = new PostService(ellipseUser, ellipsePswd, ellipsePost, ellipseDsct, urlEnviroment);
//            }

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var currentRow = TittleRow + 1;
            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var urlEnviroment = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    EFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnviroment);
                    var responseDto = EFunctions.InitiatePostConnection();

                    if (responseDto.GotErrorMessages()) return;

                    var requestXml = "<interaction>" +
                        "<actions>" +
                        "<action>" +
                        "<name>initialScreen</name>" +
                        "<data>" +
                        "<screenName>mso265</screenName>" +
                        "</data>" +
                        "<id>" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "</id>" +
                        "</action>" +
                        "</actions>" +
                        "<chains/>" +
                        "<connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                        "<application>ServiceInteraction</application>" +
                        "<applicationPage>unknown</applicationPage>" +
                        "</interaction>";

                    responseDto = EFunctions.ExecutePostRequest(requestXml);

                    var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));

                    if (!errorMessage.Equals(""))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Select();
                        _cells.GetCell(ResultColumn, currentRow).Value = errorMessage;
                        _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    }
                    else
                    {
                        var supplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
                        var accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
                        var invNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                        var invDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                        var dueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                        var currencyType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                        var invAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value);
                        var invItemDesc = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                        var invItemValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value);
                        var authBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value);
                        var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                        var branchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                        var bankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);
                        
                        requestXml = "<interaction>                                                    " +
                                         "<actions>" +
                                         "<action>" +
                                         "<name>submitScreen</name>" +
                                         "<data>" +
                                         "<inputs>" +
                                         "<screenField>" +
                                         "<name>DSTRCT_CODE1I</name>" +
                                         "<value>ICOR</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>SUPPLIER_NO1I</name>" +
                                         "<value>123327</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>SUPPLIER_NAME1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>MNEMONIC1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>GOVT_ID1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_NO1I</name>" +
                                         "<value>PSL201601260103</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_AMT_LIT1I</name>" +
                                         "<value>Invoice Amt</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_AMT1I</name>" +
                                         "<value>1,026,645.00</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ADD_TAX_LIT1I</name>" +
                                         "<value>Add Tax</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ADD_TAX_AMOUNT1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ADD_TAX_HDR_L1I</name>" +
                                         "<value>Add Tax Code</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ATAX_CODE_HDR1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ACCOUNTANT1I</name>" +
                                         "<value>ARESTR1</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ACCTNT_NAME1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ORG_INV_NO_LIT1I</name>" +
                                         "<value>Org Inv No.</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ORG_INV_NO1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>CURRENCY_TYPE1I</name>" +
                                         "<value>PES</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_COMM_TYPE1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>HANDLE_CDE1I</name>" +
                                         "<value>PN</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_DATE1I</name>" +
                                         "<value>02/04/2016</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_RCPT_DATE1I</name>" +
                                         "<value>02/04/2016</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>DUE_DATE1I</name>" +
                                         "<value>02/04/2016</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>SD_AMOUNT1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>SD_DATE1I</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>BRANCH_CODE1I</name>" +
                                         "<value>NM009</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>BANK_ACCT_NO1I</name>" +
                                         "<value>1209</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_ITEM_NO1I1</name>" +
                                         "<value>001</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_ITEM_DESC1I1</name>" +
                                         "<value>descripcion</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>INV_ITEM_VALUE1I1</name>" +
                                         "<value>1,026,645.00</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>DEDUCT_PP1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ACCT_DSTRCT1I1</name>" +
                                         "<value>ICOR</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>AUTH_BY1I1</name>" +
                                         "<value>4420</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ATAX_CODE_ITM1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ITM_TAX_LIT1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>CTAX_CODE1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>CTAX_VALUE1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ACCOUNT1I1</name>" +
                                         "<value>1988101;=79987420</value>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>WORK_ORDER1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>PROJECT_IND1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>PLANT_NO1I1</name>" +
                                         "<value/>" +
                                         "</screenField>" +
                                         "<screenField>" +
                                         "<name>ACTION1I1</name>" +
                                         "<value/>" +
                                         "</screenField>"+
                                         "</inputs>" +
                                         "<screenName>MSM265A</screenName>" +
                                         "<screenAction>TRANSMIT</screenAction>" +
                                         "</data>" +
                                         "<id>" + System.Web.Services.Ellipse.Post.Util.GetNewOperationId() + "</id>" +
                                         "</action>" +
                                         "</actions>" +
                                         "<chains/>" +
                                         "<connectionId>" + EFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                         "<application>ServiceInteraction</application>" +
                                         "<applicationPage>unknown</applicationPage>" +
                                         "</interaction>";

                        responseDto = EFunctions.ExecutePostRequest(requestXml);

                        errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
                        
                         var deSerializeDs = new DataSet();
                            try
                            {
                               deSerializeDs.ReadXmlSchema(responseDto.ToString());
                               deSerializeDs.ReadXml(responseDto.ToString(), XmlReadMode.IgnoreSchema);
                            }
                            catch (Exception ex)
                            {
                               // Handle Exceptions Here…..
                            }

                        if (errorMessage.Equals(""))
                        {
                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetCell(ResultColumn, currentRow).Value = "Creado";
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style = _cells.GetStyle(StyleConstants.Success);
                        }
                        else
                        {
                            _cells.GetCell(ResultColumn, currentRow).Select();
                            _cells.GetCell(ResultColumn, currentRow).Value = errorMessage;
                            _cells.GetRange(1, currentRow, ResultColumn, currentRow).Style =
                                _cells.GetStyle(StyleConstants.Error);
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
        }

        private void LoadCesantias()
        {
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

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
                returnWarnings = EFunctions.DebugWarnings
            };

            _cells.GetCell(1, currentRow).Select();

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) != null)
            {
                try
                {
                    _cells.GetCell(1, currentRow).Select();

                    var supplierNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, currentRow).Value);
                    var accountant = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value);
                    var invNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value);
                    var invDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value);
                    var dueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value);
                    var currencyType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, currentRow).Value);
                    var invAmount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, currentRow).Value);
                    var invItemDesc = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                    var invItemValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value);
                    var authBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, currentRow).Value);
                    var account = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, currentRow).Value);
                    var branchCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, currentRow).Value);
                    var bankAccount = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value);

                    EFunctions.RevertOperation(opSheet, proxySheet);

                    var replySheet = proxySheet.executeScreen(opSheet, "MSO265");

                    if (EFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                    }
                    else
                    {


                        if (replySheet.mapName != "MSM265A") return;
                        var arrayFields = new ArrayScreenNameValue();

                        arrayFields.Add("DSTRCT_CODE1I", _frmAuth.EllipseDsct);
                        arrayFields.Add("SUPPLIER_NO1I", supplierNo);
                        arrayFields.Add("ACCOUNTANT1I", accountant);
                        arrayFields.Add("INV_NO1I", invNo);
                        arrayFields.Add("INV_DATE1I", invDate);
                        arrayFields.Add("DUE_DATE1I", dueDate);
                        arrayFields.Add("CURRENCY_TYPE1I", currencyType);
                        arrayFields.Add("HANDLE_CDE1I", "PN");
                        arrayFields.Add("INV_AMT1I", invAmount);
                        arrayFields.Add("INV_ITEM_DESC1I1", invItemDesc);
                        arrayFields.Add("INV_ITEM_VALUE1I1", invItemValue);
                        arrayFields.Add("AUTH_BY1I1", authBy);
                        arrayFields.Add("ACCOUNT1I1", account);
                        arrayFields.Add("BRANCH_CODE1I", branchCode);
                        arrayFields.Add("BANK_ACCT_NO1I", bankAccount);
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
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(ResultColumn, currentRow).Value = replySheet.message;
                        }
                        else
                        {
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                            _cells.GetCell(ResultColumn, currentRow).Value = "Success";
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
        }

        public class SupplierInfo
        {

            public SupplierInfo()
            {

            }

            public SupplierInfo(string supplier, string enviroment)
            {
                var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", supplier, EFunctions.dbReference, EFunctions.dbLink);
                EFunctions.SetDBSettings(enviroment);

                var drSupplierInfo = EFunctions.GetQueryResult(sqlQuery);

                if (!drSupplierInfo.Read())
                {
                    Error = "No existen datos";
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
                    Error = "Success";
                }
            }

            public string SupplierNo { get; set; }
            public string TaxFileNo { get; set; }
            public string StAdress { get; set; }
            public string StBusiness { get; set; }
            public string SupplierName { get; set; }
            public string CurrencyType { get; set; }
            public string AccountName { get; set; }
            public string AccountNo { get; set; }
            public string Error { get; set; }
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
            public static string GetSupplierInvoiceInfo(string districtCode, string supplierNo, string dbReference, string dbLink)
            {
                var sqlQuery = "SELECT " +
                               "   TRIM(A.SUPPLIER_NO) SUPPLIER_NO, " +
                               "   TRIM(B.TAX_FILE_NO) TAX_FILE_NO, " +
                               "   TRIM(A.SUP_STATUS) ST_ADRESS, " +
                               "   TRIM(B.SUP_STATUS) ST_BUSINESS, " +
                               "   TRIM(A.SUPPLIER_NAME) SUPPLIER_NAME, " +
                               "   TRIM(A.CURRENCY_TYPE) CURRENCY_TYPE, " +
                               "   TRIM(B.BANK_ACCT_NAME) BANK_ACCT_NAME, " +
                               "   TRIM(B.BANK_ACCT_NO) BANK_ACCT_NO " +
                               " FROM " +
                               "   ELLIPSE.MSF200 A " +
                               " INNER JOIN ELLIPSE.MSF203 B " +
                               " ON " +
                               "   A.SUPPLIER_NO = B.SUPPLIER_NO " +
                               " AND B.DSTRCT_CODE = '" + districtCode + "' " +
                               " AND B.TAX_FILE_NO = '" + supplierNo + "' ";
                return sqlQuery;
            }
        }
    }
}

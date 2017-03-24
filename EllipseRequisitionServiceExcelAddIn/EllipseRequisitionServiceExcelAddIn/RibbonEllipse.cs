using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using System.Web.Services.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseRequisitionServiceExcelAddIn.IssueRequisitionItemStocklessService;
using EllipseRequisitionServiceExcelAddIn.Properties;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseRequisitionServiceExcelAddIn
{
    public partial class RibbonEllipse
    {
        Excel.Application _excelApp;
        ExcelStyleCells _cells;
        readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        readonly FormAuthenticate _frmAuth = new FormAuthenticate();

        private const int TitleRow = 5;
        private const int ResultColumn = 19;

        Excel.ListObject _excelSheetItems;

        private const string SheetName01 = "RequisitionService";
        private const string TableName01 = "RequisitionServiceTable";
        private const string ValidationSheet = "ValidationRequisition";

        private bool _ignoreItemError;

        private Thread _thread;


        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviromentList = EnviromentConstants.GetEnviromentList();
            foreach (var item in enviromentList)
            {
                var drpItem = Factory.CreateRibbonDropDownItem();
                drpItem.Label = item;
                drpEnviroment.Items.Add(drpItem);
            }

            drpEnviroment.SelectedItem.Label = Resources.RibbonEllipse_RibbonEllipse_Load_Productivo;
        }

        private void btnFormatNewSheet_Click(object sender, RibbonControlEventArgs e)
        {
            RequisitionServiceFormat();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(
                    @"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }
        private void btnExcecuteRequisitionService_Click(object sender, RibbonControlEventArgs e)
        {
            _ignoreItemError = false;
			//si si ya hay un thread corriendo que no se ha detenido
			if (_thread != null && _thread.IsAlive) return;
			_thread = new Thread(CreateRequisitionService);

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }
        private void btnCreateReqIgError_Click(object sender, RibbonControlEventArgs e)
        {
            _ignoreItemError = true;
			//si si ya hay un thread corriendo que no se ha detenido
			if (_thread != null && _thread.IsAlive) return;
            _thread = new Thread(CreateRequisitionService);

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }
        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
        /// <summary>
        /// Da Formato a la Hoja de Excel Creando los
        /// </summary>
        private void RequisitionServiceFormat()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                _cells.CreateNewWorksheet(ValidationSheet);
                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Name = SheetName01;

                #region FormatHeader
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");
                _cells.GetCell("B1").Value = "REQUISITION SERVICE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "G2");

                _cells.GetCell("H1").Value = "OBLIGATORIO";
                _cells.GetCell("H1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("H2").Value = "OPCIONAL";
                _cells.GetCell("H2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("H3").Value = "INFORMATIVO";
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                #endregion

                _cells.GetCell(1, TitleRow).Value = "Requested By";
                _cells.GetCell(1, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, TitleRow).Value = "Requested By Position";
                _cells.GetCell(2, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(3, TitleRow).Value = "Indicador de Serie";
                _cells.GetCell(3, TitleRow).AddComment("Indica una serie diferente para vales con encabezados comunes (Ej. Flota, número, secuencia A, etc)");
                _cells.GetCell(3, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(4, TitleRow).Value = "Requisition Number";
                _cells.GetCell(4, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(5, TitleRow).Value = "Requisition Type";
                _cells.GetCell(5, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionReqTypeList = new List<string>
                {
                    "NI - NORMAL REQUISITION",
                    "PR - PURCHASE REQUISITION",
                    "CR - CREDIT REQUISITION",
                    "LN - LOAN REQUISITION"
                };
                _cells.SetValidationList(_cells.GetCell(5, TitleRow + 1), optionReqTypeList, ValidationSheet, 1, false);
                _cells.GetCell(6, TitleRow).Value = "Transaction Type";
                _cells.GetCell(6, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var optionTransTypeList = new List<string>
                {
                    "DN - DESPACHO NO PLANEADO",
                    "DP - DESPACHO PLANEADO",
                    "CN - DEVOLUCION NO PLANEADA",
                    "CP - DEVOLUCION PLANEADA"
                };
                _cells.SetValidationList(_cells.GetCell(6, TitleRow + 1), optionTransTypeList, ValidationSheet, 2, false);
                _cells.GetCell(7, TitleRow).Value = "Required By Date";
                _cells.GetCell(7, TitleRow).AddComment("YYYYMMDD");
                _cells.GetCell(7, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(8, TitleRow).Value = "Original Warehouse";
                _cells.GetCell(8, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(9, TitleRow).Value = "Priority Code";
                _cells.GetCell(9, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var itemList = _eFunctions.GetItemCodes("PI");
                var optionPriorList = Utils.GetCodeList(itemList);
                _cells.SetValidationList(_cells.GetCell(9, TitleRow + 1), optionPriorList, ValidationSheet, 3, false);

                _cells.GetCell(10, TitleRow).Value = "Reference Type";
                _cells.GetCell(10, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                var optionRefTypeList = new List<string> {"Work Order", "Equipment No.", "Project No.", "Account Code"};
                _cells.SetValidationList(_cells.GetCell(10, TitleRow + 1), optionRefTypeList, ValidationSheet, 4);

                _cells.GetCell(11, TitleRow).Value = "Reference";
                _cells.GetCell(11, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell(12, TitleRow).Value = "Delivery Instructions"; //120 caracteres (60/60)
                _cells.GetCell(12, TitleRow).AddComment("120 caracteres");
                _cells.GetCell(12, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(13, TitleRow).Value = "Return Cause";
                _cells.GetCell(13, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(14, TitleRow).Value = "Issue Question";
                _cells.GetCell(14, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                var optionIssueList = new List<string> {"A - VENTAS", "B - RUBROS"};
                _cells.SetValidationList(_cells.GetCell(14, TitleRow + 1), optionIssueList, ValidationSheet, 5, false);

                _cells.GetCell(15, TitleRow).Value = "Partial Allowed";
                _cells.GetCell(15, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var partialAllowedList = new List<string> {"Y - YES", "N - No"};
                _cells.SetValidationList(_cells.GetCell(15, TitleRow + 1), partialAllowedList, ValidationSheet, 5, false);

                _cells.GetCell(16, TitleRow).Value = "Stock Code";
                _cells.GetCell(16, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(17, TitleRow).Value = "Unit Of Issue";
                _cells.GetCell(17, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(18, TitleRow).Value = "Quantity";
                _cells.GetCell(18, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumn, TitleRow).Value = "Result";
                _cells.GetCell(ResultColumn, TitleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, TitleRow + 1, ResultColumn, TitleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.FormatAsTable(_cells.GetRange(1, TitleRow, ResultColumn, TitleRow + 1), TableName01);

                ((Excel.Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        /// <summary>
        /// Recorre y Crea los vales de a tabla de Excel
        /// </summary>
        private void CreateRequisitionService()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn);
                _cells.ClearTableRangeColumn(TableName01, 4);

                _excelSheetItems = _cells.GetRange(TableName01).ListObject;
                //Organiza las celdas de forma que se creen la menor cantidad de vales posibles
                if (_excelSheetItems.Sort.SortFields.Count > 0)
                {
                    _excelSheetItems.Sort.SortFields.Clear();
                }

                #region SortFields
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(1, TitleRow), Excel.XlSortOn.xlSortOnValues,
                            Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(5, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(7, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(8, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(12, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(13, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(14, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                //_excelSheetItems.Sort.SortFields.Add(_cells.GetCell(16, TitleRow), Excel.XlSortOn.xlSortOnValues,
                //    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.Apply(); 
                #endregion

                //instancia del Servicio
                var proxyRequisition = new RequisitionService.RequisitionService();

                //Header
                var opRequisition = new RequisitionService.OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false);//Se asigna por 'Pooled Connection Request Timed Out'
                proxyRequisition.Url = urlService + "/RequisitionService";
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                opRequisition.district = _frmAuth.EllipseDsct;
                opRequisition.maxInstances = 100;
                opRequisition.position = _frmAuth.EllipsePost;
                opRequisition.returnWarnings = false;

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);


                var currentRow = TitleRow + 1;
                var currentRowHeader = currentRow;

                var itemList = new List<RequisitionItem>();
                RequisitionService.RequisitionServiceCreateHeaderReplyDTO headerCreateReply = null;

                RequisitionHeader prevReqHeader = null;
                var abortRequisition = false;

                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value) != null ||
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
                {
                    try
                    {
                        //obtengo los datos para el encabezado
                        var curReqHeader = new RequisitionHeader
                        {
                            AllocPcA = "100",
                            DistrictCode =
                                string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct,
                            CostDistrictA =
                                string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct,
                            RequestedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value),
                            RequiredByPos = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) == null
                                ? _frmAuth.EllipsePost
                                : _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value),
                            IndSerie = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value),
                            IreqType =
                                Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value)),
                            IssTranType =
                                Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value)),
                            RequiredByDate = _cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value),
                            OrigWhouseId = _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value),
                            PriorityCode =
                                Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value)),
                            PartIssue = true,
                            ProtectedInd = false
                        };

                        // ReSharper disable once SuggestVarOrType_BuiltInTypes
                        string switchCase = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                        var reference = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                        switch (switchCase)
                        {
                            case "Work Order":
                                curReqHeader.WorkOrderA = reference;
                                curReqHeader.WorkProjectIndA = "W"; //Solo aplica para MSO140
                                break;
                            case "Equipment No.":
                                curReqHeader.EquipmentA = reference;
                                break;
                            case "Project No.":
                                curReqHeader.ProjectA = reference;
                                curReqHeader.WorkProjectIndA = "P"; //Solo aplica para MSO140
                                break;
                            case "Account Code":
                                curReqHeader.CostCentreA = reference;
                                break;
                        }
                        curReqHeader.ProtectedInd = false;
                        curReqHeader.ProtectedIndSpecified = false;
                        curReqHeader.DelivInstrA = _cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value);
                        if (curReqHeader.DelivInstrA.Length > 80)
                            curReqHeader.DelivInstrB = curReqHeader.DelivInstrA.Substring(80);

                        curReqHeader.AnswerB = _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Length >= 2
                            ? _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Substring(0, 2).Trim()
                            : null;
                        curReqHeader.AnswerD = _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Length >= 2
                            ? _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Substring(0, 2).Trim()
                            : null;

                        //si es el primer elemento creo un nuevo encabezado
                        if (prevReqHeader == null)
                        {
                            var headerCreateRequest = curReqHeader.GetCreateRequestHeader();
                            headerCreateReply = proxyRequisition.createHeader(opRequisition, headerCreateRequest);
                            curReqHeader.IreqNo = headerCreateReply.ireqNo;
                            prevReqHeader = curReqHeader;
                            currentRowHeader = currentRow;
                            abortRequisition = false;
                        }
                        //comparo si el nuevo registro corresponde a un nuevo encabezado o si he alcanzado 99 items. Si es así, envío el encabezado anterior y creo un encabezado
                        else if (!prevReqHeader.Equals(curReqHeader) || (cbMaxItems.Checked && itemList.Count >= 99))
                        {
                            //agrego los items que tenga hasta el momento al encabezado
                            foreach (var item in itemList)
                            {
                                try
                                {
                                    var itemListDto = new List<RequisitionService.RequisitionItemDTO>
                                    {
                                        item.GetRequisitionItemDto()
                                    };

                                    var itemRequest = new RequisitionService.RequisitionServiceCreateItemRequestDTO
                                    {
                                        districtCode = prevReqHeader.DistrictCode,
                                        ireqNo = prevReqHeader.IreqNo,
                                        ireqType = prevReqHeader.IreqType,
                                        requisitionItems = itemListDto.ToArray(),

                                    };
                                    proxyRequisition.createItem(opRequisition, itemRequest);
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Value2 = itemRequest.ireqNo;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }
                                catch (Exception ex)
                                {
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 += "ERROR: " + ex.Message;
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Value2 = prevReqHeader.IreqNo;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    abortRequisition = true;
                                }
                                _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Select();
                            }

                            //aborto o finalizo según el resultado de los items
                            if (abortRequisition && !_ignoreItemError)
                            {
                                var addMessage = "";
                                try
                                {
                                    DeleteHeader(proxyRequisition, headerCreateReply, opRequisition);
                                }
                                catch (Exception ex)
                                {
                                    addMessage = ". ERROR AL ELIMINAR. " + ex.Message;
                                }

                                foreach (var item in itemList)
                                {
                                    _cells.GetCell(4, currentRowHeader + item.Index).Value2 += " - ELIMINADO";
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 += " - VALE ELIMINADO" + addMessage;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                }

                                prevReqHeader = null;
                                abortRequisition = false;
                            }
                            else
                            {
                                var finaliseRequest = new RequisitionService.RequisitionServiceFinaliseRequestDTO
                                {
                                    ireqNo = prevReqHeader.IreqNo,
                                    ireqType = prevReqHeader.IreqType,
                                    districtCode = prevReqHeader.DistrictCode
                                };

                                //Se añade este bloque try/catch porque el tiempo excesivo de finalización afecta el siguiente item de la lista. Cuando esto ocurra no afectará el proceso
                                try
                                {
                                    proxyRequisition.finalise(opRequisition, finaliseRequest);
                                }
                                catch (TimeoutException ex)
                                {
                                    _cells.GetCell(ResultColumn, currentRow - 1).Value2 = _cells.GetCell(ResultColumn, currentRow - 1).Value2 + " " + ex.Message;
                                    _cells.GetCell(ResultColumn, currentRow-1).Style = StyleConstants.Warning;
                                    _cells.GetCell(4, currentRow-1).Style = StyleConstants.Warning;
                                }
                            }

                            //creo el nuevo encabezado y reinicio variables
                            prevReqHeader = null;//no es una línea inservible. Es necesaria por si se produce una excepción al momento de creación de un nuevo encabezado
                            currentRowHeader = currentRow;
                            abortRequisition = false;
                            itemList = new List<RequisitionItem>();
                            var headerCreateRequest = curReqHeader.GetCreateRequestHeader();
                            headerCreateReply = proxyRequisition.createHeader(opRequisition, headerCreateRequest);
                            curReqHeader.IreqNo = headerCreateReply.ireqNo;
                            prevReqHeader = curReqHeader;


                        }

                        //Obtengo los datos para el item
                        var curItem = new RequisitionItem
                        {
                            Index = itemList.Count,
                            ItemType = "S",
                            PartIssueSpecified = true,
                            PartIssue =
                                Utils.IsTrue(Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)), true),
                            StockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value)
                        };
                        curItem.StockCode = (curItem.StockCode != null && curItem.StockCode.Length < 9)
                            ? curItem.StockCode.PadLeft(9, '0')
                            : curItem.StockCode;
                        curItem.UnitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(17, currentRow).Value);
                        curItem.QuantityRequired = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);

                        //Obtengo la unidad del Stock Code que voy a registrar ya que el modulo lo exige.
                        var sqlQuery = Queries.GetItemUnitOfIssue(curItem.StockCode);
                        var odr = _eFunctions.GetQueryResult(sqlQuery);
                        
                        //si se pudo obtener la Unidad
                        if (odr.Read())
                            curItem.UnitOfMeasure = "" + odr["UNIT_OF_ISSUE"];
                        else
                        {
                            abortRequisition = true;
                            _cells.GetCell(ResultColumn, currentRow).Value2 += curItem.StockCode + " NO EXISTE UNIDAD DE MEDIDA EN EL CATALOGO PARA ESTE STOCK CODE";
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            curItem.StockCode = "";//Se vacía el campo para conservar la estructura del vale, pero para que indique el error
                        }

                        //si es item de orden directa o no
                        sqlQuery = Queries.GetItemDirectOrder(curItem.StockCode);
                        odr = _eFunctions.GetQueryResult(sqlQuery);
                        if (odr.Read() && Utils.IsTrue(odr["DIRECT_ORDER_IND"]))
                        {
                            abortRequisition = true;
                            
                            _cells.GetCell(ResultColumn, currentRow).Value2 += curItem.StockCode + " ITEM DE ORDEN DIRECTA. DEBE CREAR EL VALE CON OTRO MÉTODO";
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            curItem.StockCode = "";//Se vacía el campo para conservar la estructura del vale, pero para que indique el error
                        }

                        _eFunctions.CloseConnection();
                        itemList.Add(curItem);
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn, currentRow).Value2 = ex.Message;
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }
                } //finaliza el while del proceso completo

                //para el último encabezado
                if (prevReqHeader == null) return;
                //agrego los items que tenga hasta el momento al encabezado
                foreach (var item in itemList)
                {
                    try
                    {
                        var itemListDto = new List<RequisitionService.RequisitionItemDTO> {item.GetRequisitionItemDto()};

                        var itemRequest = new RequisitionService.RequisitionServiceCreateItemRequestDTO
                        {
                            districtCode = prevReqHeader.DistrictCode,
                            ireqNo = prevReqHeader.IreqNo,
                            ireqType = prevReqHeader.IreqType,
                            requisitionItems = itemListDto.ToArray()
                        };
                        proxyRequisition.createItem(opRequisition, itemRequest);

                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        _cells.GetCell(4, currentRowHeader + item.Index).Value2 = itemRequest.ireqNo;
                        _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 +=  "ERROR: " + ex.Message;
                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRowHeader + item.Index).Value2 = prevReqHeader.IreqNo;
                        _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        abortRequisition = true;

                    }
                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Select();
                }

                //aborto o finalizo según el resultado de los items
                if (abortRequisition && !_ignoreItemError)
                {
                    var addMessage = "";
                    try
                    {
                        DeleteHeader(proxyRequisition, headerCreateReply, opRequisition);
                    }
                    catch (Exception ex)
                    {
                        addMessage = ". ERROR AL ELIMINAR. " + ex.Message;
                    }

                    foreach (var item in itemList)
                    {
                        _cells.GetCell(4, currentRowHeader + item.Index).Value2 += " - ELIMINADO";
                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 += " - VALE ELIMINADO" + addMessage;
                        _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                    }
                }
                else
                {
                    var finaliseRequest = new RequisitionService.RequisitionServiceFinaliseRequestDTO
                    {
                        ireqNo = prevReqHeader.IreqNo,
                        ireqType = prevReqHeader.IreqType,
                        districtCode = prevReqHeader.DistrictCode
                    };

                    try
                    {
                        proxyRequisition.finalise(opRequisition, finaliseRequest);
                    }
                    catch (TimeoutException ex)
                    {
                        _cells.GetCell(ResultColumn, currentRow - 1).Value2 = _cells.GetCell(ResultColumn, currentRow - 1).Value2 + " " + ex.Message;
                        _cells.GetCell(ResultColumn, currentRow - 1).Style = StyleConstants.Warning;
                        _cells.GetCell(4, currentRow - 1).Style = StyleConstants.Warning;
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
            }
        }

        /// <summary>
        /// Recorre y Crea los vales de a tabla de Excel para items catalogados como de Orden Directa
        /// </summary>
        private void CreateRequisitionScreenService()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn);
                _cells.ClearTableRangeColumn(TableName01, 4);

                #region SortItems
                _excelSheetItems = _cells.GetRange(TableName01).ListObject;
                //Organiza las celdas de forma que se creen la menor cantidad de vales posibles
                if (_excelSheetItems.Sort.SortFields.Count > 0)
                {
                    _excelSheetItems.Sort.SortFields.Clear();
                }
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(1, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(2, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(3, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(4, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(5, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(6, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(7, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(8, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(9, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(10, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(11, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(12, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(13, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.SortFields.Add(_cells.GetCell(14, TitleRow), Excel.XlSortOn.xlSortOnValues,
                    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                //_excelSheetItems.Sort.SortFields.Add(_cells.GetCell(16, TitleRow), Excel.XlSortOn.xlSortOnValues,
                //    Excel.XlOrder.xlDownThenOver, Type.Missing, Excel.XlSortDataOption.xlSortTextAsNumbers);
                _excelSheetItems.Sort.Apply();
                #endregion

                #region LoginAndScreenService
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false);//Se asigna por 'Pooled Connection Request Timed Out'
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                
                //ScreenService Opción en reemplazo de los servicios
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = _eFunctions.DebugWarnings,
                    returnWarningsSpecified = true
                };

                var proxySheet = new Screen.ScreenService {Url = urlService + "/ScreenService"};
                ////ScreenService
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                #endregion
                
                var currentRow = TitleRow + 1;
                var currentRowHeader = currentRow;

                var itemList = new List<RequisitionItem>();
                RequisitionHeader prevReqHeader = null;
                RequisitionHeader curReqHeader;
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value) != null || _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
                {
                    try
                    {
                        //obtengo los datos para el encabezado
                        curReqHeader = new RequisitionHeader
                        {
                            AllocPcA = "100",
                            DistrictCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct,
                            CostDistrictA = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? "ICOR" : _frmAuth.EllipseDsct,
                            RequestedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null ? _frmAuth.EllipseUser : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value),
                            RequiredByPos = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, currentRow).Value) == null ? _frmAuth.EllipsePost : _cells.GetEmptyIfNull(_cells.GetCell(2, currentRow).Value),
                            IndSerie = _cells.GetEmptyIfNull(_cells.GetCell(3, currentRow).Value),
                            IreqType = Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, currentRow).Value)),
                            IssTranType = Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value)),
                            RequiredByDate = _cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value),
                            OrigWhouseId = _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value),
                            PriorityCode = Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, currentRow).Value)),
                            PartIssue = true
                        };

                        string switchCase = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                        var reference = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                        switch (switchCase)
                        {
                            case "Work Order":
                                curReqHeader.WorkOrderA = reference;
                                curReqHeader.WorkProjectIndA = "W";//Solo aplica para MSO140
                                break;
                            case "Equipment No.":
                                curReqHeader.EquipmentA = reference;
                                break;
                            case "Project No.":
                                curReqHeader.ProjectA = reference;
                                curReqHeader.WorkProjectIndA = "P";//Solo aplica para MSO140
                                break;
                            case "Account Code":
                                curReqHeader.CostCentreA = reference;
                                break;
                        }

                        curReqHeader.ProtectedInd = false;
                        curReqHeader.ProtectedIndSpecified = false;
                        curReqHeader.DelivInstrA = _cells.GetEmptyIfNull(_cells.GetCell(12, currentRow).Value);
                        if (curReqHeader.DelivInstrA.Length > 80)
                            curReqHeader.DelivInstrB = curReqHeader.DelivInstrA.Substring(80);

                        curReqHeader.AnswerB = _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Length >= 2
                            ? _cells.GetEmptyIfNull(_cells.GetCell(13, currentRow).Value).Substring(0, 2).Trim()
                            : null;
                        curReqHeader.AnswerD = _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Length >= 2
                            ? _cells.GetEmptyIfNull(_cells.GetCell(14, currentRow).Value).Substring(0, 2).Trim()
                            : null;

                        //si el nuevo registro corresponde a un encabezado nuevo diferente creo el vale anterior con sus items respectivos
                        #region CreateNewRequisition
                        if ((prevReqHeader != null && !prevReqHeader.Equals(curReqHeader) && itemList.Count > 0) || (cbMaxItems.Checked && itemList.Count >= 99))
                        {
                            try
                            {
                                //Crear el encabezado en el MSO140
                                _eFunctions.RevertOperation(opSheet, proxySheet);
                                //ejecutamos el programa
                                var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                //Validamos el ingreso
                                if (reply.mapName != "MSM140A")
                                    throw new Exception("ERROR: Se ha producido un error al intentar ingresar al programa. No se puede acceder al MSO140/MSM140A. " + reply.message);

                                //se adicionan los valores a los campos
                                var arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("REQ_NO1I", "" + prevReqHeader.IreqNo);
                                arrayFields.Add("TRAN_TYPE1I", "" + prevReqHeader.IssTranType);
                                arrayFields.Add("REQ_BY_DATE1I", "" + prevReqHeader.RequiredByDate);
                                arrayFields.Add("WHOUSE_ID1I", "" + prevReqHeader.OrigWhouseId);
                                arrayFields.Add("PART_ISSUE1I", "Y");
                                arrayFields.Add("PROT_IND1I", "Y");
                                arrayFields.Add("ALLOC_PCA1I", "" + prevReqHeader.AllocPcA);
                                arrayFields.Add("WORK_PROJ_INDA1I", prevReqHeader.WorkProjectIndA);
                                arrayFields.Add("WORK_PROJA1I", prevReqHeader.WorkOrderA ?? prevReqHeader.ProjectA);
                                arrayFields.Add("COST_CENTREA1I", "" + prevReqHeader.CostCentreA);
                                arrayFields.Add("EQUIP_REFA1I", "" + prevReqHeader.EquipmentA);
                                arrayFields.Add("DELIV_INSTRA1I", "" + prevReqHeader.DelivInstrA);
                                arrayFields.Add("DELIV_INSTRB1I", "" + prevReqHeader.DelivInstrB);
                                arrayFields.Add("PRIORITY_CODE1I", "" + prevReqHeader.PriorityCode);
                                arrayFields.Add("ANSWER_B1I", "" + prevReqHeader.AnswerB);
                                arrayFields.Add("ANSWER_D1I", "" + prevReqHeader.AnswerD);
                                arrayFields.Add("REQUESTED_BY1I", "" + prevReqHeader.RequestedBy);

                                //enviar el encabezado MSM140A
                                var request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };
                                reply = proxySheet.submit(opSheet, request);
                                //Confirmar el encabezado MSM140A
                                while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
                                {
                                    request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                }

                                //no hay errores ni advertencias
                                if (reply == null || _eFunctions.CheckReplyError(reply))
                                    throw new Exception("ERROR: " + reply.message);

                                //MSM14BA
                                if (reply.mapName != "MSM14BA")
                                    throw new Exception("ERROR: Se ha producido un error al crear el encabezado. No se puede acceder al MSO140/MSM14BA. " + reply.message);

                                var parItemIndex = 0; //controla el par de items por pantalla
                                //agrego los items que tenga hasta el momento al encabezado
                                foreach (var item in itemList)
                                {
                                    //asigno número de vale si se genera
                                    var screenValues = new ArrayScreenNameValue(reply.screenFields);
                                    if (!string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                        prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                    if (parItemIndex%2 == 0)
                                        arrayFields = new ArrayScreenNameValue();
                                    arrayFields.Add("QTY_REQD1I" + (parItemIndex + 1), "" + item.QuantityRequired);
                                    arrayFields.Add("UOM1I" + (parItemIndex + 1), item.UnitOfMeasure);
                                    arrayFields.Add("TYPE1I" + (parItemIndex + 1), "S");
                                    arrayFields.Add("DESCR_A1I" + (parItemIndex + 1), item.StockCode);
                                    arrayFields.Add("PART_ISSUE1I" + (parItemIndex + 1), "Y");

                                    //envío si es el último item de la lista o si es el segundo de la pantalla
                                    if (item == itemList[itemList.Count - 1] || parItemIndex > 0)
                                    {
                                        request = new Screen.ScreenSubmitRequestDTO
                                        {
                                            screenFields = arrayFields.ToArray(),
                                            screenKey = "1"
                                        };
                                        reply = proxySheet.submit(opSheet, request);

                                        //evalúo si hay error y cancelo
                                        if (_eFunctions.CheckReplyError(reply))
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);
                                        //mientras confirmación o bodega
                                        while (reply != null && (reply.mapName == "MSM14BA" || reply.mapName == "MSM14CA"))
                                        {
                                            //evalúo error
                                            if (_eFunctions.CheckReplyError(reply))
                                                throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);

                                            //asigno número de vale si se genera
                                            screenValues = new ArrayScreenNameValue(reply.screenFields);
                                            if (screenValues.GetField("IREQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                                prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                            //si es una nueva pantalla de items
                                            if (reply.mapName == "MSM14BA" && item != itemList[itemList.Count - 1])
                                                if (screenValues.GetField("DESCR_A1I1") != null && string.IsNullOrWhiteSpace(screenValues.GetField("DESCR_A1I1").value))
                                                    break;
                                            ////MSM14CA  - Warehouse Holdings
                                            if (reply.mapName == "MSM14CA")
                                            {
                                                if (screenValues.GetField("TOTAL_REQD1I") == null)
                                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "No se ha encontrado el valor total requerido al asignar a bodega. MSM14CA");
                                                string selWarehouseIndex = "";
                                                //obtengo solo la lista de pares del objeto para actualizarla
                                                var screenArray = screenValues.ToArray();
                                                foreach (var parValue in screenArray)
                                                {
                                                    if (parValue.fieldName != null && parValue.fieldName.StartsWith("WHOUSE_ID_") && parValue.value == prevReqHeader.OrigWhouseId)
                                                        selWarehouseIndex = parValue.fieldName.Replace("WHOUSE_ID_", "");
                                                    if (parValue.fieldName != null && parValue.fieldName.StartsWith("QTY_REQD_"))
                                                        parValue.value = "";
                                                }
                                                //reingreso la lista al objeto del screen y actualizo la cantidad del w/h que quiero de acuerdo a lo realizado anteriormente
                                                if (string.IsNullOrWhiteSpace(selWarehouseIndex))
                                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "El item no está catalogado en la bodega seleccionada. MSM14CA");
                                                screenValues = new ArrayScreenNameValue(screenArray);
                                                screenValues.SetField("QTY_REQD_" + selWarehouseIndex, "" + screenValues.GetField("TOTAL_REQD1I").value);

                                                //envío el proceso
                                                request = new Screen.ScreenSubmitRequestDTO
                                                {
                                                    screenFields = screenValues.ToArray(),
                                                    screenKey = "1"
                                                };
                                                reply = proxySheet.submit(opSheet, request);
                                                continue; //continúa con el siguiente while
                                            }
                                            ////Confirm MSM14BA o cualquier otra confirmación que no requiera datos
                                            request = new Screen.ScreenSubmitRequestDTO
                                            {
                                                screenKey = "1"
                                            };
                                            reply = proxySheet.submit(opSheet, request);
                                        }
                                    }
                                    parItemIndex++;
                                    if (parItemIndex > 1)
                                        parItemIndex = 0;

                                    //OK del item
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }
                                //Confirmo la creación de todos los items. Si no llega a este punto es por algún problema presentado
                                foreach (var item in itemList)
                                {
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                                    _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                                    _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                                }
                            }
                            catch (Exception ex)
                            {
                                var addMessage = "" + ex.Message;
                                try
                                {
                                    if (string.IsNullOrWhiteSpace(prevReqHeader.IreqNo))
                                    {
                                        addMessage += " .NO SE HA REALIZADO NINGUNA ACCIÓN";
                                    }
                                    else
                                    {
                                        //Eliminación del vale por el MSO140
                                        _eFunctions.RevertOperation(opSheet, proxySheet);
                                        //ejecutamos el programa
                                        var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                        var screenValues = new ArrayScreenNameValue(reply.screenFields);

                                        if (_eFunctions.CheckReplyError(reply))
                                            throw new Exception("" + reply.message);

                                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" &&
                                            (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")) &&
                                            (screenValues.GetField("REQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("REQ_NO1I").value)))
                                        {
                                            var arrayFields = new ArrayScreenNameValue();
                                            arrayFields.Add("COMP_DEL1I", "D");

                                            var request = new Screen.ScreenSubmitRequestDTO
                                            {
                                                screenFields = arrayFields.ToArray(),
                                                screenKey = "1"
                                            };
                                            reply = proxySheet.submit(opSheet, request);
                                            screenValues = new ArrayScreenNameValue(reply.screenFields);

                                            if (_eFunctions.CheckReplyError(reply))
                                                throw new Exception(". ERROR AL ELIMINAR " + prevReqHeader.IreqNo + ": " + reply.message);
                                        }
                                        addMessage += " - VALE ELIMINADO " + prevReqHeader.IreqNo;
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    addMessage += ex2;
                                }
                                finally
                                {
                                    addMessage = addMessage.Replace("X2:0011 - INPUT REQUIRED  \"C\" TO COMPLETE OR \"D\" TO DELETE", "X2:0011 - EXISTE UNA ORDEN INCOMPLETA EN PROCESO. INGRESE AL MSO 140 PARA COMPLETARLA/ELIMINARLA");
                                    foreach (var item in itemList)
                                    {
                                        _cells.GetCell(4, currentRowHeader + item.Index).Value2 = !string.IsNullOrWhiteSpace(prevReqHeader.IreqNo) ? prevReqHeader.IreqNo + " - ELIMINADO" : "";
                                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = addMessage;
                                        _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                        _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                    }
                                }
                            }
                            finally
                            {
                                //creo el nuevo encabezado y reinicio variables
                                currentRowHeader = currentRow;
                                itemList = new List<RequisitionItem>();
                                prevReqHeader = curReqHeader;
                            }
                        }
                        #endregion
                        //Obtengo los datos para el item
                        var curItem = new RequisitionItem
                        {
                            Index = itemList.Count,
                            ItemType = "S",
                            PartIssueSpecified = true,
                            PartIssue =
                                Utils.IsTrue(Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, currentRow).Value)), true),
                            StockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value)
                        };
                        curItem.StockCode = (curItem.StockCode != null && curItem.StockCode.Length < 9)
                            ? curItem.StockCode.PadLeft(9, '0')
                            : curItem.StockCode;
                        curItem.UnitOfMeasure = _cells.GetNullOrTrimmedValue(_cells.GetCell(17, currentRow).Value);
                        curItem.QuantityRequired = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);

                        //Obtengo la unidad del Stock Code que voy a registrar ya que el modulo lo exige.
                        var sqlQuery = Queries.GetItemUnitOfIssue(curItem.StockCode);
                        var odr = _eFunctions.GetQueryResult(sqlQuery);

                        //si se pudo obtener la Unidad
                        if (odr.Read())
                            curItem.UnitOfMeasure = "" + odr["UNIT_OF_ISSUE"];
                        else
                        {
                            _cells.GetCell(ResultColumn, currentRow).Value2 +=
                                "NO EXISTE UNIDAD DE MEDIDA EN EL CATALOGO PARA ESTE STOCK CODE";
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        }
                        _eFunctions.CloseConnection();
                        itemList.Add(curItem);
                        prevReqHeader = curReqHeader;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn, currentRow).Value2 = ex.Message;
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }
                } //finaliza el while del proceso completo
                //para el último vale a crear
                #region CreateLastRequisition
                // ReSharper disable once InvertIf
                if (itemList.Count>0)
                {
                    try
                    {
                        //Crear el encabezado en el MSO140
                        _eFunctions.RevertOperation(opSheet, proxySheet);
                        //ejecutamos el programa
                        var reply = proxySheet.executeScreen(opSheet, "MSO140");
                        //Validamos el ingreso
                        if (reply.mapName != "MSM140A")
                            throw new Exception("ERROR: Se ha producido un error al intentar ingresar al programa. No se puede acceder al MSO140/MSM140A. " + reply.message);

                        //se adicionan los valores a los campos
                        var arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("REQ_NO1I", "" + prevReqHeader.IreqNo);
                        arrayFields.Add("TRAN_TYPE1I", "" + prevReqHeader.IssTranType);
                        arrayFields.Add("REQ_BY_DATE1I", "" + prevReqHeader.RequiredByDate);
                        arrayFields.Add("WHOUSE_ID1I", "" + prevReqHeader.OrigWhouseId);
                        arrayFields.Add("PART_ISSUE1I", "Y");
                        arrayFields.Add("PROT_IND1I", "Y");
                        arrayFields.Add("ALLOC_PCA1I", "" + prevReqHeader.AllocPcA);
                        arrayFields.Add("WORK_PROJ_INDA1I", prevReqHeader.WorkProjectIndA);
                        arrayFields.Add("WORK_PROJA1I", prevReqHeader.WorkOrderA ?? prevReqHeader.ProjectA);
                        arrayFields.Add("COST_CENTREA1I", "" + prevReqHeader.CostCentreA);
                        arrayFields.Add("EQUIP_REFA1I", "" + prevReqHeader.EquipmentA);
                        arrayFields.Add("DELIV_INSTRA1I", "" + prevReqHeader.DelivInstrA);
                        arrayFields.Add("DELIV_INSTRB1I", "" + prevReqHeader.DelivInstrB);
                        arrayFields.Add("PRIORITY_CODE1I", "" + prevReqHeader.PriorityCode);
                        arrayFields.Add("ANSWER_B1I", "" + prevReqHeader.AnswerB);
                        arrayFields.Add("ANSWER_D1I", "" + prevReqHeader.AnswerD);
                        arrayFields.Add("REQUESTED_BY1I", "" + prevReqHeader.RequestedBy);

                        //enviar el encabezado MSM140A
                        var request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opSheet, request);
                        //Confirmar el encabezado MSM140A
                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
                        {
                            request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenKey = "1"
                            };
                            reply = proxySheet.submit(opSheet, request);
                        }

                        //no hay errores ni advertencias
                        if (reply == null || _eFunctions.CheckReplyError(reply))
                            throw new Exception("ERROR: " + reply.message);

                        //MSM14BA
                        if (reply.mapName != "MSM14BA")
                            throw new Exception("ERROR: Se ha producido un error al crear el encabezado. No se puede acceder al MSO140/MSM14BA. " + reply.message);

                        var parItemIndex = 0; //controla el par de items por pantalla
                        //agrego los items que tenga hasta el momento al encabezado
                        foreach (var item in itemList)
                        {
                            //asigno número de vale si se genera
                            var screenValues = new ArrayScreenNameValue(reply.screenFields);
                            if (!string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                            if (parItemIndex % 2 == 0)
                                arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("QTY_REQD1I" + (parItemIndex + 1), "" + item.QuantityRequired);
                            arrayFields.Add("UOM1I" + (parItemIndex + 1), item.UnitOfMeasure);
                            arrayFields.Add("TYPE1I" + (parItemIndex + 1), "S");
                            arrayFields.Add("DESCR_A1I" + (parItemIndex + 1), item.StockCode);
                            arrayFields.Add("PART_ISSUE1I" + (parItemIndex + 1), "Y");

                            //envío si es el último item de la lista o si es el segundo de la pantalla
                            if (item == itemList[itemList.Count - 1] || parItemIndex > 0)
                            {
                                request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };
                                reply = proxySheet.submit(opSheet, request);

                                //evalúo si hay error y cancelo
                                if (_eFunctions.CheckReplyError(reply))
                                    throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);
                                //mientras confirmación o bodega
                                while (reply != null && (reply.mapName == "MSM14BA" || reply.mapName == "MSM14CA"))
                                {
                                    //evalúo error
                                    if (_eFunctions.CheckReplyError(reply))
                                        throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + reply.message);

                                    //asigno número de vale si se genera
                                    screenValues = new ArrayScreenNameValue(reply.screenFields);
                                    if (screenValues.GetField("IREQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("IREQ_NO1I").value))
                                        prevReqHeader.IreqNo = screenValues.GetField("IREQ_NO1I").value;

                                    //si es una nueva pantalla de items
                                    if (reply.mapName == "MSM14BA" && item != itemList[itemList.Count - 1])
                                        if (screenValues.GetField("DESCR_A1I1") != null && string.IsNullOrWhiteSpace(screenValues.GetField("DESCR_A1I1").value))
                                            break;
                                    ////MSM14CA  - Warehouse Holdings
                                    if (reply.mapName == "MSM14CA")
                                    {
                                        if (screenValues.GetField("TOTAL_REQD1I") == null)
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "No se ha encontrado el valor total requerido al asignar a bodega. MSM14CA");
                                        string selWarehouseIndex = "";
                                        //obtengo solo la lista de pares del objeto para actualizarla
                                        var screenArray = screenValues.ToArray();
                                        foreach (var parValue in screenArray)
                                        {
                                            if (parValue.fieldName != null && parValue.fieldName.StartsWith("WHOUSE_ID_") && parValue.value == prevReqHeader.OrigWhouseId)
                                                selWarehouseIndex = parValue.fieldName.Replace("WHOUSE_ID_", "");
                                            if (parValue.fieldName != null && parValue.fieldName.StartsWith("QTY_REQD_"))
                                                parValue.value = "";
                                        }
                                        //reingreso la lista al objeto del screen y actualizo la cantidad del w/h que quiero de acuerdo a lo realizado anteriormente
                                        if (string.IsNullOrWhiteSpace(selWarehouseIndex))
                                            throw new Exception("ERROR: Se ha producido un error al enviar el item " + item.Index + " - " + item.StockCode + ". " + "El item no está catalogado en la bodega seleccionada. MSM14CA");
                                        screenValues = new ArrayScreenNameValue(screenArray);
                                        screenValues.SetField("QTY_REQD_" + selWarehouseIndex, "" + screenValues.GetField("TOTAL_REQD1I").value);

                                        //envío el proceso
                                        request = new Screen.ScreenSubmitRequestDTO
                                        {
                                            screenFields = screenValues.ToArray(),
                                            screenKey = "1"
                                        };
                                        reply = proxySheet.submit(opSheet, request);
                                        continue; //continúa con el siguiente while
                                    }
                                    ////Confirm MSM14BA o cualquier otra confirmación que no requiera datos
                                    request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                }
                            }
                            parItemIndex++;
                            if (parItemIndex > 1)
                                parItemIndex = 0;

                            //OK del item
                            _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                            _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                            _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                            _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        }
                        //Confirmo la creación de todos los items. Si no llega a este punto es por algún problema presentado
                        foreach (var item in itemList)
                        {
                            _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = "OK";
                            _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Success;
                            _cells.GetCell(4, currentRowHeader + item.Index).Value2 = "" + prevReqHeader.IreqNo;
                            _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Success;
                        }
                    }
                    catch (Exception ex)
                    {
                        var addMessage = "" + ex.Message;
                        try
                        {
                            if (string.IsNullOrWhiteSpace(prevReqHeader.IreqNo))
                            {
                                addMessage += " .NO SE HA REALIZADO NINGUNA ACCIÓN";
                            }
                            else
                            {
                                //Eliminación del vale por el MSO140
                                _eFunctions.RevertOperation(opSheet, proxySheet);
                                //ejecutamos el programa
                                var reply = proxySheet.executeScreen(opSheet, "MSO140");
                                var screenValues = new ArrayScreenNameValue(reply.screenFields);

                                if (_eFunctions.CheckReplyError(reply))
                                    throw new Exception("" + reply.message);

                                while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM140A" && 
                                    (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")) && 
                                    (screenValues.GetField("REQ_NO1I") != null && !string.IsNullOrWhiteSpace(screenValues.GetField("REQ_NO1I").value)))
                                {
                                    var arrayFields = new ArrayScreenNameValue();
                                    arrayFields.Add("COMP_DEL1I", "D");

                                    var request = new Screen.ScreenSubmitRequestDTO
                                    {
                                        screenFields = arrayFields.ToArray(),
                                        screenKey = "1"
                                    };
                                    reply = proxySheet.submit(opSheet, request);
                                    screenValues = new ArrayScreenNameValue(reply.screenFields);

                                    if (_eFunctions.CheckReplyError(reply))
                                        throw new Exception(". ERROR AL ELIMINAR " + prevReqHeader.IreqNo + ": " + reply.message);
                                }
                                addMessage += " - VALE ELIMINADO " + prevReqHeader.IreqNo;
                            }
                        }
                        catch(Exception ex2)
                        {
                            addMessage += ex2;
                        }
                        finally
                        {
                            addMessage = addMessage.Replace("X2:0011 - INPUT REQUIRED  \"C\" TO COMPLETE OR \"D\" TO DELETE", "X2:0011 - EXISTE UNA ORDEN INCOMPLETA EN PROCESO. INGRESE AL MSO 140 PARA COMPLETARLA/ELIMINARLA");
                            foreach (var item in itemList)
                            {
                                _cells.GetCell(4, currentRowHeader + item.Index).Value2 = !string.IsNullOrWhiteSpace(prevReqHeader.IreqNo) ? prevReqHeader.IreqNo + " - ELIMINADO" : "";
                                _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Value2 = addMessage;
                                _cells.GetCell(4, currentRowHeader + item.Index).Style = StyleConstants.Error;
                                _cells.GetCell(ResultColumn, currentRowHeader + item.Index).Style = StyleConstants.Error;
                            }
                        }
                    }
                }
                #endregion
                
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
            }
        }

        /// <summary>
        /// Borra el header de un vale cuando este no se puede finalizar
        /// </summary>
        /// <param name="proxyRequisition"></param>
        /// <param name="createHeaderReply"></param>
        /// <param name="opRequisition"></param>
        private static void DeleteHeader(RequisitionService.RequisitionService proxyRequisition, RequisitionService.RequisitionServiceCreateHeaderReplyDTO createHeaderReply, RequisitionService.OperationContext opRequisition)
        {
            if (createHeaderReply == null)
                return;
            //new RequisitionService.RequisitionServiceDeleteHeaderReplyDTO();
            var deleteHeaderRequest = CreateDeleteRequestDto(createHeaderReply);

            proxyRequisition.deleteHeader(opRequisition, deleteHeaderRequest);
        }

        ///// <summary>
        ///// Borra el header de un vale cuando este no se puede finalizar usando el MSO140.
        ///// </summary>
        ///// <param name="position"></param>
        ///// <param name="requisitionHeader"></param>
        ///// <param name="urlService"></param>
        ///// <param name="district"></param>
        //private static void DeleteHeader(string urlService, string district, string position, RequisitionHeader requisitionHeader)
        //{
        //    if (requisitionHeader == null)
        //        return;
        //    //instancia del Servicio
        //    var proxyRequisition = new RequisitionService.RequisitionService();

        //    //Header
        //    var opRequisition = new RequisitionService.OperationContext();

        //    proxyRequisition.Url = urlService + "/RequisitionService";
        //    //El client conversation se ejecutó previamente en el proceso que hace llamado a este método
        //    //ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
        //    opRequisition.district = district;
        //    opRequisition.maxInstances = 100;
        //    opRequisition.position = position;
        //    opRequisition.returnWarnings = false;


        //    //new RequisitionService.RequisitionServiceDeleteHeaderReplyDTO();
        //    var deleteHeaderRequest = CreateDeleteRequestDto(requisitionHeader.GetCreateReplyHeader());

        //    proxyRequisition.deleteHeader(opRequisition, deleteHeaderRequest);
        //}

        /// <summary>
        /// Esta funcion copia el encabezado de la creacion de la requisicion en el objeto del encabezado para el borrado
        /// </summary>
        /// <param name="createHeaderReply">Encabezado de la requisicion a borrar</param>
        /// <returns></returns>
        private static RequisitionService.RequisitionServiceDeleteHeaderRequestDTO CreateDeleteRequestDto(RequisitionService.RequisitionServiceCreateHeaderReplyDTO createHeaderReply)
        {
            var deleteHeaderRequest = new RequisitionService.RequisitionServiceDeleteHeaderRequestDTO
            {
                allocPcA = createHeaderReply.allocPcA,
                allocPcB = createHeaderReply.allocPcB,
                assignToTeam = createHeaderReply.assignToTeam,
                authorisedStatusDesc = createHeaderReply.authorisedStatusDesc,
                authsdBy = createHeaderReply.authsdBy,
                authsdByName = createHeaderReply.authsdByName,
                authsdDate = createHeaderReply.authsdDate,
                authsdItmAmt = createHeaderReply.authsdItmAmt,
                authsdPosition = createHeaderReply.authsdPosition,
                authsdPositionDesc = createHeaderReply.authsdPositionDesc,
                authsdStatus = createHeaderReply.authsdStatus,
                authsdStatusDesc = createHeaderReply.authsdStatusDesc,
                authsdTime = createHeaderReply.authsdTime,
                authsdTotAmt = createHeaderReply.authsdTotAmt,
                completedDate = createHeaderReply.completedDate,
                completeItems = createHeaderReply.completeItems,
                confirmDelete = createHeaderReply.confirmDelete,
                costCentreA = createHeaderReply.costCentreA,
                costCentreB = createHeaderReply.costCentreB,
                costDistrictA = createHeaderReply.costDistrictA,
                costDistrictB = createHeaderReply.costDistrictB,
                createdBy = createHeaderReply.createdBy,
                createdByName = createHeaderReply.createdByName,
                creationDate = createHeaderReply.creationDate,
                creationTime = createHeaderReply.creationTime,
                custNo = createHeaderReply.custNo,
                custNoDesc = createHeaderReply.custNoDesc,
                delivInstrA = createHeaderReply.delivInstrA,
                delivInstrB = createHeaderReply.delivInstrB,
                delivLocation = createHeaderReply.delivLocation,
                delivLocationDesc = createHeaderReply.delivLocationDesc,
                directPurchOrd = createHeaderReply.directPurchOrd,
                districtCode = createHeaderReply.districtCode,
                districtName = createHeaderReply.districtName,
                entitlementPeriod = createHeaderReply.entitlementPeriod,
                equipmentA = createHeaderReply.equipmentA,
                equipmentB = createHeaderReply.equipmentB,
                equipmentRefA = createHeaderReply.equipmentRefA,
                equipmentRefB = createHeaderReply.equipmentRefB,
                groupClass = createHeaderReply.groupClass,
                hdr140Status = createHeaderReply.hdr140Status,
                hdr140StatusDesc = createHeaderReply.hdr140StatusDesc,
                headerType = createHeaderReply.headerType,
                inabilityDate = createHeaderReply.inabilityDate,
                inabilityRsn = createHeaderReply.inabilityRsn,
                inspectCode = createHeaderReply.inspectCode,
                inventCat = createHeaderReply.inventCat,
                inventCatDesc = createHeaderReply.inventCatDesc,
                ireqNo = createHeaderReply.ireqNo,
                ireqType = createHeaderReply.ireqType,
                issTranType = createHeaderReply.issTranType,
                issTranTypeDesc = createHeaderReply.issTranTypeDesc,
                issueRequisitionTypeDesc = createHeaderReply.issueRequisitionTypeDesc,
                lastAcqDate = createHeaderReply.lastAcqDate,
                loanDuration = createHeaderReply.loanDuration,
                loanRequisitionNo = createHeaderReply.loanRequisitionNo,
                lstAmodDate = createHeaderReply.lstAmodDate,
                lstAmodTime = createHeaderReply.lstAmodTime,
                lstAmodUser = createHeaderReply.lstAmodUser,
                matGroupCode = createHeaderReply.matGroupCode,
                matGroupCodeDesc = createHeaderReply.matGroupCodeDesc,
                moreInstr = createHeaderReply.moreInstr,
                msg140Data = createHeaderReply.msg140Data,
                narrative = createHeaderReply.narrative,
                numOfItems = createHeaderReply.numOfItems,
                orderStatusDesc = createHeaderReply.orderStatusDesc,
                origWhouseId = createHeaderReply.origWhouseId,
                origWhouseIdDesc = createHeaderReply.origWhouseIdDesc,
                partIssue = createHeaderReply.partIssue,
                partIssueSpecified = createHeaderReply.partIssueSpecified,
                password = createHeaderReply.password,
                preqNo = createHeaderReply.preqNo,
                priorityCode = createHeaderReply.priorityCode,
                projectA = createHeaderReply.projectA,
                projectB = createHeaderReply.projectB,
                protectedInd = createHeaderReply.protectedInd,
                purchaseOrdNo = createHeaderReply.purchaseOrdNo,
                purchDelivInstr = createHeaderReply.purchDelivInstr,
                purchInstr = createHeaderReply.purchInstr,
                purchInstruction = createHeaderReply.purchInstruction,
                purchOfficer = createHeaderReply.purchOfficer,
                rcvngWhouse = createHeaderReply.rcvngWhouse,
                rcvngWhouseDesc = createHeaderReply.rcvngWhouseDesc,
                relatedWhReq = createHeaderReply.relatedWhReq,
                repairRequest = createHeaderReply.repairRequest,
                requestedBy = createHeaderReply.requestedBy,
                requestedByName = createHeaderReply.requestedByName,
                requiredByDate = createHeaderReply.requiredByDate,
                requiredByPos = createHeaderReply.requiredByPos,
                requiredByPosDesc = createHeaderReply.requiredByPosDesc,
                requisitionItemStatusDesc = createHeaderReply.requisitionItemStatusDesc,
                reversePeriodStart = createHeaderReply.reversePeriodStart,
                rotnRequisitionNo = createHeaderReply.rotnRequisitionNo,
                sentType = createHeaderReply.sentType,
                sentTypeDesc = createHeaderReply.sentTypeDesc,
                statsUpdatedInd = createHeaderReply.statsUpdatedInd,
                suggestedSupp = createHeaderReply.suggestedSupp,
                surveyNo = createHeaderReply.surveyNo,
                transType = createHeaderReply.transType,
                useByDate = createHeaderReply.useByDate,
                workOrderA = createHeaderReply.workOrderA,
                workOrderB = createHeaderReply.workOrderB,
                workProjA = createHeaderReply.workProjA,
                workProjB = createHeaderReply.workProjB,
                workProjIndA = createHeaderReply.workProjIndA,
                workProjIndB = createHeaderReply.workProjIndB
            };


            return deleteHeaderRequest;
        }
        
        private void btnCleanSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
        }

        public class RequisitionHeader
        {
            public string DistrictCode;
            public string IndSerie;
            public string IreqNo;
            public string IreqType;
            public string RequestedBy;
            public string RequiredByPos;
            public string IssTranType;
            public string OrigWhouseId;
            public string PriorityCode;
           
            public string CostDistrictA;
            public string WorkOrderA;
            public string WorkProjectIndA;//Solo para el MSO140
            public string EquipmentA;
            public string ProjectA;
            public string CostCentreA;

            public string DelivInstrA;
            public string DelivInstrB;
            
            public string AllocPcA;
            public string RequiredByDate;
            public string AnswerB;
            public string AnswerD;
            
            public bool PartIssue;
            public bool PartIssueSpecified;
            public bool ProtectedInd;
            public bool PickTaskReq;
            public bool ProtectedIndSpecified;

            /// <summary>
            /// Compara el objeto encabezado RequisitionHeader con otro encabezado. Devuelve true si el encabezado es igual a objectHeader
            /// </summary>
            /// <param name="objectHeader"></param>
            /// <returns>bool: true si objectHeader es igual</returns>
            public bool Equals(RequisitionHeader objectHeader)
            {
                return DistrictCode == objectHeader.DistrictCode &&
                       IndSerie == objectHeader.IndSerie &&
                       //IreqNo == objectHeader.IreqNo && //este no se debe comparar
                       IreqType == objectHeader.IreqType &&
                       RequestedBy == objectHeader.RequestedBy &&
                       RequiredByPos == objectHeader.RequiredByPos &&
                       IssTranType == objectHeader.IssTranType &&
                       OrigWhouseId == objectHeader.OrigWhouseId &&
                       PriorityCode == objectHeader.PriorityCode &&
                       CostDistrictA == objectHeader.CostDistrictA &&
                       WorkOrderA == objectHeader.WorkOrderA &&
                       EquipmentA == objectHeader.EquipmentA &&
                       ProjectA == objectHeader.ProjectA &&
                       CostCentreA == objectHeader.CostCentreA &&
                       DelivInstrA == objectHeader.DelivInstrA &&
                       DelivInstrB == objectHeader.DelivInstrB &&
                       AllocPcA == objectHeader.AllocPcA &&
                       RequiredByDate == objectHeader.RequiredByDate &&
                       AnswerB == objectHeader.AnswerB &&
                       AnswerD == objectHeader.AnswerD &&
                       PartIssue == objectHeader.PartIssue &&
                       ProtectedInd == objectHeader.ProtectedInd &&
                       PickTaskReq == objectHeader.PickTaskReq &&
                       ProtectedIndSpecified == objectHeader.ProtectedIndSpecified;
            }

            public RequisitionService.RequisitionServiceCreateHeaderRequestDTO GetCreateRequestHeader()
            {
                var request = new RequisitionService.RequisitionServiceCreateHeaderRequestDTO
                {
                    districtCode = DistrictCode,
                    ireqNo = IreqNo,
                    ireqType = IreqType,
                    requestedBy = RequestedBy,
                    requiredByPos = RequiredByPos,
                    issTranType = IssTranType,
                    origWhouseId = OrigWhouseId,
                    priorityCode = PriorityCode,
                    costDistrictA = CostDistrictA,
                    workOrderA = GetNewWorkOrderDto(WorkOrderA),
                    equipmentA = EquipmentA,
                    projectA = ProjectA,
                    costCentreA = CostCentreA,
                    delivInstrA = DelivInstrA,
                    delivInstrB = DelivInstrB,
                    allocPcA = AllocPcA,
                    requiredByDate = RequiredByDate,
                    answerB = AnswerB,
                    answerD = AnswerD,
                    partIssue = PartIssue,
                    partIssueSpecified =  PartIssueSpecified,
                    protectedInd = ProtectedInd,
                    pickTaskReq = PickTaskReq,
                    protectedIndSpecified = ProtectedIndSpecified
                };

                return request;
            }
            public RequisitionService.RequisitionServiceCreateHeaderReplyDTO GetCreateReplyHeader()
            {
                var request = new RequisitionService.RequisitionServiceCreateHeaderReplyDTO
                {
                    districtCode = DistrictCode,
                    ireqNo = IreqNo,
                    ireqType = IreqType,
                    requestedBy = RequestedBy,
                    requiredByPos = RequiredByPos,
                    issTranType = IssTranType,
                    origWhouseId = OrigWhouseId,
                    priorityCode = PriorityCode,
                    costDistrictA = CostDistrictA,
                    workOrderA = GetNewWorkOrderDto(WorkOrderA),
                    equipmentA = EquipmentA,
                    projectA = ProjectA,
                    costCentreA = CostCentreA,
                    delivInstrA = DelivInstrA,
                    delivInstrB = DelivInstrB,
                    allocPcA = AllocPcA,
                    requiredByDate = RequiredByDate,
                    answerB = AnswerB,
                    answerD = AnswerD,
                    partIssue = PartIssue,
                    partIssueSpecified = PartIssueSpecified,
                    protectedInd = ProtectedInd,
                    pickTaskReq = PickTaskReq,
                    protectedIndSpecified = ProtectedIndSpecified
                };

                return request;
            }

            /// <summary>
            /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
            /// </summary>
            /// <param name="no">string: Número de la orden de trabajo</param>
            /// <returns>WorkOrderDTO</returns>
            public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string no)
            {
                var workOrderDto = new RequisitionService.WorkOrderDTO();
                if (string.IsNullOrWhiteSpace(no)) return workOrderDto;

                no = no.Trim();
                if (no.Length < 3)
                    throw new Exception(@"El número de orden no corresponde a una orden válida");
                workOrderDto.prefix = no.Substring(0, 2);
                workOrderDto.no = no.Substring(2, no.Length - 2);
                return workOrderDto;
            }
            /// <summary>
            /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
            /// </summary>
            /// <param name="prefix">string: prefijo de la orden de trabajo</param>
            /// <param name="no">string: Número de la orden de trabajo</param>
            /// <returns>WorkOrderDTO</returns>
            public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string prefix, string no)
            {
                var workOrderDto = new RequisitionService.WorkOrderDTO
                {
                    prefix = prefix,
                    no = no
                };

                return workOrderDto;
            }

        }

        public class RequisitionItem
        {
            public int Index;
            public string StockCode;
            public string UnitOfMeasure;
            public string ItemType;
            public decimal QuantityRequired;
            public decimal IssueRequisitionItem;
            public bool QuantityRequiredSpecified;
            public bool PartIssue;
            public bool RepairRequest;
            public bool RepairRequestProtect;
            public bool IssueDocoFlg;
            public bool PurchDocoFlg;
            public bool NarrativeExists;
            public string AlterStockCodeFlg;
            public bool IssueRequisitionItemSpecified;
            public bool RepairRequestSpecified;
            public bool RepairRequestProtectSpecified;
            public bool PartIssueSpecified;
            public bool IssueDocoFlgSpecified;
            public bool PurchDocoFlgSpecified;
            public bool NarrativeExistsSpecified;
            public bool AlterStockCodeFlgSpecified;
            
            

            public RequisitionItem()
            {
                IssueRequisitionItem = 0;
                RepairRequest = false;
                RepairRequestProtect = false;
                PartIssue = false;
                IssueDocoFlg = false;
                PurchDocoFlg = false;
                NarrativeExists = false;

                QuantityRequiredSpecified = true;

                IssueRequisitionItemSpecified = false;
                RepairRequestSpecified = false;
                RepairRequestProtectSpecified = false;
                PartIssueSpecified = false;
                IssueDocoFlgSpecified = false;
                PurchDocoFlgSpecified = false;
                NarrativeExistsSpecified = false;
                AlterStockCodeFlgSpecified = false;
            }

            public RequisitionService.RequisitionItemDTO GetRequisitionItemDto()
            {
                var item = new RequisitionService.RequisitionItemDTO
                {
                    stockCode = StockCode,
                    unitOfMeasure = UnitOfMeasure,
                    itemType = ItemType,
                    quantityRequired = QuantityRequired,
                    issueRequisitionItem = IssueRequisitionItem,
                    quantityRequiredSpecified = QuantityRequiredSpecified,
                    partIssue = PartIssue,
                    repairRequest = RepairRequest,
                    repairRequestProtect = RepairRequestProtect,
                    issueDocoFlg = IssueDocoFlg,
                    purchDocoFlg = PurchDocoFlg,
                    narrativeExists = NarrativeExists,
                    alterStockCode = AlterStockCodeFlg,
                    issueRequisitionItemSpecified = IssueRequisitionItemSpecified,
                    repairRequestSpecified = RepairRequestSpecified,
                    repairRequestProtectSpecified = RepairRequestProtectSpecified,
                    partIssueSpecified = PartIssueSpecified,
                    issueDocoFlgSpecified = IssueDocoFlgSpecified,
                    purchDocoFlgSpecified = PurchDocoFlgSpecified,
                    narrativeExistsSpecified = NarrativeExistsSpecified,
                    alterStockCodeFlgSpecified = AlterStockCodeFlgSpecified
                };


                return item;
            }
        }

        private void btnCreateReqDirectOrderItems_Click(object sender, RibbonControlEventArgs e)
        {
            _ignoreItemError = false;
            //si si ya hay un thread corriendo que no se ha detenido
            if (_thread != null && _thread.IsAlive) return;
            _thread = new Thread(CreateRequisitionScreenService);

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void btnManualCreditRequisitionMSE1VR_Click(object sender, RibbonControlEventArgs e)
        {
            _ignoreItemError = false;
            if (_thread != null && _thread.IsAlive) return;
            _thread = new Thread(ManualCreditRequisition);

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void ManualCreditRequisition()
        {
            var currentRow = TitleRow + 1;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRangeColumn(TableName01, ResultColumn);

                //instancia del Servicio
                var proxyRequisition = new IssueRequisitionItemStocklessService.IssueRequisitionItemStocklessService();

                //Header
                var opRequisition = new IssueRequisitionItemStocklessService.OperationContext();

                //Objeto para crear la coleccion de Items
                //new RequisitionService.RequisitionServiceCreateItemReplyCollectionDTO();

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                _eFunctions.SetConnectionPoolingType(false); //Se asigna por 'Pooled Connection Request Timed Out'
                proxyRequisition.Url = urlService + "/IssueRequisitionItemStocklessService";
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                opRequisition.district = _frmAuth.EllipseDsct;
                opRequisition.maxInstances = 100;
                opRequisition.position = _frmAuth.EllipsePost;
                opRequisition.returnWarnings = false;

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var headerCreateReturnReply = new IssueRequisitionItemStocklessService.ImmediateReturnStocklessDTO();


                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, currentRow).Value) != null ||
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, currentRow).Value) != null)
                {

                    try
                    {
                        string switchCase = _cells.GetNullOrTrimmedValue(_cells.GetCell(10, currentRow).Value);
                        var reference = _cells.GetNullOrTrimmedValue(_cells.GetCell(11, currentRow).Value);
                        switch (switchCase)
                        {
                            case "Work Order":
                                headerCreateReturnReply.workOrderx1 = reference;
                                break;
                            case "Equipment No.":
                                headerCreateReturnReply.equipmentReferencex1 = reference;
                                break;
                            case "Project No.":
                                headerCreateReturnReply.projectNumberx1 = reference;
                                break;
                            case "Account Code":
                                headerCreateReturnReply.costCodex1 = reference;
                                break;
                        }

                        headerCreateReturnReply.districtCode = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct)
                            ? "ICOR"
                            : _frmAuth.EllipseDsct;
                        headerCreateReturnReply.processedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.percentageAllocatedx1 = 100;
                        headerCreateReturnReply.requestedByEmployee =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.requestedByPositionId = _frmAuth.EllipsePost;
                        headerCreateReturnReply.warehouseId =
                            _cells.GetNullOrTrimmedValue(_cells.GetCell(8, currentRow).Value);
                        headerCreateReturnReply.authorisedBy =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, currentRow).Value) == null
                                ? _frmAuth.EllipseUser
                                : _cells.GetEmptyIfNull(_cells.GetCell(1, currentRow).Value);
                        headerCreateReturnReply.transactionType =
                            Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, currentRow).Value));
                        headerCreateReturnReply.requisitionNumber =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, currentRow).Value);
                        headerCreateReturnReply.processedDate =
                            DateTime.ParseExact(_cells.GetNullOrTrimmedValue(_cells.GetCell(7, currentRow).Value), "yyyyMMdd", CultureInfo.InvariantCulture);
                        headerCreateReturnReply.processedDateSpecified = true;


                        var holding = new HoldingDetailsDTO();
                        var listHolding = new List<HoldingDetailsDTO>();

                        holding.quantitySpecified = true;
                        holding.quantity = Convert.ToDecimal(_cells.GetCell(18, currentRow).Value);
                        holding.stockCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(16, currentRow).Value);

                        listHolding.Add(holding);
                        headerCreateReturnReply.holdingDetailsDTO = listHolding.ToArray();
                        
                        var result = proxyRequisition.immediateReturn(opRequisition, headerCreateReturnReply);

                        if (result.errors.Length == 0)
                        {
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Success;
                            _cells.GetCell(ResultColumn, currentRow).Value2 = "OK";
                        }
                        else
                        {
                            _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                            foreach (var e in result.errors)
                            {
                                _cells.GetCell(ResultColumn, currentRow).Value2 += " " + e.messageText ;
                            }
                        }

                        
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn, currentRow).Value2 += "ERROR: " + ex.Message;
                        _cells.GetCell(ResultColumn, currentRow).Style = StyleConstants.Error;
                        _cells.GetCell(4, currentRow).Style = StyleConstants.Error;
                    }
                    finally
                    {
                        currentRow++;
                    }

                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.SetConnectionPoolingType(true);//Se restaura por 'Pooled Connection Request Timed Out'
            }
        }
    }

    public static class Queries
    {
        public static string GetItemUnitOfIssue(string stockCode)
        {
            var sqlQuery = "SELECT UNIT_OF_ISSUE FROM ELLIPSE.MSF100 SC WHERE SC.STOCK_CODE = '" + stockCode + "' ";

            return sqlQuery;
        }

        public static string GetItemDirectOrder(string stockCode)
        {
            var sqlQuery = "SELECT SCI.DIRECT_ORDER_IND FROM ELLIPSE.MSF170 SCI WHERE STOCK_CODE = '" + stockCode + "' ";

            return sqlQuery;
        }
    }
}

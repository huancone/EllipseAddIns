using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseRequisitionClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseTransaccionesStockCodesExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName0101 = "ListaStockCodes"; //Format Requisition
        private const string SheetName0102 = "ResultadosStockCodes";
        private const string SheetName0201 = "ListaPurchaseOrders"; //Format PO
        private const string SheetName0202 = "ResultadosPurchaseOrders";
        private const string SheetName0203 = "ModificarPurchaseOrders";
        private const string SheetName0301 = "ListaPOExtended"; //Format PO Extended
        private const string SheetName0302 = "ResultadosPOExtended";
        private const string SheetName0303 = "ModificarPOExtended";
        private const string ValidationSheetName01 = "ValidationSheetSC";
        private const string ValidationSheetName02 = "ValidationSheetPO";
        private const string ValidationSheetName03 = "ValidationSheetPOExtended";

        private const int TitleRow0101 = 6;
        private const int TitleRow0102 = 5;
        private const int TitleRow0201 = 6;
        private const int TitleRow0202 = 5;
        private const int TitleRow0203 = 5;
        private const int TitleRow0301 = TitleRow0201;
        private const int TitleRow0302 = TitleRow0202;
        private const int TitleRow0303 = TitleRow0203;
        private const int ResultColumn0101 = 3;
        private const int ResultColumn0102 = 15; //aplica como indicador de ultimo registro
        private const int ResultColumn0201 = 3;
        private const int ResultColumn0202 = 23; //aplica como indicador de ultimo registro
        private const int ResultColumn0203 = 12;
        private const int ResultColumn0301 = 3;
        private const int ResultColumn0302 = 26; //aplica como indicador de ultimo registro
        private const int ResultColumn0303 = 15;
        private const string TableName0101 = "StockCodesTable";
        private const string TableName0102 = "ReviewReqSCTable";
        private const string TableName0201 = "PurchaseOrdersTable";
        private const string TableName0202 = "ReviewPOTable";
        private const string TableName0203 = "ModifyPOTable";
        private const string TableName0301 = "PurchaseOrdersExtTable";
        private const string TableName0302 = "ReviewPOExtTable";
        private const string TableName0303 = "ModifyPOExtTable";
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private ExcelStyleCells _cells;
        private Application _excelApp;

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormatRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            FormatRequisitionSheet();
        }

        private void btnFormatPurchaseOrders_Click(object sender, RibbonControlEventArgs e)
        {
            FormatPurchaseOrderSheet();
        }

        private void btnFormatPurchaseOrdersExtended_Click(object sender, RibbonControlEventArgs e)
        {
            FormatPurchaseOrderExtendedSheet();
        }

        private void btnReviewStockCodesRequisitions_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0101) ||
                _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0102))
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewReqScList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
            {
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
        }

        private void btnReviewPurchaseOrders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0201) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0202))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewPurchaseOrderList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0301) || _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0302))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewPurchaseOrderExtendedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                {
                    MessageBox.Show(
                        @"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnModifyPurchaseOrders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                if (!_cells.IsDecimalDotSeparator())
                    if (MessageBox.Show(
                            @"El separador de decimales configurado actualmente no es el punto. Usar un separador de decimales diferente puede generar errores al momento de cargar valores numéricos. ¿Está seguro que desea continuar?",
                            @"ALERTA DE SEPARADOR DE DECIMALES", MessageBoxButtons.OKCancel) != DialogResult.OK)
                        return;

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0203))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ModifyPurchaseOrderList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0303))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ModifyPurchaseOrderListExtended);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                {
                    MessageBox.Show(
                        @"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDeletePurchaseOrders_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0201) ||
                _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0203))
            {
                var dr = MessageBox.Show(
                    @"Esta acción eliminará las Órdenes de Compra existentes. ¿Está seguro que desea continuar?",
                    @"ELIMINAR PURCHASE ORDERS", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeletePurchaseOrderList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0301) ||
                     _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0303))
            {
                var dr = MessageBox.Show(
                    @"Esta acción eliminará las Órdenes de Compra existentes. ¿Está seguro que desea continuar?",
                    @"ELIMINAR PURCHASE ORDERS", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeletePurchaseOrderListExtended);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
            {
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
        }

        private void btnDeletePurchaseOrderItem_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0203))
            {
                var dr = MessageBox.Show(
                    @"Esta acción eliminará los Ítems de las Órdenes de Compra existentes. ¿Está seguro que desea continuar?",
                    @"ELIMINAR PURCHASE ORDERS ITEMS", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeletePurchaseOrderItemList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0303))
            {
                var dr = MessageBox.Show(
                    @"Esta acción eliminará los Ítems de las Órdenes de Compra existentes. ¿Está seguro que desea continuar?",
                    @"ELIMINAR PURCHASE ORDERS ITEMS", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeletePurchaseOrderItemListExtended);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
            {
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
        }

        public void FormatRequisitionSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0101;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName01);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 0101
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE TRANSACCIONES DE STOCK CODES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A4").Value = "ESTADO REQ.";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("C3").Value = "FECHA INICIAL";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("C4").Value = "FECHA FINAL";
                _cells.GetCell("C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                _cells.GetCell("E3").Value = "TIPO REQ.";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("E4").Value = "TIPO TRANS."; //TABLE_TYPE 'IT'
                _cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("G3").Value = "PRIORIDAD"; //TABLE_TYPE 'PI'
                _cells.GetCell("G3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName01, 1);

                var reqStatusList = Requisition.RequisitionStatus.GetRequisitionStatusList().Select(status => status.Value).ToList();
                reqStatusList.Add("UNCOMPLETED");
                _cells.SetValidationList(_cells.GetCell("B4"), reqStatusList, ValidationSheetName01, 2);

                var reqTypeList = Requisition.RequisitionType.GetRequisitionTypeList().Select(type => type.Key + " - " + type.Value).ToList();
                reqTypeList.Sort();
                _cells.SetValidationList(_cells.GetCell("F3"), reqTypeList, ValidationSheetName01, 3);

                var transTypeList = Requisition.TransactionType.GetTransactionTypeList(_eFunctions).Select(type => type.Key + " - " + type.Value).ToList();
                transTypeList.Sort();
                _cells.SetValidationList(_cells.GetCell("F4"), transTypeList, ValidationSheetName01, 4);

                var reqPriorityList = Requisition.PriorityCodes.GetPrioriyCodesList(_eFunctions).Select(type => type.Key + " - " + type.Value).ToList();
                reqPriorityList.Sort();
                _cells.SetValidationList(_cells.GetCell("H3"), reqPriorityList, ValidationSheetName01, 5);

                _cells.GetCell(1, TitleRow0101).Value = "STOCK_CODE";
                _cells.GetCell(1, TitleRow0101).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0101 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow0101).Value = "EVENTO";
                _cells.GetCell(2, TitleRow0101).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn0101, TitleRow0101).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0101, TitleRow0101).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0101, ResultColumn0101, TitleRow0101 + 1),
                    TableName0101);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 0102
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0102;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS STOCK CODE - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                //GENERAL
                _cells.GetRange(1, TitleRow0102, ResultColumn0102 - 1, TitleRow0102).Style =
                    StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0102).Value = "STOCK CODE";
                _cells.GetCell(2, TitleRow0102).Value = "DISTRITO";
                _cells.GetCell(3, TitleRow0102).Value = "REQ. NO";
                _cells.GetCell(4, TitleRow0102).Value = "REQ. ITEM";
                _cells.GetCell(5, TitleRow0102).Value = "REQ. TYPE";
                _cells.GetCell(6, TitleRow0102).Value = "TRAN. TYPE";
                _cells.GetCell(7, TitleRow0102).Value = "PRIORITY";
                _cells.GetCell(8, TitleRow0102).Value = "W/H";
                _cells.GetCell(9, TitleRow0102).Value = "REQD. BY.";
                _cells.GetCell(10, TitleRow0102).Value = "REQD. DATE";
                _cells.GetCell(11, TitleRow0102).Value = "QTY REQD";
                _cells.GetCell(12, TitleRow0102).Value = "PO ITEM";
                _cells.GetCell(13, TitleRow0102).Value = "ITEM STATUS";
                _cells.GetCell(14, TitleRow0102).Value = "CREATION DATE";
                _cells.GetCell(15, TitleRow0102).Value = "DELIVERY INST.";
                _cells.GetRange(1, TitleRow0102 + 1, ResultColumn0102, TitleRow0102 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0102, ResultColumn0102, TitleRow0102 + 1),
                    TableName0102);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void FormatPurchaseOrderSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.Sheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0201;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName02);

                //CONSTRUYO LA HOJA 0201
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE PURCHASE ORDERS (MSO220) - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A4").Value = "BUSCAR POR:";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Value = "PURCHASE ORDER";
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("C3").Value = "FECHA CREACIÓN INICIAL";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("C4").Value = "FECHA CREACIÓN FINAL";
                _cells.GetCell("C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                _cells.GetCell("E3").Value = "ESTADO";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F3").Style = _cells.GetStyle(StyleConstants.Select);
                //_cells.GetCell("E4").Value = "CONDICIÓN 3";
                //_cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetCell("F4").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName02, 1);
                var searchType = new List<string> {"PURCHASE ORDER", "STOCK CODE", "CONSULTAR TODO"};
                _cells.SetValidationList(_cells.GetCell("B4"), searchType, ValidationSheetName02, 2);
                //listas de validación
                var itemList1 = PurchaseOrderActions.OrderStatus.GetStatusList();
                var poStatusLust = itemList1.Select(item => item.Key + " - " + item.Value).ToList();
                _cells.SetValidationList(_cells.GetCell("F3"), poStatusLust, ValidationSheetName02, 3);

                _cells.GetCell(1, TitleRow0201).Value = "PURCH.ORD./STOCK_CODE";
                _cells.GetCell(1, TitleRow0201).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0201 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow0201).Value = "EVENTO";
                _cells.GetCell(2, TitleRow0201).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn0201, TitleRow0201).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0201, TitleRow0201).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0201, ResultColumn0201, TitleRow0201 + 1),
                    TableName0201);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO HOJA 0202
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0202;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS PURCHASE ORDERS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetRange(1, TitleRow0202, ResultColumn0202 - 1, TitleRow0202).Style =
                    StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0202).Value = "PO_NO";
                _cells.GetCell(2, TitleRow0202).Value = "PO_ITEM_NO";
                _cells.GetCell(3, TitleRow0202).Value = "PREQ_STK_CODE";
                _cells.GetCell(4, TitleRow0202).Value = "PART_NO";
                _cells.GetCell(5, TitleRow0202).Value = "MNEMONIC";
                _cells.GetCell(6, TitleRow0202).Value = "ITEM_NAME";
                _cells.GetCell(7, TitleRow0202).Value = "DESCRIPCIÓN";
                _cells.GetCell(8, TitleRow0202).Value = "CREATION_DATE";
                _cells.GetCell(9, TitleRow0202).Value = "ORDER_DATE";
                _cells.GetCell(10, TitleRow0202).Value = "ORIG_DUE_DATE";
                _cells.GetCell(11, TitleRow0202).Value = "OFST_RCPT_DATE";
                _cells.GetCell(12, TitleRow0202).Value = "ONST_RCPT_DATE";
                _cells.GetCell(8, TitleRow0202).AddComment("YYYYMMDD");
                _cells.GetCell(9, TitleRow0202).AddComment("YYYYMMDD");
                _cells.GetCell(10, TitleRow0202).AddComment("YYYYMMDD");
                _cells.GetCell(11, TitleRow0202).AddComment("YYYYMMDD");
                _cells.GetCell(12, TitleRow0202).AddComment("YYYYMMDD");
                _cells.GetCell(13, TitleRow0202).Value = "ORIG_NET_PR_I";
                _cells.GetCell(14, TitleRow0202).Value = "CURR_NET_PR_I";
                _cells.GetCell(15, TitleRow0202).Value = "ORIG_QTY_I";
                _cells.GetCell(16, TitleRow0202).Value = "CURR_QTY_I";
                _cells.GetCell(17, TitleRow0202).Value = "QTY_RCV_OFST_I";
                _cells.GetCell(18, TitleRow0202).Value = "QTY_RCV_DIR_I";
                _cells.GetCell(19, TitleRow0202).Value = "FREIGHT_CODE";
                _cells.GetCell(20, TitleRow0202).Value = "DELIV_LOCATION";
                _cells.GetCell(21, TitleRow0202).Value = "EXPEDITE_CODE";
                _cells.GetCell(22, TitleRow0202).Value = "SUPPLIER_NO";
                _cells.GetCell(23, TitleRow0202).Value = "SUPPLIER_NAME";
                _cells.GetRange(1, TitleRow0202 + 1, ResultColumn0202, TitleRow0202 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0202, ResultColumn0202, TitleRow0202 + 1),
                    TableName0202);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO HOJA 0203
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0203;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "MODIFICAR PURCHASE ORDERS MSO220/1 - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetRange(1, TitleRow0203, ResultColumn0203 - 1, TitleRow0203).Style =
                    StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0203).Value = "PO_NO";
                _cells.GetCell(2, TitleRow0203).Value = "ITEM";
                _cells.GetCell(2, TitleRow0203).AddComment("Índice de posición en la orden");
                _cells.GetCell(3, TitleRow0203).Value = "QTY_ORDERED";
                _cells.GetCell(4, TitleRow0203).Value = "DUE_DATE";
                _cells.GetCell(4, TitleRow0203).AddComment("YYYYMMDD");
                _cells.GetCell(5, TitleRow0203).Value = "FREIGHT_CODE";
                _cells.GetCell(6, TitleRow0203).Value = "DELIV_LOCATION";
                _cells.GetCell(7, TitleRow0203).Value = "EXPEDITE_CODE";
                _cells.GetCell(8, TitleRow0203).Value = "DISCOUNT 1";
                _cells.GetCell(9, TitleRow0203).Value = "SURCHARGE 1";
                _cells.GetCell(10, TitleRow0203).Value = "DISCOUNT 2";
                _cells.GetCell(11, TitleRow0203).Value = "SURCHARGE 2";
                _cells.GetCell(ResultColumn0203, TitleRow0203).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0203, TitleRow0203).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow0203 + 1, ResultColumn0203, TitleRow0203 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0203, ResultColumn0203, TitleRow0203 + 1),
                    TableName0203);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void FormatPurchaseOrderExtendedSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.Sheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0301;
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName03);

                //CONSTRUYO LA HOJA 0301
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "CONSULTA DE PURCHASE ORDERS (MSO220) - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A4").Value = "BUSCAR POR:";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Value = "PURCHASE ORDER";
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("C3").Value = "FECHA CREACIÓN INICIAL";
                _cells.GetCell("C3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("C4").Value = "FECHA CREACIÓN FINAL";
                _cells.GetCell("C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D4").AddComment("yyyyMMdd");
                _cells.GetCell("E3").Value = "ESTADO";
                _cells.GetCell("E3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("F3").Style = _cells.GetStyle(StyleConstants.Select);
                //_cells.GetCell("E4").Value = "CONDICIÓN 3";
                //_cells.GetCell("E4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetCell("F4").Style = _cells.GetStyle(StyleConstants.Select);

                //adicionamos las listas de validación
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName03, 1);
                var searchType = new List<string> {"PURCHASE ORDER", "STOCK CODE", "CONSULTAR TODO"};
                _cells.SetValidationList(_cells.GetCell("B4"), searchType, ValidationSheetName03, 2);
                //listas de validación
                var itemList1 = PurchaseOrderActions.OrderStatus.GetStatusList();
                var poStatusLust = itemList1.Select(item => item.Key + " - " + item.Value).ToList();
                _cells.SetValidationList(_cells.GetCell("F3"), poStatusLust, ValidationSheetName03, 3);

                _cells.GetCell(1, TitleRow0301).Value = "PURCH.ORD./STOCK_CODE";
                _cells.GetCell(1, TitleRow0301).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0301 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(2, TitleRow0301).Value = "EVENTO";
                _cells.GetCell(2, TitleRow0301).Style = StyleConstants.TitleInformation;

                _cells.GetCell(ResultColumn0301, TitleRow0301).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0301, TitleRow0301).Style = _cells.GetStyle(StyleConstants.TitleResult);


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0301, ResultColumn0301, TitleRow0301 + 1),
                    TableName0301);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO HOJA 0302
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0302;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESULTADO CONSULTAS PURCHASE ORDERS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetRange(1, TitleRow0302, ResultColumn0302 - 1, TitleRow0302).Style =
                    StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0302).Value = "PO_NO";
                _cells.GetCell(2, TitleRow0302).Value = "PO_ITEM_NO";
                _cells.GetCell(3, TitleRow0302).Value = "PREQ_STK_CODE";
                _cells.GetCell(4, TitleRow0302).Value = "PART_NO";
                _cells.GetCell(5, TitleRow0302).Value = "MNEMONIC";
                _cells.GetCell(6, TitleRow0302).Value = "ITEM_NAME";
                _cells.GetCell(7, TitleRow0302).Value = "DESCRIPCIÓN";
                _cells.GetCell(8, TitleRow0302).Value = "CREATION_DATE";
                _cells.GetCell(9, TitleRow0302).Value = "ORDER_DATE";
                _cells.GetCell(10, TitleRow0302).Value = "ORIG_DUE_DATE";
                _cells.GetCell(11, TitleRow0302).Value = "OFST_RCPT_DATE";
                _cells.GetCell(12, TitleRow0302).Value = "ONST_RCPT_DATE";
                _cells.GetCell(8, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(9, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(10, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(11, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(12, TitleRow0302).AddComment("YYYYMMDD");
                _cells.GetCell(13, TitleRow0302).Value = "ORIG_NET_PR_I";
                _cells.GetCell(14, TitleRow0302).Value = "CURR_NET_PR_I";
                _cells.GetCell(15, TitleRow0302).Value = "GROSS_PRICE_P";
                _cells.GetCell(16, TitleRow0302).Value = "UNIT_OF_PURCHASE";
                _cells.GetCell(17, TitleRow0302).Value = "CONV_FACTOR";
                _cells.GetCell(18, TitleRow0302).Value = "ORIG_QTY_I";
                _cells.GetCell(19, TitleRow0302).Value = "CURR_QTY_I";
                _cells.GetCell(20, TitleRow0302).Value = "QTY_RCV_OFST_I";
                _cells.GetCell(21, TitleRow0302).Value = "QTY_RCV_DIR_I";
                _cells.GetCell(22, TitleRow0302).Value = "FREIGHT_CODE";
                _cells.GetCell(23, TitleRow0302).Value = "DELIV_LOCATION";
                _cells.GetCell(24, TitleRow0302).Value = "EXPEDITE_CODE";
                _cells.GetCell(25, TitleRow0302).Value = "SUPPLIER_NO";
                _cells.GetCell(26, TitleRow0302).Value = "SUPPLIER_NAME";
                _cells.GetRange(1, TitleRow0302 + 1, ResultColumn0302, TitleRow0302 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0302, ResultColumn0302, TitleRow0302 + 1),
                    TableName0302);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO HOJA 0303
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName0303;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "MODIFICAR PURCHASE ORDERS MSO220/1 - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetRange(1, TitleRow0303, ResultColumn0303 - 1, TitleRow0303).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow0303).Value = "PO_NO";
                _cells.GetCell(2, TitleRow0303).Value = "ITEM";
                _cells.GetCell(2, TitleRow0303).AddComment("Índice de posición en la orden");
                _cells.GetCell(3, TitleRow0303).Value = "QTY_ORDERED";
                _cells.GetCell(4, TitleRow0303).Value = "DUE_DATE";
                _cells.GetCell(4, TitleRow0303).AddComment("YYYYMMDD");
                _cells.GetCell(5, TitleRow0303).Value = "GROSS_PRICE_P";
                _cells.GetCell(6, TitleRow0303).Value = "UNIT_OF_PURCHASE";
                _cells.GetCell(7, TitleRow0303).Value = "CONV_FACTOR";
                _cells.GetCell(8, TitleRow0303).Value = "FREIGHT_CODE";
                _cells.GetCell(9, TitleRow0303).Value = "DELIV_LOCATION";
                _cells.GetCell(10, TitleRow0303).Value = "EXPEDITE_CODE";
                _cells.GetCell(11, TitleRow0203).Value = "DISCOUNT 1";
                _cells.GetCell(12, TitleRow0203).Value = "SURCHARGE 1";
                _cells.GetCell(13, TitleRow0203).Value = "DISCOUNT 2";
                _cells.GetCell(14, TitleRow0203).Value = "SURCHARGE 2";
                _cells.GetCell(ResultColumn0303, TitleRow0303).Value = "RESULTADO";
                _cells.GetCell(ResultColumn0303, TitleRow0303).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow0303 + 1, ResultColumn0303, TitleRow0303 + 1).NumberFormat =
                    NumberFormatConstants.Text;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow0303, ResultColumn0303, TitleRow0303 + 1),
                    TableName0303);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja." + ex.Message);
            }
        }

        public void ReviewReqScList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var scCells = new ExcelStyleCells(_excelApp, SheetName0101);
            scCells.SetAlwaysActiveSheet(false);

            var resultCells = new ExcelStyleCells(_excelApp, SheetName0102);
            resultCells.SetAlwaysActiveSheet(false);
            resultCells.ClearTableRange(TableName0102);

            var districtCode = _cells.GetEmptyIfNull(scCells.GetCell(2, 3).Value2);
            var scStatus = _cells.GetEmptyIfNull(scCells.GetCell(2, 4).Value2);
            var startDate = _cells.GetEmptyIfNull(scCells.GetCell(4, 3).Value2);
            var endDate = _cells.GetEmptyIfNull(scCells.GetCell(4, 4).Value2);
            var reqType = _cells.GetEmptyIfNull(scCells.GetCell(6, 3).Value2);
            var transType = _cells.GetEmptyIfNull(scCells.GetCell(6, 4).Value2);
            var priorityCode = _cells.GetEmptyIfNull(scCells.GetCell(8, 3).Value2);

            if (reqType != null && reqType.Contains(" - "))
                reqType = reqType.Substring(0, reqType.IndexOf(" - ", StringComparison.Ordinal));
            if (transType != null && transType.Contains(" - "))
                transType = transType.Substring(0, transType.IndexOf(" - ", StringComparison.Ordinal));
            if (priorityCode != null && priorityCode.Contains(" - "))
                priorityCode = priorityCode.Substring(0, priorityCode.IndexOf(" - ", StringComparison.Ordinal));
            var j = TitleRow0101 + 1; //itera según cada stock code
            var i = TitleRow0102 + 1; //itera la celda para cada req-sc

            while (!string.IsNullOrEmpty("" + scCells.GetCell(1, j).Value))
                try
                {
                    var stockCode = _cells.GetEmptyIfNull(scCells.GetCell(1, j).Value2);
                    stockCode = stockCode != null && stockCode.Length < 9 ? stockCode.PadLeft(9, '0') : stockCode;

                    if (!string.IsNullOrWhiteSpace(scStatus) && !scStatus.Equals("UNCOMPLETED"))
                        scStatus = Requisition.ItemStatus.GetStatusCode(scStatus);

                    var sqlQuery = Queries.GetFetchRequisitionStockCodeQuery(_eFunctions.dbReference,
                        _eFunctions.dbLink, districtCode, stockCode, scStatus, startDate, endDate, reqType, transType,
                        priorityCode);

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {
                        resultCells.GetCell(1, i).Value = "" + odr["STOCK_CODE"];
                        resultCells.GetCell(2, i).Value = "" + odr["DSTRCT_CODE"];
                        resultCells.GetCell(3, i).Value = "" + odr["IREQ_NO"];
                        resultCells.GetCell(4, i).Value = "" + odr["IREQ_ITEM"];
                        resultCells.GetCell(5, i).Value = "" + odr["IREQ_TYPE"];
                        resultCells.GetCell(6, i).Value = "" + odr["ISS_TRAN_TYPE"];
                        resultCells.GetCell(7, i).Value = "" + odr["PRIORITY_CODE"];
                        resultCells.GetCell(8, i).Value = "" + odr["WHOUSE_ID"];
                        resultCells.GetCell(9, i).Value = "" + odr["REQUESTED_BY"];
                        resultCells.GetCell(10, i).Value = "" + odr["REQ_BY_DATE"];
                        resultCells.GetCell(11, i).Value = "" + odr["QTY_REQ"];
                        resultCells.GetCell(12, i).Value = "" + odr["PO_ITEM_NO"];
                        resultCells.GetCell(13, i).Value = "" + odr["ITEM_141_STAT"];
                        resultCells.GetCell(14, i).Value = "" + odr["CREATION_DATE"];
                        resultCells.GetCell(15, i).Value = "" + odr["DELIV_INSTR_A"] + odr["DELIV_INSTR_B"];
                        i++;
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0102))
                            resultCells.GetCell(3, i).Select();
                    }

                    scCells.GetCell(ResultColumn0101, j).Style = StyleConstants.Success;
                    scCells.GetCell(ResultColumn0101, j).Value = "SUCCESS";
                }
                catch (Exception ex)
                {
                    scCells.GetCell(ResultColumn0101, j).Style = StyleConstants.Error;
                    scCells.GetCell(ResultColumn0101, j).Value += "ERROR:" + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewReqScList()", ex.Message);
                }
                finally
                {
                    scCells.GetCell(2, j).Value = "CONSULTA DE VALES-ITEMS";
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0101))
                        scCells.GetCell(1, j).Select();
                    _eFunctions.CloseConnection();
                    j++; //aumenta SC
                }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        public void ReviewPurchaseOrderList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var poCells = new ExcelStyleCells(_excelApp, SheetName0201);
            poCells.SetAlwaysActiveSheet(false);

            var resultCells = new ExcelStyleCells(_excelApp, SheetName0202);
            resultCells.SetAlwaysActiveSheet(false);
            resultCells.ClearTableRange(TableName0202);

            var fullSearch = false; //para realizar búsquedas completas que no dependan de un PO dado
            var districtCode = _cells.GetEmptyIfNull(poCells.GetCell(2, 3).Value2);
            var searchType = _cells.GetEmptyIfNull(poCells.GetCell(2, 4).Value2);
            var startDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 3).Value2);
            var endDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 4).Value2);
            var poStatus = _cells.GetEmptyIfNull(poCells.GetCell(6, 3).Value2);

            if (poStatus != null && poStatus.Contains(" - "))
                poStatus = poStatus.Substring(0, poStatus.IndexOf(" - ", StringComparison.Ordinal));

            var j = TitleRow0201 + 1; //itera según cada stock code
            var i = TitleRow0202 + 1; //itera la celda para cada req-sc

            if (searchType != null && (string.IsNullOrWhiteSpace(searchType) || searchType.Equals("CONSULTAR TODO")))
            {
                if (string.IsNullOrWhiteSpace(startDate) && string.IsNullOrWhiteSpace(poStatus)) //TO DO
                    throw new NullReferenceException(
                        "Debe seleccionar una fecha inicial o un estado de orden para esta búsqueda");
                fullSearch = true;
            }

            while (!string.IsNullOrEmpty("" + poCells.GetCell(1, j).Value) || fullSearch)
                try
                {
                    string purchaseOrder = null;
                    string stockCode = null;
                    if (searchType.Equals("PURCHASE ORDER"))
                        purchaseOrder = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);
                    if (searchType.Equals("STOCK CODE"))
                        stockCode = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);

                    stockCode = stockCode != null && stockCode.Length < 9 ? stockCode.PadLeft(9, '0') : stockCode;

                    var sqlQuery = Queries.GetFetchPurchaseOrderQuery(_eFunctions.dbReference, _eFunctions.dbLink,
                        districtCode, purchaseOrder, stockCode, startDate, endDate, poStatus);

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {
                        resultCells.GetCell(01, i).Value = "" + odr["PO_NO"];
                        resultCells.GetCell(02, i).Value = "" + odr["PO_ITEM_NO"];
                        resultCells.GetCell(03, i).Value = "" + odr["PREQ_STK_CODE"];
                        resultCells.GetCell(04, i).Value = "" + odr["PART_NO"];
                        resultCells.GetCell(05, i).Value = "" + odr["MNEMONIC"];
                        resultCells.GetCell(06, i).Value = "" + odr["ITEM_NAME"];
                        resultCells.GetCell(07, i).Value = "" + resultCells.GetEmptyIfNull(odr["DESC_LINEX1"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX2"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX3"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX4"]);
                        resultCells.GetCell(08, i).Value = "" + odr["CREATION_DATE"];
                        resultCells.GetCell(09, i).Value = "" + odr["ORDER_DATE"];
                        resultCells.GetCell(10, i).Value = "" + odr["ORIG_DUE_DATE"];
                        resultCells.GetCell(11, i).Value = "" + odr["OFST_RCPT_DATE"];
                        resultCells.GetCell(12, i).Value = "" + odr["ONST_RCPT_DATE"];
                        resultCells.GetCell(13, i).Value = "" + odr["ORIG_NET_PR_I"];
                        resultCells.GetCell(14, i).Value = "" + odr["CURR_NET_PR_I"];
                        resultCells.GetCell(15, i).Value = "" + odr["ORIG_QTY_I"];
                        resultCells.GetCell(16, i).Value = "" + odr["CURR_QTY_I"];
                        resultCells.GetCell(17, i).Value = "" + odr["QTY_RCV_OFST_I"];
                        resultCells.GetCell(18, i).Value = "" + odr["QTY_RCV_DIR_I"];
                        resultCells.GetCell(19, i).Value = "" + odr["FREIGHT_CODE"];
                        resultCells.GetCell(20, i).Value = "" + odr["DELIV_LOCATION"];
                        resultCells.GetCell(21, i).Value = "" + odr["EXPEDITE_CODE"];
                        resultCells.GetCell(22, i).Value = "" + odr["SUPPLIER_NO"];
                        resultCells.GetCell(23, i).Value = "" + odr["SUPPLIER_NAME"];

                        i++;
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0202))
                            resultCells.GetCell(3, i).Select();
                    }

                    poCells.GetCell(ResultColumn0201, j).Style = StyleConstants.Success;
                    poCells.GetCell(ResultColumn0201, j).Value = "SUCCESS";
                }
                catch (Exception ex)
                {
                    poCells.GetCell(ResultColumn0201, j).Style = StyleConstants.Error;
                    poCells.GetCell(ResultColumn0201, j).Value += "ERROR:" + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewPurchaseOrderList()", ex.Message);
                }
                finally
                {
                    poCells.GetCell(2, j).Value = "CONSULTA DE PO-ITEMS";
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0201))
                        poCells.GetCell(1, j).Select();
                    j++; //aumenta SC
                    if (fullSearch)
                        fullSearch = false;
                }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        public void ReviewPurchaseOrderExtendedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);


            var poCells = new ExcelStyleCells(_excelApp, SheetName0301);
            poCells.SetAlwaysActiveSheet(false);

            var resultCells = new ExcelStyleCells(_excelApp, SheetName0302);
            resultCells.SetAlwaysActiveSheet(false);
            resultCells.ClearTableRange(TableName0302);

            var fullSearch = false; //para realizar búsquedas completas que no dependan de un PO dado
            var districtCode = _cells.GetEmptyIfNull(poCells.GetCell(2, 3).Value2);
            var searchType = _cells.GetEmptyIfNull(poCells.GetCell(2, 4).Value2);
            var startDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 3).Value2);
            var endDate = _cells.GetEmptyIfNull(poCells.GetCell(4, 4).Value2);
            var poStatus = _cells.GetEmptyIfNull(poCells.GetCell(6, 3).Value2);

            if (poStatus != null && poStatus.Contains(" - "))
                poStatus = poStatus.Substring(0, poStatus.IndexOf(" - ", StringComparison.Ordinal));

            var j = TitleRow0301 + 1; //itera según cada stock code
            var i = TitleRow0302 + 1; //itera la celda para cada req-sc

            if (searchType != null && (string.IsNullOrWhiteSpace(searchType) || searchType.Equals("CONSULTAR TODO")))
            {
                if (string.IsNullOrWhiteSpace(startDate) && string.IsNullOrWhiteSpace(poStatus)) //TO DO
                    throw new NullReferenceException(
                        "Debe seleccionar una fecha inicial o un estado de orden para esta búsqueda");
                fullSearch = true;
            }

            while (!string.IsNullOrEmpty("" + poCells.GetCell(1, j).Value) || fullSearch)
                try
                {
                    string purchaseOrder = null;
                    string stockCode = null;
                    if (searchType.Equals("PURCHASE ORDER"))
                        purchaseOrder = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);
                    if (searchType.Equals("STOCK CODE"))
                        stockCode = _cells.GetEmptyIfNull(poCells.GetCell(1, j).Value2);

                    stockCode = stockCode != null && stockCode.Length < 9 ? stockCode.PadLeft(9, '0') : stockCode;

                    var sqlQuery = Queries.GetFetchPurchaseOrderQuery(_eFunctions.dbReference, _eFunctions.dbLink,
                        districtCode, purchaseOrder, stockCode, startDate, endDate, poStatus);

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {
                        resultCells.GetCell(01, i).Value = "" + odr["PO_NO"];
                        resultCells.GetCell(02, i).Value = "" + odr["PO_ITEM_NO"];
                        resultCells.GetCell(03, i).Value = "" + odr["PREQ_STK_CODE"];
                        resultCells.GetCell(04, i).Value = "" + odr["PART_NO"];
                        resultCells.GetCell(05, i).Value = "" + odr["MNEMONIC"];
                        resultCells.GetCell(06, i).Value = "" + odr["ITEM_NAME"];
                        resultCells.GetCell(07, i).Value = "" + resultCells.GetEmptyIfNull(odr["DESC_LINEX1"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX2"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX3"]) +
                                                           resultCells.GetEmptyIfNull(odr["DESC_LINEX4"]);
                        resultCells.GetCell(08, i).Value = "" + odr["CREATION_DATE"];
                        resultCells.GetCell(09, i).Value = "" + odr["ORDER_DATE"];
                        resultCells.GetCell(10, i).Value = "" + odr["ORIG_DUE_DATE"];
                        resultCells.GetCell(11, i).Value = "" + odr["OFST_RCPT_DATE"];
                        resultCells.GetCell(12, i).Value = "" + odr["ONST_RCPT_DATE"];
                        resultCells.GetCell(13, i).Value = "" + odr["ORIG_NET_PR_I"];
                        resultCells.GetCell(14, i).Value = "" + odr["CURR_NET_PR_I"];
                        resultCells.GetCell(15, i).Value = "" + odr["GROSS_PRICE_P"];
                        resultCells.GetCell(16, i).Value = "" + odr["UNIT_OF_PURCH"];
                        resultCells.GetCell(17, i).Value = "" + odr["CONV_FACTOR"];
                        resultCells.GetCell(18, i).Value = "" + odr["ORIG_QTY_I"];
                        resultCells.GetCell(19, i).Value = "" + odr["CURR_QTY_I"];
                        resultCells.GetCell(20, i).Value = "" + odr["QTY_RCV_OFST_I"];
                        resultCells.GetCell(21, i).Value = "" + odr["QTY_RCV_DIR_I"];
                        resultCells.GetCell(22, i).Value = "" + odr["FREIGHT_CODE"];
                        resultCells.GetCell(23, i).Value = "" + odr["DELIV_LOCATION"];
                        resultCells.GetCell(24, i).Value = "" + odr["EXPEDITE_CODE"];
                        resultCells.GetCell(25, i).Value = "" + odr["SUPPLIER_NO"];
                        resultCells.GetCell(26, i).Value = "" + odr["SUPPLIER_NAME"];

                        i++;
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0302))
                            resultCells.GetCell(3, i).Select();
                    }

                    poCells.GetCell(ResultColumn0301, j).Style = StyleConstants.Success;
                    poCells.GetCell(ResultColumn0301, j).Value = "SUCCESS";
                }
                catch (Exception ex)
                {
                    poCells.GetCell(ResultColumn0301, j).Style = StyleConstants.Error;
                    poCells.GetCell(ResultColumn0301, j).Value += "ERROR:" + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewPurchaseOrderList()", ex.Message);
                }
                finally
                {
                    poCells.GetCell(2, j).Value = "CONSULTA DE PO-ITEMS";
                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName0301))
                        poCells.GetCell(1, j).Select();
                    j++; //aumenta SC
                    if (fullSearch)
                        fullSearch = false;
                }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        public void DeletePurchaseOrderList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0201 &&
                    _excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0203)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0201 + 1;
                    var resultC = ResultColumn0201;
                    var firstSheetActive = true;
                    if (_excelApp.ActiveSheet.Equals(ResultColumn0203))
                    {
                        i = TitleRow0203 + 1;
                        resultC = ResultColumn0203;
                        firstSheetActive = false;
                    }

                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            var resultado = DeletePurchaseOrder(opSheet, proxySheet, pOrder);


                            if (resultado)
                            {
                                if (firstSheetActive)
                                {
                                    _cells.GetCell(resultC, i).Value = "SUCCESS";
                                    _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                                }
                                else
                                {
                                    _cells.GetCell(resultC, i).Value = "ORDEN ELIMINADA";
                                }

                                _cells.GetCell(resultC, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(resultC, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                if (firstSheetActive)
                                    _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                                _cells.GetCell(resultC, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(resultC, i).Value = "ERROR: " + ex.Message;
                            if (firstSheetActive)
                                _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                            _cells.GetCell(resultC, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(resultC, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void DeletePurchaseOrderListExtended()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0301 &&
                    _excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0303)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0301 + 1;
                    var resultC = ResultColumn0301;
                    var firstSheetActive = true;
                    if (_excelApp.ActiveSheet.Equals(ResultColumn0303))
                    {
                        i = TitleRow0303 + 1;
                        resultC = ResultColumn0303;
                        firstSheetActive = false;
                    }

                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            var resultado = DeletePurchaseOrder(opSheet, proxySheet, pOrder);


                            if (resultado)
                            {
                                if (firstSheetActive)
                                {
                                    _cells.GetCell(resultC, i).Value = "SUCCESS";
                                    _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                                }
                                else
                                {
                                    _cells.GetCell(resultC, i).Value = "ORDEN ELIMINADA";
                                }

                                _cells.GetCell(resultC, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(resultC, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                if (firstSheetActive)
                                    _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                                _cells.GetCell(resultC, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(resultC, i).Value = "ERROR: " + ex.Message;
                            if (firstSheetActive)
                                _cells.GetCell(2, i).Value = "ELIMINAR ORDEN DE COMPRA";
                            _cells.GetCell(resultC, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(resultC, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void DeletePurchaseOrderItemList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0203)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0203 + 1;

                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            var item = new PurchaseOrderItem
                            {
                                Index = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2)
                            };

                            var resultado = DeletePurchaseOrderItem(opSheet, proxySheet, pOrder, item);


                            if (resultado)
                            {
                                _cells.GetCell(ResultColumn0203, i).Value = "ITEM ELIMINADO";
                                _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn0203, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn0203, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(ResultColumn0203, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void DeletePurchaseOrderItemListExtended()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0303)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0303 + 1;

                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            var item = new PurchaseOrderItem
                            {
                                Index = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2)
                            };

                            var resultado = DeletePurchaseOrderItem(opSheet, proxySheet, pOrder, item);


                            if (resultado)
                            {
                                _cells.GetCell(ResultColumn0303, i).Value = "ITEM ELIMINADO";
                                _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn0303, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn0303, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(ResultColumn0303, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeletePurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public bool DeletePurchaseOrder(Screen.OperationContext opContext, Screen.ScreenService proxySheet,
            PurchaseOrder purchaseOrder)
        {
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO220");
            //Validamos el ingreso
            if (reply.mapName != "MSM220A") return false;

            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "" + 2);
            arrayFields.Add("PO_NO1I", purchaseOrder.PurchaseNumber);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            //no hay errores ni advertencias
            if (reply == null)
                throw new Exception(@"No se ha podido realizar la eliminación");
            if (_eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            if (reply.mapName != "MSM220B")
                throw new ArgumentException("Se ha producido un fallo al realizar la acción");

            request = new Screen.ScreenSubmitRequestDTO();
            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("REPLY_22I", "Y");
            request.screenFields = arrayFields.ToArray();
            request.screenKey = "1";
            reply = proxySheet.submit(opContext, request);

            while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM220B" &&
                   (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" ||
                    reply.functionKeys.StartsWith("XMIT-WARNING")))
            {
                request.screenFields = arrayFields.ToArray();
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }

            if (!_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply)) return true;
            if (reply != null)
                throw new ArgumentException(reply.message);
            throw new Exception(@"No se ha podido obtener respuesta del servidor");
        }

        public bool DeletePurchaseOrderItem(Screen.OperationContext opContext, Screen.ScreenService proxySheet,
            PurchaseOrder purchaseOrder, PurchaseOrderItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Index))
                throw new NullReferenceException("Debe ingresar el índice del item del vale que desea eliminar");
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO220");
            //Validamos el ingreso
            if (reply.mapName != "MSM220A") return false;

            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "" + 2);
            arrayFields.Add("PO_NO1I", purchaseOrder.PurchaseNumber);
            arrayFields.Add("PO_ITEM_NO1I", item.Index);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            //no hay errores ni advertencias
            if (reply == null)
                throw new Exception(@"No se ha podido realizar la eliminación");
            if (_eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            if (reply.mapName != "MSM22CA")
                throw new ArgumentException("Se ha producido un fallo al realizar la acción");

            request = new Screen.ScreenSubmitRequestDTO();
            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("REPLY_11I", "Y");
            request.screenFields = arrayFields.ToArray();
            request.screenKey = "1";
            reply = proxySheet.submit(opContext, request);

            while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM22CA" &&
                   (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" ||
                    reply.functionKeys.StartsWith("XMIT-WARNING")))
            {
                request.screenFields = arrayFields.ToArray();
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }

            if (!_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply)) return true;
            if (reply != null)
                throw new ArgumentException(reply.message);
            throw new Exception(@"No se ha podido obtener respuesta del servidor");
        }

        public void ModifyPurchaseOrderList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            try
            {
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0203)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0203 + 1;
                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            pOrder.FreightCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value2);
                            pOrder.DeliveryLocation = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value2);
                            var item = new PurchaseOrderItem
                            {
                                Index = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2),
                                Quantity = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value2),
                                DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value2),
                                ExpediteCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value2),
                                Discount1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value2),
                                Surcharge1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value2),
                                Discount2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value2),
                                Surcharge2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value2)
                            };

                            item.Discount1 = item.Discount1 == "0" ? "" : item.Discount1;
                            item.Discount2 = item.Discount2 == "0" ? "" : item.Discount2;
                            item.Surcharge1 = item.Surcharge1 == "0" ? "" : item.Surcharge1;
                            item.Surcharge2 = item.Surcharge2 == "0" ? "" : item.Surcharge2;

                            var resultado = ModifyPurchaseOrder(opSheet, proxySheet, pOrder, item);

                            if (resultado)
                            {
                                _cells.GetCell(ResultColumn0203, i).Value = "SUCCESS";
                                _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn0203, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn0203, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn0203, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:ModifyPurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(ResultColumn0203, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:ModifyPurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void ModifyPurchaseOrderListExtended()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName0303)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnvironment.SelectedItem.Label != null && !drpEnvironment.SelectedItem.Label.Equals(""))
                {
                    var i = TitleRow0303 + 1;
                    while ("" + _cells.GetCell(1, i).Value != "")
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                returnWarnings = Debugger.DebugWarnings
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                            // ReSharper disable once UseObjectOrCollectionInitializer
                            var pOrder = new PurchaseOrder();
                            pOrder.PurchaseNumber = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                            pOrder.FreightCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value2);
                            pOrder.DeliveryLocation = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value2);
                            var item = new PurchaseOrderItem
                            {
                                Index = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2),
                                Quantity = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value2),
                                DueDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value2),
                                GrossPrice = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value2),
                                UnitOfPurchase = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value2),
                                ConversionFactor = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value2),
                                ExpediteCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value2),
                                Discount1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value2),
                                Surcharge1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value2),
                                Discount2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value2),
                                Surcharge2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value2)
                            };

                            item.Discount1 = item.Discount1 == "0" ? "" : item.Discount1;
                            item.Discount2 = item.Discount2 == "0" ? "" : item.Discount2;
                            item.Surcharge1 = item.Surcharge1 == "0" ? "" : item.Surcharge1;
                            item.Surcharge2 = item.Surcharge2 == "0" ? "" : item.Surcharge2;

                            var resultado = ModifyPurchaseOrder(opSheet, proxySheet, pOrder, item);

                            if (resultado)
                            {
                                _cells.GetCell(ResultColumn0303, i).Value = "SUCCESS";
                                _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Success;
                            }
                            else
                            {
                                _cells.GetCell(ResultColumn0303, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Warning;
                            }
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn0303, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn0303, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:ModifyPurchaseOrderList()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            _cells.GetCell(ResultColumn0303, i).Select();
                            i++;
                        }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:ModifyPurchaseOrderList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public bool ModifyPurchaseOrder(Screen.OperationContext opContext, Screen.ScreenService proxySheet,
            PurchaseOrder purchaseOrder, PurchaseOrderItem item)
        {
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO220");
            //Validamos el ingreso
            if (reply.mapName != "MSM220A") return false;

            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "" + 1);
            arrayFields.Add("PO_NO1I", purchaseOrder.PurchaseNumber);
            arrayFields.Add("PO_ITEM_NO1I", item.Index);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            //no hay errores ni advertencias
            if (reply == null)
                throw new Exception(@"No se ha podido realizar ninguna acción");
            if (_eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            if (reply.mapName != "MSM22CA")
                throw new ArgumentException("Se ha producido un fallo al intentar ejecutar el programa");

            arrayFields = new ArrayScreenNameValue();

            if (reply.mapName == "MSM220B")
                throw new ArgumentException("Debe ingresar el índice del item a modificar");

            if (reply.mapName == "MSM22CA")
            {
                if (item.Quantity != null) arrayFields.Add("CURR_QTY_P1I", item.Quantity);
                if (item.DueDate != null) arrayFields.Add("DUE_DATE1I", item.DueDate);
                if (item.GrossPrice != null) arrayFields.Add("GROSS_PR_UOP1I", item.GrossPrice);
                if (item.UnitOfPurchase != null) arrayFields.Add("UNIT_OF_PURCH1I", item.UnitOfPurchase);
                if (item.ConversionFactor != null) arrayFields.Add("CONV_FACTOR1I", item.ConversionFactor);
                if (purchaseOrder.FreightCode != null) arrayFields.Add("FREIGHT_CODE1I", purchaseOrder.FreightCode);
                if (purchaseOrder.DeliveryLocation != null)
                    arrayFields.Add("DELIV_LOCATION1I", purchaseOrder.DeliveryLocation);
                if (item.ExpediteCode != null) arrayFields.Add("EXPEDITE_CODE1I", item.ExpediteCode);
                if (item.Discount1 != null) arrayFields.Add("DISCOUNT_A1I", item.Discount1);
                if (item.Surcharge1 != null) arrayFields.Add("SURCHARGE_A1I", item.Surcharge1);
                if (item.Discount2 != null) arrayFields.Add("DISCOUNT_B1I", item.Discount2);
                if (item.Surcharge2 != null) arrayFields.Add("SURCHARGE_B1I", item.Surcharge2);
            }

            request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM22CA" &&
                   (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" ||
                    reply.functionKeys.StartsWith("XMIT-WARNING")))
            {
                request.screenFields = arrayFields.ToArray();
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }

            if (!_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply)) return true;
            if (reply != null)
                throw new ArgumentException(reply.message);
            throw new Exception(@"No se ha podido obtener respuesta del servidor");
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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    public class PurchaseOrder
    {
        public string AuthorizedStatus;
        public string Currency;
        public string DeliveryLocation;
        public string FreightCode;

        public List<PurchaseOrderItem> Items;
        public string Medium;
        public string NumberOfItems;
        public string OrderDate;
        public string OrderStatus;
        public string OrderType;
        public string OriginCode;
        public string PurchaseNumber;
        public string PurchaseOfficer;
        public string PurchaseTeam;
        public string SupplierName;
        public string SupplierNumber;
        public string TotalEstimatedValue;
    }

    public class PurchaseOrderItem
    {
        public string ConversionFactor;
        public string Discount1;
        public string Discount2;
        public string DueDate;
        public string ExpediteCode;
        public string GrossPrice;
        public string Index;
        public string Quantity;
        public string Surcharge1;
        public string Surcharge2;
        public string UnitOfPurchase;
    }

    public static class PurchaseOrderActions
    {
        public static class OrderStatus
        {
            public static string UnprintedCode = "0";
            public static string Unprinted = "UNPRINTED";
            public static string PrintedCode = "1";
            public static string Printed = "PRINTED";
            public static string ModifiedCode = "2";
            public static string Modified = "MODIFIED";
            public static string CancelledCode = "3";
            public static string Cancelled = "CANCELLED";
            public static string CompletedCode = "9";
            public static string Completed = "COMPLETED";
            public static string UncompletedCode = "U";
            public static string Uncompleted = "UNCOMPLETED";

            public static Dictionary<string, string> GetStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {UnprintedCode, Unprinted},
                    {PrintedCode, Printed},
                    {ModifiedCode, Modified},
                    {CancelledCode, Cancelled},
                    {CompletedCode, Completed},
                    {UncompletedCode, Uncompleted}
                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsValue(statusName)
                    ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key
                    : null;
            }

            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
    }

    public static class Queries
    {
        public static string GetFetchRequisitionStockCodeQuery(string dbReference, string dbLink, string districtCode,
            string stockCode, string scStatus, string startDate, string finishDate, string reqType, string transType,
            string priorityCode)
        {
            if (!string.IsNullOrWhiteSpace(districtCode))
                districtCode = " AND SC.DSTRCT_CODE = '" + districtCode + "'";
            if (!string.IsNullOrWhiteSpace(scStatus))
            {
                if (scStatus.Equals("UNCOMPLETED"))
                    scStatus = " AND SC.ITEM_141_STAT <> '" + Requisition.ItemStatus.CompleteCode + "'";
                else
                    scStatus = " AND SC.ITEM_141_STAT = '" + scStatus + "'";
            }

            if (!string.IsNullOrWhiteSpace(startDate))
                startDate = " AND RQ.CREATION_DATE >= " + startDate;
            if (!string.IsNullOrWhiteSpace(finishDate))
                finishDate = " AND RQ.CREATION_DATE <= " + finishDate;
            if (!string.IsNullOrWhiteSpace(reqType))
                reqType = " AND RQ.IREQ_TYPE = '" + reqType + "'";
            if (!string.IsNullOrWhiteSpace(transType))
                transType = " AND RQ.ISS_TRAN_TYPE = '" + transType + "'";
            if (!string.IsNullOrWhiteSpace(priorityCode))
                priorityCode = " AND RQ.PRIORITY_CODE = '" + priorityCode + "'";

            var sqlQuery = "" +
                           " SELECT " +
                           "   SC.DSTRCT_CODE, SC.IREQ_NO, RQ.IREQ_TYPE, RQ.ISS_TRAN_TYPE, SC.STOCK_CODE, SC.IREQ_ITEM," +
                           "   RQ.AUTHSD_STATUS, RQ.HDR_140_STATUS, SC.ITEM_141_STAT," +
                           "   RQ.PRIORITY_CODE, SC.WHOUSE_ID, RQ.REQUESTED_BY, RQ.CREATION_DATE, RQ.REQ_BY_DATE, RQ.DELIV_INSTR_A, RQ.DELIV_INSTR_B," +
                           "   SC.QTY_REQ, SC.PO_ITEM_NO" +
                           " FROM ELLIPSE.MSF141 SC" +
                           " JOIN ELLIPSE.MSF140 RQ" +
                           " ON SC.IREQ_NO         = RQ.IREQ_NO" +
                           " WHERE STOCK_CODE      = '" + stockCode + "'" +
                           districtCode +
                           scStatus +
                           startDate +
                           finishDate +
                           reqType +
                           transType +
                           priorityCode;
            return sqlQuery;
        }

        public static string GetFetchPurchaseOrderQuery(string dbReference, string dbLink, string districtCode,
            string purchaseOrder, string stockCode, string startDate, string finishDate, string poStatus)
        {
            if (!string.IsNullOrWhiteSpace(districtCode)
            ) //muchos stockcodes no tienen registrado distrito en los parte número
                districtCode =
                    " AND PO.DSTRCT_CODE = '" + districtCode + "'"; // + " AND PN.DSTRCT_CODE = '" + districtCode + "'";
            if (!string.IsNullOrWhiteSpace(startDate))
                startDate = " AND PO.CREATION_DATE >= " + startDate;
            if (!string.IsNullOrWhiteSpace(finishDate))
                finishDate = " AND PO.CREATION_DATE <= " + finishDate;
            if (!string.IsNullOrWhiteSpace(purchaseOrder))
                purchaseOrder = " PO.PO_NO = '" + purchaseOrder + "'";
            if (!string.IsNullOrWhiteSpace(stockCode))
            {
                stockCode = " POI.PREQ_STK_CODE = '" + stockCode + "'";
                if (!string.IsNullOrWhiteSpace(purchaseOrder))
                    stockCode = " AND " + stockCode;
            }

            if (!string.IsNullOrWhiteSpace(poStatus))
            {
                if (poStatus.Equals("U"))
                    poStatus = " AND PO.STATUS_220 NOT IN ('3', '9')";
                else
                    poStatus = " AND PO.STATUS_220 = '" + poStatus + "'";
            }


            var sqlQuery = "" +
                           " WITH POITEMS AS(" +
                           "  SELECT" +
                           "    POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE, SC.ITEM_NAME, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, PN.PART_NO, PN.MNEMONIC, " +
                           "    POI.GROSS_PRICE_P, POI.UNIT_OF_PURCH, POI.CONV_FACTOR, " +
                           "    PN.PREF_PART_IND, MIN(PN.PREF_PART_IND) OVER (PARTITION BY POI.PREQ_STK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE ORDER BY POI.PREQ_STK_CODE, PN.PREF_PART_IND ASC) ROWPPI, " +
                           "    PO.STATUS_220, PO.CREATION_DATE, PO.ORDER_DATE, POI.ORIG_DUE_DATE, POI.ORIG_NET_PR_I, POI.CURR_NET_PR_I, POI.ORIG_QTY_I, POI.CURR_QTY_I, POI.QTY_RCV_OFST_I, POI.OFST_RCPT_DATE, POI.QTY_RCV_DIR_I, POI.ONST_RCPT_DATE, PO.FREIGHT_CODE, PO.DELIV_LOCATION, POI.EXPEDITE_CODE, PO.SUPPLIER_NO, SUP.SUPPLIER_NAME, PO.PO_MEDIUM_IND, PO.ORIGIN_CODE, PO.PURCH_OFFICER, PO.TEAM_ID" +
                           "  FROM" +
                           "    ELLIPSE.MSF220 PO JOIN ELLIPSE.MSF221 POI ON PO.PO_NO = POI.PO_NO LEFT JOIN ELLIPSE.MSF100 SC ON POI.PREQ_STK_CODE = SC.STOCK_CODE LEFT JOIN ELLIPSE.MSF110 PN ON POI.PREQ_STK_CODE = PN.STOCK_CODE LEFT JOIN ELLIPSE.MSF200 SUP ON PO.SUPPLIER_NO = SUP.SUPPLIER_NO" +
                           "  WHERE" +
                           purchaseOrder +
                           stockCode +
                           districtCode +
                           startDate +
                           finishDate +
                           poStatus +
                           "    AND PN.STATUS_CODES = 'V'" +
                           "  ORDER BY POI.PO_ITEM_NO" +
                           "  )" +
                           "  SELECT * FROM POITEMS WHERE POITEMS.PREF_PART_IND = POITEMS.MINPPI AND ROWPPI = 1";

            sqlQuery = sqlQuery.Replace("WHERE AND", "WHERE");
            return sqlQuery;
        }
    }
}
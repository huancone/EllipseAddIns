using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = EllipseCommonsClassLibrary.ScreenService;


namespace EllipseMSO17EExcelAddIn
{
    public partial class RibbonEllipse
    {

        private static readonly EllipseFunctions EFunctions = new EllipseFunctions();
        private const int TittleRow = 7;
        private const int ResultColumn = 8;
        private const int MaxRows = 10000;
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private readonly string _sheetName01 = "MSO17E PRICE ADJUSTEMENT";
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var enviroments = EnviromentConstants.GetEnviromentList();
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

        private void btnCheckStocks_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLoadData_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                Worksheet excelSheet = excelBook.ActiveSheet;

                excelSheet.Name = _sheetName01;

                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Style = _cells.GetStyle(StyleConstants.Normal);
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearFormats();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).ClearComments();
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).Clear();

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "INVENTORY PRICE ADJUSTEMENT";
                _cells.GetRange("B1", "D2").Merge();
                _cells.GetRange("B1", "D2").WrapText = true;

                //STOCK_CODE	Total Available	Current_Inventory Price1 (DOL)	Current_Inventory Price2 (PES)	Change_Inventory Price1 (DOL)	Change_Inventory Price2(PES)


                _cells.GetCell(1, TittleRow).Value = "Stock Code";
                _cells.GetCell(2, TittleRow).Value = "Stock Description";
                _cells.GetCell(3, TittleRow).Value = "Unit of Issue";
                _cells.GetCell(3, TittleRow).Value = "Total Available";
                _cells.GetCell(4, TittleRow).Value = "Current Inventory Price1 (DOL)";
                _cells.GetCell(5, TittleRow).Value = "Current Inventory Price1 (PES)";
                _cells.GetCell(6, TittleRow).Value = "Change_Inventory Price1 (DOL)";
                _cells.GetCell(7, TittleRow).Value = "Change_Inventory Price2(PES)";
                _cells.GetCell(ResultColumn, TittleRow).Value = "Resultado";

                #region Styles

                _cells.GetRange(1, TittleRow, ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ResultColumn, TittleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetRange(1, TittleRow + 1, ResultColumn, MaxRows).NumberFormat = "@";

                #endregion

                _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow+1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                _excelSheetItems.Name = "MSO17EData";

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();

                _cells.GetRange(1, TittleRow + 1, 2, MaxRows).NumberFormat = "@";
                _cells.GetRange(6, TittleRow + 1, 6, MaxRows).NumberFormat = "#.##0,00 $";
            }
            catch (Exception error)
            {
                _cells.GetCell(ResultColumn, TittleRow).Value += " Error " + error.Message;
            }
        }

        class StockCode
        {
            public StockCode(string stockCode, string districtCode)
            {
                try
                {
                    if (string.IsNullOrEmpty(districtCode))
                    {
                        Error = "Transaccion Invalida";
                        return;
                    }
                    var sqlQuery = Queries.GetStockInfo(stockCode, districtCode, EFunctions.dbReference, EFunctions.dbLink);

                    var drStockInfo = EFunctions.GetQueryResult(sqlQuery);

                    if (drStockInfo != null && !drStockInfo.IsClosed && drStockInfo.HasRows)
                    {
                        while (drStockInfo.Read())
                        {
                            stockCode = drStockInfo["STOCK_CODE"].ToString();
                            stockDesc = drStockInfo["STK_DESC"].ToString();
                            unitOfIssue = drStockInfo["PROJECT_NO"].ToString();
                            stockOnHand = drStockInfo["IND"].ToString();
                            currentInventoryPrice1 = drStockInfo["TRAN_AMOUNT"].ToString();
                            currentInventoryPrice2 = drStockInfo["TRAN_AMOUNT_S"].ToString();
                            error = drStockInfo["TRAN_AMOUNT_S"].ToString();

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

            public string stockCode                 { get; set; }
            public string stockDesc                 { get; set; }
            public string unitOfIssue               { get; set; }
            public string stockOnHand               { get; set; }
            public string currentInventoryPrice1 { get; set; }
            public string currentInventoryPrice2 { get; set; }
            public string error                     { get; set; }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    

    public static class Queries
    {
        public static string GetStockInfo(string stockCode, string districtCode,  string dbReference, string dbLink)
        {
            var query = " " +
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
            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
    }
}

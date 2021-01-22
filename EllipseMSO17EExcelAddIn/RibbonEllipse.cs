using System;
using SharedClassLibrary;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse.Connections;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService;


namespace EllipseMSO17EExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const int TittleRow = 7;
        private const int ResultColumn = 8;
        private const int MaxRows = 10000;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private const string SheetName01 = "MSO17E PRICE ADJUSTEMENT";
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private ListObject _excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            var settings = new Settings();
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

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
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

                excelSheet.Name = SheetName01;

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

        //class StockCode
        //{
        //    public StockCode(string stockCode, string districtCode)
        //    {
        //        try
        //        {
        //            if (string.IsNullOrEmpty(districtCode))
        //            {
        //                Error = "Transaccion Invalida";
        //                return;
        //            }
        //            var sqlQuery = Queries.GetStockInfo(stockCode, districtCode, EFunctions.dbReference, EFunctions.dbLink);

        //            var drStockInfo = EFunctions.GetQueryResult(sqlQuery);

        //            if (drStockInfo != null && !drStockInfo.IsClosed && drStockInfo.HasRows)
        //            {
        //                while (drStockInfo.Read())
        //                {
        //                    stockCode = drStockInfo["STOCK_CODE"].ToString();
        //                    stockDesc = drStockInfo["STK_DESC"].ToString();
        //                    unitOfIssue = drStockInfo["PROJECT_NO"].ToString();
        //                    stockOnHand = drStockInfo["IND"].ToString();
        //                    currentInventoryPrice1 = drStockInfo["TRAN_AMOUNT"].ToString();
        //                    currentInventoryPrice2 = drStockInfo["TRAN_AMOUNT_S"].ToString();
        //                    error = drStockInfo["TRAN_AMOUNT_S"].ToString();

        //                }
        //            }
        //            else
        //            {
        //                Error = "La Transaccion no Existe";
        //            }
        //        }
        //        catch (Exception error)
        //        {
        //            Error = error.Message;
        //        }
        //    }

        //    public string stockCode                 { get; set; }
        //    public string stockDesc                 { get; set; }
        //    public string unitOfIssue               { get; set; }
        //    public string stockOnHand               { get; set; }
        //    public string currentInventoryPrice1 { get; set; }
        //    public string currentInventoryPrice2 { get; set; }
        //    public string error                     { get; set; }
        //}

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    

    //internal static class Queries
    //{
    //    public static string GetStockInfo(string stockCode, string districtCode,  string dbReference, string dbLink)
    //    {
    //        var query = " " +
    //                       "SELECT DISTINCT " +
    //                       "  EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE " +
    //                       "FROM " +
    //                       "  " + dbReference + ".MSF870" + dbLink + " POS " +
    //                       "INNER JOIN " + dbReference + ".MSF878" + dbLink + " EMPOS " +
    //                       "ON" +
    //                       "  EMPOS.POSITION_ID = POS.POSITION_ID " +
    //                       "AND " +
    //                       "  (" +
    //                       "    EMPOS.POS_STOP_DATE > TO_CHAR ( SYSDATE, 'YYYYMMDD' ) " +
    //                       "  OR EMPOS.POS_STOP_DATE = '00000000' " +
    //                       "  ) " +
    //                       "INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP " +
    //                       "ON " +
    //                       "  EMPOS.EMPLOYEE_ID = EMP.EMPLOYEE_ID " +
    //                       "WHERE " +
    //                       "EMPOS.EMPLOYEE_ID = '" + employeeId + "' ";
    //        query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
    //        return query;
    //    }
    //}
}

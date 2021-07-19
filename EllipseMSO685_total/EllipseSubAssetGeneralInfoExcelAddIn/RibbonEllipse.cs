using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Web.Services.Ellipse;
using EllScreen = EllipseScreenLibrary;
using System.Drawing;
using EllipseEnviroment;

namespace EllipseSubAssetGeneralInfoExcelAddIn
{
    public partial class RibbonEllipse
    {
        public static String elliseUser = "";
        public static String ellisePswd = "";
        public static String ellisePost = "";
        public static String elliseDsct = "";

        excel.Application excelApp;
        excel.Workbook excelBook;
        excel.Worksheet excelSheet, excelSheet2, excelSheet3, excelSheet4, excelSheet5, excelSheet6;
        excel.ListObject excelSheetItems, excelSheetItems2, excelSheetItems3, excelSheetItems4, excelSheetItems5, excelSheetItems6;

        //SHEETVC
        public static String sheetNameOP1vc = "MSO685_Opc1_VC";
        public static String sheetNameOP1vl = "MSO685_Opc1_VL";
        public static String sheetNameOP3vc = "MSO685_Op3_VC";
        public static String sheetNameOP3vl = "MSO685_Op3_VL";
        public static String sheetNameOP4vc = "MSO685_Op4_VC";
        public static String sheetNameOP4vl = "MSO685_Op4_VL";
        //COLUMNS
        public static String beginColumnOP1vc = "A";
        public static String endColumnOP1vc = "G";

        // vlarga
        public static String beginColumn2 = "A";
        public static String endColumn2 = "Z";
        //ROWS

        public static String beginColumn3 = "A";
        public static String endColumn3 = "D";


        public static String beginColumn4 = "A";
        public static String endColumn4 = "S";


        //mso685 op4vc
        public static String beginColumn5 = "A";
        public static String endColumn5 = "G";


        //mso685op4vl
        public static String beginColumn6 = "A";
        public static String endColumn6 = "Q";

        public static int headerRowOP1vc = 4;
        public static int dataRowOP1vc = 5;

        public static int headerRow2 = 4;
        public static int dataRow2 = 5;


        //op3
        public static int headerRow3 = 2;
        public static int dataRow3 = 3;

        //op3larga
        public static int headerRow4 = 4;
        public static int dataRow4 = 5;

        //mso685 op4 vc
        public static int headerRow5 = 2;
        public static int dataRow5 = 3;

        public static int headerRow6 = 4;
        public static int dataRow6 = 5;







        //MESSAGE CELL1
        public static string messageProcess = "Processing...";
        public static string messageUploaded = "Uploaded";
        public static string messageModified = "Modified";
        public static string messageStatusloaded = "Status loaded";
        public static string messageDataRequired = "Data required";
        //MESSAGE BOX1
        public static string messageTitle = "Message";
        public static string messageTitleError = "Error";
        public static string messageRequiredFields = "\n\rYou should fill the table with valid information for processing";
        public static string messageProcessFinished = "\n\rProcess finished";
        public static string messageSelectOption = "\n\rPlease Select a Env. Option";



        ///mesagge3
        ///
        public static string errorAuthUser3 = "USER PROFILE NOT FOUND";
        //MESSAGE CELL
        public static string messageProcess3 = "Processing...";
        public static string messageUploaded3 = "Uploaded";
        //MESSAGE BOX
        public static string messageTitle3 = "Message";
        public static string messageTitleError3 = "Error";
        public static string messageRequiredFields3 = "\n\rYou should fill the table with valid information for processing";
        public static string messageProcessFinished3 = "\n\rProcess finished";
        public static string messageUserPassincorrect3 = "\n\rThe user does not exist or the password is incorrect";
        public static string messageSelectOption3 = "\n\rPlease Select a Env. Option";
        public static string messageWebServiceOK3 = "Asset No assigned:";

        

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            drpSubAssetGeneralInfoEnv.SelectedItemIndex = System.Web.Services.Ellipse.Util.GetEnvironment(global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EnvDefault);
        }

        private void btnSubAssetGeneralInfoFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Cargar formato MSO685 Opcion 1
                loadFormatMSO();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnSubAssetGeneralInfoExecute_Click(object sender, RibbonControlEventArgs e)
        {
         
        }

        private void loadFormatMSO()
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                excelBook = excelApp.Workbooks.Add();




                ///6
                ///mo 685op4vl


                excelSheet6 = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet6.Name = sheetNameOP4vl;

                excel.Range RangeMaintItem6 = excelSheet6.get_Range(beginColumn6 + "1:" + endColumn6 + "1");
                RangeMaintItem6.Font.Bold = true;
                RangeMaintItem6.Merge();
                RangeMaintItem6.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeMaintItem6.Font.Color = Color.White;
                RangeMaintItem6.Value = "MSO685 Opcion 4 VL Maintain Sub-Asset Valuation Details - ELLIPSE Loader";
                RangeMaintItem6.WrapText = true;
                RangeMaintItem6 = excelSheet6.get_Range(beginColumn6 + "1:" + endColumn6 + "1");
                RangeMaintItem6.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem6.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem6.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem6.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem6.Borders.Color = Color.Black;

                excel.Range rangeItemTitle6 = excelSheet6.get_Range("A2");
                rangeItemTitle6.Font.Color = Color.Green;
                rangeItemTitle6.Value = "## - Borrar";
                rangeItemTitle6 = excelSheet6.get_Range("B2");
                rangeItemTitle6.Font.Color = Color.Green;
                rangeItemTitle6.Value = "Vacio - No se modifica";

                //Encabezado MSM685A
                excel.Range RangeItem6 = excelSheet6.get_Range("A" + headerRow6);
                RangeItem6.Value = "Asset Reference *";
                RangeItem6 = excelSheet6.get_Range("B" + headerRow6);
                RangeItem6.Value = "Sub Asset Number *";
                RangeItem6 = excelSheet6.get_Range("C" + headerRow6);
                RangeItem6.Value = "Book Type *";

                RangeItem6 = excelSheet6.get_Range("D" + headerRow6);
                RangeItem6.Value = "Function";
                RangeItem6 = excelSheet6.get_Range("E" + headerRow6);
                RangeItem6.Value = "Adjustment Date (YYYYMMDD)";
                RangeItem6 = excelSheet6.get_Range("F" + headerRow6);
                RangeItem6.Value = "Sub Asset Diary";
                RangeItem6 = excelSheet6.get_Range("G" + headerRow6);
                RangeItem6.Value = "District";
                RangeItem6 = excelSheet6.get_Range("H" + headerRow6);
                RangeItem6.Value = "Offset Account";
                //Apply Adjustment
                RangeItem6 = excelSheet6.get_Range("I" + headerRow6);
                RangeItem6.Value = "Y/N 1";
                RangeItem6 = excelSheet6.get_Range("J" + headerRow6);
                RangeItem6.Value = "Y/N 2";
                RangeItem6 = excelSheet6.get_Range("K" + headerRow6);
                RangeItem6.Value = "Y/N 3";
                //Apply Adjustment
                RangeItem6 = excelSheet6.get_Range("I3:K3");
                RangeItem6.Merge();
                RangeItem6.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem6.Font.Color = Color.White;
                RangeItem6.Font.Bold = true;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders.Color = Color.Black;
                RangeItem6.Value = "Apply Adjustment";

                //Function: CAP
                //Current Values
                RangeItem6 = excelSheet6.get_Range("L" + headerRow6);
                RangeItem6.Value = "Capital";
                RangeItem6 = excelSheet6.get_Range("M" + headerRow6);
                RangeItem6.Value = "Accum Dep (Prior Year)";
                RangeItem6 = excelSheet6.get_Range("N" + headerRow6);
                RangeItem6.Value = "Accum Dep (This Year)";
                RangeItem6 = excelSheet6.get_Range("O" + headerRow6);
                RangeItem6.Value = "Adjusted WDV";
                RangeItem6 = excelSheet6.get_Range("P" + headerRow6);
                RangeItem6.Value = "Action";
                //Current Values
                RangeItem6 = excelSheet6.get_Range("L3:P3");
                RangeItem6.Merge();
                RangeItem6.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem6.Font.Color = Color.White;
                RangeItem6.Font.Bold = true;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders.Color = Color.Black;
                RangeItem6.Value = "Current Values";

                RangeItem6 = excelSheet6.get_Range("Q" + headerRow6);
                RangeItem6.Value = "Message";

                //Se aplica formato de texto al subactivo
                excelSheet6.get_Range(beginColumn6 + dataRow6 + ":" + endColumn6 + "100000").NumberFormat = "@";

                RangeItem6 = excelSheet6.get_Range("T1");
                RangeItem6.Value = "Function";
                RangeItem6 = excelSheet6.get_Range("T2");
                RangeItem6.Value = "CAP";
                RangeItem6 = excelSheet6.get_Range("U2");
                RangeItem6.Value = "Capitalization";
                RangeItem6 = excelSheet6.get_Range("T3");
                RangeItem6.Value = "DEP";
                RangeItem6 = excelSheet6.get_Range("U3");
                RangeItem6.Value = "Depreciation Adjustment";
                RangeItem6 = excelSheet6.get_Range("T4");
                RangeItem6.Value = "IMP";
                RangeItem6 = excelSheet6.get_Range("U4");
                RangeItem6.Value = "Impairment";
                RangeItem6 = excelSheet6.get_Range("T5");
                RangeItem6.Value = "IMR";
                RangeItem6 = excelSheet6.get_Range("U5");
                RangeItem6.Value = "Impairment to Revalued Asset";
                RangeItem6 = excelSheet6.get_Range("T6");
                RangeItem6.Value = "JNL";
                RangeItem6 = excelSheet6.get_Range("U6");
                RangeItem6.Value = "Journal";
                RangeItem6 = excelSheet6.get_Range("T7");
                RangeItem6.Value = "RMP";
                RangeItem6 = excelSheet6.get_Range("U7");
                RangeItem6.Value = "Reverse Impairment";
                RangeItem6 = excelSheet6.get_Range("T8");
                RangeItem6.Value = "RVL";
                RangeItem6 = excelSheet6.get_Range("U8");
                RangeItem6.Value = "Revaluation";
                RangeItem6 = excelSheet6.get_Range("T9");
                RangeItem6.Value = "SXP";
                RangeItem6 = excelSheet6.get_Range("U9");
                RangeItem6.Value = "Sub Expenditure";

                RangeItem6 = excelSheet6.get_Range("T1:U1");
                RangeItem6.Merge();
                RangeItem6.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem6.Font.Color = Color.White;
                RangeItem6.Font.Bold = true;

                RangeItem6 = excelSheet6.get_Range("T1:U9");
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem6.Borders.Color = Color.Black;

                RangeItem6 = excelSheet6.get_Range("D" + dataRow6 + ":D100000");
                RangeItem6.Validation.Delete();
                RangeItem6.Validation.Add(XlDVType.xlValidateList,
                                    XlDVAlertStyle.xlValidAlertStop,
                                    XlFormatConditionOperator.xlBetween,
                                    "=$T$2:$T$9",
                                    Type.Missing);
                RangeItem6.Validation.IgnoreBlank = true;
                RangeItem6.Validation.InCellDropdown = true;

                excelSheetItems6 = excelSheet6.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet6.get_Range(beginColumn6 + headerRow6 + ":" + endColumn6 + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

                ///f6
                ///

                //oipcion 4 version corta
       
                excelSheet5 = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet5.Name = sheetNameOP4vc;

                excel.Range RangeMaintItem5 = excelSheet5.get_Range(beginColumn5 + "1:" + endColumn5+ "1");
                RangeMaintItem5.Font.Bold = true;
                RangeMaintItem5.Merge();
                RangeMaintItem5.Value = "MSO685 Sub-Asset Valuation Details";
                RangeMaintItem5.WrapText = true;
                RangeMaintItem5.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                excel.Range RangeItemTitle5 = excelSheet5.get_Range("A" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Asset Reference";
                RangeItemTitle5 = excelSheet5.get_Range("A:A");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("B" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Sub Asset Number";
                RangeItemTitle5 = excelSheet5.get_Range("B:B");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("C" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Book Type";
                RangeItemTitle5 = excelSheet5.get_Range("C:C");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("D" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Function";
                RangeItemTitle5 = excelSheet5.get_Range("D:D");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("E" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Adjustment Date (AAAAMMDD)";
                RangeItemTitle5 = excelSheet5.get_Range("E:E");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("F" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Capital";
                RangeItemTitle5 = excelSheet5.get_Range("F:F");
                RangeItemTitle5.NumberFormat = "@";

                RangeItemTitle5 = excelSheet5.get_Range("G" + headerRow5);
                RangeItemTitle5.Font.Bold = true;
                RangeItemTitle5.Value = "Message";
                RangeItemTitle5.ColumnWidth = 80;
                RangeItemTitle5 = excelSheet5.get_Range("G:G");
                RangeItemTitle5.NumberFormat = "@";

                excelSheetItems5 = excelSheet5.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet5.get_Range(beginColumn5 + headerRow5 + ":" + endColumn5 + "40"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);





                //5
                //////MSO 685 Op4 VC
                ///

                //excelSheet5 = (excel.Worksheet)excelBook.Sheets.Add();

                //excelSheet5.Name = sheetNameOP4vc;

                //excel.Range RangeMaintItem5 = excelSheet5.get_Range(beginColumn5 + "1:" + endColumn5 + "1");
                //RangeMaintItem5.Font.Bold = true;
                //RangeMaintItem5.Merge();
                //RangeMaintItem5.Value = "MSO685 Option 3 VC - Sub-Asset Depreciation Details";
                //RangeMaintItem5.WrapText = true;
                //RangeMaintItem5.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                //excel.Range RangeItemTitle5 = excelSheet5.get_Range("A" + headerRow5);
                //RangeItemTitle5.Value = "Asset Reference";
                //RangeItemTitle5 = excelSheet5.get_Range("B" + headerRow5);
                //RangeItemTitle5.Value = "Sub Asset Number";
                //RangeItemTitle5 = excelSheet5.get_Range("C" + headerRow5);
                //RangeItemTitle5.Value = "Book Type";

                //RangeItemTitle5 = excelSheet5.get_Range("D" + headerRow5);
                //RangeItemTitle5.Value = "Message";

                //RangeItemTitle5 = excelSheet5.get_Range(beginColumn5 + dataRow5 + ":" + endColumn5 + "100000");
                //RangeItemTitle5.NumberFormat = "@";

                //excelSheetItems5 = excelSheet5.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet5.get_Range(beginColumn5 + headerRow5 + ":" + endColumn5 + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);
                ////f5


                ///4

                ////MSO685 OP3 V_LARGA


                excelSheet4 = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet4.Name = sheetNameOP3vl;

                excel.Range RangeMaintItem4 = excelSheet4.get_Range(beginColumn4 + "1:" + endColumn4 + "1");
                RangeMaintItem4.Font.Bold = true;
                RangeMaintItem4.Merge();
                RangeMaintItem4.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeMaintItem4.Font.Color = Color.White;
                RangeMaintItem4.Value = "MSO685 Opcion 3 Maintain Sub-Asset Depreciation Details - ELLIPSE Loader";
                RangeMaintItem4.WrapText = true;
                RangeMaintItem4 = excelSheet4.get_Range(beginColumn4 + "1:" + endColumn4 + "1");
                RangeMaintItem4.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem4.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem4.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem4.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem4.Borders.Color = Color.Black;

                excel.Range rangeItemTitle4 = excelSheet4.get_Range("A2");
                rangeItemTitle4.Font.Color = Color.Green;
                rangeItemTitle4.Value = "## - Borrar";
                rangeItemTitle4 = excelSheet4.get_Range("B2");
                rangeItemTitle4.Font.Color = Color.Green;
                rangeItemTitle4.Value = "Vacio - No se modifica";

                //Encabezado MSM685A
                excel.Range RangeItem4 = excelSheet4.get_Range("A" + headerRow4);
                RangeItem4.Value = "Asset Reference *";
                RangeItem4 = excelSheet4.get_Range("B" + headerRow4);
                RangeItem4.Value = "Sub Asset Number *";
                RangeItem4 = excelSheet4.get_Range("C" + headerRow4);
                RangeItem4.Value = "Book Type *";

                //MSM685C
                //Depreciation Details
                RangeItem4 = excelSheet4.get_Range("D3:P3");
                RangeItem4.Merge();
                RangeItem4.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem4.Font.Color = Color.White;
                RangeItem4.Font.Bold = true;
                RangeItem4.Value = "Depreciation Details";
                RangeItem4 = excelSheet4.get_Range("D3:P3");
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders.Color = Color.Black;
                //Depreciation Elements
                RangeItem4 = excelSheet4.get_Range("D" + headerRow4);
                RangeItem4.Value = "Depreciation Method *";
                RangeItem4 = excelSheet4.get_Range("E" + headerRow4);
                RangeItem4.Value = "Depreciation Rate";
                RangeItem4 = excelSheet4.get_Range("F" + headerRow4);
                RangeItem4.Value = "Manual Period Depn";
                RangeItem4 = excelSheet4.get_Range("G" + headerRow4);
                RangeItem4.Value = "Until Period";
                RangeItem4 = excelSheet4.get_Range("H" + headerRow4);
                RangeItem4.Value = "Accelerated Depn Rate";
                RangeItem4 = excelSheet4.get_Range("I" + headerRow4);
                RangeItem4.Value = "Until Period";
                RangeItem4 = excelSheet4.get_Range("J" + headerRow4);
                RangeItem4.Value = "Rate Table";
                RangeItem4 = excelSheet4.get_Range("K" + headerRow4);
                RangeItem4.Value = "Recovery Period";
                RangeItem4 = excelSheet4.get_Range("L" + headerRow4);
                RangeItem4.Value = "Dividend Statistic";
                RangeItem4 = excelSheet4.get_Range("M" + headerRow4);
                RangeItem4.Value = "Divisor Statistic";
                RangeItem4 = excelSheet4.get_Range("N" + headerRow4);
                RangeItem4.Value = "Estimated Life (months)";
                RangeItem4 = excelSheet4.get_Range("O" + headerRow4);
                RangeItem4.Value = "Useful Life Group Code";
                RangeItem4 = excelSheet4.get_Range("P" + headerRow4);
                RangeItem4.Value = "Est Retirement Value - Local";

                //Sub Asset Movement Summary
                RangeItem4 = excelSheet4.get_Range("Q3:R3");
                RangeItem4.Merge();
                RangeItem4.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem4.Font.Color = Color.White;
                RangeItem4.Font.Bold = true;
                RangeItem4.Value = "Sub Asset Movement Summary";
                RangeItem4 = excelSheet4.get_Range("Q3:R3");
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem4.Borders.Color = Color.Black;
                //Capitalization Details
                RangeItem4 = excelSheet4.get_Range("Q" + headerRow4);
                RangeItem4.Value = "Foreign Currency Cost";
                RangeItem4 = excelSheet4.get_Range("R" + headerRow4);
                RangeItem4.Value = "Foreign Currency Type";

                RangeItem4 = excelSheet4.get_Range("S" + headerRow4);
                RangeItem4.Value = "Message";

                //Se aplica formato de texto al subactivo
                excelSheet4.get_Range(beginColumn4 + dataRow4 + ":" + endColumn4 + "100000").NumberFormat = "@";

                excelSheetItems4 = excelSheet4.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet4.get_Range(beginColumn4 + headerRow4 + ":" + endColumn4 + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

                ///F4


                ////3

                excelSheet3 = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet3.Name = sheetNameOP3vc;

                excel.Range RangeMaintItem3 = excelSheet3.get_Range(beginColumn3 + "1:" + endColumn3 + "1");
                RangeMaintItem3.Font.Bold = true;
                RangeMaintItem3.Merge();
                RangeMaintItem3.Value = "MSO685 Option 3 VC - Sub-Asset Depreciation Details";
                RangeMaintItem3.WrapText = true;
                RangeMaintItem3.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                excel.Range RangeItemTitle3 = excelSheet3.get_Range("A" + headerRow3);
                RangeItemTitle3.Value = "Asset Reference";
                RangeItemTitle3 = excelSheet3.get_Range("B" + headerRow3);
                RangeItemTitle3.Value = "Sub Asset Number";
                RangeItemTitle3 = excelSheet3.get_Range("C" + headerRow3);
                RangeItemTitle3.Value = "Book Type";

                RangeItemTitle3 = excelSheet3.get_Range("D" + headerRow3);
                RangeItemTitle3.Value = "Message";

                RangeItemTitle3 = excelSheet3.get_Range(beginColumn3 + dataRow3 + ":" + endColumn3 + "100000");
                RangeItemTitle3.NumberFormat = "@";

                excelSheetItems3 = excelSheet3.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet3.get_Range(beginColumn3 + headerRow3 + ":" + endColumn3 + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

                ///F3


                ///2

                excelSheet2 = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet2.Name = sheetNameOP1vl;

                excel.Range RangeMaintItem2 = excelSheet2.get_Range(beginColumn2 + "1:" + endColumn2 + "1");
                RangeMaintItem2.Font.Bold = true;
                RangeMaintItem2.Merge();
                RangeMaintItem2.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeMaintItem2.Font.Color = Color.White;
                RangeMaintItem2.Value = "MSO685 Opcion 1 VL Modify Maintain General Information - Ellipse 9 Loader";
                RangeMaintItem2.WrapText = true;
                RangeMaintItem2 = excelSheet2.get_Range(beginColumn2 + "1:" + endColumn2 + "1");
                RangeMaintItem2.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem2.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem2.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem2.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItem2.Borders.Color = Color.Black;

                //Encabezado
                excel.Range RangeItem2 = excelSheet2.get_Range("A" + headerRow2);
                RangeItem2.Value = "Asset Reference  *";
                RangeItem2 = excelSheet2.get_Range("B" + headerRow2);
                RangeItem2.Value = "Sub Asset Number *";

                //Encabezado MSM685B
                RangeItem2 = excelSheet2.get_Range("C" + headerRow2);
                RangeItem2.Value = "Serial No";
                RangeItem2 = excelSheet2.get_Range("D" + headerRow2);
                RangeItem2.Value = "Sub-Asset Description";

                excel.Range rangeItemTitle2 = excelSheet2.get_Range("A2");
                rangeItemTitle2.Font.Color = Color.Green;
                rangeItemTitle2.Value = "## - Borrar";
                rangeItemTitle2 = excelSheet2.get_Range("B2");
                rangeItemTitle2.Font.Color = Color.Green;
                rangeItemTitle2.Value = "Vacio - No se modifica";

                //Asset Details
                RangeItem2 = excelSheet2.get_Range("E3:O3");
                RangeItem2.Merge();
                RangeItem2.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem2.Font.Color = Color.White;
                RangeItem2.Font.Bold = true;
                RangeItem2.Value = "Asset Details";
                RangeItem2 = excelSheet2.get_Range("E3:O3");
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders.Color = Color.Black;

                RangeItem2 = excelSheet2.get_Range("E" + headerRow2);
                RangeItem2.Value = "Asset Class *";
                RangeItem2 = excelSheet2.get_Range("F" + headerRow2);
                RangeItem2.Value = "Depreciate  *";
                RangeItem2 = excelSheet2.get_Range("G" + headerRow2);
                RangeItem2.Value = "Capitalization Type";
                RangeItem2 = excelSheet2.get_Range("H" + headerRow2);
                RangeItem2.Value = "Revaluation SubClass";
                RangeItem2 = excelSheet2.get_Range("I" + headerRow2);
                RangeItem2.Value = "Cash Generating Unit";
                RangeItem2 = excelSheet2.get_Range("J" + headerRow2);
                RangeItem2.Value = "Reporting Code";
                //Update Depreciation Details
                RangeItem2 = excelSheet2.get_Range("K" + headerRow2);
                RangeItem2.Value = "BK";
                RangeItem2 = excelSheet2.get_Range("L" + headerRow2);
                RangeItem2.Value = "T1";
                RangeItem2 = excelSheet2.get_Range("M" + headerRow2);
                RangeItem2.Value = "T2";
                RangeItem2 = excelSheet2.get_Range("N" + headerRow2);
                RangeItem2.Value = "T3";
                RangeItem2 = excelSheet2.get_Range("O" + headerRow2);
                RangeItem2.Value = "T4";

                //Asset Classification
                RangeItem2 = excelSheet2.get_Range("P3:W3");
                RangeItem2.Merge();
                RangeItem2.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem2.Font.Color = Color.White;
                RangeItem2.Font.Bold = true;
                RangeItem2.Value = "Asset Classification";
                RangeItem2 = excelSheet2.get_Range("P3:W3");
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders.Color = Color.Black;

                RangeItem2 = excelSheet2.get_Range("P" + headerRow2);
                RangeItem2.Value = "Plant Assets";
                RangeItem2 = excelSheet2.get_Range("Q" + headerRow2);
                RangeItem2.Value = "Early Access Clasi";
                RangeItem2 = excelSheet2.get_Range("R" + headerRow2);
                RangeItem2.Value = "Fixed plant";
                RangeItem2 = excelSheet2.get_Range("S" + headerRow2);
                RangeItem2.Value = "Mobile equipment";
                RangeItem2 = excelSheet2.get_Range("T" + headerRow2);
                RangeItem2.Value = "Distribucion de Co";
                RangeItem2 = excelSheet2.get_Range("U" + headerRow2);
                RangeItem2.Value = "Computer equipment";
                RangeItem2 = excelSheet2.get_Range("V" + headerRow2);
                RangeItem2.Value = "Asset Classificati";
                RangeItem2 = excelSheet2.get_Range("W" + headerRow2);
                RangeItem2.Value = "Asset Classificati";

                //Account Profile
                RangeItem2 = excelSheet2.get_Range("X3:Y3");
                RangeItem2.Merge();
                RangeItem2.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeItem2.Font.Color = Color.White;
                RangeItem2.Font.Bold = true;
                RangeItem2.Value = "Account Profile";
                RangeItem2 = excelSheet2.get_Range("X3:Y3");
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeItem2.Borders.Color = Color.Black;


                RangeItem2 = excelSheet2.get_Range("X" + headerRow2);
                RangeItem2.Value = "Balance Sheet Profile";
                RangeItem2 = excelSheet2.get_Range("Y" + headerRow2);
                RangeItem2.Value = "Profit & Loss Profile";

                RangeItem2 = excelSheet2.get_Range("Z" + headerRow2);
                RangeItem2.Value = "Message";

                //Se aplica formato de texto al subactivo
                excelSheet2.get_Range(beginColumn2 + dataRow2 + ":" + endColumn2 + "100000").NumberFormat = "@";

                excelSheetItems2 = excelSheet2.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet2.get_Range(beginColumn2 + headerRow2 + ":" + endColumn2 + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

                ///F2




                //formatear op 685 opvc
                excelSheet = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet.Name = "MSO685_Opc1_VC";

                excel.Range RangeMaintItemOP1vc = excelSheet.get_Range(beginColumnOP1vc + "1:" + endColumnOP1vc + "1");
                RangeMaintItemOP1vc.Font.Bold = true;
                RangeMaintItemOP1vc.Merge();
                RangeMaintItemOP1vc.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeMaintItemOP1vc.Font.Color = Color.White;
                RangeMaintItemOP1vc.Value = "MSO685 Opcion 1 VC Maintain General Information - Ellipse 9 LoaderX";
                RangeMaintItemOP1vc.WrapText = true;
                RangeMaintItemOP1vc = excelSheet.get_Range(beginColumnOP1vc + "1:" + endColumnOP1vc + "1");
                RangeMaintItemOP1vc.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItemOP1vc.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItemOP1vc.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItemOP1vc.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                RangeMaintItemOP1vc.Borders.Color = Color.Black;

                //Encabezado
                excel.Range RangeItemOP1vc = excelSheet.get_Range("A" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Asset Reference  *";

                RangeItemOP1vc = excelSheet.get_Range("B" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Sub Asset Number *";

                //MODIFICADO
                RangeItemOP1vc = excelSheet.get_Range("C" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Depreciate  *";

                RangeItemOP1vc = excelSheet.get_Range("D" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Capitalization Type";


                RangeItemOP1vc = excelSheet.get_Range("E" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Balance Sheet Profile";

                RangeItemOP1vc = excelSheet.get_Range("F" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Profit & Loss Profile";


                RangeItemOP1vc = excelSheet.get_Range("G" + headerRowOP1vc);
                RangeItemOP1vc.Value = "Message";





                //Se aplica formato de texto al subactivo
                excelSheet.get_Range(beginColumnOP1vc + dataRowOP1vc + ":" + endColumnOP1vc + "100000").NumberFormat = "@";

                excelSheetItems = excelSheet.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet.get_Range(beginColumnOP1vc + headerRowOP1vc + ":" + endColumnOP1vc + "100000"), XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

              


               



              

            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }


        /// <summary>
        /// 


        ///

        private void executeMSO685op4vl()
        {
            EllScreen.Ellipse screen = null;
            try
            {
                if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        //elliseUser = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
                        elliseUser = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
                        ellisePost = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.PosDefault;
                        elliseDsct = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.DstrDefault;
                    }

                    FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.ShowDialog();

                    if (frm.Auth.Authenticated)
                    {
                        elliseUser = frm.Auth.Username;
                        ellisePost = frm.Auth.Position;
                        elliseDsct = frm.Auth.District;
                        ellisePswd = frm.Auth.Password;

                        if (excelSheet == null)
                        {
                            excelApp = Globals.ThisAddIn.Application;
                            excelBook = excelApp.Workbooks.Item[1];
                            excelSheet6 = (excel.Worksheet)excelBook.Sheets[sheetNameOP4vl];
                        }

                        string url = "";

                        EllipseEnviroment.EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                        if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = conf.UrlProd;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = conf.UrlCont;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = conf.UrlDesa;
                        }
                        else
                        {
                            url = conf.UrlTest;
                        }

                        string error;
                        int currentRow = dataRow6;

                        String campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range(beginColumn6 + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            excelSheet6.Select();

                            screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                            ClientConversation.authenticate(elliseUser, ellisePswd);
                            screen.InitMSOInstance(url);

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    screen.ExecuteScreen("MSO685", "MSM685A");
                                    if (!screen.GetMSOError().Equals(""))
                                    {
                                        throw new Exception(screen.GetMSOError());
                                    }

                                    excelSheet6.get_Range(endColumn6 + currentRow).Select();
                                    excelSheet6.get_Range(endColumn6 + currentRow).Value = messageProcess;

                                    if (screen.MSO.mapName.Equals("MSM685A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "4");
                                        screen.SetMSOFieldValue("ASSET_REF1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUB_ASSET_NO1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("B" + currentRow).Value)));
                                        screen.SetMSOFieldValue("BOOK_OR_TAX1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("C" + currentRow).Value)));
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOError())
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (screen.MSO.mapName.Equals("MSM68BA"))
                                    {
                                        screen.InitScreenFields();

                                        //Depreciation Details
                                        string deprMethod = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("D" + currentRow).Value));
                                        if (!"".Equals(deprMethod) && !"##".Equals(deprMethod))
                                        {
                                            screen.SetMSOFieldValue("FUNCTION1I", deprMethod);
                                        }
                                        else if ("##".Equals(deprMethod))
                                        {
                                            screen.SetMSOFieldValue("FUNCTION1I", "");
                                        }

                                        string adjustDate = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("E" + currentRow).Value));
                                        if (!"".Equals(adjustDate) && !"##".Equals(adjustDate))
                                        {
                                            //try
                                            //{
                                            //    DateTime dateValue = DateTime.ParseExact(adjustDate, "MM/dd/yyyy", null);
                                            //}
                                            //catch
                                            //{
                                            //    throw new Exception("Formato fecha invalido ' Adjustment Date (MM/DD/YYYY) '");
                                            //}
                                            screen.SetMSOFieldValue("ADJUST_DATE1I", adjustDate);
                                        }
                                        else if ("##".Equals(adjustDate))
                                        {
                                            screen.SetMSOFieldValue("ADJUST_DATE1I", "");
                                        }

                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);


                                        if (screen.IsMSOError())
                                        {
                                            if (screen.GetMSOError().Contains("W1:A287"))
                                            {
                                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOError())
                                                {
                                                    error = screen.GetMSOError();
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                            }
                                            else if(screen.GetMSOError().Contains("W2:A225"))
                                            {
                                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOError())
                                                {
                                                    error = screen.GetMSOError();
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                            }
                                            else
                                            {
                                                error = screen.GetMSOError();
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                    }

                                    if (screen.MSO.mapName.Equals("MSM68BA"))
                                    {
                                        screen.InitScreenFields();

                                        string subAsset = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("F" + currentRow).Value));
                                        if (!"".Equals(subAsset) && !"##".Equals(subAsset))
                                            screen.SetMSOFieldValue("SUB_ASSET_DIARY1I", subAsset);
                                        else if ("##".Equals(subAsset))
                                            screen.SetMSOFieldValue("SUB_ASSET_DIARY1I", "");

                                        string district = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("G" + currentRow).Value));
                                        if (!"".Equals(district) && !"##".Equals(district))
                                            screen.SetMSOFieldValue("DISTRICT1I", district);
                                        else if ("##".Equals(district))
                                            screen.SetMSOFieldValue("DISTRICT1I", "");

                                        string offSet = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("H" + currentRow).Value));
                                        if (!"".Equals(offSet) && !"##".Equals(offSet))
                                            screen.SetMSOFieldValue("OFFSET_ACCT1I", offSet);
                                        else if ("##".Equals(offSet))
                                            screen.SetMSOFieldValue("OFFSET_ACCT1I", "");

                                        string applyAdj1 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("I" + currentRow).Value));
                                        if (!"".Equals(applyAdj1) && !"##".Equals(applyAdj1))
                                            screen.SetMSOFieldValue("APPLY_ADJ_1_1I", applyAdj1);
                                        else if ("##".Equals(applyAdj1))
                                            screen.SetMSOFieldValue("APPLY_ADJ_1_1I", "");

                                        string applyAdj2 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("J" + currentRow).Value));
                                        if (!"".Equals(applyAdj2) && !"##".Equals(applyAdj2))
                                            screen.SetMSOFieldValue("APPLY_ADJ_2_1I", applyAdj2);
                                        else if ("##".Equals(applyAdj2))
                                            screen.SetMSOFieldValue("APPLY_ADJ_2_1I", "");

                                        string applyAdj3 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("K" + currentRow).Value));
                                        if (!"".Equals(applyAdj3) && !"##".Equals(applyAdj3))
                                            screen.SetMSOFieldValue("APPLY_ADJ_3_1I", applyAdj3);
                                        else if ("##".Equals(applyAdj3))
                                            screen.SetMSOFieldValue("APPLY_ADJ_3_1I", "");

                                        //Current Account
                                        string adjCapLocal = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("L" + currentRow).Value));
                                        if (!"".Equals(adjCapLocal) && !"##".Equals(adjCapLocal))
                                            screen.SetMSOFieldValue("ADJ_CAP_LOCAL1I", adjCapLocal);
                                        else if ("##".Equals(adjCapLocal))
                                            screen.SetMSOFieldValue("ADJ_CAP_LOCAL1I", "");

                                        string adjDepPrev = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("M" + currentRow).Value));
                                        if (!"".Equals(adjDepPrev) && !"##".Equals(adjDepPrev))
                                            screen.SetMSOFieldValue("ADJ_DEP_PREV_LOCAL1I", adjDepPrev);
                                        else if ("##".Equals(adjDepPrev))
                                            screen.SetMSOFieldValue("ADJ_DEP_PREV_LOCAL1I", "");

                                        string adjDepThis = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("N" + currentRow).Value));
                                        if (!"".Equals(adjDepThis) && !"##".Equals(adjDepThis))
                                            screen.SetMSOFieldValue("ADJ_DEP_THIS_LOCAL1I", adjDepThis);
                                        else if ("##".Equals(adjDepThis))
                                            screen.SetMSOFieldValue("ADJ_DEP_THIS_LOCAL1I", "");

                                        string adjWdvLocal = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("O" + currentRow).Value));
                                        if (!"".Equals(adjWdvLocal) && !"##".Equals(adjWdvLocal))
                                            screen.SetMSOFieldValue("ADJ_WDV_LOCAL1I", adjWdvLocal);
                                        else if ("##".Equals(adjWdvLocal))
                                            screen.SetMSOFieldValue("ADJ_WDV_LOCAL1I", "");

                                        string action = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range("P" + currentRow).Value));
                                        if (!"".Equals(action) && !"##".Equals(action))
                                            screen.SetMSOFieldValue("ACTION1I", action);
                                        else if ("##".Equals(action))
                                            screen.SetMSOFieldValue("ACTION1I", "");

                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);

                                        if(screen.MSO.functionKeys.Contains("XMIT-Confirm"))
                                        {
                                            screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                            if (screen.IsMSOError())
                                            {
                                                error = screen.GetMSOError();
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                            else
                                            {
                                                excelSheet6.get_Range(endColumn6 + currentRow).Value = messageModified;
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            }
                                        }
                                        else
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());

                                        }

                                        //if (screen.IsMSOError())
                                        //{
                                        //    if (screen.GetMSOError().Contains("confirm"))
                                        //    {
                                        //        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        //        if (screen.IsMSOError())
                                        //        {
                                        //            error = screen.GetMSOError();
                                        //            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        //            throw new Exception(error.Trim());
                                        //        }
                                        //        else
                                        //        {
                                        //            excelSheet6.get_Range(endColumn6 + currentRow).Value = messageModified;
                                        //            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        //        }
                                        //    }
                                        //    else
                                        //    {
                                        //        error = screen.GetMSOError();
                                        //        screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        //        throw new Exception(error.Trim());
                                        //    }
                                        //}
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet6.get_Range(endColumn6 + currentRow).Value = errorEx.Message.Trim();
                                    //WorkAround de Excepción de Programa MSO685 detectada
                                    if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                        currentRow--;
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet6.get_Range(beginColumn6 + currentRow).Value));
                                }
                            }
                            excelSheet6.Cells.Columns.AutoFit();
                            excelSheet6.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            screen.ClosedConnection();
                        }
                    }
                    else
                    {
                        MessageBox.Show(messageSelectOption, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    screen.ClosedConnection();
                    MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception errorCatch)
            {
                screen.ClosedConnection();
                MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        ////

        private void executeMSOmso685op4vc()
        {
            // opcion corta mso685 op4
            EllScreen.Ellipse screen = null;
            if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
            {
                FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                frm.StartPosition = FormStartPosition.CenterScreen;
                frm.ShowDialog();


                elliseUser = frm.Auth.Username;
                ellisePost = frm.Auth.Position;
                elliseDsct = frm.Auth.District;
                ellisePswd = frm.Auth.Password;
                try
                {
                    if (excelSheet5 == null)
                    {
                        excelApp = Globals.ThisAddIn.Application;
                        excelBook = excelApp.Workbooks.Item[1];
                        excelSheet5 = (excel.Worksheet)excelBook.Sheets["MSO685 OP4VC"];
                    }

                    string url = "";

                    EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                    if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                    {
                        url = conf.UrlProd;
                    }
                    else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                    {
                        url = conf.UrlCont;
                    }
                    else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                    {
                        url = conf.UrlDesa;
                    }
                    else
                    {
                        url = conf.UrlTest;
                    }

                    screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                    ClientConversation.authenticate(elliseUser, ellisePswd);
                    screen.InitMSOInstance(url);

                    if (!screen.GetMSOError().Equals(""))
                    {
                        throw new Exception(screen.GetMSOError());
                    }
                    

                    int currentRow = 3;
                    System.Array MyValues = (System.Array)excelSheet5.get_Range("A" + currentRow.ToString(), "G" + currentRow.ToString()).Cells.Value;

                    while (MyValues.GetValue(1, 1) != null)
                    {
                        try
                        {
                            screen.ExecuteScreen("MSO685", "MSM685A");
                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            excelSheet5.get_Range("G" + currentRow.ToString()).Select();
                            excelSheet5.get_Range("G" + currentRow.ToString()).Value = "Processing..";

                            if (screen.MSO.mapName.Equals("MSM685A"))
                            {
                                screen.InitScreenFields();
                                screen.SetMSOFieldValue("OPTION1I", "4");
                                screen.SetMSOFieldValue("ASSET_REF1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("A" + currentRow).Value)));
                                screen.SetMSOFieldValue("SUB_ASSET_NO1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("B" + currentRow).Value)));
                                screen.SetMSOFieldValue("BOOK_OR_TAX1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("C" + currentRow).Value)));
                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);    //Ejecuta
                                if (screen.IsMSOError())
                                {
                                    throw new Exception(screen.GetMSOError());
                                }


                            }

                            if (screen.MSO.mapName.Equals("MSM68BA"))
                            {
                                string date = Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("E" + currentRow).Value));
                                screen.InitScreenFields();
                                screen.SetMSOFieldValue("FUNCTION1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("D" + currentRow).Value)));
                                screen.SetMSOFieldValue("ADJUST_DATE1I", date);

                                //"11/30/2015"

                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);    //Ejecuta
                                //if (screen.IsMSOError())
                                //{
                                //    throw new Exception(screen.GetMSOError());
                                //}


                                if (screen.IsMSOError())
                                {
                                    if (screen.GetMSOError().Contains("W1:A287"))
                                    {
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        if (screen.IsMSOError())
                                        {
                                            
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(screen.GetMSOError());
                                        }
                                    }
                                    else if (screen.GetMSOError().Contains("W2:A225"))
                                    {
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        if (screen.IsMSOError())
                                        {
                                            
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(screen.GetMSOError());
                                        }
                                    }
                                    else
                                    {
                                        
                                       // screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        throw new Exception(screen.GetMSOError());
                                    }
                                }

                                if (screen.IsMSOError())
                                {
                                    screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                }

                                screen.InitScreenFields();
                                screen.SetMSOFieldValue("ADJ_CAP_LOCAL1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet5.get_Range("F" + currentRow).Value)));
                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);    //Ejecuta
                                if (screen.MSO.functionKeys.Contains("XMIT-Confirm"))

                                {
                                    screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                }

                                if (screen.IsMSOError())
                                {
                                    throw new Exception(screen.GetMSOError());
                                }

                                screen.InitScreenFields();
                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                if (screen.IsMSOError())
                                {
                                    throw new Exception(screen.GetMSOError());
                                }
                                else
                                {
                                    excelSheet5.get_Range(endColumn5 + currentRow).Value = "".Equals(screen.GetMSOError()) ? messageUploaded : screen.GetMSOError();
                                }
                            }
                        }
                        catch (Exception errorEx)
                        {
                            excelSheet5.get_Range(endColumn5 + currentRow).Value = errorEx.Message.Trim();
                            //WorkAround de Excepción de Programa MSO685 detectada
                            if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                currentRow--;
                        }
                        currentRow++;
                        MyValues = (System.Array)excelSheet5.get_Range("A" + currentRow.ToString(), "G" + currentRow.ToString()).Cells.Value;
                        screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                    }

                    excelSheet5.Cells.Columns.AutoFit();
                    excelSheet5.Cells.Rows.AutoFit();
                    MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    screen.ClosedConnection();


                }
                catch (Exception errorCatch)
                {
                    screen.ClosedConnection();
                    var messageBox = MessageBox.Show("\n\rMessage:" + errorCatch.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                screen.ClosedConnection();
                MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ////
        }


        /// </summary>

        private void execute_mso685op3vL()
        {

            EllScreen.Ellipse screen = null;
            //

            try
            {
                if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        elliseUser = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
                        ellisePost = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.PosDefault;
                        elliseDsct = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.DstrDefault;
                    }

                    FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.ShowDialog();

                    if (frm.Auth.Authenticated)
                    {
                        elliseUser = frm.Auth.Username;
                        ellisePost = frm.Auth.Position;
                        elliseDsct = frm.Auth.District;
                        ellisePswd = frm.Auth.Password;

                        if (excelSheet == null)
                        {
                            excelApp = Globals.ThisAddIn.Application;
                            excelBook = excelApp.Workbooks.Item[1];
                            excelSheet = (excel.Worksheet)excelBook.Sheets[sheetNameOP3vl];
                        }

                        string url = "";

                        EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                        if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = conf.UrlProd;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                        { 
                        
                            url = conf.UrlCont;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = conf.UrlDesa;
                        }
                        else
                        {
                            url = conf.UrlTest;
                        }

                        string error;
                        int currentRow = dataRow4;

                        String campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range(beginColumn4 + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            excelSheet4.Select();


                            screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                            ClientConversation.authenticate(elliseUser, ellisePswd);
                            screen.InitMSOInstance(url);

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    screen.ExecuteScreen("MSO685", "MSM685A");
                                    if (!screen.GetMSOError().Equals(""))
                                    {
                                        throw new Exception(screen.GetMSOError());
                                    }

                                    excelSheet4.get_Range(endColumn4 + currentRow).Select();
                                    excelSheet4.get_Range(endColumn4 + currentRow).Value = messageProcess;

                                    if (screen.MSO.mapName.Equals("MSM685A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "3");
                                        screen.SetMSOFieldValue("ASSET_REF1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUB_ASSET_NO1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("B" + currentRow).Value)));
                                        screen.SetMSOFieldValue("BOOK_OR_TAX1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("C" + currentRow).Value)));
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOError())
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (screen.MSO.mapName.Equals("MSM685C"))
                                    {
                                        screen.InitScreenFields();
                                        //Depreciation Details
                                        string deprMethod = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("D" + currentRow).Value));
                                        if (!"".Equals(deprMethod) && !"##".Equals(deprMethod))
                                            screen.SetMSOFieldValue("DEPR_METHOD3I", deprMethod);
                                        else if ("##".Equals(deprMethod))
                                            screen.SetMSOFieldValue("DEPR_METHOD3I", "");

                                        string deprRate = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("E" + currentRow).Value));
                                        if (!"".Equals(deprRate) && !"##".Equals(deprRate))
                                            screen.SetMSOFieldValue("DEPR_RATE3I", deprRate);
                                        else if ("##".Equals(deprRate))
                                            screen.SetMSOFieldValue("DEPR_RATE3I", "");

                                        string manPer = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("F" + currentRow).Value));
                                        if (!"".Equals(manPer) && !"##".Equals(manPer))
                                            screen.SetMSOFieldValue("MAN_PER_DEPR3I", manPer);
                                        else if ("##".Equals(manPer))
                                            screen.SetMSOFieldValue("MAN_PER_DEPR3I", "");

                                        string finMan = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("G" + currentRow).Value));
                                        if (!"".Equals(finMan) && !"##".Equals(finMan))
                                            screen.SetMSOFieldValue("FIN_MAN_PER3I", finMan);
                                        else if ("##".Equals(finMan))
                                            screen.SetMSOFieldValue("FIN_MAN_PER3I", "");

                                        string accelDepr = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("H" + currentRow).Value));
                                        if (!"".Equals(accelDepr) && !"##".Equals(accelDepr))
                                            screen.SetMSOFieldValue("ACCEL_DEPR_RT3I", accelDepr);
                                        else if ("##".Equals(accelDepr))
                                            screen.SetMSOFieldValue("ACCEL_DEPR_RT3I", "");

                                        string finAccel = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("I" + currentRow).Value));
                                        if (!"".Equals(finAccel) && !"##".Equals(finAccel))
                                            screen.SetMSOFieldValue("FIN_ACCEL_PER3I", finAccel);
                                        else if ("##".Equals(finAccel))
                                            screen.SetMSOFieldValue("FIN_ACCEL_PER3I", "");

                                        string rateTable = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("J" + currentRow).Value));
                                        if (!"".Equals(rateTable) && !"##".Equals(rateTable))
                                            screen.SetMSOFieldValue("RATE_TABLE3I", rateTable);
                                        else if ("##".Equals(rateTable))
                                            screen.SetMSOFieldValue("RATE_TABLE3I", "");

                                        string recovPeriod = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("K" + currentRow).Value));
                                        if (!"".Equals(rateTable) && !"##".Equals(rateTable))
                                            screen.SetMSOFieldValue("RECOV_PERIOD3I", rateTable);
                                        else if ("##".Equals(rateTable))
                                            screen.SetMSOFieldValue("RECOV_PERIOD3I", "");

                                        string dividendStat = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("L" + currentRow).Value));
                                        if (!"".Equals(dividendStat) && !"##".Equals(dividendStat))
                                            screen.SetMSOFieldValue("DIVIDEND_STAT3I", dividendStat);
                                        else if ("##".Equals(dividendStat))
                                            screen.SetMSOFieldValue("DIVIDEND_STAT3I", "");

                                        string divisorStat = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("M" + currentRow).Value));
                                        if (!"".Equals(divisorStat) && !"##".Equals(divisorStat))
                                            screen.SetMSOFieldValue("DIVISOR_STAT3I", divisorStat);
                                        else if ("##".Equals(divisorStat))
                                            screen.SetMSOFieldValue("DIVISOR_STAT3I", "");

                                        string estMn = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("N" + currentRow).Value));
                                        if (!"".Equals(estMn) && !"##".Equals(estMn))
                                            screen.SetMSOFieldValue("EST_MN_LIFE3I", estMn);
                                        else if ("##".Equals(estMn))
                                            screen.SetMSOFieldValue("EST_MN_LIFE3I", "");

                                        string lifeGrp = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("O" + currentRow).Value));
                                        if (!"".Equals(lifeGrp) && !"##".Equals(lifeGrp))
                                            screen.SetMSOFieldValue("LIFE_GRP_CODE3I", lifeGrp);
                                        else if ("##".Equals(lifeGrp))
                                            screen.SetMSOFieldValue("LIFE_GRP_CODE3I", "");

                                        string estDispos = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("P" + currentRow).Value));
                                        if (!"".Equals(estDispos) && !"##".Equals(estDispos))
                                            screen.SetMSOFieldValue("EST_DISPOS_VAL3I", estDispos);
                                        else if ("##".Equals(estDispos))
                                            screen.SetMSOFieldValue("EST_DISPOS_VAL3I", "");

                                        //Sub Asset Movement Summary
                                        string forCurr = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("Q" + currentRow).Value));
                                        if (!"".Equals(forCurr) && !"##".Equals(forCurr))
                                            screen.SetMSOFieldValue("FOR_CURR_AMT3I", forCurr);
                                        else if ("##".Equals(forCurr))
                                            screen.SetMSOFieldValue("FOR_CURR_AMT3I", "");

                                        string foreignCurr = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range("R" + currentRow).Value));
                                        if (!"".Equals(foreignCurr) && !"##".Equals(foreignCurr))
                                            screen.SetMSOFieldValue("FOREIGN_CURR3I", foreignCurr);
                                        else if ("##".Equals(foreignCurr))
                                            screen.SetMSOFieldValue("FOREIGN_CURR3I", "");

                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);
                                        if (screen.MSO.functionKeys.Contains("XMIT-Confirm"))

                                        {
                                            screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        }

                                        if (screen.IsMSOError())
                                        {
                                            if (screen.GetMSOError().Contains("confirm"))
                                            {
                                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOError())
                                                {
                                                    error = screen.GetMSOError();
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                                else
                                                {
                                                    excelSheet4.get_Range(endColumn4 + currentRow).Value = messageModified;
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                }
                                            }
                                            else
                                            {
                                                error = screen.GetMSOError();
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                        else
                                        {
                                            excelSheet4.get_Range(endColumn4 + currentRow).Value = messageModified;
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        }
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet4.get_Range(endColumn4 + currentRow).Value = errorEx.Message.Trim();
                                    //WorkAround de Excepción de Programa MSO685 detectada
                                    if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                        currentRow--;
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet4.get_Range(beginColumn4 + currentRow).Value));
                                }
                            }
                            excelSheet4.Cells.Columns.AutoFit();
                            excelSheet4.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            screen.ClosedConnection();
                        }
                    }
                    else
                    {
                        MessageBox.Show(messageSelectOption, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    screen.ClosedConnection();
                }
            }
            catch (Exception errorCatch)
            {
                screen.ClosedConnection();
                MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //


        }

        private void btnmso685op4vl_Click(object sender, RibbonControlEventArgs e)
        {
            executeMSO685op4vl();
        }

        private void executeMSO685op3cv()

        {
            EllScreen.Ellipse screen = null;
            if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
            {
                FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                frm.StartPosition = FormStartPosition.CenterScreen;
                frm.ShowDialog();


                elliseUser = frm.Auth.Username;
                ellisePost = frm.Auth.Position;
                elliseDsct = frm.Auth.District;
                ellisePswd = frm.Auth.Password;
                try
                {
                    if (excelSheet == null)
                    {
                        excelApp = Globals.ThisAddIn.Application;
                        excelBook = excelApp.Workbooks.Item[1];
                        excelSheet3 = (excel.Worksheet)excelBook.Sheets[sheetNameOP3vc];
                    }

                    string url = "";
                    String error = "";
                    var confUrl = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory;
                    confUrl = "c:\\ellipse\\EllipseConfiguration.xml";
                    EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(confUrl);

                    if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                    {
                        url = conf.UrlProd;
                    }
                    else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                    {
                        url = conf.UrlCont;
                    }
                    else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                    {
                        url = conf.UrlDesa;
                    }
                    else
                    {
                        url = conf.UrlTest;
                    }

                    int currentRow = dataRow3;
                    string assetRef = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("A" + currentRow).Value));

                    if (assetRef.Equals(""))
                    {
                        MessageBox.Show(messageRequiredFields, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string subAssetNumber = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("B" + currentRow).Value));
                        string bookType = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("C" + currentRow).Value));

                        screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                        ClientConversation.authenticate(elliseUser, ellisePswd);
                        screen.InitMSOInstance(url);

                        if (!screen.GetMSOError().Equals(""))
                        {
                            throw new Exception(screen.GetMSOError());
                        }

                        

                        while (!assetRef.Equals(""))
                        {
                            try
                            {
                                screen.ExecuteScreen("MSO685", "MSM685A");
                                if (!screen.GetMSOError().Equals(""))
                                {
                                    throw new Exception(screen.GetMSOError());
                                }
                                excelSheet3.get_Range(endColumn3 + currentRow).Select();
                                excelSheet3.get_Range(endColumn3 + currentRow).Value = messageProcess3;

                                if (screen.IsScreenNameCorrect("MSM685A"))
                                {
                                    screen.InitScreenFields();
                                    screen.SetMSOFieldValue("OPTION1I", "3");
                                    screen.SetMSOFieldValue("ASSET_REF1I", assetRef);
                                    screen.SetMSOFieldValue("SUB_ASSET_NO1I", subAssetNumber);
                                    screen.SetMSOFieldValue("BOOK_OR_TAX1I", bookType);
                                    screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);

                                    if (screen.MSO.functionKeys.Contains("XMIT-Confirm"))

                                    {
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                    }

                                    if (screen.IsMSOError())
                                    {
                                        error = screen.GetMSOError();
                                        screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        throw new Exception(error.Trim());
                                    }

                                    if (screen.IsScreenNameCorrect("MSM685C"))
                                    {
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        if (screen.IsMSOError())
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                        else
                                        {
                                            excelSheet3.get_Range(endColumn3 + currentRow).Value = "".Equals(screen.GetMSOError()) ? messageUploaded3 : screen.GetMSOError();
                                        }
                                    }
                                    else
                                    {
                                        excelSheet3.get_Range(endColumn3 + currentRow).Value = screen.GetMSOError();
                                    }
                                }
                                else
                                {
                                    excelSheet3.get_Range(endColumn3 + currentRow).Value = screen.GetMSOError();
                                }
                            }
                            catch (Exception errorEx)
                            {
                                excelSheet3.get_Range(endColumn3 + currentRow).Value = errorEx.Message;
                                //WorkAround de Excepción de Programa MSO685 detectada
                                if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                    currentRow--;
                            }
                            finally
                            {
                                currentRow++;
                                assetRef = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("A" + currentRow).Value));
                                subAssetNumber = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("B" + currentRow).Value));
                                bookType = Utils.formatearCeldaACadena(Convert.ToString(excelSheet3.get_Range("C" + currentRow).Value));
                            }
                        }

                        excelSheet3.Cells.Columns.AutoFit();
                        excelSheet3.Cells.Rows.AutoFit();
                        MessageBox.Show(messageProcessFinished3, messageTitle3, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        screen.ClosedConnection();
                    }
                }
                catch (Exception errorCatch)
                {
                    screen.ClosedConnection();
                    MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, messageTitleError3, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(messageSelectOption3, messageTitleError3, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        

    }
        private void executeMSO_685op1_vc() {
            EllScreen.Ellipse screen = null;
            try
            {
                if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        elliseUser = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
                        ellisePost = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.PosDefault;
                        elliseDsct = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.DstrDefault;
                    }

                    FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.ShowDialog();

                    if (frm.Auth.Authenticated)
                    {
                        elliseUser = frm.Auth.Username;
                        ellisePost = frm.Auth.Position;
                        elliseDsct = frm.Auth.District;
                        ellisePswd = frm.Auth.Password;

                        if (excelSheet == null)
                        {
                            excelApp = Globals.ThisAddIn.Application;
                            excelBook = excelApp.Workbooks.Item[1];
                            excelSheet = (excel.Worksheet)excelBook.Sheets[sheetNameOP1vc];
                        }

                        string url = "";


                        var confUrl = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory;
                        
                        EllipseEnviroment.EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(confUrl);

                        if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = conf.UrlProd;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = conf.UrlCont;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = conf.UrlDesa;
                        }
                        else
                        {
                            url = conf.UrlTest;
                        }

                        string error;
                        int currentRow = dataRowOP1vc;                        

                        String campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range(beginColumnOP1vc + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            excelSheet.Select();

                            screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                            ClientConversation.authenticate(elliseUser, ellisePswd);
                            screen.InitMSOInstance(url);
                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    screen.ExecuteScreen("MSO685", "MSM685A");
                                    if (!screen.GetMSOError().Equals(""))
                                    {
                                        throw new Exception(screen.GetMSOError());
                                    }
                                    excelSheet.get_Range(endColumnOP1vc + currentRow).Select();
                                    excelSheet.get_Range(endColumnOP1vc + currentRow).Value = messageProcess;

                                    if (screen.MSO.mapName.Equals("MSM685A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "1");
                                        screen.SetMSOFieldValue("ASSET_REF1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUB_ASSET_NO1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value)));
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOError())
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (screen.MSO.mapName.Equals("MSM685B"))
                                    {
                                        screen.InitScreenFields();
                                       

                                        string depr = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("C" + currentRow).Value));
                                        if (!"".Equals(depr) && !"##".Equals(depr))
                                            screen.SetMSOFieldValue("DEPR_IND2I", depr);
                                        else if ("##".Equals(depr))
                                            screen.SetMSOFieldValue("DEPR_IND2I", "");

                                        string capitalisation = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("D" + currentRow).Value));
                                        if (!"".Equals(capitalisation) && !"##".Equals(capitalisation))
                                            screen.SetMSOFieldValue("CAPITALISATION_TY2I", capitalisation);
                                        else if ("##".Equals(capitalisation))
                                            screen.SetMSOFieldValue("CAPITALISATION_TY2I", "");

                                      
                                        //Account Profile
                                        string acct = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E" + currentRow).Value));
                                        if (!"".Equals(acct) && !"##".Equals(acct))
                                            screen.SetMSOFieldValue("ACCT_PROFILE2I", acct);
                                        else if ("##".Equals(acct))
                                            screen.SetMSOFieldValue("ACCT_PROFILE2I", "");

                                        string deprExp = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("F" + currentRow).Value));
                                        if (!"".Equals(deprExp) && !"##".Equals(deprExp))
                                            screen.SetMSOFieldValue("DEPR_EXP_CODE2I", deprExp);
                                        else if ("##".Equals(deprExp))
                                            screen.SetMSOFieldValue("DEPR_EXP_CODE2I", "");

                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);

                                        if (screen.MSO.functionKeys.Contains("XMIT-Confirm"))

                                        {
                                            screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        }

                                        //if (screen.IsMSOMessage())
                                        if (screen.IsMSOError())
                                        {
                                            if (screen.GetMSOError().Contains("confirm"))
                                            {
                                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOError())
                                                {
                                                    error = screen.GetMSOError();
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                                else
                                                {
                                                    excelSheet.get_Range(endColumnOP1vc + currentRow).Value = messageModified;
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                }
                                            }
                                            else {
                                                error = screen.GetMSOError();
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                        else
                                        {
                                            excelSheet.get_Range(endColumnOP1vc + currentRow).Value = messageModified;
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        }
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet.get_Range(endColumnOP1vc + currentRow).Value = errorEx.Message.Trim();
                                    //WorkAround de Excepción de Programa MSO685 detectada
                                    if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                        currentRow--;
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range(beginColumnOP1vc + currentRow).Value));
                                }
                            }
                            excelSheet.Cells.Columns.AutoFit();
                            excelSheet.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            screen.ClosedConnection();

                        }
                    }
                    else
                    {
                        MessageBox.Show(messageSelectOption, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    screen.ClosedConnection();
                    MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception errorCatch)
            {
                screen.ClosedConnection();
                MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_op3vl_Click(object sender, RibbonControlEventArgs e)
        {
            
            try
                {
                //Ejecutar MSO685
                execute_mso685op3vL();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnOp4vc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                executeMSOmso685op4vc();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                
                executeMSO_685op1_vc();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            

            try
            {
                executeMSO685op3cv();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btn_MSO685OP1VL_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                executeMSO685op1vl();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Ejecutar MSO685 - Opcion 1
                executeMSO_685op1_vc();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void executeMSO685op1vl()
        {
            EllScreen.Ellipse screen = null;
            try
            {
                if (drpSubAssetGeneralInfoEnv.Label != null && !drpSubAssetGeneralInfoEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        elliseUser = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
                        ellisePost = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.PosDefault;
                        elliseDsct = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.DstrDefault;
                    }

                    FormAuthenticate frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.ShowDialog();

                    if (frm.Auth.Authenticated)
                    {
                        elliseUser = frm.Auth.Username;
                        ellisePost = frm.Auth.Position;
                        elliseDsct = frm.Auth.District;
                        ellisePswd = frm.Auth.Password;

                        if (excelSheet == null)
                        {
                            excelApp = Globals.ThisAddIn.Application;
                            excelBook = excelApp.Workbooks.Item[1];
                            excelSheet2 = (excel.Worksheet)excelBook.Sheets[sheetNameOP1vl];
                        }

                        string url = "";

                        EllipseEnviroment.EllipseConfiguration conf = EllipseEnviroment.Util.GetEllipseConfiguration(global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                        if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = conf.UrlProd;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = conf.UrlCont;
                        }
                        else if (drpSubAssetGeneralInfoEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = conf.UrlDesa;
                        }
                        else
                        {
                            url = conf.UrlTest;
                        }

                        string error;
                        int currentRow = dataRow2;

                        String campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range(beginColumn2 + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            excelSheet2.Select();

                            screen = new EllScreen.Ellipse(elliseDsct, ellisePost, 100, true);
                            ClientConversation.authenticate(elliseUser, ellisePswd);
                            screen.InitMSOInstance(url);

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    screen.ExecuteScreen("MSO685", "MSM685A");
                                    if (!screen.GetMSOError().Equals(""))
                                    {
                                        throw new Exception(screen.GetMSOError());
                                    }

                                    excelSheet2.get_Range(endColumn2 + currentRow).Select();
                                    excelSheet2.get_Range(endColumn2 + currentRow).Value = messageProcess;

                                    if (screen.MSO.mapName.Equals("MSM685A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "1");
                                        screen.SetMSOFieldValue("ASSET_REF1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUB_ASSET_NO1I", Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("B" + currentRow).Value)));
                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);
                                        //if (screen.IsMSOMessage())
                                        if (screen.IsMSOError())
                                        {
                                            error = screen.GetMSOError();
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (screen.MSO.mapName.Equals("MSM685B"))
                                    {
                                        screen.InitScreenFields();
                                        string serial = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("C" + currentRow).Value));
                                        if (!"".Equals(serial) && !"##".Equals(serial))
                                            screen.SetMSOFieldValue("SERIAL_EQUIP2I", serial);
                                        else if ("##".Equals(serial))
                                            screen.SetMSOFieldValue("SERIAL_EQUIP2I", "");

                                        string subAsset = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("D" + currentRow).Value));
                                        if (!"".Equals(subAsset) && !"##".Equals(subAsset))
                                            screen.SetMSOFieldValue("SUB_ASSET_DESC2I", subAsset);
                                        else if ("##".Equals(subAsset))
                                            screen.SetMSOFieldValue("SUB_ASSET_DESC2I", "");

                                        //Asset Details
                                        string assetClassif = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("E" + currentRow).Value));
                                        if (!"".Equals(assetClassif) && !"##".Equals(assetClassif))
                                            screen.SetMSOFieldValue("ASSET_CLASSIF2I", assetClassif);
                                        else if ("##".Equals(assetClassif))
                                            screen.SetMSOFieldValue("ASSET_CLASSIF2I", "");

                                        string depr = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("F" + currentRow).Value));
                                        if (!"".Equals(depr) && !"##".Equals(depr))
                                            screen.SetMSOFieldValue("DEPR_IND2I", depr);
                                        else if ("##".Equals(depr))
                                            screen.SetMSOFieldValue("DEPR_IND2I", "");

                                        string capitalisation = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("G" + currentRow).Value));
                                        if (!"".Equals(capitalisation) && !"##".Equals(capitalisation))
                                            screen.SetMSOFieldValue("CAPITALISATION_TY2I", capitalisation);
                                        else if ("##".Equals(capitalisation))
                                            screen.SetMSOFieldValue("CAPITALISATION_TY2I", "");

                                        string revn = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("H" + currentRow).Value));
                                        if (!"".Equals(revn) && !"##".Equals(revn))
                                            screen.SetMSOFieldValue("REVN_SUBCLASS2I", revn);
                                        else if ("##".Equals(revn))
                                            screen.SetMSOFieldValue("REVN_SUBCLASS2I", "");

                                        string cash = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("I" + currentRow).Value));
                                        if (!"".Equals(revn) && !"##".Equals(revn))
                                            screen.SetMSOFieldValue("CASH_GEN_UNIT2I", revn);
                                        else if ("##".Equals(revn))
                                            screen.SetMSOFieldValue("CASH_GEN_UNIT2I", "");

                                        string report = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("J" + currentRow).Value));
                                        if (!"".Equals(report) && !"##".Equals(report))
                                            screen.SetMSOFieldValue("REPORT_CODE2I", report);
                                        else if ("##".Equals(report))
                                            screen.SetMSOFieldValue("REPORT_CODE2I", "");

                                        //Update depreciation Details
                                        string book = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("K" + currentRow).Value));
                                        if (!"".Equals(book) && !"##".Equals(book))
                                            screen.SetMSOFieldValue("BOOK2I", book);
                                        else if ("##".Equals(book))
                                            screen.SetMSOFieldValue("BOOK2I", "");

                                        string tax = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("L" + currentRow).Value));
                                        if (!"".Equals(tax) && !"##".Equals(tax))
                                            screen.SetMSOFieldValue("TAX12I", tax);
                                        else if ("##".Equals(tax))
                                            screen.SetMSOFieldValue("TAX12I", "");

                                        string tax2 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("M" + currentRow).Value));
                                        if (!"".Equals(tax2) && !"##".Equals(tax2))
                                            screen.SetMSOFieldValue("TAX22I", tax2);
                                        else if ("##".Equals(tax2))
                                            screen.SetMSOFieldValue("TAX22I", "");

                                        string tax3 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("N" + currentRow).Value));
                                        if (!"".Equals(tax3) && !"##".Equals(tax3))
                                            screen.SetMSOFieldValue("TAX32I", tax3);
                                        else if ("##".Equals(tax3))
                                            screen.SetMSOFieldValue("TAX32I", "");

                                        string tax4 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("O" + currentRow).Value));
                                        if (!"".Equals(tax4) && !"##".Equals(tax4))
                                            screen.SetMSOFieldValue("TAX42I", tax4);
                                        else if ("##".Equals(tax4))
                                            screen.SetMSOFieldValue("TAX42I", "");

                                        //Asset Clasification
                                        string deprCode = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("P" + currentRow).Value));
                                        if (!"".Equals(deprCode) && !"##".Equals(deprCode))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I1", deprCode);
                                        else if ("##".Equals(deprCode))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I1", "");

                                        string deprCode2 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("Q" + currentRow).Value));
                                        if (!"".Equals(deprCode2) && !"##".Equals(deprCode2))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I2", deprCode2);
                                        else if ("##".Equals(deprCode2))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I2", "");

                                        string deprCode3 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("R" + currentRow).Value));
                                        if (!"".Equals(deprCode3) && !"##".Equals(deprCode3))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I3", deprCode3);
                                        else if ("##".Equals(deprCode3))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I3", "");

                                        string deprCode4 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("S" + currentRow).Value));
                                        if (!"".Equals(deprCode4) && !"##".Equals(deprCode4))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I4", deprCode4);
                                        else if ("##".Equals(deprCode4))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I4", "");

                                        string deprCode5 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("T" + currentRow).Value));
                                        if (!"".Equals(deprCode5) && !"##".Equals(deprCode5))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I5", deprCode5);
                                        else if ("##".Equals(deprCode5))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I5", "");

                                        string deprCode6 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("U" + currentRow).Value));
                                        if (!"".Equals(deprCode6) && !"##".Equals(deprCode6))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I6", deprCode6);
                                        else if ("##".Equals(deprCode6))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I6", "");

                                        string deprCode7 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("V" + currentRow).Value));
                                        if (!"".Equals(deprCode7) && !"##".Equals(deprCode7))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I7", deprCode7);
                                        else if ("##".Equals(deprCode7))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I7", "");

                                        string deprCode8 = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("W" + currentRow).Value));
                                        if (!"".Equals(deprCode8) && !"##".Equals(deprCode8))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I8", deprCode8);
                                        else if ("##".Equals(deprCode8))
                                            screen.SetMSOFieldValue("DEPR_CODE_A2I8", "");

                                        //Account Profile
                                        string acct = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("X" + currentRow).Value));
                                        if (!"".Equals(acct) && !"##".Equals(acct))
                                            screen.SetMSOFieldValue("ACCT_PROFILE2I", acct);
                                        else if ("##".Equals(acct))
                                            screen.SetMSOFieldValue("ACCT_PROFILE2I", "");

                                        string deprExp = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range("Y" + currentRow).Value));
                                        if (!"".Equals(deprExp) && !"##".Equals(deprExp))
                                            screen.SetMSOFieldValue("DEPR_EXP_CODE2I", deprExp);
                                        else if ("##".Equals(deprExp))
                                            screen.SetMSOFieldValue("DEPR_EXP_CODE2I", "");

                                        screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, true);

                                        if (screen.MSO.functionKeys.Contains("XMIT-Confirm"))

                                        {
                                            screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                        }

                                        if (screen.IsMSOError())
                                        {
                                            if (screen.GetMSOError().Contains("confirm"))
                                            {
                                                screen.ExecuteMSO(EllScreen.Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOError())
                                                {
                                                    error = screen.GetMSOError();
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                                else
                                                {
                                                    excelSheet2.get_Range(endColumn2 + currentRow).Value = messageModified;
                                                    screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                }
                                            }
                                            else
                                            {
                                                error = screen.GetMSOError();
                                                screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                        else
                                         {
                                            excelSheet2.get_Range(endColumn2 + currentRow).Value = messageModified;
                                            screen.ExecuteMSO(EllScreen.Ellipse.F3_KEY, false);
                                        }
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet2.get_Range(endColumn2 + currentRow).Value = errorEx.Message.Trim();
                                    //WorkAround de Excepción de Programa MSO685 detectada
                                    if (errorEx.Message.Equals("Fault: com.mincom.ellipse.ejra.jca.EllipseProgramExecutionException"))
                                        currentRow--;
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = Utils.formatearCeldaACadena(Convert.ToString(excelSheet2.get_Range(beginColumn2 + currentRow).Value));
                                }
                            }
                            excelSheet2.Cells.Columns.AutoFit();
                            excelSheet2.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            screen.ClosedConnection();
                        }
                    }
                    else
                    {
                        MessageBox.Show(messageSelectOption, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    screen.ClosedConnection();
                    MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception errorCatch)
            {
                screen.ClosedConnection();
                MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    
    }
}

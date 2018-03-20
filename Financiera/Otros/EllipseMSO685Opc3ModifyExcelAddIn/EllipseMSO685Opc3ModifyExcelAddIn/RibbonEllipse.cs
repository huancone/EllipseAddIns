using System;
using System.Drawing;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Cerrejon.Screen.Post;
using EllipseMSO685Opc3ModifyExcelAddIn;
using EllipseMSO685Opc3ModifyExcelAddIn.Properties;
using Microsoft.Office.Tools.Ribbon;
using excel = Microsoft.Office.Interop.Excel;
using Util = System.Web.Services.Ellipse.Util;


namespace EllipseMSO685Opc3ModifyExcelAddIn
{
    public partial class RibbonEllipse
    {
        public static string ElliseUser = "";
        public static string EllisePswd = "";
        public static string EllisePost = "";
        public static string ElliseDsct = "";
        //SHEET
        public static string SheetName = "MSO685_Opc3_Modify";
        //COLUMNS
        public static string BeginColumn = "A";
        public static string EndColumn = "S";
        //ROWS
        public static int HeaderRow = 4;
        public static int DataRow = 5;
        //MESSAGE CELL
        public static string MessageProcess = "Processing...";
        public static string MessageUploaded = "Uploaded";
        public static string MessageModified = "Modified";
        public static string MessageStatusloaded = "Status loaded";
        public static string MessageDataRequired = "Data required";
        //MESSAGE BOX
        public static string MessageTitle = "Message";
        public static string MessageTitleError = "Error";

        public static string MessageRequiredFields =
            "\n\rYou should fill the table with valid information for processing";

        public static string MessageProcessFinished = "\n\rProcess finished";
        public static string MessageSelectOption = "\n\rPlease Select a Env. Option";
        private excel.Application _excelApp;
        private excel.Workbook _excelBook;
        private excel.Worksheet _excelSheet;
        internal excel.ListObject ExcelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            drpEnv.SelectedItemIndex = Util.GetEnvironment(Settings.Default.EnvDefault);
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Cargar formato MSO685 Opcion 3
                LoadFormatMso();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void btnExecute_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Ejecutar MSO685 - Opcion 3
                ExecuteMso();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void LoadFormatMso()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelBook = _excelApp.Workbooks.Add();
                _excelSheet = (excel.Worksheet) _excelBook.Sheets.Add();

                _excelSheet.Name = SheetName;

                var rangeMaintItem = _excelSheet.Range[BeginColumn + "1:" + EndColumn + "1"];
                rangeMaintItem.Font.Bold = true;
                rangeMaintItem.Merge();
                rangeMaintItem.Interior.Color = Color.FromArgb(79, 129, 189);
                rangeMaintItem.Font.Color = Color.White;
                rangeMaintItem.Value = "MSO685 Opcion 3 Maintain Sub-Asset Depreciation Details - Ellipse 8 Loader";
                rangeMaintItem.WrapText = true;
                rangeMaintItem = _excelSheet.Range[BeginColumn + "1:" + EndColumn + "1"];
                rangeMaintItem.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeMaintItem.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeMaintItem.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeMaintItem.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeMaintItem.Borders.Color = Color.Black;

                var rangeItemTitle1 = _excelSheet.Range["A2"];
                rangeItemTitle1.Font.Color = Color.Green;
                rangeItemTitle1.Value = "## - Borrar";
                rangeItemTitle1 = _excelSheet.Range["B2"];
                rangeItemTitle1.Font.Color = Color.Green;
                rangeItemTitle1.Value = "Vacio - No se modifica";

                //Encabezado MSM685A
                var rangeItem = _excelSheet.Range["A" + HeaderRow];
                rangeItem.Value = "Asset Reference *";
                rangeItem = _excelSheet.Range["B" + HeaderRow];
                rangeItem.Value = "Sub Asset Number *";
                rangeItem = _excelSheet.Range["C" + HeaderRow];
                rangeItem.Value = "Book Type *";

                //MSM685C
                //Depreciation Details
                rangeItem = _excelSheet.Range["D3:P3"];
                rangeItem.Merge();
                rangeItem.Interior.Color = Color.FromArgb(79, 129, 189);
                rangeItem.Font.Color = Color.White;
                rangeItem.Font.Bold = true;
                rangeItem.Value = "Depreciation Details";
                rangeItem = _excelSheet.Range["D3:P3"];
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders.Color = Color.Black;
                //Depreciation Elements
                rangeItem = _excelSheet.Range["D" + HeaderRow];
                rangeItem.Value = "Depreciation Method *";
                rangeItem = _excelSheet.Range["E" + HeaderRow];
                rangeItem.Value = "Depreciation Rate";
                rangeItem = _excelSheet.Range["F" + HeaderRow];
                rangeItem.Value = "Manual Period Depn";
                rangeItem = _excelSheet.Range["G" + HeaderRow];
                rangeItem.Value = "Until Period";
                rangeItem = _excelSheet.Range["H" + HeaderRow];
                rangeItem.Value = "Accelerated Depn Rate";
                rangeItem = _excelSheet.Range["I" + HeaderRow];
                rangeItem.Value = "Until Period";
                rangeItem = _excelSheet.Range["J" + HeaderRow];
                rangeItem.Value = "Rate Table";
                rangeItem = _excelSheet.Range["K" + HeaderRow];
                rangeItem.Value = "Recovery Period";
                rangeItem = _excelSheet.Range["L" + HeaderRow];
                rangeItem.Value = "Dividend Statistic";
                rangeItem = _excelSheet.Range["M" + HeaderRow];
                rangeItem.Value = "Divisor Statistic";
                rangeItem = _excelSheet.Range["N" + HeaderRow];
                rangeItem.Value = "Estimated Life (months)";
                rangeItem = _excelSheet.Range["O" + HeaderRow];
                rangeItem.Value = "Useful Life Group Code";
                rangeItem = _excelSheet.Range["P" + HeaderRow];
                rangeItem.Value = "Est Retirement Value - Local";

                //Sub Asset Movement Summary
                rangeItem = _excelSheet.Range["Q3:R3"];
                rangeItem.Merge();
                rangeItem.Interior.Color = Color.FromArgb(79, 129, 189);
                rangeItem.Font.Color = Color.White;
                rangeItem.Font.Bold = true;
                rangeItem.Value = "Sub Asset Movement Summary";
                rangeItem = _excelSheet.Range["Q3:R3"];
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItem.Borders.Color = Color.Black;
                //Capitalization Details
                rangeItem = _excelSheet.Range["Q" + HeaderRow];
                rangeItem.Value = "Foreign Currency Cost";
                rangeItem = _excelSheet.Range["R" + HeaderRow];
                rangeItem.Value = "Foreign Currency Type";

                rangeItem = _excelSheet.Range["S" + HeaderRow];
                rangeItem.Value = "Message";

                //Se aplica formato de texto al subactivo
                _excelSheet.Range[BeginColumn + DataRow + ":" + EndColumn + "100000"].NumberFormat = "@";

                ExcelSheetItems = _excelSheet.ListObjects.AddEx(excel.XlListObjectSourceType.xlSrcRange,
                    _excelSheet.Range[BeginColumn + HeaderRow + ":" + EndColumn + "100000"],
                    XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void ExecuteMso()
        {
            try
            {
                if (drpEnv.Label != null && !drpEnv.Label.Equals(""))
                {
                    if (ElliseUser.Equals(""))
                    {
                        ElliseUser = Settings.Default.UserDefault;
                        EllisePost = Settings.Default.PosDefault;
                        ElliseDsct = Settings.Default.DstrDefault;
                    }

                    var frm = new FormAuthenticate(ElliseUser, ElliseDsct, EllisePost)
                    {
                        StartPosition = FormStartPosition.CenterScreen
                    };
                    frm.ShowDialog();

                    if (frm.Auth.Authenticated)
                    {
                        ElliseUser = frm.Auth.Username;
                        EllisePost = frm.Auth.Position;
                        ElliseDsct = frm.Auth.District;
                        EllisePswd = frm.Auth.Password;

                        if (_excelSheet == null)
                        {
                            _excelApp = Globals.ThisAddIn.Application;
                            _excelBook = _excelApp.Workbooks.Item[1];
                            _excelSheet = (excel.Worksheet) _excelBook.Sheets[SheetName];
                        }

                        string url;

                        var ell = Cerrejon.Screen.Post.Util.GetEllipseConfiguration(Settings.Default.EllipseDirectory);

                        if (drpEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = ell.UrlProd;
                        }
                        else if (drpEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = ell.UrlCont;
                        }
                        else if (drpEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = ell.UrlDesa;
                        }
                        else
                        {
                            url = ell.UrlTest;
                        }

                        var currentRow = DataRow;

                        string campoRequerido =
                            MyUtilities.FormatearCeldaACadena(
                                Convert.ToString(_excelSheet.Range[BeginColumn + currentRow].Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(MessageRequiredFields, MessageTitle, MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            _excelSheet.Select();

                            var screen = new Ellipse(ElliseUser, EllisePswd, EllisePost, ElliseDsct, url);
                            screen.InitConexion();

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            screen.ExecuteScreen("MSO685", "MSM685A");
                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    _excelSheet.Range[EndColumn + currentRow].Select();
                                    _excelSheet.Range[EndColumn + currentRow].Value = MessageProcess;

                                    string error;
                                    if (screen.MSO.MapName.Equals("MSM685A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "3");
                                        screen.SetMSOFieldValue("ASSET_REF1I",
                                            MyUtilities.FormatearCeldaACadena(
                                                Convert.ToString(_excelSheet.Range["A" + currentRow].Value)));
                                        screen.SetMSOFieldValue("SUB_ASSET_NO1I",
                                            MyUtilities.FormatearCeldaACadena(
                                                Convert.ToString(_excelSheet.Range["B" + currentRow].Value)));
                                        screen.SetMSOFieldValue("BOOK_OR_TAX1I",
                                            MyUtilities.FormatearCeldaACadena(
                                                Convert.ToString(_excelSheet.Range["C" + currentRow].Value)));
                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOMessage())
                                        {
                                            error = screen.GetMSOMessage();
                                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (!screen.MSO.MapName.Equals("MSM685C")) continue;
                                    screen.InitScreenFields();
                                    //Depreciation Details
                                    string deprMethod =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["D" + currentRow].Value));
                                    if (!"".Equals(deprMethod) && !"##".Equals(deprMethod))
                                        screen.SetMSOFieldValue("DEPR_METHOD3I", deprMethod);
                                    else if ("##".Equals(deprMethod))
                                        screen.SetMSOFieldValue("DEPR_METHOD3I", "");

                                    string deprRate =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["E" + currentRow].Value));
                                    if (!"".Equals(deprRate) && !"##".Equals(deprRate))
                                        screen.SetMSOFieldValue("DEPR_RATE3I", deprRate);
                                    else if ("##".Equals(deprRate))
                                        screen.SetMSOFieldValue("DEPR_RATE3I", "");

                                    string manPer =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["F" + currentRow].Value));
                                    if (!"".Equals(manPer) && !"##".Equals(manPer))
                                        screen.SetMSOFieldValue("MAN_PER_DEPR3I", manPer);
                                    else if ("##".Equals(manPer))
                                        screen.SetMSOFieldValue("MAN_PER_DEPR3I", "");

                                    string finMan =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["G" + currentRow].Value));
                                    if (!"".Equals(finMan) && !"##".Equals(finMan))
                                        screen.SetMSOFieldValue("FIN_MAN_PER3I", finMan);
                                    else if ("##".Equals(finMan))
                                        screen.SetMSOFieldValue("FIN_MAN_PER3I", "");

                                    string accelDepr =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["H" + currentRow].Value));
                                    if (!"".Equals(accelDepr) && !"##".Equals(accelDepr))
                                        screen.SetMSOFieldValue("ACCEL_DEPR_RT3I", accelDepr);
                                    else if ("##".Equals(accelDepr))
                                        screen.SetMSOFieldValue("ACCEL_DEPR_RT3I", "");

                                    string finAccel =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["I" + currentRow].Value));
                                    if (!"".Equals(finAccel) && !"##".Equals(finAccel))
                                        screen.SetMSOFieldValue("FIN_ACCEL_PER3I", finAccel);
                                    else if ("##".Equals(finAccel))
                                        screen.SetMSOFieldValue("FIN_ACCEL_PER3I", "");

                                    string rateTable =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["J" + currentRow].Value));
                                    if (!"".Equals(rateTable) && !"##".Equals(rateTable))
                                        screen.SetMSOFieldValue("RATE_TABLE3I", rateTable);
                                    else if ("##".Equals(rateTable))
                                        screen.SetMSOFieldValue("RATE_TABLE3I", "");

                                    MyUtilities.FormatearCeldaACadena(
                                        Convert.ToString(_excelSheet.Range["K" + currentRow].Value));
                                    if (!"".Equals(rateTable) && !"##".Equals(rateTable))
                                        screen.SetMSOFieldValue("RECOV_PERIOD3I", rateTable);
                                    else if ("##".Equals(rateTable))
                                        screen.SetMSOFieldValue("RECOV_PERIOD3I", "");

                                    string dividendStat =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["L" + currentRow].Value));
                                    if (!"".Equals(dividendStat) && !"##".Equals(dividendStat))
                                        screen.SetMSOFieldValue("DIVIDEND_STAT3I", dividendStat);
                                    else if ("##".Equals(dividendStat))
                                        screen.SetMSOFieldValue("DIVIDEND_STAT3I", "");

                                    string divisorStat =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["M" + currentRow].Value));
                                    if (!"".Equals(divisorStat) && !"##".Equals(divisorStat))
                                        screen.SetMSOFieldValue("DIVISOR_STAT3I", divisorStat);
                                    else if ("##".Equals(divisorStat))
                                        screen.SetMSOFieldValue("DIVISOR_STAT3I", "");

                                    string estMn =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["N" + currentRow].Value));
                                    if (!"".Equals(estMn) && !"##".Equals(estMn))
                                        screen.SetMSOFieldValue("EST_MM_LIFE3I", estMn);
                                    else if ("##".Equals(estMn))
                                        screen.SetMSOFieldValue("EST_MM_LIFE3I", "");

                                    string lifeGrp =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["O" + currentRow].Value));
                                    if (!"".Equals(lifeGrp) && !"##".Equals(lifeGrp))
                                        screen.SetMSOFieldValue("LIFE_GRP_CODE3I", lifeGrp);
                                    else if ("##".Equals(lifeGrp))
                                        screen.SetMSOFieldValue("LIFE_GRP_CODE3I", "");

                                    string estDispos =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["P" + currentRow].Value));
                                    if (!"".Equals(estDispos) && !"##".Equals(estDispos))
                                        screen.SetMSOFieldValue("EST_DISPOS_VAL3I", estDispos);
                                    else if ("##".Equals(estDispos))
                                        screen.SetMSOFieldValue("EST_DISPOS_VAL3I", "");

                                    //Sub Asset Movement Summary
                                    string forCurr =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["Q" + currentRow].Value));
                                    if (!"".Equals(forCurr) && !"##".Equals(forCurr))
                                        screen.SetMSOFieldValue("FOR_CURR_AMT3I", forCurr);
                                    else if ("##".Equals(forCurr))
                                        screen.SetMSOFieldValue("FOR_CURR_AMT3I", "");

                                    string foreignCurr =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range["R" + currentRow].Value));
                                    if (!"".Equals(foreignCurr) && !"##".Equals(foreignCurr))
                                        screen.SetMSOFieldValue("FOREIGN_CURR3I", foreignCurr);
                                    else if ("##".Equals(foreignCurr))
                                        screen.SetMSOFieldValue("FOREIGN_CURR3I", "");

                                    screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                    
                                    if (screen.IsMSOMessage())
                                    {
                                        if (screen.GetMSOInformation().Contains("confirm"))
                                        {
                                            screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                            if (screen.IsMSOMessage())
                                            {
                                                error = screen.GetMSOMessage();
                                                screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                            _excelSheet.Range[EndColumn + currentRow].Value = MessageModified;
                                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                        }
                                        else
                                        {
                                            error = screen.GetMSOMessage();
                                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }
                                    else
                                    {
                                        _excelSheet.Range[EndColumn + currentRow].Value = MessageModified;
                                        screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    _excelSheet.Range[EndColumn+ currentRow].Value = errorEx.Message.Trim();
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido =
                                        MyUtilities.FormatearCeldaACadena(
                                            Convert.ToString(_excelSheet.Range[BeginColumn + currentRow].Value));
                                }
                            }
                            _excelSheet.Cells.Columns.AutoFit();
                            _excelSheet.Cells.Rows.AutoFit();
                            MessageBox.Show(MessageProcessFinished, MessageTitle, MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show(MessageSelectOption, MessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show(MessageSelectOption, MessageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception errorCatch)
            {
                MessageBox.Show(
                    "\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" +
                    errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
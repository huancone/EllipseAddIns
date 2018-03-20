using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Web.Services.Ellipse;
using Cerrejon.Screen.Post;

namespace EllipseModifyInvoicesExcelAddIn
{
    public partial class RibbonEllipse
    {
        public static String elliseUser = "";
        public static String ellisePswd = "";
        public static String ellisePost = "";
        public static String elliseDsct = "";

        excel.Application excelApp;
        excel.Workbook excelBook;
        excel.Worksheet excelSheet;
        excel.ListObject excelSheetItems;
        //SHEET
        public static String sheetName = "MSO261";
        //COLUMNS
        public static String beginColumn = "A";
        public static String endColumn = "K";
        //ROWS
        public static int headerRow = 2;
        public static int dataRow = 3;

        //MESSAGE CELL
        public static string messageProcess = "Processing...";
        public static string messageUploaded = "Uploaded";
        public static string messageStatusloaded = "Status loaded";
        //MESSAGE BOX
        public static string messageTitle = "Message";
        public static string messageTitleError = "Error";
        public static string messageRequiredFields = "\n\rYou should fill the table with valid information for processing";
        public static string messageProcessFinished = "\n\rProcess finished";
        public static string messageUserPassincorrect = "\n\rThe user does not exist or the password is incorrect";
        public static string messageSelectOption = "\n\rPlease Select a Env. Option";

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            drpModifyInvoicesEnv.SelectedItemIndex = System.Web.Services.Ellipse.Util.GetEnvironment(global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.EnvDefault);
        }

        private void btnModifyInvoicesLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //EJECUTAR MSO: MSO261
                loadStatusInvoicesMSO();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnModifyInvoicesFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                excelBook = excelApp.Workbooks.Add();
                excelSheet = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet.Name = "Planilla";

                excel.Range RangeMaintItem = excelSheet.get_Range(beginColumn + "1:" + endColumn + "1");
                RangeMaintItem.Font.Bold = true;
                RangeMaintItem.Merge();
                RangeMaintItem.Value = "MSO261- Modify Invoices";
                RangeMaintItem.WrapText = true;
                RangeMaintItem.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                excel.Range RangeItemTitle = excelSheet.get_Range("A" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "District";
                RangeItemTitle = excelSheet.get_Range("A:A");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("B" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Supplier";
                RangeItemTitle = excelSheet.get_Range("B:B");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("C" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Invoice No";
                RangeItemTitle = excelSheet.get_Range("C:C");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("D" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Payment Status";
                RangeItemTitle = excelSheet.get_Range("D:D");
                RangeItemTitle.NumberFormat = "@";                

                RangeItemTitle = excelSheet.get_Range("E" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Due Date (AAAAMMDD)";
                RangeItemTitle = excelSheet.get_Range("E:E");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("F" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Bank Branch";
                RangeItemTitle = excelSheet.get_Range("F:F");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("G" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Bank Account";
                RangeItemTitle = excelSheet.get_Range("G:G");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("H" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Handling Code";
                RangeItemTitle = excelSheet.get_Range("H:H");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("I" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Settlement Discount";
                RangeItemTitle = excelSheet.get_Range("I:I");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("J" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "Discount Date";
                RangeItemTitle = excelSheet.get_Range("J:J");
                RangeItemTitle.NumberFormat = "@";

                RangeItemTitle = excelSheet.get_Range("K" + headerRow);
                RangeItemTitle.Font.Bold = true;
                RangeItemTitle.Value = "message";
                RangeItemTitle = excelSheet.get_Range("K:K");
                RangeItemTitle.NumberFormat = "@";

                excelSheetItems = excelSheet.ListObjects.AddEx(SourceType: excel.XlListObjectSourceType.xlSrcRange, Source: excelSheet.Range[beginColumn + headerRow + ":" + endColumn + "40"], XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnModifyInvoicesExecute_Click(object sender, RibbonControlEventArgs e)
        {
            if (drpModifyInvoicesEnv.Label != null && !drpModifyInvoicesEnv.Label.Equals(""))
            {
                if (elliseUser.Equals(""))
                {
                    elliseUser = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.UserDefault;
                    ellisePost = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.PosDefault;
                    elliseDsct = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.DstrDefault;
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

                    try
                    {
                        if (excelSheet == null)
                        {
                            excelApp = Globals.ThisAddIn.Application;
                            excelBook = excelApp.Workbooks.Item[1];
                            excelSheet = (excel.Worksheet)excelBook.Sheets[sheetName];
                        }

                        string url = "";

                        EllipseConfiguration ell = Cerrejon.Screen.Post.Util.GetEllipseConfiguration(global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                        if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = ell.UrlProd;
                        }
                        else if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = ell.UrlCont;
                        }
                        else if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = ell.UrlDesa;
                        }
                        else
                        {
                            url = ell.UrlTest;
                        }

                        string error;
                        int currentRow = dataRow;
                        
                        String campoRequerido = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle);
                        }
                        else
                        {
                            excelSheet.Select();
                            excelSheet.get_Range(endColumn + currentRow.ToString()).Value = messageProcess;

                            Ellipse screen = new Ellipse(elliseUser, ellisePswd, ellisePost, elliseDsct, url);
                            screen.InitConexion();

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            screen.ExecuteScreen("MSO261", "MSM261A");
                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    //System.Array MyValues = (System.Array)excelSheet.get_Range("A" + currentRow.ToString(), "K" + currentRow.ToString()).Cells.Value;
                                    excelSheet.get_Range(endColumn + currentRow.ToString()).Select();
                                    excelSheet.get_Range(endColumn + currentRow.ToString()).Value = messageProcess;

                                    if (screen.MSO.MapName.Equals("MSM261A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "1");
                                        screen.SetMSOFieldValue("DSTRCT_CODE1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUPPLIER_NO1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value)));
                                        screen.SetMSOFieldValue("INV_NO1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("C" + currentRow).Value)));
                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOMessage())
                                        {
                                            error = screen.GetMSOMessage();
                                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                    }

                                    if (screen.MSO.MapName.Equals("MSM261B"))
                                    {
                                        screen.InitScreenFields();
                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("D" + currentRow).Value))))
                                            screen.SetMSOFieldValue("PMT_STATUS2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("D" + currentRow).Value)));
                                        
                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E" + currentRow).Value))))
                                            screen.SetMSOFieldValue("DUE_DATE2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E" + currentRow).Value)));

                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("F" + currentRow).Value))))
                                            screen.SetMSOFieldValue("BRANCH_CODE2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("F" + currentRow).Value)));

                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("G" + currentRow).Value))))
                                            screen.SetMSOFieldValue("BANK_ACCT_NO2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("G" + currentRow).Value)));

                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("H" + currentRow).Value))))
                                            screen.SetMSOFieldValue("HANDLE_CDE2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("H" + currentRow).Value)));

                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("I" + currentRow).Value))))
                                            screen.SetMSOFieldValue("SD_AMOUNT2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("I" + currentRow).Value)));

                                        if (!"".Equals(MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("J" + currentRow).Value))))
                                            screen.SetMSOFieldValue("SD_DATE2I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("J" + currentRow).Value)));

                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true); //Confirma
                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true); //Ejecuta
                                        if (screen.IsMSOMessage())
                                        {
                                            error = screen.GetMSOMessage();
                                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                            throw new Exception(error.Trim());
                                        }
                                        else
                                        {
                                            if (screen.MSO.MapName.Equals("MSM261A"))
                                            {
                                                excelSheet.get_Range(endColumn + currentRow).Value = messageUploaded; 
                                            }
                                        }
                                    }
                                    else if (screen.MSO.MapName.Equals("MSM261A"))
                                    {
                                        excelSheet.get_Range(endColumn + currentRow).Value = "".Equals(screen.GetMSOError()) ? messageUploaded : screen.GetMSOError().Trim();
                                    }
                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet.get_Range(endColumn + currentRow).Value = errorEx.Message.Trim();
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value));
                                    screen.ExecuteMSO(Ellipse.F3_KEY, true);
                                }                        
                            }
                            excelSheet.Cells.Columns.AutoFit();
                            excelSheet.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception errorCatch)
                    {
                        var messageBox = MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show(messageSelectOption, "T", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(messageSelectOption, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }           

        private void loadStatusInvoicesMSO()
        {
            try
            {
                if (drpModifyInvoicesEnv.Label != null && !drpModifyInvoicesEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        elliseUser = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.UserDefault;
                        ellisePost = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.PosDefault;
                        elliseDsct = global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.DstrDefault;
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
                            excelSheet = (excel.Worksheet)excelBook.Sheets[sheetName];
                        }

                        string url = "";

                        EllipseConfiguration ell = Cerrejon.Screen.Post.Util.GetEllipseConfiguration(global::EllipseModifyInvoicesExcelAddIn.Properties.Settings.Default.EllipseDirectory);

                        if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = ell.UrlProd;
                        }
                        else if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = ell.UrlCont;
                        }
                        else if (drpModifyInvoicesEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = ell.UrlDesa;
                        }
                        else
                        {
                            url = ell.UrlTest;
                        }
                        
                        string error;
                        int currentRow = dataRow;
                        
                        String campoRequerido = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value));

                        if (campoRequerido.Equals(""))
                        {
                            MessageBox.Show(messageRequiredFields, messageTitle);
                        }
                        else
                        {
                            excelSheet.Select();
                            excelSheet.get_Range(endColumn + currentRow.ToString()).Value = messageProcess;

                            Ellipse screen = new Ellipse(elliseUser, ellisePswd, ellisePost, elliseDsct, url);
                            screen.InitConexion();

                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            screen.ExecuteScreen("MSO261", "MSM261A");
                            if (!screen.GetMSOError().Equals(""))
                            {
                                throw new Exception(screen.GetMSOError());
                            }

                            while (!"".Equals(campoRequerido))
                            {
                                try
                                {
                                    excelSheet.get_Range("D" + currentRow.ToString()).Select();
                                    excelSheet.get_Range(endColumn + currentRow.ToString()).Value = messageProcess;

                                    if (screen.MSO.MapName.Equals("MSM261A"))
                                    {
                                        screen.InitScreenFields();
                                        screen.SetMSOFieldValue("OPTION1I", "1");
                                        screen.SetMSOFieldValue("DSTRCT_CODE1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("A" + currentRow).Value)));
                                        screen.SetMSOFieldValue("SUPPLIER_NO1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value)));
                                        screen.SetMSOFieldValue("INV_NO1I", MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("C" + currentRow).Value)));

                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                        if (screen.IsMSOMessage())
                                        {
                                            if (!"".Equals(screen.GetMSOWarning()))
                                            {
                                                screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                                if (screen.IsMSOMessage())
                                                {
                                                    error = screen.GetMSOMessage();
                                                    screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                                    throw new Exception(error.Trim());
                                                }
                                            }
                                            else
                                            {
                                                error = screen.GetMSOMessage();
                                                screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                    }

                                    if (screen.MSO.MapName.Equals("MSM261B"))
                                    {
                                        if (!"".Equals(screen.GetMSOFieldValue("PMT_STATUS2I")))
                                            excelSheet.get_Range("D" + currentRow).Value = screen.GetMSOFieldValue("PMT_STATUS2I");

                                        excelSheet.get_Range(endColumn + currentRow).Value = messageStatusloaded;
                                    }
                                    screen.ExecuteMSO(Ellipse.F3_KEY, false);

                                }
                                catch (Exception errorEx)
                                {
                                    excelSheet.get_Range(endColumn + currentRow).Value = errorEx.Message.Trim();
                                }
                                finally
                                {
                                    currentRow++;
                                    campoRequerido = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("B" + currentRow).Value));
                                    screen.ExecuteMSO(Ellipse.F3_KEY, true);
                                }
                            }
                            excelSheet.Cells.Columns.AutoFit();
                            excelSheet.Cells.Rows.AutoFit();
                            MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show(messageSelectOption, "T", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    var messageBox = MessageBox.Show("Please Select a Env. Option", messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception errorCatch)
            {
                var messageBox = MessageBox.Show("\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" + errorCatch.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
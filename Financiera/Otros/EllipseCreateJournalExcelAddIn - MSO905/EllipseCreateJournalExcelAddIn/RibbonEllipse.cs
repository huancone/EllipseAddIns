using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Odbc;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Cerrejon.Screen.Post;
using EllipseCreateJournalExcelAddIn.Properties;
using Microsoft.Office.Tools.Ribbon;
using excel = Microsoft.Office.Interop.Excel;
using Message = Cerrejon.Screen.Post.Message;
using Util = Cerrejon.Screen.Post.Util;

namespace EllipseCreateJournalExcelAddIn
{
    public partial class RibbonEllipse
    {
        public static string elliseUser = "";
        public static string ellisePswd = "";
        public static string ellisePost = "";
        public static string elliseDsct = "";
        //SHEET
        public static string sheetName = "MSO905_Opc3_CreateJournal";
        //COLUMNS
        public static string beginColumn = "A";
        public static string endColumn = "I";
        public static string messageEndColumn = "J";
        //ROWS
        public static int headerRow = 4;
        public static int dataRow = 9;
        //MESSAGE
        public static string errorAuthUser = "USER PROFILE NOT FOUND";
        //MESSAGE CELL
        public static string messageProcess = "Processing...";
        public static string messageUploaded = "Uploaded";
        public static string messageRead = "Read";
        //MESSAGE BOX
        public static string messageTitle = "Message";
        public static string messageTitleError = "Error";

        public static string messageRequiredFields =
            "\n\rYou should fill the table with valid information for processing";

        public static string messageProcessFinished = "\n\rProcess finished";
        public static string messageUserPassincorrect = "\n\rThe user does not exist or the password is incorrect";
        public static string messageSelectOption = "\n\rPlease Select a Env. Option";
        public static string messageWebServiceOK = "Asset No assigned:";
        public static string messageWebWarn2074 = "2074: FOREIGN CURR VAR BAL TRANS";
        public static string messageWebJournalNumber = "JOURNAL NUMBER";
        public static string messageWebWarn3578 = "3578: SECONDARY CURRENCY";
        public static string messageODBC = "\n\rYou need to configure the ODBC connection";
        private excel.Application excelApp;
        private excel.Workbook excelBook;
        private excel.Worksheet excelSheet;
        private excel.ListObject excelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnCreateJournalFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //CARGAR FORMATO DEL MSO: MSO905 Opcion 3
                loadFormat();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnCreateJournalExecute_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //EJECUTAR MSO: MSO905 Opcion 3
                executeMSO();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void loadFormat()
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                excelBook = excelApp.Workbooks.Add();
                excelSheet = (excel.Worksheet)excelBook.Sheets.Add();

                excelSheet.Name = sheetName;

                var RangeTitle = excelSheet.Range[beginColumn + "1:" + endColumn + "1"];
                RangeTitle.Font.Bold = true;
                RangeTitle.Interior.Color = Color.FromArgb(79, 129, 189);
                RangeTitle.Font.Color = Color.White;
                RangeTitle.Merge();
                RangeTitle.Value = "MSO905 Opcion 3 Create a Journal - Ellipse 8 Loader";
                RangeTitle.WrapText = true;
                RangeTitle.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                var rangePeriodo = excelSheet.Range["A3"];
                rangePeriodo.Value = "Accounting Period (MM/YY) *";
                rangePeriodo.Interior.Color = Color.FromArgb(79, 129, 189);
                rangePeriodo.Font.Color = Color.White;
                rangePeriodo.Font.Bold = true;

                rangePeriodo = excelSheet.Range["B3"];
                rangePeriodo.NumberFormat = "@";

                //Bordes
                excel.Range rangeItemBordes = null;
                rangeItemBordes = excelSheet.Range["A3:B3"];
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders.Color = Color.Black;

                rangePeriodo = excelSheet.Range["D3"];
                rangePeriodo.Value = "Journal         *";
                rangePeriodo = excelSheet.Range["D4"];
                rangePeriodo.Value = "Jnl Type";
                rangePeriodo = excelSheet.Range["D5"];
                rangePeriodo.Value = "Description     *";
                rangePeriodo = excelSheet.get_Range("D6");
                rangePeriodo.Value = "Accrual Journal *";

                rangePeriodo = excelSheet.get_Range("E3:E6");
                rangePeriodo.NumberFormat = "@";

                rangePeriodo = excelSheet.get_Range("D3:D6");
                rangePeriodo.Font.Bold = true;
                rangePeriodo.Interior.Color = Color.FromArgb(79, 129, 189);
                rangePeriodo.Font.Color = Color.White;

                //Bordes
                rangeItemBordes = excelSheet.get_Range("D3:E6");
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders.Color = Color.Black;

                //Rate
                rangePeriodo = excelSheet.get_Range("F6");
                rangePeriodo.Value = "Rate";
                rangePeriodo.Font.Bold = true;
                rangePeriodo.Interior.Color = Color.FromArgb(79, 129, 189);
                rangePeriodo.Font.Color = Color.White;
                rangePeriodo = excelSheet.get_Range("G6");
                rangePeriodo.NumberFormat = "@";
                //Rate Bordes
                rangeItemBordes = excelSheet.get_Range("F6:G6");
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeLeft].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeRight].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeTop].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders[excel.XlBordersIndex.xlEdgeBottom].LineStyle = excel.XlLineStyle.xlContinuous;
                rangeItemBordes.Borders.Color = Color.Black;

                var RangeItemTitle1 = excelSheet.get_Range("A8");
                RangeItemTitle1.Value = "Account Code";
                RangeItemTitle1 = excelSheet.get_Range("B8");
                RangeItemTitle1.Value = "W/Order Or Project";
                RangeItemTitle1 = excelSheet.get_Range("C8");
                RangeItemTitle1.Value = "W/P";
                RangeItemTitle1 = excelSheet.get_Range("D8");
                RangeItemTitle1.Value = "Journal Description";
                RangeItemTitle1 = excelSheet.get_Range("E8");
                RangeItemTitle1.Value = "Amount (+/-) Pesos";
                RangeItemTitle1 = excelSheet.get_Range("F8");
                RangeItemTitle1.Value = "Document Ref";
                RangeItemTitle1 = excelSheet.get_Range("G8");
                RangeItemTitle1.Value = "Foreign";
                RangeItemTitle1 = excelSheet.get_Range("H8");
                RangeItemTitle1.Value = "Dolars";
                RangeItemTitle1 = excelSheet.get_Range("I8");
                RangeItemTitle1.Value = "Message";
                RangeItemTitle1 = excelSheet.get_Range("J8");
                RangeItemTitle1.Value = "Message validate";

                RangeItemTitle1 = excelSheet.get_Range(beginColumn + dataRow + ":" + endColumn + "100000");
                RangeItemTitle1.NumberFormat = "@";

                excelSheetItems = excelSheet.ListObjects.AddEx(excel.XlListObjectSourceType.xlSrcRange,
                    excelSheet.get_Range(beginColumn + "8:" + messageEndColumn + "100000"),
                    XlListObjectHasHeaders: excel.XlYesNoGuess.xlYes);

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();

                cargarTasaCambio();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void executeMSO()
        {
            try
            {
                if (drpCreateJournalEnv.Label != null && !drpCreateJournalEnv.Label.Equals(""))
                {
                    if (elliseUser.Equals(""))
                    {
                        elliseUser = Settings.Default.UserDefault;
                        ellisePost = Settings.Default.PosDefault;
                        elliseDsct = Settings.Default.DstrDefault;
                    }

                    var frm = new FormAuthenticate(elliseUser, elliseDsct, ellisePost);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.ShowDialog();

                    if (!frm.Auth.Authenticated) return;
                    elliseUser = frm.Auth.Username;
                    ellisePost = frm.Auth.Position;
                    elliseDsct = frm.Auth.District;
                    ellisePswd = frm.Auth.Password;

                    if (excelSheet == null)
                    {
                        excelApp = Globals.ThisAddIn.Application;
                        excelBook = excelApp.Workbooks.Item[1];
                        excelSheet = (excel.Worksheet)excelBook.Sheets.Item[sheetName];
                    }

                    var url = "";
                    var error = "";

                    var ell = Util.GetEllipseConfiguration(Settings.Default.EllipseDirectory);

                    if (drpCreateJournalEnv.SelectedItem.Label.Equals("Productivo"))
                    {
                        url = ell.UrlProd;
                    }
                    else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Contingencia"))
                    {
                        url = ell.UrlCont;
                    }
                    else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Desarrollo"))
                    {
                        url = ell.UrlDesa;
                    }
                    else
                    {
                        url = ell.UrlTest;
                    }

                    var currentRow = dataRow;
                    string periodo = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.Range["B3"].Value));
                    string descPer = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.Range["E5"].Value));
                    string journalAsig =
                        MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.Range["E3"].Value));

                    string campoRequerido =
                        MyUtilities.formatearCeldaACadena(
                            Convert.ToString(excelSheet.Range[beginColumn + currentRow].Value));

                    if (campoRequerido.Equals("") || periodo.Equals(""))
                    {
                        MessageBox.Show(messageRequiredFields, messageTitle);
                    }
                    else
                    {
                        excelSheet.Select();
                        excelSheet.get_Range(endColumn + currentRow).Select();
                        excelSheet.get_Range(endColumn + currentRow).Value = messageProcess;

                        var screen = new Ellipse(elliseUser, ellisePswd, ellisePost, elliseDsct, url);
                        screen.InitConexion();

                        if (!screen.GetMSOError().Equals(""))
                        {
                            throw new Exception(screen.GetMSOError());
                        }

                        screen.ExecuteScreen("MSO905", "MSM905A");
                        if (!screen.GetMSOError().Equals(""))
                        {
                            throw new Exception(screen.GetMSOError());
                        }

                        try
                        {
                            /* PANTALLA MSM905A*/
                            if (screen.IsScreenNameCorrect("MSM905A"))
                            {
                                excelSheet.get_Range(endColumn + currentRow).Select();
                                excelSheet.get_Range(endColumn + currentRow).Value = "";

                                screen.InitScreenFields();
                                screen.SetMSOFieldValue("OPTION1I", "3");
                                screen.SetMSOFieldValue("ACCT_PERIOD1I", periodo);
                                screen.SetMSOFieldValue("FOREIGN_IND1I", "Y");
                                screen.SetMSOFieldValue("MAN_JNL_NO1I", journalAsig);

                                screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                if (screen.IsMSOMessage())
                                {
                                    error = screen.GetMSOMessage();
                                    screen.ExecuteMSO(Ellipse.F3_KEY, false);
                                    throw new Exception(error.Trim());
                                }

                                //Procesa en caso de no presentar error                                
                                // GRILLA DE DATOS
                                string journalDescVal =
                                    MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E5").Value));
                                string journalTypeVal =
                                    MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E4").Value));
                                string reversAutoVal =
                                    MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E6").Value));

                                if (screen.IsScreenNameCorrect("MSM907A"))
                                {
                                    screen.InitScreenFields();
                                    screen.SetMSOFieldValue("JOURNAL_DESC1I", journalDescVal);
                                    screen.SetMSOFieldValue("JOURNAL_TYPE1I", journalTypeVal);
                                    screen.SetMSOFieldValue("ACCOUNTANT1I", elliseUser);
                                    screen.SetMSOFieldValue("ACCRUAL_IND1I", reversAutoVal);
                                    screen.SetMSOFieldValue("APPROVAL_STAT1I", "Y");

                                    var MyValuesGrilla = (Array)excelSheet.get_Range("A" + currentRow, "H" + currentRow).Cells.Value;

                                    var i = 1;
                                    var fin = true;
                                    string messageEnd = MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range(endColumn + currentRow).Value));

                                    while (!MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 1))).Equals(""))
                                    {
                                        fin = false;
                                        //Primer lote de registros                                            
                                        excelSheet.Range[endColumn + currentRow].Select();
                                        excelSheet.Range[endColumn + currentRow].Value = "Read";
                                        screen.SetMSOFieldValue("ACCOUNT_CODE1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 1))));
                                        screen.SetMSOFieldValue("WORK_PROJ1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 2))).Length >= 8 ? MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 2))).Substring(0, 8).Trim(): MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 2))));
                                        screen.SetMSOFieldValue("WORK_PROJ_IND1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 3))).Length >= 1? MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 3))).Substring(0, 1).Trim(): MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 3))));
                                        screen.SetMSOFieldValue("JNL_DESC1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 4))).Length >= 40? MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 4))).Substring(0, 40).Trim(): MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 4))));
                                        screen.SetMSOFieldValue("TRAN_AMOUNT1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 8))));screen.SetMSOFieldValue("DOCUMENT_REF1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 6))).Length >= 8? MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 6))).Substring(0, 8).Trim(): MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 6))));
                                        screen.SetMSOFieldValue("FOREIGN_CURR1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 7))));screen.SetMSOFieldValue("MEMO_AMOUNT1I" + i,MyUtilities.formatearCeldaACadena(Convert.ToString(MyValuesGrilla.GetValue(1, 5))));

                                        if (i == 3)
                                        {
                                            screen.ExecuteMSO(Ellipse.TRANSMIT, true);
                                            if (screen.IsMSOError())
                                            {
                                                error = screen.GetMSOMessage();
                                                throw new Exception(error.Trim());
                                            }
                                            screen.InitScreenFields();
                                            screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                            i = 1;
                                            fin = true;
                                        }
                                        else
                                        {
                                            i++;
                                        }
                                        currentRow++;
                                        MyValuesGrilla =(Array)excelSheet.get_Range("A" + currentRow, "H" + currentRow).Cells.Value;
                                    }

                                    currentRow--;
                                    if (fin)
                                    {
                                        screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                        var mensaje = getMessage(screen.MSO.MessageInfomations);
                                        mensaje = mensaje.Trim();
                                        if (!mensaje.Equals(""))
                                        {
                                            excelSheet.get_Range(endColumn + currentRow).Value =
                                                getMessage(screen.MSO.MessageInfomations).Trim();
                                        }
                                        else
                                        {
                                            if (screen.GetMSOInformation().Contains(messageWebWarn2074))
                                            {
                                                screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                                if (screen.GetMSOWarning().Contains(messageWebJournalNumber))
                                                {
                                                    excelSheet.get_Range(endColumn + currentRow).Value =
                                                        getMessage(screen.MSO.MessageInfomations).Trim();
                                                }
                                                else if (screen.IsMSOMessage())
                                                {
                                                    excelSheet.get_Range(endColumn + currentRow).Value =
                                                        screen.GetMSOError().Trim();
                                                }
                                                else
                                                {
                                                    if (screen.IsScreenNameCorrect("MSM905A"))
                                                    {
                                                        excelSheet.get_Range(endColumn + currentRow).Value =
                                                            messageUploaded;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                excelSheet.get_Range(endColumn + currentRow).Value =
                                                    screen.GetMSOError().Trim();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        screen.ExecuteMSO(Ellipse.TRANSMIT, true);

                                        if (screen.IsMSOMessage())
                                        {
                                            if (screen.IsMSOInformation())
                                            {
                                                screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                                if (screen.GetMSOInformation().Contains(messageWebJournalNumber))
                                                {
                                                    excelSheet.get_Range(endColumn + currentRow).Value =
                                                        getMessage(screen.MSO.MessageInfomations).Trim();
                                                }
                                                else if (screen.GetMSOWarning().Contains(messageWebWarn2074))
                                                {
                                                    screen.ExecuteMSO(Ellipse.TRANSMIT, false);
                                                    if (screen.GetMSOInformation().Contains(messageWebJournalNumber))
                                                    {
                                                        excelSheet.get_Range(endColumn + currentRow).Value =
                                                            getMessage(screen.MSO.MessageInfomations).Trim();
                                                    }
                                                    else if (screen.IsMSOMessage())
                                                    {
                                                        excelSheet.get_Range(endColumn + currentRow).Value =
                                                            screen.GetMSOError().Trim();
                                                    }
                                                    else
                                                    {
                                                        if (screen.IsScreenNameCorrect("MSM905A"))
                                                        {
                                                            excelSheet.get_Range(endColumn + currentRow).Value =
                                                                messageUploaded;
                                                        }
                                                    }
                                                }
                                                else if (screen.IsMSOMessage())
                                                {
                                                    error = screen.GetMSOMessage();
                                                    throw new Exception(error.Trim());
                                                }
                                                else
                                                {
                                                    if (screen.IsScreenNameCorrect("MSM905A"))
                                                    {
                                                        excelSheet.get_Range(endColumn + currentRow).Value =
                                                            messageUploaded;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                error = screen.GetMSOMessage();
                                                throw new Exception(error.Trim());
                                            }
                                        }
                                    }
                                    //Fin - Valida si el mensaje esta vacio para continuar.
                                }
                            }
                        }
                        catch (Exception errorCatch)
                        {
                            excelSheet.get_Range(endColumn + currentRow).Value = errorCatch.Message;
                        }
                        finally
                        {
                            screen.ExecuteMSO(Ellipse.F3_KEY, false);
                            currentRow++;
                            campoRequerido =
                                MyUtilities.formatearCeldaACadena(
                                    Convert.ToString(excelSheet.get_Range(beginColumn + currentRow).Value));
                        }

                        excelSheet.Cells.Columns.AutoFit();
                        excelSheet.Cells.Rows.AutoFit();
                        MessageBox.Show(messageProcessFinished, messageTitle, MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
                else
                {
                    var messageBox = MessageBox.Show("Please Select a Env. Option", messageTitleError,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception errorCatch)
            {
                excelSheet.get_Range(endColumn + dataRow).Value = "";
                if (errorAuthUser.Equals(errorCatch.Message.ToUpper()))
                    MessageBox.Show(messageUserPassincorrect, messageTitleError, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                else
                    MessageBox.Show(
                        "\n\rMessage:" + errorCatch.Message + "\n\rSource:" + errorCatch.Source + "\n\rStackTrace:" +
                        errorCatch.StackTrace, messageTitleError, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string getMessage(List<Message> MessageInfomations)
        {
            var message = "";
            foreach (var m in MessageInfomations)
            {
                message += m.Text;
            }
            return message;
        }

        private void cargarTasaCambio()
        {
            OdbcConnection conn = null;

            try
            {
                if (excelSheet == null)
                {
                    excelApp = Globals.ThisAddIn.Application;
                    excelBook = excelApp.Workbooks.Item[1];
                    excelSheet = (excel.Worksheet)excelBook.Sheets.Item[sheetName];
                }
                excelSheet.get_Range("G6").Value = "Loading...";

                OdbcCommand cmd = null;
                OdbcDataReader reader = null;
                string querySQL;

                conn = new OdbcConnection();
                if (drpCreateJournalEnv.SelectedItem.Label.Equals("Productivo"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseProdConnectionString"].ConnectionString;
                }
                else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Contingencia"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseCtgConnectionString"].ConnectionString;
                }
                else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Desarrollo"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseDesaConnectionString"].ConnectionString;
                }
                else
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseTestConnectionString"].ConnectionString;
                }

                conn.Open();

                querySQL = "SELECT ELLIPSE.GET_TASA_CONVERSION('USD', to_char(sysdate,'YYYYMMDD'), 'PES') TASA" +
                           " FROM DUAL;";

                cmd = new OdbcCommand(querySQL, conn);
                reader = cmd.ExecuteReader();
                var valor = "";

                while (reader.Read())
                {
                    valor = reader.GetValue(0).ToString().Trim();
                    excelSheet.get_Range("G6").Value = reader.GetValue(0).ToString().Trim();
                }

                reader.Close();
                cmd.Dispose();
            }
            catch (Exception error)
            {
                if (error.Message.Contains("No se encuentra el nombre del origen de datos"))
                    MessageBox.Show(messageODBC, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show(
                        "\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" +
                        error.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                conn.Close();
            }
        }

        private void convertirPesos()
        {
            try
            {
                if (excelSheet == null)
                {
                    excelApp = Globals.ThisAddIn.Application;
                    excelBook = excelApp.Workbooks.Item[1];
                    excelSheet = (excel.Worksheet)excelBook.Sheets.Item[sheetName];
                }

                var x = "";
                double dolar, pesos, tasa;
                var currentRow = dataRow;

                x = MyUtilities.formatearCeldaACadena(excelSheet.get_Range(beginColumn + currentRow).Value);
                tasa = MyUtilities.formatearCeldaADouble(excelSheet.get_Range("G6").Value);
                while (!"".Equals(x))
                {
                    if (
                        "".Equals(
                            MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("H" + currentRow).Value))))
                    {
                        dolar =
                            MyUtilities.formatearCeldaADouble(Convert.ToString(excelSheet.get_Range("E" + currentRow).Value));
                        pesos = Math.Round(dolar / tasa, 2);
                        excelSheet.get_Range("H" + currentRow).Value = pesos;
                    }
                    currentRow++;
                    x = MyUtilities.formatearCeldaACadena(excelSheet.get_Range(beginColumn + currentRow).Value);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(
                    "\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" + error.StackTrace,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void convertirDolares()
        {
            try
            {
                if (excelSheet == null)
                {
                    excelApp = Globals.ThisAddIn.Application;
                    excelBook = excelApp.Workbooks.Item[1];
                    excelSheet = (excel.Worksheet)excelBook.Sheets.Item[sheetName];
                }

                var x = "";
                double dolar, pesos, tasa;
                var currentRow = dataRow;

                x = MyUtilities.formatearCeldaACadena(excelSheet.get_Range(beginColumn + currentRow).Value);
                tasa = MyUtilities.formatearCeldaADouble(excelSheet.get_Range("G6").Value);
                while (!"".Equals(x))
                {
                    if (
                        "".Equals(
                            MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("E" + currentRow).Value))))
                    {
                        dolar =
                            MyUtilities.formatearCeldaADouble(Convert.ToString(excelSheet.get_Range("H" + currentRow).Value));
                        pesos = Math.Round(dolar * tasa, 2);
                        excelSheet.get_Range("E" + currentRow).Value = pesos;
                    }
                    currentRow++;
                    x = MyUtilities.formatearCeldaACadena(excelSheet.get_Range(beginColumn + currentRow).Value);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(
                    "\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" + error.StackTrace,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDolares_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Convertir pesos a dolares
                convertirPesos();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnPesos_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Convertir pesos a dolares
                convertirDolares();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void btnValidateNit_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Validar NIT
                validarNit();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show(error.Message);
            }
        }

        private void validarNit()
        {
            OdbcConnection conn = null;

            try
            {
                if (excelSheet == null)
                {
                    excelApp = Globals.ThisAddIn.Application;
                    excelBook = excelApp.Workbooks.Item[1];
                    excelSheet = (excel.Worksheet)excelBook.Sheets.Item[sheetName];
                }

                OdbcCommand cmd = null;
                OdbcDataReader reader = null;
                string querySQL;

                conn = new OdbcConnection();
                if (drpCreateJournalEnv.SelectedItem.Label.Equals("Productivo"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseProdConnectionString"].ConnectionString;
                }
                else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Contingencia"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseCtgConnectionString"].ConnectionString;
                }
                else if (drpCreateJournalEnv.SelectedItem.Label.Equals("Desarrollo"))
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseDesaConnectionString"].ConnectionString;
                }
                else
                {
                    conn.ConnectionString =
                        ConfigurationManager.ConnectionStrings["EllipseTestConnectionString"].ConnectionString;
                }

                conn.Open();
                excelSheet.Select();

                var currentRow = dataRow;
                string descripcion =
                    MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("D" + currentRow).Value));
                var nit = "";
                var expresion = @"(#)(\d+\-\d{1}|\d+)(\W|_)";

                while (!"".Equals(descripcion))
                {
                    excelSheet.get_Range(messageEndColumn + currentRow).Select();
                    excelSheet.get_Range(messageEndColumn + currentRow).Value = messageProcess;
                    descripcion = Regex.Match(descripcion, expresion).Value;
                    if (!"".Equals(descripcion))
                    {
                        nit = descripcion.Substring(1, descripcion.Length - 2);
                        querySQL = " SELECT NIT" +
                                   " FROM (select tax_file_no nit" +
                                   " from ellipse.msf203" +
                                   " where TRIM(tax_file_no) = '" + nit + "'" +
                                   " Union" +
                                   " select GOVT_ID_NO nit" +
                                   " from ellipse.msf503" +
                                   " where TRIM(GOVT_ID_NO) = '" + nit + "'" +
                                   " Union" +
                                   " select TABLE_CODE nit" +
                                   " from ellipse.msf010" +
                                   " where table_type = '+NIT' AND TRIM(TABLE_CODE) = '" + nit + "');";
                        cmd = new OdbcCommand(querySQL, conn);
                        reader = cmd.ExecuteReader();
                    }
                    if (reader != null)
                    {
                        if (reader.HasRows)
                            excelSheet.get_Range(messageEndColumn + currentRow).Value = "Ok";
                        else
                            excelSheet.get_Range(messageEndColumn + currentRow).Value = "No existe en ellipse";

                        reader.Close();
                        reader = null;
                    }
                    else
                    {
                        excelSheet.get_Range(messageEndColumn + currentRow).Value = "No existe en ellipse";
                    }

                    currentRow++;
                    descripcion =
                        MyUtilities.formatearCeldaACadena(Convert.ToString(excelSheet.get_Range("D" + currentRow).Value));
                }

                if (reader != null)
                    reader.Close();
                if (cmd != null)
                    cmd.Dispose();

                excelSheet.Cells.Columns.AutoFit();
                excelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                if (error.Message.Contains("No se encuentra el nombre del origen de datos"))
                    MessageBox.Show(messageODBC, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show(
                        "\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" +
                        error.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }

        private void btnValidateAccountCode_Click(object sender, RibbonControlEventArgs e)
        {
        }
    }
}
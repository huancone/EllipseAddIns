using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Globalization;
using System.Web.Services.Ellipse.Post;
using System.Threading;

namespace EllipseMSE140DeleteExcelAddIn
{
    public partial class RibbonEllipse
    {
        public static string ElliseUser = "";
        public static string EllisePswd = "";
        public static string EllisePost = "";
        public static string ElliseDsct = "";
        private static string SheetName = "MSE140D";
        private static string HeaderRow = "2";
        private static string BeginColumn = "A";
        private static string EndColumn = "I";
        private static string MessageColumn = "J";
        private static int NumRows = 1000;
        private static bool InExecution = false;

        Excel.Application ExcelApp;
        Excel.Workbook ExcelBook;
        Excel.Worksheet ExcelSheet;
        Excel.ListObject ExcelSheetItems;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnEllipseFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ExcelApp = Globals.ThisAddIn.Application;
                ExcelBook = ExcelApp.Workbooks.Add();
                ExcelSheet = ExcelBook.Sheets.Add();

                ExcelSheet.Name = SheetName;

                Excel.Range RangeTitle = ExcelSheet.get_Range(BeginColumn + "1:" + MessageColumn + "1");
                RangeTitle.Merge();
                RangeTitle.Interior.Color = Color.RoyalBlue;
                RangeTitle.Font.Color = Color.LavenderBlush;
                RangeTitle.Font.Size = 12;
                RangeTitle.WrapText = true;
                RangeTitle.Font.FontStyle = "Bold";
                RangeTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                RangeTitle.Value = "Loader Delete Item In Requisition (MSE140) ELLIPSE";

                Excel.Range RangeItemTitle = null;
                RangeItemTitle = ExcelSheet.get_Range("A" + HeaderRow);
                RangeItemTitle.Value = "District";
                RangeItemTitle = ExcelSheet.get_Range("B" + HeaderRow);
                RangeItemTitle.Value = "Requesition";
                RangeItemTitle = ExcelSheet.get_Range("C" + HeaderRow);
                RangeItemTitle.Value = "Item";
                RangeItemTitle = ExcelSheet.get_Range("D" + HeaderRow);
                RangeItemTitle.Value = "Warehouse";
                RangeItemTitle = ExcelSheet.get_Range("E" + HeaderRow);
                RangeItemTitle.Value = "Req. Type";
                RangeItemTitle = ExcelSheet.get_Range("F" + HeaderRow);
                RangeItemTitle.Value = "Item Type";
                RangeItemTitle = ExcelSheet.get_Range("G" + HeaderRow);
                RangeItemTitle.Value = "Stock Code";
                RangeItemTitle = ExcelSheet.get_Range("H" + HeaderRow);
                RangeItemTitle.Value = "Qty. Required";
                RangeItemTitle = ExcelSheet.get_Range("I" + HeaderRow);
                RangeItemTitle.Value = "Unit Of Measure ";
                RangeItemTitle = ExcelSheet.get_Range("J" + HeaderRow);
                RangeItemTitle.Value = "Message";

                int ContentRowBegin = int.Parse(HeaderRow) + 1;

                ExcelSheet.get_Range(BeginColumn + ContentRowBegin.ToString() + ":" + EndColumn + NumRows.ToString()).NumberFormat = "@";

                ExcelSheetItems = ExcelSheet.ListObjects.AddEx(SourceType: Excel.XlListObjectSourceType.xlSrcRange, Source: ExcelSheet.get_Range(BeginColumn + HeaderRow + ":" + MessageColumn + NumRows.ToString()), XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes);

                ExcelSheet.Cells.Columns.AutoFit();
                ExcelSheet.Cells.Rows.AutoFit();
            }
            catch (Exception error)
            {
                var messageBox = MessageBox.Show("\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" + error.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Execute() {
            InExecution = true;
            if (drpEllipseEnv.SelectedItem.Label != null && !drpEllipseEnv.SelectedItem.Label.Equals(""))
            {
                if (ElliseUser.Equals(""))
                {
                    ElliseUser = AppConfiguration.GetConfiguration("UserDefault");
                    EllisePost = AppConfiguration.GetConfiguration("PosDefault");
                    ElliseDsct = AppConfiguration.GetConfiguration("DstrDefault");
                }
                System.Web.Services.Ellipse.FormAuthenticate frm = new System.Web.Services.Ellipse.FormAuthenticate(ElliseUser, ElliseDsct, EllisePost);
                frm.StartPosition = FormStartPosition.CenterScreen;
                frm.ShowDialog();
                if (frm.Auth.Authenticated)
                {
                    ElliseUser = frm.Auth.Username;
                    EllisePost = frm.Auth.Position;
                    ElliseDsct = frm.Auth.District;
                    EllisePswd = frm.Auth.Password;

                    try
                    {
                        if (ExcelSheet == null)
                        {
                            ExcelApp = Globals.ThisAddIn.Application;
                            ExcelBook = ExcelApp.Workbooks.Item[1];
                            ExcelSheet = (Excel.Worksheet) ExcelBook.Sheets.Item[SheetName];
                        }

                        string url = "";

                        EllipseConfiguration conf = System.Web.Services.Ellipse.Post.Util.GetEllipseConfiguration(AppConfiguration.GetConfiguration("EllipseDirectory"));
                        if (drpEllipseEnv.SelectedItem.Label.Equals("Productivo"))
                        {
                            url = conf.UrlProd;
                        }
                        else if (drpEllipseEnv.SelectedItem.Label.Equals("Contingencia"))
                        {
                            url = conf.UrlCont;
                        }
                        else if (drpEllipseEnv.SelectedItem.Label.Equals("Desarrollo"))
                        {
                            url = conf.UrlDesa;
                        }
                        else
                        {
                            url = conf.UrlTest;
                        }

                        int index = int.Parse(HeaderRow) + 1;
                        PostService proxy = new PostService(ElliseUser, EllisePswd, EllisePost, ElliseDsct, url);
                        ResponseDTO resp = proxy.InitConexion();
                        if (!resp.GotErrorMessages())
                        {
                            while (ExcelSheet.get_Range(BeginColumn + index.ToString()).Value != null)
                            {
                                ExcelSheet.get_Range(MessageColumn + index.ToString()).Value = "Procesando...";
                                System.Array MyValues = (System.Array)ExcelSheet.get_Range(BeginColumn + index.ToString(), EndColumn + index.ToString()).Cells.Value;
                                try
                                {
                                    Requisition req = new Requisition(MyValues);
                                    ExcelSheet.get_Range(MessageColumn + index.ToString()).Value = DeleteItemRequisitionPost(proxy, req);
                                }
                                catch (Exception error)
                                {
                                    ExcelSheet.get_Range(MessageColumn + index.ToString()).Value = error.Message;
                                }
                                index++;
                            }
                        }
                        else
                        {
                            string error = "";
                            foreach (System.Web.Services.Ellipse.Post.Message msg in resp.Errors)
                            {
                                error += msg.Text;
                            }
                            var messageBox = MessageBox.Show(error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        ExcelSheet.Cells.Columns.AutoFit();
                        ExcelSheet.Cells.Rows.AutoFit();
                    }
                    catch (Exception error)
                    {
                        var messageBox = MessageBox.Show("\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" + error.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    var messageBox = MessageBox.Show("User not Authenticated!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                var messageBox = MessageBox.Show("Please Select a Env. Option", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            InExecution = false;
        }

        private void btnEllipseExecute_Click(object sender, RibbonControlEventArgs e)
        {
            if (!InExecution)
            {
                try
                {
                    Thread thread = new Thread(new ThreadStart(Execute));
                    thread.Start();
                }
                catch (Exception)
                {
                }
                finally
                {
                    InExecution = false;
                }
            }
            else
            {
                var messageBox = MessageBox.Show("Execution In Progress...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string DeleteItemRequisitionPost(PostService proxy, Requisition req)
        {
            string message = "";
            Dictionary<string, string> mapReponse = new Dictionary<string, string>();
            ResponseDTO resp = null;
            if (req.Warehouse != null)
            {
                req.Warehouse = req.Warehouse.Trim();
            }
            StringBuilder requestXML = new StringBuilder("");
            requestXML.Append("<interaction>");
            requestXML.Append("<actions>");
            requestXML.Append("<action>");
            requestXML.Append("<name>service</name>");
            requestXML.Append("<data>");
            requestXML.Append("<name>com.mincom.enterpriseservice.ellipse.requisition.RequisitionService</name>");
            requestXML.Append("<operation>multipleDeleteItem</operation>");
            requestXML.Append("<returnWarnings>false</returnWarnings>");
            requestXML.Append("<dto>");
            requestXML.Append("<dto uuid=\"");
            requestXML.Append(System.Web.Services.Ellipse.Post.Util.GetNewOperationId());
            requestXML.Append("\" deleted=\"true\" modified=\"false\">");
            requestXML.Append("<gPASelected>false</gPASelected>");
            requestXML.Append("<iRequisitionItemType>C</iRequisitionItemType>");
            requestXML.Append("<activityCounter>000</activityCounter>");
            requestXML.Append("<alterStockCodeFlg>false</alterStockCodeFlg>");
            requestXML.Append("<directOrderQuantity>0</directOrderQuantity>");
            requestXML.Append("<directOrderReceived>0</directOrderReceived>");
            requestXML.Append("<directPurchOrdItem>0</directPurchOrdItem>");
            requestXML.Append("<estimatedPrice editable=\"false\" />");
            requestXML.Append("<issChangeReason>DES1</issChangeReason>");
            requestXML.Append("<issueDistrictCode>");
            requestXML.Append(req.District);
            requestXML.Append("</issueDistrictCode>");
            requestXML.Append("<issueDocoFlg>false</issueDocoFlg>");
            requestXML.Append("<issueRequisitionItem>");
            requestXML.Append(req.Item);
            requestXML.Append("</issueRequisitionItem>");
            requestXML.Append("<issueWarehouseId>");
            requestXML.Append(req.Warehouse);
            requestXML.Append("</issueWarehouseId>");
            requestXML.Append("<itemType>");
            requestXML.Append(req.ItemType);
            requestXML.Append("</itemType>");
            requestXML.Append("<leadTimeComp>true</leadTimeComp>");
            requestXML.Append("<moreDescExists>false</moreDescExists>");
            requestXML.Append("<narrativeExists>N</narrativeExists>");
            requestXML.Append("<partIssue>N</partIssue>");
            requestXML.Append("<stockCode>");
            requestXML.Append(req.StockCode);
            requestXML.Append("</stockCode>");
            requestXML.Append("<quantityRequired>");
            requestXML.Append(req.QtyReq);
            requestXML.Append("</quantityRequired>");
            requestXML.Append("<unitOfMeasure>");
            requestXML.Append(req.UnitOfMeasure);
            requestXML.Append("</unitOfMeasure>");
            requestXML.Append("<districtCode>");
            requestXML.Append(req.District);
            requestXML.Append("</districtCode>");
            requestXML.Append("<ireqNo>");
            requestXML.Append(req.Requesition);
            requestXML.Append("</ireqNo>");
            requestXML.Append("<ireqType>");
            requestXML.Append(req.ReqType);
            requestXML.Append("</ireqType>");
            requestXML.Append("</dto>");
            requestXML.Append("</dto>");
            requestXML.Append("</data>");
            requestXML.Append("<id>");
            requestXML.Append(System.Web.Services.Ellipse.Post.Util.GetNewOperationId());
            requestXML.Append("</id>");
            requestXML.Append("</action>");
            requestXML.Append("</actions>");
            requestXML.Append("<chains />");
            requestXML.Append("<connectionId>");
            requestXML.Append(proxy.ConnectionId);
            requestXML.Append("</connectionId>");
            requestXML.Append("<application>mse140</application>");
            requestXML.Append("<applicationPage>read</applicationPage>");
            requestXML.Append("<transaction>true</transaction>");
            requestXML.Append("</interaction>");
            resp = proxy.ExecutePostRequest(requestXML.ToString());
            if (resp.GotErrorMessages())
            {
                foreach (System.Web.Services.Ellipse.Post.Message msg in resp.Errors)
                {
                    message += msg.Field + " " + msg.Text;
                }
            }
            if (!message.Equals(""))
            {
                throw new Exception(message);
            }
            message = "Requisition Marked To Be Delete.";
            return message;
        }

        private void RetrieveData()
        {
            InExecution = true;
            if (drpEllipseEnv.SelectedItem.Label != null && !drpEllipseEnv.SelectedItem.Label.Equals(""))
            {
                System.Data.Odbc.OdbcConnection conn = null;
                System.Data.Odbc.OdbcCommand cmd = null;
                System.Data.Odbc.OdbcDataReader reader = null;
                try
                {
                    string EllipseDS = "";
                    if (ExcelSheet == null)
                    {
                        ExcelApp = Globals.ThisAddIn.Application;
                        ExcelBook = ExcelApp.Workbooks.Item[1];
                        ExcelSheet = (Excel.Worksheet) ExcelBook.Sheets.Item[SheetName];
                    }

                    if (drpEllipseEnv.SelectedItem.Label.Equals("Productivo"))
                    {
                        EllipseDS = AppConfiguration.GetConfiguration("EllipseDSProd");
                    }
                    else if (drpEllipseEnv.SelectedItem.Label.Equals("Contingencia"))
                    {
                        EllipseDS = AppConfiguration.GetConfiguration("EllipseDSProd");
                    }
                    else if (drpEllipseEnv.SelectedItem.Label.Equals("Desarrollo"))
                    {
                        EllipseDS = AppConfiguration.GetConfiguration("EllipseDSDesa");
                    }
                    else
                    {
                        EllipseDS = AppConfiguration.GetConfiguration("EllipseDSTest");
                    }

                    conn = new System.Data.Odbc.OdbcConnection();
                    conn.ConnectionString = EllipseDS;
                    conn.Open();

                    int Index = int.Parse(HeaderRow) + 1;
                    while (ExcelSheet.get_Range(BeginColumn + Index.ToString()).Value != null)
                    {
                        try
                        {
                            ExcelSheet.get_Range(MessageColumn + Index.ToString()).Value = "Procesando...";
                            System.Array MyValues = (System.Array)ExcelSheet.get_Range(BeginColumn + Index.ToString(), EndColumn + Index.ToString()).Cells.Value;
                            Requisition req = new Requisition(MyValues);
                            cmd = new System.Data.Odbc.OdbcCommand("SELECT MSF141.QTY_REQ, MSF141.STOCK_CODE, MSF141.WHOUSE_ID, MSF140.IREQ_TYPE, MSF100.UNIT_OF_ISSUE FROM ELLIPSE.MSF140 INNER JOIN ELLIPSE.MSF141 ON MSF140.DSTRCT_CODE = MSF141.DSTRCT_CODE AND MSF140.IREQ_NO = MSF141.IREQ_NO INNER JOIN ELLIPSE.MSF100 ON MSF141.STOCK_CODE = MSF100.STOCK_CODE WHERE MSF141.DSTRCT_CODE = '" + req.District + "' AND MSF141.IREQ_NO = RPAD('" + req.Requesition + "', 6, ' ') AND MSF141.IREQ_ITEM = LPAD('" + req.Item + "', 4, '0')", conn);
                            reader = cmd.ExecuteReader();
                            if (reader.Read())
                            {
                                ExcelSheet.get_Range("D" + Index).Value = reader["WHOUSE_ID"].ToString();
                                ExcelSheet.get_Range("E" + Index).Value = reader["IREQ_TYPE"].ToString();
                                ExcelSheet.get_Range("F" + Index).Value = "S";
                                ExcelSheet.get_Range("G" + Index).Value = reader["STOCK_CODE"].ToString();
                                ExcelSheet.get_Range("H" + Index).Value = reader["QTY_REQ"].ToString();
                                ExcelSheet.get_Range("I" + Index).Value = reader["UNIT_OF_ISSUE"].ToString();
                            }
                            ExcelSheet.get_Range(MessageColumn + Index.ToString()).Value = "";
                            reader.Close();
                            cmd.Dispose();
                        }
                        catch (Exception err)
                        {
                            ExcelSheet.get_Range(MessageColumn + Index.ToString()).Value = err.Message;
                        }
                        Index++;
                    }
                }
                catch (Exception error)
                {
                    var messageBox = MessageBox.Show("\n\rMessage:" + error.Message + "\n\rSource:" + error.Source + "\n\rStackTrace:" + error.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                }
            }
            else
            {
                var messageBox = MessageBox.Show("Please Select a Env. Option", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            InExecution = false;
        }

        private void btnLoadData_Click(object sender, RibbonControlEventArgs e)
        {
            if (!InExecution)
            {
                try
                {
                    Thread thread = new Thread(new ThreadStart(RetrieveData));
                    thread.Start();
                }
                catch (Exception)
                {
                }
                finally
                {
                    InExecution = false;
                }
            }
            else
            {
                var messageBox = MessageBox.Show("Execution In Progress...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

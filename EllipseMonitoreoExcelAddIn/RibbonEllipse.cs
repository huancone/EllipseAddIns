using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMonitoreoExcelAddIn.CondMeasurementService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse;
using VarEncript = SharedClassLibrary.Utilities.Encryption;

using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
//using Authenticator = EllipseMonitoreoExcelAddIn.AuthenticatorService;
namespace EllipseMonitoreoExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells Cells;
        EllipseFunctions EFunctions = new EllipseFunctions();
        FormAuthenticate frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;
        String SheetName01 = "Monitoreo";
        String ColHeader = "AK";
        String ColFinal = "AL";
        int ColFin = 38;
        String ColOcultar = "AG1";
        int RowCabezera = 7;
        int RowInicial = 8;
        int maxRow = 10000;

        //Variables de Conexion 
        private string SQL;
        private string DataBase;
        private string User; //Ej. SIGCON, CONSULBO
        private string Pw;
        // ReSharper disable once InconsistentNaming
        public string DbLink; //Ej. @DBLMIMS

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            List<string> enviroments = new List<string>();
            enviroments.Add("Productivo");
            enviroments.Add("Productivox");
            enviroments.Add("Test");
            enviroments.Add("Desarrollo");
            enviroments.Add("Contingencia");
            enviroments.Add("EL9CONV");
            //var enviroments = Environments.GetEnvironmentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }


        public bool ConexionDataBase(string enviroments)
        {
            if (enviroments == "Productivox")
            {
                //Sigman
                DataBase = VarEncript.Encryption.Decrypt("CrOkubls0sZ8lj8iUOR+QY18P9jBSp7MV17Q1hMCt0zpW2WGmMHYV5XXc8j/FdQQNSMJhAHs3GXzbxU0zB+CNt5K1PIiJBvP7RlVJqPn+vHh1mLdhaACGMniPn234d2s");
                User = VarEncript.Encryption.Decrypt("x4yNNf5qsgLpNdA1xUaBM1GaKhwrINqfzNsmDA7rZmZWVx8308y12p1zvsIuEzx+yszVVnhqhQ1cFWL+lBB8yYb53Yx1kBkvdWcXspKfG8buz4RuwCjtXcXkvGOQwdzw");
                Pw = VarEncript.Encryption.Decrypt("M8/fjRkEAGaWFKtzyECz8mlJluF8xZevecMTrJ8tf0uboneZPAzICqYYB1WWx23w6sF5AXHDY3MtMZNJVGJ1ALO2D76lFq0M9fLmnU8Q8aOYcANWnlQCQzpX/EqnO8Ow");
                DbLink = "";
            }
            else if (enviroments == "Test")
            {
                //ELLIPSE TEST
                DataBase = VarEncript.Encryption.Decrypt("ZmuwRdpIqQDXlgbVCTMap4/2rae2TEeElYs0dwdKpLs79OD+0DB5C7PU+YfACBxYW8/EhT71lw+UWXMm0dHrecEAbgruixwRCesj/kZdhcqMKnJmfHjYVx/kzfnBZ+ff");
                User = VarEncript.Encryption.Decrypt("KDWGvC9euLoOV0/ut9uidRLNVNu29uqivJHh717JZUlB37WRHYrqg71B99RW6YbpI/8cikLxMoaFp+phLJxiMQdwWx7LxjgztWhi9FlYUbgqLyYzYn1pnSpSXBfCfWRG");
                Pw = VarEncript.Encryption.Decrypt("M8/fjRkEAGaWFKtzyECz8mlJluF8xZevecMTrJ8tf0uboneZPAzICqYYB1WWx23w6sF5AXHDY3MtMZNJVGJ1ALO2D76lFq0M9fLmnU8Q8aOYcANWnlQCQzpX/EqnO8Ow");
                DbLink = "";
            }
            else if (enviroments == "Desarrollo")
            {
                //Ellipse Desarrollo
                DataBase = VarEncript.Encryption.Decrypt("1IKfU5uJXMSEmagte2It5Yo4RKspvU8kDY8JRRFZZ2EaEci7t5HhQ7KMsVFKx8WbfiCEHKAy6h6woQTNKC7cly4Nsjae4WCgI/BdHj8+47L3Ux2xZqVCSELXVqzEdZRN");
                User = VarEncript.Encryption.Decrypt("KDWGvC9euLoOV0/ut9uidRLNVNu29uqivJHh717JZUlB37WRHYrqg71B99RW6YbpI/8cikLxMoaFp+phLJxiMQdwWx7LxjgztWhi9FlYUbgqLyYzYn1pnSpSXBfCfWRG");
                Pw = VarEncript.Encryption.Decrypt("CnybQg6aRmqDpzwekCgGJkT58UpCIdmMt7br1TUhchrC0D+mG1z+pchSBUsXfklz1wBONoZoxtdLnKJ9T30PTvZzmCrbhE+MkmiN96CU3zORPXddVL6aPxysDNthpP3Z");
                DbLink = "";
            }
            else if (enviroments == "Contingencia")
            {
                //Ellipse Contingencia
                DataBase = VarEncript.Encryption.Decrypt("brw6hTk7tyzbWMnkgOAGm7T5ISbOxIDZzSuf/5nvKn94VsLindO9npazUR8CDo7/5YX0KUYHtN+VxayBURC3BPWpjIhFlX+hVWYxVGV3FBoO5gv6XYTiHcXupsZ5bm5S");
                User = VarEncript.Encryption.Decrypt("KDWGvC9euLoOV0/ut9uidRLNVNu29uqivJHh717JZUlB37WRHYrqg71B99RW6YbpI/8cikLxMoaFp+phLJxiMQdwWx7LxjgztWhi9FlYUbgqLyYzYn1pnSpSXBfCfWRG");
                Pw = VarEncript.Encryption.Decrypt("CnybQg6aRmqDpzwekCgGJkT58UpCIdmMt7br1TUhchrC0D+mG1z+pchSBUsXfklz1wBONoZoxtdLnKJ9T30PTvZzmCrbhE+MkmiN96CU3zORPXddVL6aPxysDNthpP3Z");
                DbLink = "";
            }
            else if (enviroments == "Productivo")
            {
                //Ellipse Productivo
                DataBase = VarEncript.Encryption.Decrypt("brw6hTk7tyzbWMnkgOAGm7T5ISbOxIDZzSuf/5nvKn94VsLindO9npazUR8CDo7/5YX0KUYHtN+VxayBURC3BPWpjIhFlX+hVWYxVGV3FBoO5gv6XYTiHcXupsZ5bm5S");
                User = VarEncript.Encryption.Decrypt("x4yNNf5qsgLpNdA1xUaBM1GaKhwrINqfzNsmDA7rZmZWVx8308y12p1zvsIuEzx+yszVVnhqhQ1cFWL+lBB8yYb53Yx1kBkvdWcXspKfG8buz4RuwCjtXcXkvGOQwdzw");
                Pw = VarEncript.Encryption.Decrypt("Td/V9ZKxqcRFLUfFZD15bv4qZwZIHI0IhNQjdK3EoZQL+8ZJb0vhv5x/XhxtfrN6TxiMJud/+TWSgU6GOTq5YiKRDVJMlSV+f8dswzHxZJ7xjfL8fjyYpd0rFQRMCK41");
                DbLink = "";
            }
            else if (enviroments == "EL9CONV")
            {
                DataBase = VarEncript.Encryption.Decrypt("wCxxnrgxkVOTvIjT7zGOrrnDMwfV5bUHRia1bbl4uaBst2/ndU2Rx/U9QZxazU40TmchLcacJPNXsdUcp/ba8qmO5klx9Fi40kr6gmxJ2/ScoVHzn5W/clZexU62cCYh");
                User = VarEncript.Encryption.Decrypt("p9M5h3knGEbvXqCtwljSTTMeymUMVDXGs1K215lYDLM6zmOe9KCeZw6dIkK2Pv+QYh2cG1iyE7ydQanSYAegh7iqU7RJTGxwv55Eic4VGdcqEIGtdqTuA6bhpNMWQ2b4");
                Pw = VarEncript.Encryption.Decrypt("QfGhOi0/Ub+iepNKjtMpykKmHOyIDM+UTrJa9yhsXihPynUYJO44/6X7+hrgT4cKbeEFUUxIBGJI0Rs0NggyKe9mte1EXfItITbaJVS0dVUwFo2C1ppDCGK2kc5EXskd");
                DbLink = "";
            }
            else if (enviroments == "SIGMAN")
            {
                DataBase = VarEncript.Encryption.Decrypt("YaS6sILu9wwCxRMZK92xpsTUAZbnqJ/xiBrWqSTJIYFjrssEx3Gkj6b+NAK2Prt0HaUEyM6Zn09flO1ZourRTDdMWEBDjybYBh7li16Zsz5DQitq6IpSchv9sLETaHRg");
                User = VarEncript.Encryption.Decrypt("Hxz6bYgtmxCYA+K7R3r8enU3TPoj2/zp0/mM1g8GX2Pq7VK5cSdsWpplCyX8pyVPFdSgjkRl9n0w8tiaIJWeRzzWw7W/Li7fayALDleCBFBbJvR8ae7ZgS0HX3fR03PF");
                Pw = VarEncript.Encryption.Decrypt("C6OLJREhoROT/aF3OvsMfB1IflGSaypP9bSdh6Gubi+aQ9ex+4EsYnKrVzSLKMAmCdO/GLJLxBgZTedVG+OdFFLdcD5/xLI7hmzO/mbRbAL6BQs7tmJBA73saotLWL83");
                DbLink = "";
            }
            else
            {
                throw new NullReferenceException("NO SE PUEDE ENCONTRAR LA BASE DE DATOS SELECCIONADA");
            }
            return true;
        }

        private void formatear_Click(object sender, RibbonControlEventArgs e)
        {
            setSheetHeaderData();
            Centrar();
        }

        public void setSheetHeaderData()
        {
            try
            {
                this._excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                this._excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                if (this.Cells == null)

                    this.Cells = new ExcelStyleCells(this._excelApp);

                Cells.GetCell(ColFinal + "1").Value = "OBLIGATORIO";
                Cells.GetCell(ColFinal + "1").Style = Cells.GetStyle(StyleConstants.TitleRequired);
                Cells.GetCell(ColFinal + "1").Style = Cells.GetStyle(StyleConstants.TitleRequired);
                Cells.GetCell(ColFinal + "2").Value = "OPCIONAL";
                Cells.GetCell(ColFinal + "2").Style = Cells.GetStyle(StyleConstants.TitleOptional);
                Cells.GetCell(ColFinal + "3").Value = "INFORMATIVO";
                Cells.GetCell(ColFinal + "3").Style = Cells.GetStyle(StyleConstants.TitleInformation);
                Cells.GetCell(ColFinal + "4").Value = "ACCIÓN A REALIZAR";
                Cells.GetCell(ColFinal + "4").Style = Cells.GetStyle(StyleConstants.TitleAction);
                Cells.GetCell(ColFinal + "5").Value = "REQUERIDO ADICIONAL";
                Cells.GetCell(ColFinal + "5").Style = Cells.GetStyle(StyleConstants.TitleAdditional);

                Cells.GetRange(ColOcultar, "XFD1048576").Columns.Hidden = true;

                Cells.GetCell("A" + RowCabezera).Value = "FECHA";
                Cells.GetCell("A" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);
                Cells.GetCell("A7").AddComment("MMDDYYYY");

                Cells.GetCell("B" + RowCabezera).Value = "MUESTRA";
                Cells.GetCell("B" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleInformation);

                Cells.GetCell("C" + RowCabezera).Value = "EQUIPO";
                Cells.GetCell("C" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);
                Cells.GetCell("C7").NumberFormat = "@";

                Cells.GetCell("D" + RowCabezera).Value = "COMPAR";
                Cells.GetCell("D" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleInformation);

                Cells.GetCell("E" + RowCabezera).Value = "HOROM";
                Cells.GetCell("E" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleInformation);

                Cells.GetCell("F" + RowCabezera).Value = "RECHEQ";
                Cells.GetCell("F" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleInformation);

                Cells.GetCell("G" + RowCabezera).Value = "PB";
                Cells.GetCell("G" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("H" + RowCabezera).Value = "CU";
                Cells.GetCell("H" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("I" + RowCabezera).Value = "FE";
                Cells.GetCell("I" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("J" + RowCabezera).Value = "CR";
                Cells.GetCell("J" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("K" + RowCabezera).Value = "AL";
                Cells.GetCell("K" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("L" + RowCabezera).Value = "SI";
                Cells.GetCell("L" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("M" + RowCabezera).Value = "MO";
                Cells.GetCell("M" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("N" + RowCabezera).Value = "NA";
                Cells.GetCell("N" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("O" + RowCabezera).Value = "B";
                Cells.GetCell("O" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("P" + RowCabezera).Value = "HOLLIN";
                Cells.GetCell("P" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("Q" + RowCabezera).Value = "DI";
                Cells.GetCell("Q" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("R" + RowCabezera).Value = "H2";
                Cells.GetCell("R" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("S" + RowCabezera).Value = "VI";
                Cells.GetCell("S" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("T" + RowCabezera).Value = "CAL";
                Cells.GetCell("T" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("U" + RowCabezera).Value = "MG";
                Cells.GetCell("U" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("V" + RowCabezera).Value = "OXIDA";
                Cells.GetCell("V" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("W" + RowCabezera).Value = "NITRA";
                Cells.GetCell("W" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("X" + RowCabezera).Value = "SULFA";
                Cells.GetCell("X" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("Y" + RowCabezera).Value = "P";
                Cells.GetCell("Y" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("Z" + RowCabezera).Value = "ZN";
                Cells.GetCell("Z" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AA" + RowCabezera).Value = "NI";
                Cells.GetCell("AA" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AB" + RowCabezera).Value = "SN";
                Cells.GetCell("AB" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AC" + RowCabezera).Value = "TI";
                Cells.GetCell("AC" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AD" + RowCabezera).Value = "V";
                Cells.GetCell("AD" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AE" + RowCabezera).Value = "CADMIO";
                Cells.GetCell("AE" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AF" + RowCabezera).Value = "BARIO";
                Cells.GetCell("AF" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AG" + RowCabezera).Value = "COMPONENTE";
                Cells.GetCell("AG" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AH" + RowCabezera).Value = "MOD";
                Cells.GetCell("AH" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AI" + RowCabezera).Value = "ZFDM";
                Cells.GetCell("AI" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AJ" + RowCabezera).Value = "ISO>4";
                Cells.GetCell("AJ" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AK" + RowCabezera).Value = "ISO>6";
                Cells.GetCell("AK" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);

                Cells.GetCell("AL" + RowCabezera).Value = "ISO>14";
                Cells.GetCell("AL" + RowCabezera).Style = Cells.GetStyle(StyleConstants.TitleRequired);


                Cells.GetCell("A1").Value = "CERREJÓN";
                Cells.GetCell("A1").Style = Cells.GetStyle(StyleConstants.HeaderDefault);
                Cells.MergeCells("A1", "A5");
                Cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.GetRange("A1", "A5").Borders.Weight = "2";

                Cells.GetCell("B1").Value = "MONITOREO - ELLIPSE && ELLIPSE 9";
                Cells.GetCell("B1").Style = Cells.GetStyle(StyleConstants.HeaderDefault);
                Cells.MergeCells("B1", ColHeader + "5");
                Cells.GetRange("B1", ColHeader + "5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.GetRange("B1", ColHeader + "5").Borders.Weight = "2";
                /*Cells.MergeCells("C6", "L11");
                Cells.GetRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.GetRange("C6", "L11").Borders.Weight = "2";
                
                */
                Cells.MergeCells("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1));
                Cells.GetRange("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.GetRange("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1)).Borders.Weight = "2";

                this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                Cells.GetCell("A" + RowInicial).Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show("Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void Centrar()
        {

            Cells.GetCell("B" + RowInicial + ":" + ColFinal + maxRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Cells.GetCell("A" + RowInicial + ":" + ColFinal + maxRow).NumberFormat = "@";

        }

        private void cargar_Click(object sender, RibbonControlEventArgs e)
        {
            if (drpEnviroment.SelectedItem.Label != "EL9CONV")
            {
                frmAuth.StartPosition = FormStartPosition.CenterScreen;
                frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;

                if (frmAuth.ShowDialog() == DialogResult.OK)
                {
                    /* if (true)
                     {
                         frmAuth.EllipseDstrct = "ICOR";
                         frmAuth.EllipsePost = "";
                         frmAuth.EllipseUser = "";
                         frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + Cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            if (i == 33)
                            {
                                i = i + 2;
                            }
                            String Medicion = "" + Cells.GetCell(i, CurrentRow).Value;
                            Medicion = Medicion.Trim();
                            if (!string.IsNullOrEmpty(Medicion))
                            {
                                String Fecha = "" + Cells.GetCell("A" + CurrentRow).Value;
                                Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                                String Equipo = "" + Cells.GetCell("C" + CurrentRow).Value;
                                Equipo = Equipo.Trim();
                                String TipoMon = "AA";
                                String CompMon = "" + Cells.GetCell("AG" + CurrentRow).Value;
                                CompMon = CompMon.Trim();
                                String ModMon = "" + Cells.GetCell("AH" + CurrentRow).Value;
                                ModMon = ModMon.Trim();
                                String Elemento = "" + Cells.GetCell(i, RowCabezera).Value;
                                Elemento = Elemento.Trim();

                                CondMeasurementService.CondMeasurementService proxySheet = new CondMeasurementService.CondMeasurementService();

                                CondMeasurementService.OperationContext opSheet = new CondMeasurementService.OperationContext();

                                try
                                {
                                    CondMeasurementService.CondMeasurementServiceCreateRequestDTO requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();
                                    CondMeasurementService.CondMeasurementServiceCreateReplyDTO replySheet = new CondMeasurementServiceCreateReplyDTO();

                                    proxySheet.Url = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                                    opSheet.district = frmAuth.EllipseDstrct;
                                    opSheet.position = frmAuth.EllipsePost;
                                    opSheet.maxInstances = 100;
                                    opSheet.returnWarnings = Debugger.DebugWarnings;

                                    ClientConversation.authenticate(frmAuth.EllipseUser, frmAuth.EllipsePswd);

                                    requestParamsSheet.equipmentRef = Equipo;
                                    requestParamsSheet.condMonType = TipoMon;
                                    requestParamsSheet.compCode = CompMon;
                                    requestParamsSheet.compModCode = ModMon;
                                    requestParamsSheet.measureDate = Fecha;
                                    requestParamsSheet.condMonMeas = Elemento;

                                    if (Elemento == "HOLLIN" || Elemento == "OXIDA" || Elemento == "NITRA" || Elemento == "SULFA")
                                    {
                                        //Medicion = Medicion.Replace(".", ",");
                                        requestParamsSheet.measureValue = Math.Round(Convert.ToDecimal(Medicion));
                                        requestParamsSheet.measureValueSpecified = true;

                                    }
                                    else
                                    {
                                        requestParamsSheet.measureValue = Convert.ToDecimal(Medicion);
                                        requestParamsSheet.measureValueSpecified = true;
                                    }

                                    replySheet = proxySheet.create(opSheet, requestParamsSheet);

                                    Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAdditional);
                                    Cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                                }
                                catch (Exception ex)
                                {
                                    Cells.GetCell(i, CurrentRow).ClearComments();
                                    Cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                    Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAction);
                                    Cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                                    //   ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                                }

                            }
                        }

                        CurrentRow++;
                    }
                }
            }
            else
            {
                frmAuth.StartPosition = FormStartPosition.CenterScreen;
                frmAuth.SelectedEnvironment = "Contingencia";

                if (frmAuth.ShowDialog() == DialogResult.OK)
                {
                    /* if (true)
                     {
                         frmAuth.EllipseDstrct = "ICOR";
                         frmAuth.EllipsePost = "";
                         frmAuth.EllipseUser = "";
                         frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + Cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            if (i == 33)
                            {
                                i = i + 2;
                            }
                            String Medicion = "" + Cells.GetCell(i, CurrentRow).Value;
                            Medicion = Medicion.Trim();
                            if (!string.IsNullOrEmpty(Medicion))
                            {
                                String Fecha = "" + Cells.GetCell("A" + CurrentRow).Value;
                                Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                                String Equipo = "" + Cells.GetCell("C" + CurrentRow).Value;
                                Equipo = Equipo.Trim();
                                String TipoMon = "AA";
                                String CompMon = "" + Cells.GetCell("AG" + CurrentRow).Value;
                                CompMon = CompMon.Trim();
                                String ModMon = "" + Cells.GetCell("AH" + CurrentRow).Value;
                                ModMon = ModMon.Trim();
                                String Elemento = "" + Cells.GetCell(i, RowCabezera).Value;
                                Elemento = Elemento.Trim();

                                CondMeasurementService9.CondMeasurementService proxySheet = new CondMeasurementService9.CondMeasurementService();

                                CondMeasurementService9.OperationContext opSheet = new CondMeasurementService9.OperationContext();

                                try
                                {
                                    CondMeasurementService9.CondMeasurementServiceCreateRequestDTO requestParamsSheet = new CondMeasurementService9.CondMeasurementServiceCreateRequestDTO();
                                    CondMeasurementService9.CondMeasurementServiceCreateReplyDTO replySheet = new CondMeasurementService9.CondMeasurementServiceCreateReplyDTO();

                                    proxySheet.Url = "http://ews-eamprd.lmnerp01.cerrejon.com/ews/services" + "/CondMeasurementService";

                                    opSheet.district = frmAuth.EllipseDstrct;
                                    opSheet.position = frmAuth.EllipsePost;
                                    opSheet.maxInstances = 100;
                                    opSheet.returnWarnings = Debugger.DebugWarnings;

                                    ClientConversation.authenticate(frmAuth.EllipseUser, frmAuth.EllipsePswd);

                                    requestParamsSheet.equipmentRef = Equipo;
                                    requestParamsSheet.condMonType = TipoMon;
                                    requestParamsSheet.compCode = CompMon;
                                    requestParamsSheet.compModCode = ModMon;
                                    requestParamsSheet.measureDate = Fecha;
                                    requestParamsSheet.condMonMeas = Elemento;

                                    if (Elemento == "HOLLIN" || Elemento == "OXIDA" || Elemento == "NITRA" || Elemento == "SULFA")
                                    {
                                        //Medicion = Medicion.Replace(".", ",");
                                        requestParamsSheet.measureValue = Math.Round(Convert.ToDecimal(Medicion));
                                        requestParamsSheet.measureValueSpecified = true;

                                    }
                                    else
                                    {
                                        requestParamsSheet.measureValue = Convert.ToDecimal(Medicion);
                                        requestParamsSheet.measureValueSpecified = true;
                                    }

                                    replySheet = proxySheet.create(opSheet, requestParamsSheet);

                                    Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAdditional);
                                    Cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                                }
                                catch (Exception ex)
                                {
                                    Cells.GetCell(i, CurrentRow).ClearComments();
                                    Cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                    Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAction);
                                    Cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                                    //   ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                                }

                            }
                        }

                        CurrentRow++;
                    }
                }
            }

                    MessageBox.Show("Proceso Finalizado");

               // }
            
        }

        private void Limpiar()
        {
            Cells.GetCell("A" + RowInicial + ":" + ColFinal + maxRow).Clear();
        }

        private void buttonLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            Limpiar();
        }

        private void borrar_Click(object sender, RibbonControlEventArgs e)
        {
            if (drpEnviroment.SelectedItem.Label != "EL9CONV")
            {
                frmAuth.StartPosition = FormStartPosition.CenterScreen;
                frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;

                if (frmAuth.ShowDialog() == DialogResult.OK)
                {
                    /*if (true)
                    {
                        frmAuth.EllipseDstrct = "";
                        frmAuth.EllipsePost = "";
                        frmAuth.EllipseUser = "";
                        frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + Cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            /*if (i == 33)
                            {
                                i = i + 2;
                            }*/

                            String Fecha = "" + Cells.GetCell("A" + CurrentRow).Value;
                            Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                            String Equipo = "" + Cells.GetCell("C" + CurrentRow).Value;
                            Equipo = Equipo.Trim();
                            String TipoMon = "AA";
                            String CompMon = "" + Cells.GetCell("AG" + CurrentRow).Value;
                            CompMon = CompMon.Trim();
                            String ModMon = "" + Cells.GetCell("AH" + CurrentRow).Value;
                            ModMon = ModMon.Trim();
                            String Elemento = "" + Cells.GetCell(i, RowCabezera).Value;
                            Elemento = Elemento.Trim();

                            CondMeasurementService.CondMeasurementService proxySheet = new CondMeasurementService.CondMeasurementService();

                            CondMeasurementService.OperationContext opSheet = new CondMeasurementService.OperationContext();

                            try
                            {

                                CondMeasurementService.CondMeasurementServiceDeleteRequestDTO requestParamsSheet = new CondMeasurementServiceDeleteRequestDTO();
                                CondMeasurementService.CondMeasurementServiceDeleteReplyDTO replySheet = new CondMeasurementServiceDeleteReplyDTO();

                                proxySheet.Url = EFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                                opSheet.district = frmAuth.EllipseDstrct;
                                opSheet.position = frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(frmAuth.EllipseUser, frmAuth.EllipsePswd);

                                requestParamsSheet.equipmentRef = Equipo;
                                requestParamsSheet.condMonType = TipoMon;
                                requestParamsSheet.compCode = CompMon;
                                requestParamsSheet.compModCode = ModMon;
                                requestParamsSheet.measureDate = Fecha;
                                requestParamsSheet.condMonMeas = Elemento;

                                replySheet = proxySheet.delete(opSheet, requestParamsSheet);

                                Cells.GetCell(i, CurrentRow).Clear();
                                Cells.GetCell(i, CurrentRow).Select();

                            }
                            catch (Exception ex)
                            {
                                Cells.GetCell(i, CurrentRow).ClearComments();
                                Cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAction);
                                Cells.GetCell(i, CurrentRow).Clear();
                                Cells.GetCell(i, CurrentRow).Select();
                                // ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                            }

                        }

                        CurrentRow++;
                    }
                }
            }
            else
            {
                frmAuth.StartPosition = FormStartPosition.CenterScreen;
                frmAuth.SelectedEnvironment = "Contingencia";

                if (frmAuth.ShowDialog() == DialogResult.OK)
                {
                    /*if (true)
                    {
                        frmAuth.EllipseDstrct = "";
                        frmAuth.EllipsePost = "";
                        frmAuth.EllipseUser = "";
                        frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + Cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            /*if (i == 33)
                            {
                                i = i + 2;
                            }*/

                            String Fecha = "" + Cells.GetCell("A" + CurrentRow).Value;
                            Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                            String Equipo = "" + Cells.GetCell("C" + CurrentRow).Value;
                            Equipo = Equipo.Trim();
                            String TipoMon = "AA";
                            String CompMon = "" + Cells.GetCell("AG" + CurrentRow).Value;
                            CompMon = CompMon.Trim();
                            String ModMon = "" + Cells.GetCell("AH" + CurrentRow).Value;
                            ModMon = ModMon.Trim();
                            String Elemento = "" + Cells.GetCell(i, RowCabezera).Value;
                            Elemento = Elemento.Trim();

                            CondMeasurementService9.CondMeasurementService proxySheet = new CondMeasurementService9.CondMeasurementService();

                            CondMeasurementService9.OperationContext opSheet = new CondMeasurementService9.OperationContext();

                            try
                            {

                                CondMeasurementService9.CondMeasurementServiceDeleteRequestDTO requestParamsSheet = new CondMeasurementService9.CondMeasurementServiceDeleteRequestDTO();
                                CondMeasurementService9.CondMeasurementServiceDeleteReplyDTO replySheet = new CondMeasurementService9.CondMeasurementServiceDeleteReplyDTO();

                                proxySheet.Url = "http://ews-eamprd.lmnerp01.cerrejon.com/ews/services" + "/CondMeasurementService";

                                opSheet.district = frmAuth.EllipseDstrct;
                                opSheet.position = frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(frmAuth.EllipseUser, frmAuth.EllipsePswd);

                                requestParamsSheet.equipmentRef = Equipo;
                                requestParamsSheet.condMonType = TipoMon;
                                requestParamsSheet.compCode = CompMon;
                                requestParamsSheet.compModCode = ModMon;
                                requestParamsSheet.measureDate = Fecha;
                                requestParamsSheet.condMonMeas = Elemento;

                                replySheet = proxySheet.delete(opSheet, requestParamsSheet);

                                Cells.GetCell(i, CurrentRow).Clear();
                                Cells.GetCell(i, CurrentRow).Select();

                            }
                            catch (Exception ex)
                            {
                                Cells.GetCell(i, CurrentRow).ClearComments();
                                Cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                Cells.GetCell(i, CurrentRow).Style = Cells.GetStyle(StyleConstants.TitleAction);
                                Cells.GetCell(i, CurrentRow).Clear();
                                Cells.GetCell(i, CurrentRow).Select();
                                // ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                            }

                        }

                        CurrentRow++;
                    }
                }
            }

                    MessageBox.Show("Proceso Finalizado");

            //    }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn("Gustavo Vargas Lopez", "").ShowDialog();
        }
        private void drpEnviroment_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}


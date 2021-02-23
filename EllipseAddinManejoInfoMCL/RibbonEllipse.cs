using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.Globalization;
using System.Data;
using data = System.Data;
using System.Drawing;
using System.Threading;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Vsto.Excel;
using VarEncript = SharedClassLibrary.Utilities.Encryption;
using System.Web.Services.Ellipse;
using System.Web.Services;
using EllipseWorkOrdersClassLibrary;
using Debugger = SharedClassLibrary.Utilities.Debugger;
using SharedClassLibrary.Classes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.IO;
using System.Runtime.InteropServices;
//using Oracle.ManagedDataAccess.Client;


namespace EllipseAddinManejoInfoMCL
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions;
        private Excel.Application _excelApp;

        //private OracleConnection _sqlOracleConn;

        //private Worksheet _excelSheet;
        private Thread _thread;
        private bool _progressUpdate = true;
        //Selecionar objecto por default
        static object useDefault = Type.Missing;
        //CONEXION CADENA
        //private Excel.Application _excelApp;

        //-------------------------------------------------------------------------------------------PARAMETROS DE MOVIMIENTO DE OBJECTOS EN LA HOJA DE CALCULO---------------------------
        //INICIO DE COLUMNA Y FILA DE -----------------------------------------------IMAGEN
        public Int32 StartColImg = 2;
        public Int32 StartRowImg = 2;
        public Int32 EndColImg = 3;
        public Int32 EndRowImg = 2;
        //INICIO DE COLUMNA Y FILA DE -----------------------------------------------TITULO
        public Int32 StartColTitulo = 5;
        public Int32 StartRowTitulo = 2;
        public Int32 EndColTitulo = 12;
        public Int32 EndRowTitulo = 2;
        //INICIOS DE LA COLUMNA Y FILA DE LOS ---------------------------------------INPUT
        public Int32 StartColInputMenu = 5;
        public Int32 StartRowInputMenu = 1;
        //INICIO DE COLUMNA Y FILA DE LA --------------------------------------------TABLA
        public Int32 StartColTable = 1;
        public Int32 StartRowTable = 5;
        //UTILIDADES PARA LOS MOVIMIENTOS DINAMICOS
        public Int32 Mayor = 0;
        public Int32 CntIndicador = 0;
        //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        public SqlConnection cnnx;
        public String Sql = "";
        //Variables de Conexion 
        private string SQL;
        private string DataBase;
        private string User; //Ej. SIGCON, CONSULBO
        private string Pw;
        // ReSharper disable once InconsistentNaming
        public string DbLink; //Ej. @DBLMIMS

        //private const string ValidationSheetName01 = "ValidationSheetEventos";
        //private const string ValidationSheetName02 = "ValidationSheetCargas";
        private const string SheetName01 = "PLAN COMBUSTIBLE";
        private const string SheetName02 = "PERSONAL";
        private const string SheetName03 = "Eventos_Pivot";

        private const string tableName01 = "_01Eventos";
        private const string tableName02 = "_01Cargas";
        private const string tableName03 = "Pivot_Eventos";
        //private const int titleRow = 8;


        //OracleConnection Conexion;
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            //_frmAuth = new FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;
            _excelApp.EnableEvents = true;
            //var environments = Environments.GetEnvironmentList();
            var environments = new string[] { "lmnsql01", "lmnsql03" };
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
            else if (enviroments == "lmnsql01")
            {
                DataBase = VarEncript.Encryption.Decrypt("IiN2an9lhojMR+mXSw/g2+JTY43Jq25tDuzEP1pxbciQIdYiq7TQaTAZku+D7SjuUSzjdKGZb3QuMqT9ws4+vRM5kKRLj57H9uD9I1yhZpd2K1KGYOk4TNgv//v8NiqK");
                User = VarEncript.Encryption.Decrypt("blJAiE1+5yu5twbORVc/BGoYwVn9hs4KJFe9+czOn2UNJW9E0+LVIha4jQrbwpZkbmLezwru+ppsvlXBkkuv5THz/j/eELhWWiGcdYsHZmCNvpjF/fzZOcsYVhGIPUTy");
                Pw = VarEncript.Encryption.Decrypt("BsXJzxjwCPrJm5jaBhqOrgGUyatkO7e8BDaPfPp9WwKN8WBD4jHenvkTr2QU/hsOAORa9J3tlTHcAYbM59Dv/IGl30tOi1i6uYi/yUUQwnkmB5T46WxT5e8q9qLjFw0g");
                DbLink = "";
            }
            else if (enviroments == "lmnsql03")
            {
                DataBase = VarEncript.Encryption.Decrypt("IiN2an9lhojMR+mXSw/g2+JTY43Jq25tDuzEP1pxbciQIdYiq7TQaTAZku+D7SjuUSzjdKGZb3QuMqT9ws4+vRM5kKRLj57H9uD9I1yhZpd2K1KGYOk4TNgv//v8NiqK");
                User = VarEncript.Encryption.Decrypt("blJAiE1+5yu5twbORVc/BGoYwVn9hs4KJFe9+czOn2UNJW9E0+LVIha4jQrbwpZkbmLezwru+ppsvlXBkkuv5THz/j/eELhWWiGcdYsHZmCNvpjF/fzZOcsYVhGIPUTy");
                Pw = VarEncript.Encryption.Decrypt("BsXJzxjwCPrJm5jaBhqOrgGUyatkO7e8BDaPfPp9WwKN8WBD4jHenvkTr2QU/hsOAORa9J3tlTHcAYbM59Dv/IGl30tOi1i6uYi/yUUQwnkmB5T46WxT5e8q9qLjFw0g");
                DbLink = "";
            }
            else
            {
                throw new NullReferenceException("NO SE PUEDE ENCONTRAR LA BASE DE DATOS SELECCIONADA");
            }
            return true;
        }

        private void btnFormatear_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = true;
            try
            {
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                Formatear("PLAN DE TANQUEO DE COMBUSTIBLE", SheetName01, false);
                Formatear("PERSONAL", SheetName02, false);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:Formatear()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null)
                    _cells.SetCursorDefault();
                BorrarSheets();
                _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                _excelApp.ScreenUpdating = true;
                _excelApp.DisplayAlerts = true;
            }
        }

        public data.DataTable getdata(string SQL, Int32 SW = 0)
        {
            if (SW == 0)
            {
                var dbi = Environments.GetDatabaseItem(Environments.EllipseProductivo);
                dbi.DbUser = "consulbo";
                dbi.DbPassword = "ventyx15";
                _eFunctions.SetDBSettings(dbi.DbName, dbi.DbUser, dbi.DbPassword);
                _eFunctions.SetConnectionTimeOut(0);
            }
            else
            {
                var dbi = Environments.GetDatabaseItem(Environments.SigmanProductivo);
                dbi.DbUser = "sigman";
                dbi.DbPassword = "sig0679";
                _eFunctions.SetDBSettings(dbi.DbName, dbi.DbUser, dbi.DbPassword);
                _eFunctions.SetConnectionTimeOut(0);
            }



            //ConexionDataBase(drpEnviroment.SelectedItem.Label);
            //VerificarConexion(DataBase, User, Pw, DbLink);
            //_eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var dat = _eFunctions.GetQueryResult(SQL);
            data.DataTable DATA = new data.DataTable();
            DATA.Load(dat);
            _eFunctions.CloseConnection();
            return DATA;
        }
        public void ExecuteQuery(string sqlQuery, string NameHoja, Int32 T = 0)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                data.DataTable table;
                table = getdata(sqlQuery, T);
                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("NO SE ENCONTRO INFORMACION");
                    return;
                }
                //hacemos estatica la primer fila y aplicamos filtros asi,
                _excelApp.Application.ActiveWindow.SplitRow = StartRowTable;
                _excelApp.Application.ActiveWindow.FreezePanes = true;
                int i = 0;
                string[,] data = new string[table.Rows.Count, table.Columns.Count];
                foreach (data.DataRow row in table.Rows)
                {
                    int j = 0;
                    //Columnas de la consulta
                    foreach (data.DataColumn col in table.Columns)
                    {
                        data[i, j] = row[col].ToString();
                        j++;
                    }
                    i++;
                    //format row
                    /*if (i % 2 == 0)
                    {
                        _cells.GetRange(StartColTable, (StartRowTable + i), table.Columns.Count, (StartRowTable + i)).Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(221, 235, 247)));
                    }*/
                }
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).NumberFormat = "@";
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = data;
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value;

                Encabezado(table, _excelApp.ActiveWorkbook.ActiveSheet.Name);
                FormatTable(_cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable), NameHoja, 1, 1);
                if(_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                {
                    List<string> lista = new List<string>();
                    lista.Add("C");
                    lista.Add("M");
                    lista.Add("D");
                    //for (int F = 0; F < table.Rows.Count; F++)
                    //{
                        _cells.GetRange("D" + 6, "D1000").Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), lista), Type.Missing);
                    //}
                }
                table = null;

                _excelApp.ActiveWindow.Zoom = 80;
                _excelApp.Columns.AutoFit();
                _excelApp.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ExecuteQuery()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _eFunctions.CloseConnection();
                _cells.SetCursorDefault();
                _excelApp.ScreenUpdating = true;

                //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                //_excelApp.ScreenUpdating = true;
                //_excelApp.DisplayAlerts = true;
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
            }

        }

        public data.DataTable getdataSql(string SQL)
        {
            //if (ConexionDataBase(drpEnvironment.SelectedItem.Label))
            ConexionDataBase(drpEnvironment.SelectedItem.Label);
            cnnx = new SqlConnection("server=" + drpEnvironment.SelectedItem.Label + ";database= " + DataBase + "; uid=" + User + "; pwd =" + Pw + ";Connection Timeout=0");
            data.DataTable DATA = new data.DataTable();
            SqlDataAdapter dat = new SqlDataAdapter(SQL, cnnx);
            dat.Fill(DATA);
            return DATA;
            
        }

        List<string> ListaDatos(Int32 Tipo, string Param1 = "", string Param2 = "", String ORDEN = "")
        {
            List<string> listRange = new List<string>();
            data.DataTable table = null;
            if (Tipo == 1)
            {
                Sql = (@"select
                          max(shiftindex)
                        from
                          PowerView.dbo.hist_exproot " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 2)
            {
                Sql = (@"select
                          CASE WHEN shift# = 1 THEN 'A1' ELSE 'A2' END AS TURNO
                        from
                          PowerView.dbo.hist_exproot 
                        Where
                          shiftindex = " + Param1 + " " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 3)
            {
                Sql = (@"select
                          RTRIM(crew) AS GRUPO
                        from
                          PowerView.dbo.hist_exproot 
                        Where
                          shiftindex = " + Param1 + " " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 4)
            {
                Sql = (@"SELECT
                          PRS.NAME
                        FROM
                          SIGMAN.APP_PTC_PERSONAL PRS
                        WHERE
                          PRS.TYPE = '" + Param1 + "' " + ORDEN);
                table = getdata(Sql,1);
            }
            else if(Tipo == 5)
            {
                Sql = (@"SELECT
                            hist_statusevents.eqmt--,
					        --Count(hist_statusevents.category) EV_DISPONIBLES
				        FROM
					        PowerView.dbo.hist_statusevents hist_statusevents
                            INNER JOIN PowerView.dbo.hist_eqmtlist hist_eqmtlist on(hist_statusevents.shiftindex=hist_eqmtlist.shiftindex AND hist_statusevents.eqmt=hist_eqmtlist.eqmtid)
                        WHERE
					        hist_statusevents.shiftindex = '" + Param2 + @"'
                            AND hist_statusevents.category = 2
                            AND hist_eqmtlist.unit = '" + Param1 + @"'
                        GROUP BY
                            hist_statusevents.eqmt 
                        HAVING Count(hist_statusevents.category) > 0 ORDER BY 1 " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 6)
            {
                Sql = (@"SELECT
                          icr_tajos.Inicial + '-' + LTRIM(Tajo)
                        FROM
                          PowerView.dbo.icr_tajos
                        WHERE
                          icr_tajos.Inicial <> 'REM' ORDER BY icr_tajos.Inicial " + ORDEN);
                table = getdataSql(Sql);
            }
            int i = 0;
            string[,] data = new string[table.Rows.Count, table.Columns.Count];
            //Filas de la consulta
            foreach (data.DataRow row in table.Rows)
            {
                int j = 0;
                //Columnas de la consulta
                foreach (data.DataColumn col in table.Columns)
                {
                    //data[i, j] = Convert.ToDouble(row[c].ToString()).ToString("N02");
                    listRange.Add(row[col].ToString());
                    j++;
                }
                i++;
            }
            return listRange;
        }

        public void Formatear(string Titulo = "", string NombreHoja = "", bool SubEncab = false)
        {
            CntIndicador = CntIndicador + 1;
            //_eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            #region CONSTRUYO LA HOJA 1 y 2
            //while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
           
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            if (CntIndicador == 1)
            {
                _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.Windows.Application.Sheets[_excelApp.Windows.Application.Sheets.Count]);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = NombreHoja;



                FormatCamposMenu(_cells.GetRange("A1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("A1", "S2"));
                FormatCamposMenu(_cells.GetRange("T1", "V4"), true, true, true, "CUMPLIMIENTO PLANES", "", 11, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("T1", "V4"));

                //_cells.SetValidationList(_cells.GetRange(ColResCode, StartRowTable + 1, ColResCode, FinRowForFormat), ListaDatos(3, "ASC"));


                //3 Y 4 FILA DESDE A HASTA S
                FormatCamposMenu(_cells.GetRange("A3", "A4"), true, true, true, "FECHA", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("A3", "A4"));
                FormatCamposMenu(_cells.GetRange("B3", "C4"), true, true, true, DateTime.Now.ToString("yyyy-MM-dd"), "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B3", "C4"));
                FormatCamposMenu(_cells.GetRange("D3", "E3"), true, true, true, "SHIFTINDEX", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D3", "E3"));
                FormatCamposMenu(_cells.GetRange("D4", "E4"), true, true, true, ListaDatos(1)[0], "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D4", "E4"));
                FormatCamposMenu(_cells.GetRange("F3", "F4"), true, true, true, "TURNO", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("F3", "F4"));
                FormatCamposMenu(_cells.GetRange("G3", "G4"), true, true, true, ListaDatos(2, _cells.GetCell("D4").Value)[0], "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("G3", "G4"));
                FormatCamposMenu(_cells.GetRange("H3", "H4"), true, true, true, "GRUPO", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("H3", "H4"));
                FormatCamposMenu(_cells.GetRange("I3", "L4"), true, true, true, ListaDatos(3, _cells.GetCell("D4").Value)[0], "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("I3", "L4"));
                FormatCamposMenu(_cells.GetRange("M3", "N3"), true, true, true, "DESPACHADOR:", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("M3", "N3"));
                FormatCamposMenu(_cells.GetRange("M4", "N4"), true, true, true, "SUPERVISOR:", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("M4", "N4"));
                FormatCamposMenu(_cells.GetRange("O3", "S3"), true, true, true, "", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
                _cells.GetCell("O3").Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), ListaDatos(4, "DP")), Type.Missing);
                FormatBordes(_cells.GetRange("O3", "S3"));
                FormatCamposMenu(_cells.GetRange("O4", "S4"), true, true, true, "", "", 14, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
                _cells.GetCell("O4").Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), ListaDatos(4, "SP")), Type.Missing);
                FormatBordes(_cells.GetRange("O4", "S4"));
                //Primera fila Palas Indicadores
                FormatCamposMenu(_cells.GetRange("W1", "X1"), true, true, true, "PALAS:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("W1", "X1"));
                FormatCamposMenu(_cells.GetCell("Y1"), true, false, false, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Y1"));
                FormatCamposMenu(_cells.GetCell("Z1"), true, false, true, "SIN COMB:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Z1"));
                FormatCamposMenu(_cells.GetCell("AA1"), true, false, false, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AA1"));
                FormatCamposMenu(_cells.GetCell("AB1"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AB1"));
                FormatCamposMenu(_cells.GetRange("AC1", "AD1"), true, true, true, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AC1", "AD1"));
                FormatCamposMenu(_cells.GetCell("AE1"), true, false, true, "% de cump.", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AE1"));
                FormatCamposMenu(_cells.GetCell("AF1"), true, false, true, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AF1"));

                //Segunda fila Auxiliares Indicadores
                FormatCamposMenu(_cells.GetRange("W2", "X2"), true, true, true, "AUXILIARES:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("W2", "X2"));
                FormatCamposMenu(_cells.GetCell("Y2"), true, false, false, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Y2"));
                FormatCamposMenu(_cells.GetCell("Z2"), true, false, true, "SIN COMB:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Z2"));
                FormatCamposMenu(_cells.GetCell("AA2"), true, false, false, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AA2"));
                FormatCamposMenu(_cells.GetCell("AB2"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AB2"));
                FormatCamposMenu(_cells.GetRange("AC2", "AD2"), true, true, true, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AC2", "AD2"));
                FormatCamposMenu(_cells.GetCell("AE2"), true, false, true, "% de cump.", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AE2"));
                FormatCamposMenu(_cells.GetCell("AF2"), true, false, true, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AF2"));

                //Tercera fila Auxiliares Indicadores
                FormatCamposMenu(_cells.GetRange("W3", "X3"), true, true, true, "CARGADORES:", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("W3", "X3"));
                FormatCamposMenu(_cells.GetCell("Y3"), true, false, false, "", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Y3"));
                FormatCamposMenu(_cells.GetCell("Z3"), true, false, true, "SIN COMB:", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Z3"));
                FormatCamposMenu(_cells.GetCell("AA3"), true, false, false, "", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AA3"));
                FormatCamposMenu(_cells.GetCell("AB3"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AB3"));
                FormatCamposMenu(_cells.GetRange("AC3", "AD3"), true, true, true, "", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AC3", "AD3"));
                FormatCamposMenu(_cells.GetCell("AE3"), true, false, true, "% de cump.", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AE3"));
                FormatCamposMenu(_cells.GetCell("AF3"), true, false, true, "", "", 11, Rf: 198, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AF3"));

                //Cuarta fila Auxiliares Indicadores
                FormatCamposMenu(_cells.GetRange("W4", "X4"), true, true, true, "CARGADORES:", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("W4", "X4"));
                FormatCamposMenu(_cells.GetCell("Y4"), true, false, false, "", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Y4"));
                FormatCamposMenu(_cells.GetCell("Z4"), true, false, true, "SIN COMB:", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("Z4"));
                FormatCamposMenu(_cells.GetCell("AA4"), true, false, false, "", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AA4"));
                FormatCamposMenu(_cells.GetCell("AB4"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AB4"));
                FormatCamposMenu(_cells.GetRange("AC4", "AD4"), true, true, true, "", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AC4", "AD4"));
                FormatCamposMenu(_cells.GetCell("AE4"), true, false, true, "% de cump.", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AE4"));
                FormatCamposMenu(_cells.GetCell("AF4"), true, false, true, "", "", 11, Rf: 142, Gf: 169, Bf: 219, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetCell("AF4"));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS
                FormatCamposMenu(_cells.GetCell("A5"), true, false, true, "TALADROS", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("B5"), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("C5"), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("D5"), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("E5"), true, false, true, "GL-300-600", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("F5"), true, false, true, "HORA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);

                var Datos1 = ListaDatos(5, "Perforadora          ", _cells.GetCell("D4").Value, "ASC");
                var Tam = Datos1.Count;
                Int32 F = 0;
                foreach (var Result in Datos1)
                {
                    FormatCamposMenu(_cells.GetCell("A" + (6 + F)), true, false, true, Result, "ASC", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                //var Tam = F;
                FormatBordes(_cells.GetRange("A5", "F" + (F + 5)));

                var Datos2 = ListaDatos(6, null, null, "ASC");
                for (int i = 0; i < Tam; i++)
                {
                    _cells.GetCell("C" + (6 + i)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                }

                List<string> lista = new List<string>();
                lista.Add("OK");
                lista.Add("Sin Combustible");
                lista.Add("Reprogramado");
                for (int i = 0; i < Tam; i++)
                {
                    _cells.GetCell("D" + (6 + i)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), lista), Type.Missing);
                }

            }
            if(CntIndicador == 2)
            {
                _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.Windows.Application.Sheets[_excelApp.Windows.Application.Sheets.Count]);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = NombreHoja;

                FormatCamposMenu(_cells.GetRange("A1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("A1", "S2"));
            }






            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();

            #endregion
            //_excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }

        private void FormatCamposMenu(Excel.Range Celda, bool Col, bool Merge, bool Negrita, String Texto = "", String Comentario = "", /*bool Bords, */Int32 TamLetra = 9, Int32 Rf = 91, Int32 Gf = 155, Int32 Bf = 213, Int32 Rl = 255, Int32 Gl = 255, Int32 Bl = 255, string Aline = "")
        {

            Celda.NumberFormat = "@";
            if(Negrita)
            {
                Celda.Font.Bold = true;
            }
            Celda.Font.Size = TamLetra;
            /*if (Bords)
            {
                FormatBordes(Celda);
            }*/
            if (Col)
            {
                Celda.Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(Rf, Gf, Bf)));
                Celda.Font.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(Rl, Gl, Bl/*Color.White*/)));
                if (Aline == "C")
                {
                    Celda.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Celda.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                else if (Aline == "L")
                {
                    Celda.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    Celda.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                else if (Aline == "R")
                {
                    Celda.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Celda.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                else
                {
                    Celda.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Celda.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
            }
            if(Merge)
            {
                Celda.Merge();
            }
            if (Texto != "")
            {
                Celda.Value = Texto;
            }
            if (Comentario != "")
            {
                Celda.AddComment(Comentario);
            }

        }
        private void FormatBordes(Excel.Range Rango, Excel.XlBorderWeight Weight = Excel.XlBorderWeight.xlThin)
        {
            Rango.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            Rango.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //Asigno los bordes ubicados a la izquierda
            Excel.Border B1 = Rango.Borders[Excel.XlBordersIndex.xlEdgeLeft];
            B1.LineStyle = Excel.XlLineStyle.xlContinuous;
            B1.ColorIndex = 0;
            B1.TintAndShade = 0;
            B1.Weight = Weight;
            //Asigno los bordes ubicados en la parte superior o Top
            Excel.Border B2 = Rango.Borders[Excel.XlBordersIndex.xlEdgeTop];
            B2.LineStyle = Excel.XlLineStyle.xlContinuous;
            B2.ColorIndex = 0;
            B2.TintAndShade = 0;
            B2.Weight = Weight;
            //Asigno los bordes ubicados en el Boton
            Excel.Border B3 = Rango.Borders[Excel.XlBordersIndex.xlEdgeBottom];
            B3.LineStyle = Excel.XlLineStyle.xlContinuous;
            B3.ColorIndex = 0;
            B3.TintAndShade = 0;
            B3.Weight = Weight;
            //Asigno los bordes ubicados a la Derecha
            Excel.Border B4 = Rango.Borders[Excel.XlBordersIndex.xlEdgeRight];
            B4.LineStyle = Excel.XlLineStyle.xlContinuous;
            B4.ColorIndex = 0;
            B4.TintAndShade = 0;
            B4.Weight = Weight;
            //Asigno los bordes ubicados en la parte Vertical de la cenda
            Excel.Border B5 = Rango.Borders[Excel.XlBordersIndex.xlInsideVertical];
            B5.LineStyle = Excel.XlLineStyle.xlContinuous;
            B5.ColorIndex = 0;
            B5.TintAndShade = 0;
            B5.Weight = Weight;
            //Asigno los bordes ubicados en la parte horizontal de la cenda
            Excel.Border B6 = Rango.Borders[Excel.XlBordersIndex.xlInsideHorizontal];
            B6.LineStyle = Excel.XlLineStyle.xlContinuous;
            B6.ColorIndex = 0;
            B6.TintAndShade = 0;
            B6.Weight = Weight;
        }
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBoxExcelAddIn About = new AboutBoxExcelAddIn("Gustavo Vargas", "GAVL-SOFT");
            About.ShowDialog();
        }

        private void BorrarSheets(String Hoja = "")
        {
            //_excelApp.DisplayAlerts = false;
            if (Hoja != "")
            {
                for (int v = 1; v <= _excelApp.Windows.Application.Sheets.Count; v++)
                {
                    var wkSheet = _excelApp.Windows.Application.Sheets[v];
                    if (wkSheet.Name == Hoja)
                    {
                        wkSheet.Delete();
                        break;
                    }
                }
            }
            else
            {
                string HojaIdioma = LanguageExcel();
                for (int v = _excelApp.ActiveWorkbook.Worksheets.Count; v > 0; v--)
                {
                    Excel.Worksheet wkSheet = (Excel.Worksheet)_excelApp.ActiveWorkbook.Worksheets[v];
                    if (wkSheet.Name == HojaIdioma + 1 || wkSheet.Name == HojaIdioma + 2 || wkSheet.Name == HojaIdioma + 3)
                    {
                        wkSheet.Delete();
                    }
                }
            }
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
        }


        private string LanguageExcel()
        {
            string Hoja = "";
            Int32 CodLanguage = _excelApp.Application.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI];

            if (CodLanguage == 1033)
            {
                Hoja = "Sheet";
            }
            else
            {
                Hoja = "Hoja";
            }
            return Hoja;
        }




        private string Separador()
        {
            string separator;
            //si uso los separadores del sistema
            if (_excelApp.UseSystemSeparators)
            {
                separator = LanguageSettingConstants.ListSeparator;
                //si el separador de lista y el separador decimal son iguales
                if (LanguageSettingConstants.ListSeparator.Equals(LanguageSettingConstants.DecimalSeparator))
                    separator = LanguageSettingConstants.DecimalSeparator.Equals(",") ? ";" : ",";
            }
            else
            {
                separator = _excelApp.DecimalSeparator.Equals(",") ? ";" : ",";

            }
            return separator;
        }





        private void Encabezado(data.DataTable table, String Hoja)
        {
            //Formateando columnas de encabezado
            //_excelApp.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, _cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, StartRowTable), Type.Missing, Excel.XlYesNoGuess.xlNo, Type.Missing).Name = "TiTul01";
            int cont = StartColTable;
            for (var i = 0; i < table.Columns.Count; i++)
            {

                _cells.GetCell(cont, StartRowTable).Value = table.Columns[i].ColumnName.Trim();
                cont++;
            }

        }





        private void FormatTable(Excel.Range Rango, string HojaName, Int32 StyleText = 0, Int32 TypeTable = 0)
        {
            //_excelApp.ActiveWorkbook.ActiveSheet.Select();
            //Rango.Select();
            String NameTable = "01";
            NameTable = NameTable + Convert.ToString(HojaName);
            //Rango.Select();
            if (StyleText == 1)
            {
                //Formato de letras
                Rango.Font.FontStyle = "Negrita";
                Rango.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                Rango.Font.Size = 10;
                Rango.Font.TintAndShade = 0;
                //Rango.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontMinor;
            }
            else if (StyleText == 2)
            {
            }

            if (TypeTable == 1)
            {
                //CREO NOMBRE A LA TABLA
                Excel.ListObject TableFiltro = _excelApp.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, Rango, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
                TableFiltro.Name = NameTable;
                if (HojaName == SheetName01)
                {
                    TableFiltro.ShowHeaders = false;
                }
                //Rango.AutoFilter(StartRowTable, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            }
            else if (TypeTable == 2)
            {

            }
            FormatBordes(Rango);

        }

        private void btnConsultar_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 1), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ExecuteQuery()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show(@"Se ha producido un error: " + ex.Message);
                }
                finally
                {
                    if (_cells != null)
                        //_eFunctions.CloseConnection();
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    //_excelApp.ScreenUpdating = true;
                    //_excelApp.DisplayAlerts = true;
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }



        public string Consulta(Int32 Tipe, Int32 Hoja, string P1 = "", string P2 = "", string P3 = "")
        {
            string sqlQuery = "";
            if (Hoja == 1)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  --PRS.ID, 
                                  PRS.NAME, 
                                  PRS.TYPE, 
                                  PRS.CEDULA,
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_PERSONAL PRS";
                }
            }
            else if (Hoja == 2)
            {
                if (Tipe == 1)
                {
                }

            }
            return sqlQuery;
        }





   }
}

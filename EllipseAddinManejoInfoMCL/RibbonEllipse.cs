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
using Microsoft.Office.Tools.Excel;
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
        private const string SheetName03 = "RUTAS";
        private const string SheetName04 = "CUMPLI";
        private const string SheetName05 = "DEMORAS";
        private const string SheetName06 = "PROG_PALAS";


        private const string tableName01 = "xA";
        private const string tableName02 = "_01PERSONAL";
        private const string tableName03 = "_01RUTAS";
        private const string tableName04 = "_01CUMPLI";
        private const string tableName05 = "_01DEMORAS";
        private const string tableName06 = "_01PROG_PALAS";

        private const string RangeOne = "Select1";
        public Int32 Tam1 = 0;
        public Int32 Tam2 = 0;
        public Int32 Tam3 = 0;
        public Int32 Tam4 = 0;
        public Int32 Tam5 = 0;
        public Int32 Tam6 = 0;
        public Int32 Tam7 = 0;
        public Int32 Tam8 = 0;
        public Int32 Tam9 = 0;
        public Int32 Tam10 = 0;
        public Int32 Tam11 = 0;
        public Int32 Tam12 = 0;
        public Int32 Tam13 = 0;
        public Int32 Tam14 = 0;
        public Int32 Tam15 = 0;
        //public event EventHandler SelectionChangeCommitted;
        //public event Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler Change;
        //public event Microsoft.Office.Interop.Excel.WorkbookEvents_SheetChangeEventHandler SheetChange;
        Microsoft.Office.Tools.Excel.NamedRange changesRange;
        Microsoft.Office.Tools.Excel.NamedRange changesRange2;
        Microsoft.Office.Tools.Excel.NamedRange changesRange3;
        Microsoft.Office.Tools.Excel.NamedRange changesRange4;
        Microsoft.Office.Tools.Excel.NamedRange changesRange5;
        Microsoft.Office.Tools.Excel.NamedRange changesRange6;
        Microsoft.Office.Tools.Excel.NamedRange changesRange7;
        Microsoft.Office.Tools.Excel.NamedRange changesRange8;
        Microsoft.Office.Tools.Excel.NamedRange changesRange9;
        Microsoft.Office.Tools.Excel.NamedRange changesRange10;
        Microsoft.Office.Tools.Excel.NamedRange changesRange11;
        Microsoft.Office.Tools.Excel.NamedRange changesRange12;
        Microsoft.Office.Tools.Excel.NamedRange changesRange13;

        //private const int titleRow = 8;


        OracleConnection Conexion;

        public object Controls { get; private set; }

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

        public bool EjecutarSql(string sqlQuery)
        {
            //int ConnectionTimeOut = 15;
            //bool PoolingDataBase = true;
            Conexion = new OracleConnection();
            var connectionString = "Data Source=" + DataBase + ";User ID=" + User + ";Password=" + Pw;
            Conexion.ConnectionString = connectionString;
            Conexion.Open();
            //OracleConnection Cmd = Conexion.CreateCommand();
            OracleCommand Cmd = Conexion.CreateCommand();
            Cmd.CommandText = sqlQuery;
            OracleDataReader Datos = Cmd.ExecuteReader();
            Conexion.Close();
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
                Formatear("RUTAS", SheetName03, false);
                Formatear("COMENTARIOS CUMPLIMIENTO", SheetName04, false);
                Formatear("DEMORAS", SheetName05, false);
                Formatear("HORAS PROGRAMADAS PALAS", SheetName06, false);
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
                //_excelApp.ActiveSheet.Select();
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
                dbi.DbUser = VarEncript.Encryption.Decrypt("MIWSjri3U2kHeLVJQtTgpSDsNVy3d+x+8cbt5wy9iw/Vq56wBADnB1PWt1ZXCp7jlSWzLfb4Fzi5zh46FLSXFJeC9w6M1MV62N7bCG19JDdidzqeut+Dqno0pPwSF3GZ");
                dbi.DbPassword = VarEncript.Encryption.Decrypt("M72B5zDxwGSZdaW44Hiu4LbzjiqGD8tfoQKmbw78ELAByRyXujj1tK6FSK2/G6KdZPayz7DCqbT0cfvPxIbJUfju3aH1iTj0l618bm3bnEdAYlSS3whP01s4T6vmSxxw");
                _eFunctions.SetDBSettings(dbi.DbName, dbi.DbUser, dbi.DbPassword);
                _eFunctions.SetConnectionTimeOut(0);
            }
            else
            {
                var dbi = Environments.GetDatabaseItem(Environments.SigmanProductivo);
                dbi.DbUser = VarEncript.Encryption.Decrypt("VDppSMCRaK7ZTG63w9k5WKc3ON0rTcnAf7+eEDM+a+HpZfC3DRpODpJ2KzkZjufVFle/R7LRdw2wLoNTourt1qr96ckLHV4E2uMR+ROoMrLppzAm6xZaiuP7bLRTZm65");
                dbi.DbPassword = VarEncript.Encryption.Decrypt("Wx9o0zzjjw2vjAmhUD/nb/qCTqK9pD6rXg1JxePXdxCnVQXlrAZZAEliXG3O8/yHXtt3TyUrzpGv3YaeBwqnRd02y6ovBHnPny8ikERW2fRXKvDbMxnUC2GIX4dWQjCT");
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
            _excelApp.ScreenUpdating = true;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                borrarTabla();
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
                //if(_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                //{
                List<string> lista = new List<string>();
                //lista.Add("M");
                lista.Add("C");
                lista.Add("U");
                lista.Add("D");
                //for (int F = 0; F < table.Rows.Count; F++)
                //{
                _cells.GetRange((table.Columns.Count), StartRowTable + 1, (table.Columns.Count), (StartRowTable + 1000)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), lista), Type.Missing);
                //}

                /*Excel.Range D = _cells.ActiveSheet.Cells;
                D.Select();
                D.Locked = false;
                D.FormulaHidden = false;
                _cells.GetRange("A6","A9").Select();
                _cells.GetRange("A6", "A9").Locked = true;
                _cells.GetRange("A6", "A9").FormulaHidden = true;
                _excelApp.ActiveSheet.Protect(DrawingObjects: true, Contents: true, Scenarios: true);
                //_cells.GetRange("A:A").FormulaHidden = false;
                _cells.GetRange("G5").Select();
                */
                //}
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

        List<string> ListaDatos(Int32 Tipo, string Param1 = "", string Param2 = "", string Param3 = "unit= '", String ORDEN = "")
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
                          TRIM(PRS.NAME)
                        FROM
                          SIGMAN.APP_PTC_PERSONAL PRS
                        WHERE
                          PRS.TYPE = '" + Param1 + "' ORDER BY ID " + ORDEN);
                table = getdata(Sql, 1);
            }
            else if (Tipo == 5)
            {
                Sql = (@"SELECT
                            hist_statusevents.eqmt--,
                            --DescEv.categoria,
                            --DescEv.[Status],
                            --DescEv.Descripcion
                        FROM
                          PowerView.dbo.hist_statusevents hist_statusevents
                          INNER JOIN PowerView.dbo.hist_eqmtlist hist_eqmtlist on(hist_statusevents.shiftindex=hist_eqmtlist.shiftindex AND hist_statusevents.eqmt=hist_eqmtlist.eqmtid)
                          --LEFT OUTER JOIN  PowerView.dbo.icr_codigoscategoria_200502 DescEv ON hist_statusevents.reason = DescEv.codigo AND hist_statusevents.status = DescEv.statusnum
                        WHERE
                          hist_statusevents.shiftindex = " + Param2 + @" 
                          AND hist_statusevents.starttime = (SELECT MAX(st.starttime) FROM PowerView.dbo.hist_statusevents st WHERE st.shiftindex = " + Param2 + @" AND st.eqmt = hist_statusevents.eqmt)
                          AND hist_statusevents.category Not In(7)
                          --AND hist_statusevents.category = 2
                          AND hist_eqmtlist." + Param3 + Param1 + @"'   
                        ORDER BY 1 " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 6)
            {
                Sql = (@"SELECT
                          TRIM(APP_PTC_RUTA.RUTA)
                        FROM
                          SIGMAN.APP_PTC_RUTA " + ORDEN);
                table = getdata(Sql, 1);
            }
            else if (Tipo == 7)
            {
                Sql = (@"SELECT
                          TRIM(APP_PTC_CUMPLI.ESTADO)
                        FROM
                          SIGMAN.APP_PTC_CUMPLI " + ORDEN);
                table = getdata(Sql, 1);
            }
            else if (Tipo == 8)
            {
                Sql = (@"SELECT
                          MAX(DATEADD(SECOND,(hist_statusevents.endtime+hist_turnos.start),Convert(datetime,hist_turnos.shiftdate,112))) AS Fecha_Hr_End
                        FROM
                          PowerView.dbo.hist_statusevents hist_statusevents
                          INNER JOIN PowerView.dbo.hist_turnos ON(hist_statusevents.shiftindex=hist_turnos.shiftindex)
                                WHERE
                          hist_statusevents.shiftindex =
                          (
                            SELECT
                              MAX(hist_statusevents.shiftindex)
                            FROM
                              PowerView.dbo.hist_statusevents hist_statusevents
                            WHERE
                              hist_statusevents.reason = 421
                              AND hist_statusevents.eqmt = '" + Param1 + @"' 
                          )
                          AND hist_statusevents.eqmt = '" + Param1 + @"' 
                          AND hist_statusevents.reason = 421 " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 9)
            {
                Sql = (@"SELECT
                            ROUND(SUM(hist_statusevents.duration)/3600,2) AS HRS_M_PM      
                            --DATEADD(SECOND, (hist_statusevents.starttime + hist_turnos.start), Convert(datetime, hist_turnos.shiftdate, 112))
                        FROM
                            dbo.hist_statusevents
                            LEFT JOIN dbo.hist_exproot hist_turnos ON (hist_statusevents.shiftindex=hist_turnos.shiftindex )
                        WHERE
                            DATEADD(SECOND, (hist_statusevents.starttime + hist_turnos.start), Convert(datetime, hist_turnos.shiftdate, 112)) >= '" + Param2 + @"' 
                            AND hist_statusevents.eqmt = '" + Param1 + @"'
                            AND hist_statusevents.category IN('2', '5') " + ORDEN);
                table = getdataSql(Sql);
            }
            else if (Tipo == 10)
            {
                Sql = (@"SELECT
                          TRIM(APP_PTC_DEMORAS.DEMORA)
                        FROM
                          SIGMAN.APP_PTC_DEMORAS " + ORDEN);
                table = getdata(Sql,1);
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
            _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.Windows.Application.Sheets[_excelApp.Windows.Application.Sheets.Count]);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = NombreHoja;
            if (CntIndicador == 1)
            {

                var InfoGuardado = new List<string> { "GUARDAR", "ACTUALIZAR", "ELIMINAR", "BUSCAR" };
                _cells.GetCell("A1").Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), InfoGuardado), Type.Missing);
                FormatCamposMenu(_cells.GetRange("A1", "B2"), true, true, true, "GUARDAR", "", 18, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("A1", "B2"));


                FormatCamposMenu(_cells.GetRange("C1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("C1", "S2"));
                FormatCamposMenu(_cells.GetRange("T1", "V4"), true, true, true, "CUMPLIMIENTO PLANES", "", 11, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("T1", "V4"));

                //_cells.SetValidationList(_cells.GetRange(ColResCode, StartRowTable + 1, ColResCode, FinRowForFormat), ListaDatos(3, "ASC"));


                //3 Y 4 FILA DESDE A HASTA S
                FormatCamposMenu(_cells.GetRange("A3", "A4"), true, true, true, "FECHA", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("A3", "A4"));
                FormatCamposMenu(_cells.GetRange("B3", "C4"), true, true, true, DateTime.Now.ToString("yyyy-MM-dd"), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B3", "C4"));
                FormatCamposMenu(_cells.GetRange("D3", "E3"), true, true, true, "SHIFTINDEX", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D3", "E3"));
                FormatCamposMenu(_cells.GetRange("D4", "E4"), true, true, true, ListaDatos(1)[0], "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D4", "E4"));
                FormatCamposMenu(_cells.GetRange("F3", "F4"), true, true, true, "TURNO", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("F3", "F4"));
                FormatCamposMenu(_cells.GetRange("G3", "G4"), true, true, true, ListaDatos(2, _cells.GetCell("D4").Value)[0], "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("G3", "G4"));
                FormatCamposMenu(_cells.GetRange("H3", "H4"), true, true, true, "GRUPO", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("H3", "H4"));
                FormatCamposMenu(_cells.GetRange("I3", "L4"), true, true, true, ListaDatos(3, _cells.GetCell("D4").Value)[0], "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("I3", "L4"));
                FormatCamposMenu(_cells.GetRange("M3", "N3"), true, true, true, "DESPACHADOR:", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
                FormatBordes(_cells.GetRange("M3", "N3"));
                FormatCamposMenu(_cells.GetRange("M4", "N4"), true, true, true, "SUPERVISOR:", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
                FormatBordes(_cells.GetRange("M4", "N4"));
                FormatCamposMenu(_cells.GetRange("O3", "S3"), true, true, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
                _cells.GetCell("O3").Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), ListaDatos(4, "DP")), Type.Missing);
                FormatBordes(_cells.GetRange("O3", "S3"));
                FormatCamposMenu(_cells.GetRange("O4", "S4"), true, true, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0, Aline: "L");
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



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS PERFORADORAS
                FormatCamposMenu(_cells.GetCell("A5"), true, false, true, "TALADROS", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("B5"), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("C5"), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("D5"), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("E5"), true, false, true, "GL-300-600", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("F5"), true, false, true, "HORA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("G5"), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("H5"), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);

                var Perforadora = ListaDatos(5, "Perforadora          ", _cells.GetCell("D4").Value, ORDEN: "ASC");
                var Datos2 = ListaDatos(4, "OP");
                var Datos3 = ListaDatos(6);
                var Datos4 = ListaDatos(7);


                var Datos5 = ListaDatos(10);
                Tam1 = Perforadora.Count;
                Int32 F = 0;
                foreach (var Result in Perforadora)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("A" + (6 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("B" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("C" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("D" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("G" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("A" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("H" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("A" + (6 + F)).Value, _cells.GetCell("G" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                //var Tam1 = F;
                FormatCamposMenu(_cells.GetCell("A" + (F + 6)), true, false, true, Tam1.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(1, (F + 6), 8, (F + 6)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("A" + (F + 7), "C" + (F + 7)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("D" + (F + 7), "H" + (F + 7)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("A5", "H" + (F + 7)));


                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS Tractor de LLantas 854
                FormatCamposMenu(_cells.GetCell("I5"), true, false, true, "LLANTA 854", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("J5"), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("K5"), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("L5"), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("M5"), true, false, true, "GLS 413  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("N5"), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("O5"), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("P5"), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);

                var TractoresLlantas = ListaDatos(5, "T Llantas 854%", _cells.GetCell("D4").Value, "eqmttype LIKE '", ORDEN: "ASC");
                Tam2 = TractoresLlantas.Count;

                F = 0;
                foreach (var Result in TractoresLlantas)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("I" + (6 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("J" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("K" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("L" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("I" + (F + 6)), true, false, true, Tam2.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(9, (F + 6), 16, (F + 6)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("I" + (F + 7), "K" + (F + 7)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("L" + (F + 7), "P" + (F + 7)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("I5", "P" + (F + 6)));





                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS ORUGA D9T
                FormatCamposMenu(_cells.GetCell("Q5"), true, false, true, "ORUGA D9T", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("R5"), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("S5"), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("T5"), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("U5"), true, false, true, "GLS 235  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("V5"), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("W5"), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("X5"), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                var ORUGA_D9T = ListaDatos(5, "TOruga D9T     ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam3 = ORUGA_D9T.Count;

                F = 0;
                foreach (var Result in ORUGA_D9T)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("Q" + (6 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("R" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("S" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("T" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("Q" + (F + 6)), true, false, true, Tam3.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(17, (F + 6), 24, (F + 6)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("Q" + (F + 7), "S" + (F + 7)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("T" + (F + 7), "X" + (F + 7)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("Q5", "X" + (F + 6)));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS ORUGA D8T
                FormatCamposMenu(_cells.GetCell("Y5"), true, false, true, "ORUGA D8T", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Z5"), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AA5"), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AB5"), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AC5"), true, false, true, "GLS 170  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AD5"), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AE5"), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AF5"), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);


                var ORUGA_D8T = ListaDatos(5, "TOruga D8T     ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam4 = ORUGA_D8T.Count;

                F = 0;
                foreach (var Result in ORUGA_D8T)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("Y" + (6 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("Z" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AA" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AB" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("Y" + (F + 6)), true, false, true, Tam4.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(25, (F + 6), 32, (F + 6)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("Y" + (F + 7), "AA" + (F + 7)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AB" + (F + 7), "AF" + (F + 7)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("Y5", "AF" + (F + 6)));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS HIT EX5500
                FormatCamposMenu(_cells.GetCell("AG5"), true, false, true, "HIT EX5500", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AH5"), true, false, true, "OPERADOR", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AI5"), true, false, true, "RUTA", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AJ5"), true, false, true, "CUMP", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AK5"), true, false, true, "GLS 2960 ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AL5"), true, false, true, "HORA  ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AM5"), true, false, true, "HORA PROG. D", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AN5"), true, false, true, "DEMORAS", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AO5"), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AP5"), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);


                var HIT_EX5500 = ListaDatos(5, "Hitachi EX5500 ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam5 = HIT_EX5500.Count;

                F = 0;
                foreach (var Result in HIT_EX5500)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("AG" + (6 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("AH" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AI" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AJ" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    _cells.GetCell("AN" + (6 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos5), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("AG" + (F + 6)), true, false, true, Tam5.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(33, (F + 6), 42, (F + 6)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("AG" + (F + 7), "AI" + (F + 7)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AJ" + (F + 7), "AP" + (F + 7)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("AG5", "AP" + (F + 6)));




                /*
                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS TALADRO
                FormatCamposMenu(_cells.GetRange("A" + (8 + Tam1)), true, false, true, "TALADROS", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("B" + (8 + Tam1), "H" + (8 + Tam1)), true, true, true, "RETANQUEO DE TALADROS", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);

                FormatCamposMenu(_cells.GetCell("A" + (9 + Tam1)), true, false, true, "500-230", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("A" + (10 + Tam1)), true, false, true, "500-251", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("A" + (11 + Tam1)), true, false, true, "500-254", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("A" + (12 + Tam1)), true, false, true, "500-256", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("A" + (13 + Tam1)), true, false, true, "500-258", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);

                Tam14 = 9;
                //FormatCamposMenu(_cells.GetCell("A" + (9 + Tam1)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                _cells.GetRange("B" + (9 + Tam1), "B" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                _cells.GetRange("C" + (9 + Tam1), "C" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                _cells.GetRange("D" + (9 + Tam1), "D" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);



                FormatCamposMenu(_cells.GetCell("A" + (9 + Tam1 + (Tam14 + 1))), true, false, true, "5", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("B" + (9 + Tam1 + (Tam14 + 1)), "H" + (9 + Tam1 + (Tam14 + 1))), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("A" + (9 + Tam1 + (Tam14 + 2)), "C" + (9 + Tam1 + (Tam14 + 2))), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("D" + (9 + Tam1 + (Tam14 + 2)), "H" + (9 + Tam1 + (Tam14 + 2))), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("A5", "H" + (9 + Tam1 + (Tam14 + 2))));

                */





                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS LLANTA 834
                FormatCamposMenu(_cells.GetCell("I" + (Tam2 + 8)), true, false, true, "LLANTA 834", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("J" + (Tam2 + 8)), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("K" + (Tam2 + 8)), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("L" + (Tam2 + 8)), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("M" + (Tam2 + 8)), true, false, true, "GLS 209  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("N" + (Tam2 + 8)), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("O" + (Tam2 + 8)), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("P" + (Tam2 + 8)), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);


                var LL834 = ListaDatos(5, "Tractor 834H   ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam6 = LL834.Count;

                F = 0;
                foreach (var Result in LL834)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("I" + ((Tam2 + 9) + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("J" + ((Tam2 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("K" + ((Tam2 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("L" + ((Tam2 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }

                FormatCamposMenu(_cells.GetCell("I" + ((Tam2 + 9) + F)), true, false, true, Tam6.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(9, ((Tam2 + 9) + F), 16, ((Tam2 + 9) + F)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("I" + ((Tam2 + 9) + F + 1), "K" + ((Tam2 + 9) + F + 1)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("L" + ((Tam2 + 9) + F + 1), "P" + ((Tam2 + 9) + F + 1)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("I5", "P" + ((Tam2 + 9) + F + 1)));




                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS ORUGA D10T
                FormatCamposMenu(_cells.GetCell("Q" + (Tam3 + 8)), true, false, true, "ORUGA D10T", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("R" + (Tam3 + 8)), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("S" + (Tam3 + 8)), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("T" + (Tam3 + 8)), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("U" + (Tam3 + 8)), true, false, true, "GLS 318  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("V" + (Tam3 + 8)), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("W" + (Tam3 + 8)), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("X" + (Tam3 + 8)), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                var ORUGA_D10T = ListaDatos(5, "TOruga D10T    ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam7 = ORUGA_D10T.Count;

                F = 0;
                foreach (var Result in ORUGA_D10T)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("Q" + ((Tam3 + 9) + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("R" + ((Tam3 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("S" + ((Tam3 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("T" + ((Tam3 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("Q" + ((Tam3 + 9) + F)), true, false, true, Tam7.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(17, ((Tam3 + 9) + F), 24, ((Tam3 + 9) + F)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("Q" + ((Tam3 + 9) + F + 1), "S" + ((Tam3 + 9) + F + 1)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("T" + ((Tam3 + 9) + F + 1), "X" + ((Tam3 + 9) + F + 1)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("Q5", "X" + ((Tam3 + 9) + F + 1)));


                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS ORUGA D11T
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + 8))), true, false, true, "ORUGA D11T", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Z" + ((Tam4 + 8))), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AA" + ((Tam4 + 8))), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AB" + ((Tam4 + 8))), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AC" + ((Tam4 + 8))), true, false, true, "GLS 520  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AD" + ((Tam4 + 8))), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AE" + ((Tam4 + 8))), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AF" + ((Tam4 + 8))), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);


                var ORUGA_D11T = ListaDatos(5, "TOruga D11T    ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam8 = ORUGA_D11T.Count;

                F = 0;
                foreach (var Result in ORUGA_D11T)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + 9) + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("Z" + ((Tam4 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AA" + ((Tam4 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AB" + ((Tam4 + 9) + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + 9) + F)), true, false, true, Tam7.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(25, ((Tam4 + 9) + F), 32, ((Tam4 + 9) + F)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("Y" + ((Tam4 + 9) + F + 1), "AA" + ((Tam4 + 9) + F + 1)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AB" + ((Tam4 + 9) + F + 1), "AF" + ((Tam4 + 9) + F + 1)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("Y5", "AF" + ((Tam4 + 9) + F + 1)));




                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS HIT EX3600
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + 8))), true, false, true, "HIT EX3600", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AH" + ((Tam5 + 8))), true, false, true, "OPERADOR", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AI" + ((Tam5 + 8))), true, false, true, "RUTA", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AJ" + ((Tam5 + 8))), true, false, true, "CUMP", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AK" + ((Tam5 + 8))), true, false, true, "GLS 2960 ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AL" + ((Tam5 + 8))), true, false, true, "HORA  ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AM" + ((Tam5 + 8))), true, false, true, "HORA PROG. D", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AN" + ((Tam5 + 8))), true, false, true, "DEMORAS", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AO" + ((Tam5 + 8))), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AP" + ((Tam5 + 8))), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);


                var HIT_EX3600 = ListaDatos(5, "Hit EX3600R    ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam9 = HIT_EX3600.Count;

                F = 0;
                foreach (var Result in HIT_EX3600)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("AG" + +((Tam5 + 9 + F))), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("AH" + ((Tam5 + 9 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AI" + ((Tam5 + 9 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AJ" + ((Tam5 + 9 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    _cells.GetCell("AN" + ((Tam5 + 9 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos5), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + 9 + F))), true, false, true, Tam9.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(33, ((Tam5 + 9 + F)), 42, ((Tam5 + 9 + F))), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("AG" + ((Tam5 + 10 + F)), "AI" + ((Tam5 + 10 + F))), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AJ" + ((Tam5 + 10 + F)), "AP" + ((Tam5 + 10 + F))), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("AG5", "AP" + ((Tam5 + 10 + F))));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS CARGADORES
                FormatCamposMenu(_cells.GetCell("A" + (Tam1 + 8)), true, false, true, "CARGADORES", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("B" + (Tam1 + 8)), true, false, true, "OPERADOR", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("C" + (Tam1 + 8)), true, false, true, "RUTA", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("D" + (Tam1 + 8)), true, false, true, "CUMP", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("E" + (Tam1 + 8)), true, false, true, "GLS 900  ", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("F" + (Tam1 + 8)), true, false, true, "HORA  ", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("G" + (Tam1 + 8)), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("H" + (Tam1 + 8)), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 189, Gf: 89, Bf: 17, Rl: 0, Gl: 0, Bl: 0);

                var CARGADORES = ListaDatos(5, "L1350          ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam10 = CARGADORES.Count;

                F = 0;
                foreach (var Result in CARGADORES)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("A" + (Tam1 + 9 + F)), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("B" + (Tam1 + 9 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("C" + (Tam1 + 9 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("D" + (Tam1 + 9 + F)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("A" + (Tam1 + F + 9)), true, false, true, Tam10.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("B" + (Tam1 + F + 9), "H" + (Tam1 + F + 9)), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("A" + (Tam1 + F + 10), "B" + (Tam1 + F + 10)), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("C" + (Tam1 + 10 + F), "H" + (Tam1 + 10 + F)), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("A" + (Tam1  + 9), "H" + (Tam1 + Tam10 + 9)));




                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS OTROS
                FormatCamposMenu(_cells.GetCell("I" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "OTROS", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("J" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "OPERADOR", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("K" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "RUTA", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("L" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "CUMP", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("M" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "GALONES", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("N" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "HORA  ", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("O" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("P" + ((Tam2 + Tam6) + 8 + 3)), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 255, Gf: 0, Bf: 0, Rl: 0, Gl: 0, Bl: 0);

                Tam15 = 10;

                _cells.GetRange("J" + ((Tam2 + Tam6) + 8 + 1 + 3), "J" + ((Tam2 + Tam6) + 8 + 3 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                _cells.GetRange("K" + ((Tam2 + Tam6) + 8 + 1 + 3), "K" + ((Tam2 + Tam6) + 8 + 3 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                _cells.GetRange("L" + ((Tam2 + Tam6) + 8 + 1 + 3), "L" + ((Tam2 + Tam6) + 8 + 3 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);

                FormatBordes(_cells.GetRange("I" + ((Tam2 + Tam6) + 8 + 3), "P" + ((Tam2 + Tam6) + 8 + 3 + Tam15)));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS GAICO
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "GAICO", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Z" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "OPERADOR", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AA" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "RUTA", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AB" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "CUMP", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AC" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "GALONES", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AD" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "HORA  ", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AE" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AF" + ((Tam4 + Tam8) + 8 + 3)), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 155, Gf: 194, Bf: 230, Rl: 0, Gl: 0, Bl: 0);

                Tam11 = 25;


                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 1 + 3)), true, false, true, "174", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 2 + 3)), true, false, true, "160", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 3 + 3)), true, false, true, "170", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 4 + 3)), true, false, true, "169", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 5 + 3)), true, false, true, "211", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 6 + 3)), true, false, true, "180", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 7 + 3)), true, false, true, "692", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 8 + 3)), true, false, true, "162", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 9 + 3)), true, false, true, "159", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 10 + 3)), true, false, true, "44", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 11 + 3)), true, false, true, "161", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 12 + 3)), true, false, true, "164", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 13 + 3)), true, false, true, "182", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 14 + 3)), true, false, true, "159", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 15 + 3)), true, false, true, "174", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("Y" + ((Tam4 + Tam8) + 8 + 16 + 3)), true, false, true, "171", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);



                _cells.GetRange("Z" + ((Tam4 + Tam8) + 8 + 1 + 3), "Z" + ((Tam4 + Tam8) + 8 + 3 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                _cells.GetRange("AA" + ((Tam4 + Tam8) + 8 + 1 + 3), "AA" + ((Tam4 + Tam8) + 8 + 3 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                _cells.GetRange("AB" + ((Tam4 + Tam8) + 8 + 1 + 3), "AB" + ((Tam4 + Tam8) + 8 + 3 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);

                FormatBordes(_cells.GetRange("Y" + ((Tam4 + Tam8) + 8 + 3), "AF" + ((Tam4 + Tam8) + 8 + 3 + Tam11)));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS PC4000
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "PC4000", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AH" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "OPERADOR", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AI" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "RUTA", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AJ" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "CUMP", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AK" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "GLS 1690 ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AL" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "HORA  ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AM" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "HORA PROG. D", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AN" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "DEMORAS", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AO" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AP" + ((Tam5 + Tam9 + 8 + 3))), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);


                var PC4000 = ListaDatos(5, "PC4000         ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam12 = PC4000.Count;

                F = 0;
                foreach (var Result in PC4000)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("AH" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AI" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AJ" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    _cells.GetCell("AN" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos5), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + 8 + 3 + 1 + F))), true, false, true, Tam12.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange(33, ((Tam5 + Tam9 + 8 + 3 + 1 + F)), 42, ((Tam5 + Tam9 + 8 + 3 + 1 + F))), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("AG" + ((Tam5 + Tam9 + 8 + 3 + 2 + F)), "AI" + ((Tam5 + Tam9 + 8 + 3 + 2 + F))), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AJ" + ((Tam5 + Tam9 + 8 + 3 + 2 + F)), "AP" + ((Tam5 + Tam9 + 8 + 3 + 2 + F))), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("AG5", "AP" + ((Tam5 + Tam9 + Tam12 + 8 + F))));



                //ENCABEZADOS DE TITULOS DE CUMPLIMIENTOS LIEBHERR
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "LIEBHERR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AH" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "OPERADOR", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AI" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "RUTA", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AJ" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "CUMP", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AK" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "GLS 1690 ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AL" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "HORA  ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AM" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "FECHA-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetCell("AN" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3))), true, false, true, "HRS-ULT-TANQ", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);


                var LIEBHERR = ListaDatos(5, "LIE984C        ", _cells.GetCell("D4").Value, "eqmttype = '", ORDEN: "ASC");
                Tam13 = LIEBHERR.Count;

                F = 0;
                foreach (var Result in LIEBHERR)
                {
                    //_cells.GetCell("G" + (6 + F)).NumberFormat = "@";
                    FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3 + 1 + F))), true, false, true, Result, "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    _cells.GetCell("AH" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos2), Type.Missing);
                    _cells.GetCell("AI" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos3), Type.Missing);
                    _cells.GetCell("AJ" + ((Tam5 + Tam9 + Tam12 + 8 + 3 + 3 + 1 + F))).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Datos4), Type.Missing);
                    //FormatCamposMenu(_cells.GetCell("O" + (6 + F)), true, false, true, ListaDatos(8, _cells.GetCell("I" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    //FormatCamposMenu(_cells.GetCell("P" + (6 + F)), true, false, true, ListaDatos(9, _cells.GetCell("I" + (6 + F)).Value, _cells.GetCell("O" + (6 + F)).Value)[0], "", 8, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    F++;
                }
                FormatCamposMenu(_cells.GetCell("AG" + ((Tam5 + Tam9 + Tam12 + 15 + F))), true, false, true, Tam13.ToString(), "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("AH" + ((Tam5 + Tam9 + Tam12 + 15 + F)), "AN" + ((Tam5 + Tam9 + Tam12 + 15 + F))), true, false, true, "", "", 11, Rf: 166, Gf: 166, Bf: 166, Rl: 0, Gl: 0, Bl: 0);
                FormatCamposMenu(_cells.GetRange("AG" + ((Tam5 + Tam9 + Tam12 + 16 + F)), "AH" + ((Tam5 + Tam9 + Tam12 + 16 + F))), true, true, true, "TOTAL SERVICIOS", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatCamposMenu(_cells.GetRange("AI" + ((Tam5 + Tam9 + Tam12 + 16 + F)), "AN" + ((Tam5 + Tam9 + Tam12 + 16 + F))), true, false, true, "", "", 11, Rf: 0, Gf: 32, Bf: 96, Rl: 255, Gl: 255, Bl: 255);
                FormatBordes(_cells.GetRange("AG5", "AN" + ((Tam5 + Tam9 + Tam12 + 16 + F))));




                Recargar();
                //Para detectar el evento de cambio en un rango especifico
                NotifyChanges();
                NotifyChanges2();
                //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();

                _excelApp.ActiveWindow.Zoom = 60;
                _excelApp.Columns.AutoFit();
                _excelApp.Rows.AutoFit();
            }
            if (CntIndicador == 2)
            {
                FormatCamposMenu(_cells.GetRange("B1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B1", "S2"));
            }
            if (CntIndicador == 3)
            {
                FormatCamposMenu(_cells.GetRange("B1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B1", "S2"));
            }
            if (CntIndicador == 4)
            {
                FormatCamposMenu(_cells.GetRange("B1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B1", "S2"));
            }
            if (CntIndicador == 5)
            {
                FormatCamposMenu(_cells.GetRange("B1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B1", "S2"));
            }
            if (CntIndicador == 6)
            {
                FormatCamposMenu(_cells.GetRange("B1", "S2"), true, true, true, Titulo, "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B1", "S2"));
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
            if (Negrita)
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
            if (Merge)
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
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 2), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ExecuteQuery()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show(@"Se ha producido un error: " + ex.Message);
                }
                finally
                {
                    if (_cells != null)
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
                    //_excelApp.DisplayAlerts = true;
                }
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 3), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
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
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Select();
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Lockeed = false;
                        //_excelApp.Worksheets[SheetName02].Application.Cells.FormulaHidden = false;
                        //Excel.Range D = _cells.WorkingWorksheet.Cells;
                        //_cells.GetRange("E" + 6, "E1000").Locked = true;
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
                    //_excelApp.DisplayAlerts = true;
                }
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 4), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
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
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Select();
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Lockeed = false;
                        //_excelApp.Worksheets[SheetName02].Application.Cells.FormulaHidden = false;
                        //Excel.Range D = _cells.WorkingWorksheet.Cells;
                        //_cells.GetRange("E" + 6, "E1000").Locked = true;
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
                    //_excelApp.DisplayAlerts = true;
                }
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 5), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
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
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Select();
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Lockeed = false;
                        //_excelApp.Worksheets[SheetName02].Application.Cells.FormulaHidden = false;
                        //Excel.Range D = _cells.WorkingWorksheet.Cells;
                        //_cells.GetRange("E" + 6, "E1000").Locked = true;
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
                    //_excelApp.DisplayAlerts = true;
                }
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
            {
                try
                {
                    ExecuteQuery(Consulta(1, 6), _excelApp.ActiveWorkbook.ActiveSheet.Name, T: 1);
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
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Select();
                        //_excelApp.Worksheets[SheetName02].Application.Cells.Lockeed = false;
                        //_excelApp.Worksheets[SheetName02].Application.Cells.FormulaHidden = false;
                        //Excel.Range D = _cells.WorkingWorksheet.Cells;
                        //_cells.GetRange("E" + 6, "E1000").Locked = true;
                        _cells.SetCursorDefault();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
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
                    sqlQuery = @"";
                }
            }
            else if (Hoja == 2)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  PRS.ID, 
                                  TRIM(PRS.NAME), 
                                  PRS.TYPE, 
                                  PRS.CEDULA,
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_PERSONAL PRS ORDER BY 1";
                }
            }
            else if (Hoja == 3)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  RUTA.ID,
                                  TRIM(RUTA.RUTA),
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_RUTA RUTA ORDER BY 1 ";
                }

            }
            else if (Hoja == 4)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  CUMPLI.ID,
                                  TRIM(CUMPLI.ESTADO),
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_CUMPLI CUMPLI ORDER BY 1";
                }

            }
            else if (Hoja == 5)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  DEMORAS.ID,
                                  TRIM(DEMORAS.DEMORA),
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_DEMORAS DEMORAS ORDER BY 1";
                }

            }
            else if (Hoja == 6)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  PROG.ID,
                                  PROG.SHIFTINDEX,
                                  PROG.EQUIP_NO,
                                  PROG.HR_PROGRAMAR,
                                  PROG.FECHA,
                                  'M' AS ACCION
                                FROM
                                  SIGMAN.APP_PTC_PROG_PAL PROG ORDER BY 1";
                }

            }
            return sqlQuery;
        }

        private void BtnAcciones_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = true;
            //if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            //{
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(AccionesPers);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:AccionesPers()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                if (_cells != null)
                    _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                _excelApp.ScreenUpdating = true;
                _excelApp.DisplayAlerts = true;
            }
            //}
            //else
            //MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void AccionesPers()
        {
            _excelApp.ScreenUpdating = false;
            _cells.GetCell("A3").Select();
            //if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            //{
            data.DataTable Table;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //Table = getdata(Consulta(1, 1), 1);
                string action = _cells.GetCell("A1").Value;
                ProcesoHoja1(action);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                Table = getdata(Consulta(1, 2), 1);
                ProcesoAcciones(Table);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                Table = getdata(Consulta(1, 3), 1);
                ProcesoAcciones(Table);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
            {
                Table = getdata(Consulta(1, 4), 1);
                ProcesoAcciones(Table);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
            {
                Table = getdata(Consulta(1, 5), 1);
                ProcesoAcciones(Table);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
            {
                Table = getdata(Consulta(1, 6), 1);
                ProcesoAcciones(Table);
            }
            else
            {
                MessageBox.Show(@"Hoja No Definida");
                return;
            }

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                Reload_Info_Select(_cells.GetRange("O3", "S3"), ListaDatos(4, "DP"));
                Reload_Info_Select(_cells.GetRange("O4", "S4"), ListaDatos(4, "SP"));
                var Operadores = ListaDatos(4, "OP");
                //var Ruta = ListaDatos(6);
                //var Cumplimiento = ListaDatos(7);

                //Primera Linea
                _cells.GetRange("B" + 6, "B" + (Tam1 + 5)).Clear();
                _cells.GetRange("B" + 6, "B" + (Tam1 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("B" + 6, "B" + (Tam1 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B" + 6, "B" + (Tam1 + 5)));

                _cells.GetRange("J" + 6, "J" + (Tam2 + 5)).Clear();
                _cells.GetRange("J" + 6, "J" + (Tam2 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("J" + 6, "J" + (Tam2 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("J" + 6, "J" + (Tam2 + 5)));

                _cells.GetRange("R" + 6, "R" + (Tam3 + 5)).Clear();
                _cells.GetRange("R" + 6, "R" + (Tam3 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("R" + 6, "R" + (Tam3 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("R" + 6, "R" + (Tam3 + 5)));
                if (Tam4 != 0)
                {
                    _cells.GetRange("Z" + 6, "Z" + (Tam4 + 5)).Clear();
                    _cells.GetRange("Z" + 6, "Z" + (Tam4 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                    FormatCamposMenu(_cells.GetRange("Z" + 6, "Z" + (Tam4 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    FormatBordes(_cells.GetRange("Z" + 6, "Z" + (Tam4 + 5)));
                }
                _cells.GetRange("AH" + 6, "AH" + (Tam5 + 5)).Clear();
                _cells.GetRange("AH" + 6, "AH" + (Tam5 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AH" + 6, "AH" + (Tam5 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AH" + 6, "AH" + (Tam5 + 5)));


                //Segunda Linea
                /*_cells.GetRange("B" + (9 + Tam1), "B" + (9 + Tam1 + Tam14)).Clear();
                _cells.GetRange("B" + (9 + Tam1), "B" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("B" + (9 + Tam1), "B" + (9 + Tam1 + Tam14)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B" + (9 + Tam1), "B" + (9 + Tam1 + Tam14)));
                */
                _cells.GetRange("J" + (9 + Tam2), "J" + (Tam2 + Tam6 + 8)).Clear();
                _cells.GetRange("J" + (9 + Tam2), "J" + (Tam2 + Tam6 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("J" + (9 + Tam2), "J" + (Tam2 + Tam6 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("J" + (9 + Tam2), "J" + (Tam2 + Tam6 + 8)));


                _cells.GetRange("R" + (9 + Tam3), "R" + (Tam3 + Tam7 + 8)).Clear();
                _cells.GetRange("R" + (9 + Tam3), "R" + (Tam3 + Tam7 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("R" + (9 + Tam3), "R" + (Tam3 + Tam7 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("R" + (9 + Tam3), "R" + (Tam3 + Tam7 + 8)));

                _cells.GetRange("Z" + (9 + Tam4), "Z" + (Tam4 + Tam8 + 8)).Clear();
                _cells.GetRange("Z" + (9 + Tam4), "Z" + (Tam4 + Tam8 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("Z" + (9 + Tam4), "Z" + (Tam4 + Tam8 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("Z" + (9 + Tam4), "Z" + (Tam4 + Tam8 + 8)));

                _cells.GetRange("AH" + (9 + Tam5), "AH" + (Tam5 + Tam9 + 8)).Clear();
                _cells.GetRange("AH" + (9 + Tam5), "AH" + (Tam5 + Tam9 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AH" + (9 + Tam5), "AH" + (Tam5 + Tam9 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AH" + (9 + Tam5), "AH" + (Tam5 + Tam9 + 8)));


                //Tercera Linea
                _cells.GetRange("B" + (Tam1 + 9), "B" + (Tam1 + 8 + Tam10)).Clear();
                _cells.GetRange("B" + (Tam1 + 9), "B" + (Tam1 + 8 + Tam10)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("B" + (Tam1 + 9), "B" + (Tam1 + 8 + Tam10)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("B" + (Tam1 + 9), "B" + (Tam1 + 8 + Tam10)));

                _cells.GetRange("J" + (Tam2 + Tam6 + 12), "J" + (Tam2 + Tam6 + 11 + Tam15)).Clear();
                _cells.GetRange("J" + (Tam2 + Tam6 + 12), "J" + (Tam2 + Tam6 + 11 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("J" + (Tam2 + Tam6 + 12), "J" + (Tam2 + Tam6 + 11 + Tam15)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("J" + (Tam2 + Tam6 + 12), "J" + (Tam2 + Tam6 + 11 + Tam15)));

                _cells.GetRange("Z" + (Tam4 + Tam8 + 12), "Z" + (Tam4 + Tam8 + 11 + Tam11)).Clear();
                _cells.GetRange("Z" + (Tam4 + Tam8 + 12), "Z" + (Tam4 + Tam8 + 11 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("Z" + (Tam4 + Tam8 + 12), "Z" + (Tam4 + Tam8 + 11 + Tam11)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("Z" + (Tam4 + Tam8 + 12), "Z" + (Tam4 + Tam8 + 11 + Tam11)));

                _cells.GetRange("AH" + (Tam5 + Tam9 + 12), "AH" + (Tam5 + Tam9 + 11 + Tam12)).Clear();
                _cells.GetRange("AH" + (Tam5 + Tam9 + 12), "AH" + (Tam5 + Tam9 + 11 + Tam12)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AH" + (Tam5 + Tam9 + 12), "AH" + (Tam5 + Tam9 + 11 + Tam12)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AH" + (Tam5 + Tam9 + 12), "AH" + (Tam5 + Tam9 + 11 + Tam12)));

                _cells.GetRange("AH" + (Tam5 + Tam9 + Tam12 + 15), "AH" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Clear();
                _cells.GetRange("AH" + (Tam5 + Tam9 + Tam12 + 15), "AH" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Operadores), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AH" + (Tam5 + Tam9 + Tam12 + 15), "AH" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AH" + (Tam5 + Tam9 + Tam12 + 15), "AH" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)));



                /*for (int x = 0; x < Tam1; x++)
                {
                    Reload_Info_Select(_cells.GetCell("B" + (6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam2; x++)
                {
                    Reload_Info_Select(_cells.GetCell("J" + (6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam3; x++)
                {
                    Reload_Info_Select(_cells.GetCell("R" + (6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam4; x++)
                {
                    Reload_Info_Select(_cells.GetCell("Z" + (6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam5; x++)
                {
                    Reload_Info_Select(_cells.GetCell("AH" + (6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                
                
                //Segunda Linea
                for (int x = 0; x <= Tam14; x++)
                {
                    Reload_Info_Select(_cells.GetCell("B" + (9 + x + Tam1)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam6; x++)
                {
                    Reload_Info_Select(_cells.GetCell("J" + (9 + x + Tam2)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam7; x++)
                {
                    Reload_Info_Select(_cells.GetCell("R" + (9 + x + Tam3)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam8; x++)
                {
                    Reload_Info_Select(_cells.GetCell("Z" + (9 + x + Tam4)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam9; x++)
                {
                    Reload_Info_Select(_cells.GetCell("AH" + (9 + x + Tam5)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }




                //Tercera Linea
                for (int x = 0; x < Tam10; x++)
                {
                    Reload_Info_Select(_cells.GetCell("B" + (Tam1 + Tam14 + 13 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam15; x++)
                {
                    Reload_Info_Select(_cells.GetCell("J" + (12 + Tam2 + Tam6 + x)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam11; x++)
                {
                    Reload_Info_Select(_cells.GetCell("Z" + (12 + x + Tam4 + Tam8)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam12; x++)
                {
                    Reload_Info_Select(_cells.GetRange("AH" + (12 + x + Tam5 + Tam9)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                for (int x = 0; x < Tam13; x++)
                {
                    Reload_Info_Select(_cells.GetRange("AH" + (15 + x + Tam5 + Tam9 + Tam12)), Operadores, 8, 255, 255, 255, 0, 0, 0);
                }
                */
                _excelApp.ActiveWorkbook.Sheets[SheetName02].Select();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                //var Operadores = ListaDatos(4, "OP");
                var Ruta = ListaDatos(6);
                //var Cumplimiento = ListaDatos(7);


                //Primera Linea
                _cells.GetRange("C" + 6, "C" + (Tam1 + 5)).Clear();
                _cells.GetRange("C" + 6, "C" + (Tam1 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("C" + 6, "C" + (Tam1 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("C" + 6, "C" + (Tam1 + 5)));

                _cells.GetRange("K" + 6, "K" + (Tam2 + 5)).Clear();
                _cells.GetRange("K" + 6, "K" + (Tam2 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("K" + 6, "K" + (Tam2 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("K" + 6, "K" + (Tam2 + 5)));

                _cells.GetRange("S" + 6, "S" + (Tam3 + 5)).Clear();
                _cells.GetRange("S" + 6, "S" + (Tam3 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("S" + 6, "S" + (Tam3 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("S" + 6, "S" + (Tam3 + 5)));
                if (Tam4 != 0)
                {
                    _cells.GetRange("AA" + 6, "AA" + (Tam4 + 5)).Clear();
                    _cells.GetRange("AA" + 6, "AA" + (Tam4 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                    FormatCamposMenu(_cells.GetRange("AA" + 6, "AA" + (Tam4 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    FormatBordes(_cells.GetRange("AA" + 6, "AA" + (Tam4 + 5)));
                }
                _cells.GetRange("AI" + 6, "AI" + (Tam5 + 5)).Clear();
                _cells.GetRange("AI" + 6, "AI" + (Tam5 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AI" + 6, "AI" + (Tam5 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AI" + 6, "AI" + (Tam5 + 5)));


                //Segunda Linea
                /*_cells.GetRange("C" + (9 + Tam1), "C" + (9 + Tam1 + Tam14)).Clear();
                _cells.GetRange("C" + (9 + Tam1), "C" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("C" + (9 + Tam1), "C" + (9 + Tam1 + Tam14)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("C" + (9 + Tam1), "C" + (9 + Tam1 + Tam14)));
                */
                _cells.GetRange("K" + (9 + Tam2), "K" + (Tam2 + Tam6 + 8)).Clear();
                _cells.GetRange("K" + (9 + Tam2), "K" + (Tam2 + Tam6 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("K" + (9 + Tam2), "K" + (Tam2 + Tam6 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("K" + (9 + Tam2), "K" + (Tam2 + Tam6 + 8)));


                _cells.GetRange("S" + (9 + Tam3), "S" + (Tam3 + Tam7 + 8)).Clear();
                _cells.GetRange("S" + (9 + Tam3), "S" + (Tam3 + Tam7 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("S" + (9 + Tam3), "S" + (Tam3 + Tam7 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("S" + (9 + Tam3), "S" + (Tam3 + Tam7 + 8)));

                _cells.GetRange("AA" + (9 + Tam4), "AA" + (Tam4 + Tam8 + 8)).Clear();
                _cells.GetRange("AA" + (9 + Tam4), "AA" + (Tam4 + Tam8 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AA" + (9 + Tam4), "AA" + (Tam4 + Tam8 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AA" + (9 + Tam4), "AA" + (Tam4 + Tam8 + 8)));

                _cells.GetRange("AI" + (9 + Tam5), "AI" + (Tam5 + Tam9 + 8)).Clear();
                _cells.GetRange("AI" + (9 + Tam5), "AI" + (Tam5 + Tam9 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AI" + (9 + Tam5), "AI" + (Tam5 + Tam9 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AI" + (9 + Tam5), "AI" + (Tam5 + Tam9 + 8)));


                //Tercera Linea
                _cells.GetRange("C" + (Tam1 + 9), "C" + (Tam1 + 8 + Tam10)).Clear();
                _cells.GetRange("C" + (Tam1 + 9), "C" + (Tam1 + 8 + Tam10)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("C" + (Tam1 + 9), "C" + (Tam1 + 8 + Tam10)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("C" + (Tam1 + 9), "C" + (Tam1 + 8 + Tam10)));

                _cells.GetRange("K" + (Tam2 + Tam6 + 12), "K" + (Tam2 + Tam6 + 11 + Tam15)).Clear();
                _cells.GetRange("K" + (Tam2 + Tam6 + 12), "K" + (Tam2 + Tam6 + 11 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("K" + (Tam2 + Tam6 + 12), "K" + (Tam2 + Tam6 + 11 + Tam15)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("K" + (Tam2 + Tam6 + 12), "K" + (Tam2 + Tam6 + 11 + Tam15)));

                _cells.GetRange("AA" + (Tam4 + Tam8 + 12), "AA" + (Tam4 + Tam8 + 11 + Tam11)).Clear();
                _cells.GetRange("AA" + (Tam4 + Tam8 + 12), "AA" + (Tam4 + Tam8 + 11 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AA" + (Tam4 + Tam8 + 12), "AA" + (Tam4 + Tam8 + 11 + Tam11)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AA" + (Tam4 + Tam8 + 12), "AA" + (Tam4 + Tam8 + 11 + Tam11)));

                _cells.GetRange("AI" + (Tam5 + Tam9 + 12), "AI" + (Tam5 + Tam9 + 11 + Tam12)).Clear();
                _cells.GetRange("AI" + (Tam5 + Tam9 + 12), "AI" + (Tam5 + Tam9 + 11 + Tam12)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AI" + (Tam5 + Tam9 + 12), "AI" + (Tam5 + Tam9 + 11 + Tam12)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AI" + (Tam5 + Tam9 + 12), "AI" + (Tam5 + Tam9 + 11 + Tam12)));

                _cells.GetRange("AI" + (Tam5 + Tam9 + Tam12 + 15), "AI" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Clear();
                _cells.GetRange("AI" + (Tam5 + Tam9 + Tam12 + 15), "AI" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Ruta), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AI" + (Tam5 + Tam9 + Tam12 + 15), "AI" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AI" + (Tam5 + Tam9 + Tam12 + 15), "AI" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)));

                _excelApp.ActiveWorkbook.Sheets[SheetName03].Select();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
            {
                _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                //var Operadores = ListaDatos(4, "OP");
                //var Ruta = ListaDatos(6);
                var Cumplimiento = ListaDatos(7);

                //Primera Linea
                _cells.GetRange("D" + 6, "D" + (Tam1 + 5)).Clear();
                _cells.GetRange("D" + 6, "D" + (Tam1 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("D" + 6, "D" + (Tam1 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D" + 6, "D" + (Tam1 + 5)));

                _cells.GetRange("L" + 6, "L" + (Tam2 + 5)).Clear();
                _cells.GetRange("L" + 6, "L" + (Tam2 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("L" + 6, "L" + (Tam2 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("L" + 6, "L" + (Tam2 + 5)));

                _cells.GetRange("T" + 6, "T" + (Tam3 + 5)).Clear();
                _cells.GetRange("T" + 6, "T" + (Tam3 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("T" + 6, "T" + (Tam3 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("T" + 6, "T" + (Tam3 + 5)));
                if (Tam4 != 0)
                {
                    _cells.GetRange("AB" + 6, "AB" + (Tam4 + 5)).Clear();
                    _cells.GetRange("AB" + 6, "AB" + (Tam4 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                    FormatCamposMenu(_cells.GetRange("AB" + 6, "AB" + (Tam4 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                    FormatBordes(_cells.GetRange("AB" + 6, "AB" + (Tam4 + 5)));
                }
                _cells.GetRange("AJ" + 6, "AJ" + (Tam5 + 5)).Clear();
                _cells.GetRange("AJ" + 6, "AJ" + (Tam5 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AJ" + 6, "AJ" + (Tam5 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AJ" + 6, "AJ" + (Tam5 + 5)));


                //Segunda Linea
                /*_cells.GetRange("D" + (9 + Tam1), "D" + (9 + Tam1 + Tam14)).Clear();
                _cells.GetRange("D" + (9 + Tam1), "D" + (9 + Tam1 + Tam14)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("D" + (9 + Tam1), "D" + (9 + Tam1 + Tam14)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D" + (9 + Tam1), "D" + (9 + Tam1 + Tam14)));
                */
                _cells.GetRange("L" + (9 + Tam2), "L" + (Tam2 + Tam6 + 8)).Clear();
                _cells.GetRange("L" + (9 + Tam2), "L" + (Tam2 + Tam6 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("L" + (9 + Tam2), "L" + (Tam2 + Tam6 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("L" + (9 + Tam2), "L" + (Tam2 + Tam6 + 8)));


                _cells.GetRange("T" + (9 + Tam3), "T" + (Tam3 + Tam7 + 8)).Clear();
                _cells.GetRange("T" + (9 + Tam3), "T" + (Tam3 + Tam7 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("T" + (9 + Tam3), "T" + (Tam3 + Tam7 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("T" + (9 + Tam3), "T" + (Tam3 + Tam7 + 8)));

                _cells.GetRange("AB" + (9 + Tam4), "AB" + (Tam4 + Tam8 + 8)).Clear();
                _cells.GetRange("AB" + (9 + Tam4), "AB" + (Tam4 + Tam8 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AB" + (9 + Tam4), "AB" + (Tam4 + Tam8 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AB" + (9 + Tam4), "AB" + (Tam4 + Tam8 + 8)));

                _cells.GetRange("AJ" + (9 + Tam5), "AJ" + (Tam5 + Tam9 + 8)).Clear();
                _cells.GetRange("AJ" + (9 + Tam5), "AJ" + (Tam5 + Tam9 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AJ" + (9 + Tam5), "AJ" + (Tam5 + Tam9 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AJ" + (9 + Tam5), "AJ" + (Tam5 + Tam9 + 8)));


                //Tercera Linea
                _cells.GetRange("D" + (Tam1 + 9), "D" + (Tam1 + 8 + Tam10)).Clear();
                _cells.GetRange("D" + (Tam1 + 9), "D" + (Tam1 + 8 + Tam10)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("D" + (Tam1 + 9), "D" + (Tam1 + 8 + Tam10)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("D" + (Tam1 + 9), "D" + (Tam1 + 8 + Tam10)));

                _cells.GetRange("L" + (Tam2 + Tam6 + 12), "L" + (Tam2 + Tam6 + 11 + Tam15)).Clear();
                _cells.GetRange("L" + (Tam2 + Tam6 + 12), "L" + (Tam2 + Tam6 + 11 + Tam15)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("L" + (Tam2 + Tam6 + 12), "L" + (Tam2 + Tam6 + 11 + Tam15)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("L" + (Tam2 + Tam6 + 12), "L" + (Tam2 + Tam6 + 11 + Tam15)));

                _cells.GetRange("AB" + (Tam4 + Tam8 + 12), "AB" + (Tam4 + Tam8 + 11 + Tam11)).Clear();
                _cells.GetRange("AB" + (Tam4 + Tam8 + 12), "AB" + (Tam4 + Tam8 + 11 + Tam11)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AB" + (Tam4 + Tam8 + 12), "AB" + (Tam4 + Tam8 + 11 + Tam11)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AB" + (Tam4 + Tam8 + 12), "AB" + (Tam4 + Tam8 + 11 + Tam11)));

                _cells.GetRange("AJ" + (Tam5 + Tam9 + 12), "AJ" + (Tam5 + Tam9 + 11 + Tam12)).Clear();
                _cells.GetRange("AJ" + (Tam5 + Tam9 + 12), "AJ" + (Tam5 + Tam9 + 11 + Tam12)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AJ" + (Tam5 + Tam9 + 12), "AJ" + (Tam5 + Tam9 + 11 + Tam12)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AJ" + (Tam5 + Tam9 + 12), "AJ" + (Tam5 + Tam9 + 11 + Tam12)));

                _cells.GetRange("AJ" + (Tam5 + Tam9 + Tam12 + 15), "AJ" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Clear();
                _cells.GetRange("AJ" + (Tam5 + Tam9 + Tam12 + 15), "AJ" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AJ" + (Tam5 + Tam9 + Tam12 + 15), "AJ" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AJ" + (Tam5 + Tam9 + Tam12 + 15), "AJ" + (Tam5 + Tam9 + Tam12 + 14 + Tam13)));


                _excelApp.ActiveWorkbook.Sheets[SheetName04].Select();
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
            {
                _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                //var Operadores = ListaDatos(4, "OP");
                //var Ruta = ListaDatos(6);
                var Cumplimiento = ListaDatos(10);


                _cells.GetRange("AN" + 6, "AN" + (Tam5 + 5)).Clear();
                _cells.GetRange("AN" + 6, "AN" + (Tam5 + 5)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AN" + 6, "AN" + (Tam5 + 5)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AN" + 6, "AN" + (Tam5 + 5)));

                _cells.GetRange("AN" + (9 + Tam5), "AN" + (Tam5 + Tam9 + 8)).Clear();
                _cells.GetRange("AN" + (9 + Tam5), "AN" + (Tam5 + Tam9 + 8)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AN" + (9 + Tam5), "AN" + (Tam5 + Tam9 + 8)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AN" + (9 + Tam5), "AN" + (Tam5 + Tam9 + 8)));

                _cells.GetRange("AN" + (Tam5 + Tam9 + 12), "AN" + (Tam5 + Tam9 + 11 + Tam12)).Clear();
                _cells.GetRange("AN" + (Tam5 + Tam9 + 12), "AN" + (Tam5 + Tam9 + 11 + Tam12)).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Cumplimiento), Type.Missing);
                FormatCamposMenu(_cells.GetRange("AN" + (Tam5 + Tam9 + 12), "AN" + (Tam5 + Tam9 + 11 + Tam12)), true, false, true, "", "", 11, Rf: 255, Gf: 255, Bf: 255, Rl: 0, Gl: 0, Bl: 0);
                FormatBordes(_cells.GetRange("AN" + (Tam5 + Tam9 + 12), "AN" + (Tam5 + Tam9 + 11 + Tam12)));



                _excelApp.ActiveWorkbook.Sheets[SheetName05].Select();
            }
            //Nombre = null;
            _excelApp.ScreenUpdating = true;
        }


        private void ProcesoAcciones(DataTable Table)
        {
            ConexionDataBase("SIGMAN");
            /*string[] Nombre = new string[Table.Columns.Count];
            for (Int32 H = 0; H < Table.Columns.Count; H++)
            {
                Nombre[H] = "Var" + H;
            }*/
            _cells.GetRange(Table.Columns.Count + 1, StartRowTable, Table.Columns.Count + 1, StartRowTable).Value = "Resultado";
            _cells.GetRange(Table.Columns.Count + 1, StartRowTable, Table.Columns.Count + 1, StartRowTable).Style = StyleConstants.TitleResult;
            string Var1 = "";
            string Var2 = "";
            string Var3 = "";
            string Var4 = "";
            string Var5 = "";
            Int32 i = 1;
            while (!string.IsNullOrEmpty("" + _cells.GetRange(Table.Columns.Count, StartRowTable + i, Table.Columns.Count, StartRowTable + i).Value))
            {
                string action2 = _cells.GetEmptyIfNull(_cells.GetRange(Table.Columns.Count, StartRowTable + i, Table.Columns.Count, StartRowTable + i).Value);
                try
                {
                    if (_cells.GetRange(Table.Columns.Count, StartRowTable + i, Table.Columns.Count, StartRowTable + i).Value != "M")
                    {
                        string action = _cells.GetEmptyIfNull(_cells.GetRange(Table.Columns.Count, StartRowTable + i, Table.Columns.Count, StartRowTable + i).Value);


                        Var1 = _cells.GetEmptyIfNull(_cells.GetCell("A" + (StartRowTable + i)).Value);
                        Var2 = _cells.GetEmptyIfNull(_cells.GetCell("B" + (StartRowTable + i)).Value);
                        Var3 = _cells.GetEmptyIfNull(_cells.GetCell("C" + (StartRowTable + i)).Value);
                        Var4 = _cells.GetEmptyIfNull(_cells.GetCell("D" + (StartRowTable + i)).Value);
                        Var5 = _cells.GetEmptyIfNull(_cells.GetCell("E" + (StartRowTable + i)).Value);
                        if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                        {
                            if (string.IsNullOrWhiteSpace(action))
                                continue;
                            else if (action.Equals("C"))
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_PERSONAL (NAME, TYPE, CEDULA) VALUES ('" + Var2.Trim().ToUpper() + "', '" + Var3.Trim().ToUpper() + "', '" + Var4.Trim() + "')");
                            }
                            else if (action.Equals("U"))
                            {
                                EjecutarSql("UPDATE SIGMAN.APP_PTC_PERSONAL SET NAME = '" + Var2.Trim().ToUpper() + "', TYPE = '" + Var3.Trim().ToUpper() + "', CEDULA = '" + Var4.Trim() + "' WHERE ID = '" + Var1 + "'");
                            }
                            else if (action.Equals("D"))
                            {
                                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_PERSONAL WHERE ID = '" + Var1 + "' ");
                            }
                        }
                        else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
                        {
                            if (string.IsNullOrWhiteSpace(action))
                                continue;
                            else if (action.Equals("C"))
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_RUTA (RUTA) VALUES ('" + Var2.Trim().ToUpper() + "')");
                            }
                            else if (action.Equals("U"))
                            {
                                EjecutarSql("UPDATE SIGMAN.APP_PTC_RUTA SET RUTA = '" + Var2.Trim().ToUpper() + "' WHERE ID = '" + Var1 + "'");
                            }
                            else if (action.Equals("D"))
                            {
                                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_RUTA WHERE ID = '" + Var1 + "' ");
                            }
                        }
                        else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
                        {
                            if (string.IsNullOrWhiteSpace(action))
                                continue;
                            else if (action.Equals("C"))
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_CUMPLI (ESTADO) VALUES ('" + Var2.Trim().ToUpper() + "')");
                            }
                            else if (action.Equals("U"))
                            {
                                EjecutarSql("UPDATE SIGMAN.APP_PTC_CUMPLI SET ESTADO = '" + Var2.Trim().ToUpper() + "' WHERE ID = '" + Var1 + "'");
                            }
                            else if (action.Equals("D"))
                            {
                                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_CUMPLI WHERE ID = '" + Var1 + "' ");
                            }
                        }
                        else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
                        {
                            if (string.IsNullOrWhiteSpace(action))
                                continue;
                            else if (action.Equals("C"))
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_DEMORAS (DEMORA) VALUES ('" + Var2.Trim().ToUpper() + "')");
                            }
                            else if (action.Equals("U"))
                            {
                                EjecutarSql("UPDATE SIGMAN.APP_PTC_DEMORAS SET DEMORA = '" + Var2.Trim().ToUpper() + "' WHERE ID = '" + Var1 + "'");
                            }
                            else if (action.Equals("D"))
                            {
                                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_DEMORAS WHERE ID = '" + Var1 + "' ");
                            }
                        }
                        else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
                        {
                            if (string.IsNullOrWhiteSpace(action))
                                continue;
                            else if (action.Equals("C"))
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_PROG_PAL (SHIFTINDEX, EQUIP_NO, HR_PROGRAMAR, FECHA) VALUES ('" + Var2.Trim().ToUpper() + "', '" + Var3.Trim().ToUpper() + "', '" + Var4.Trim() + "', '" + Var5.Trim() + "')");
                            }
                            else if (action.Equals("U"))
                            {
                                EjecutarSql("UPDATE SIGMAN.APP_PTC_PROG_PAL SET SHIFTINDEX = '" + Var2.Trim().ToUpper() + "', EQUIP_NO = '" + Var3.Trim().ToUpper() + "', HR_PROGRAMAR = '" + Var4.Trim() + "', FECHA = '" + Var5.Trim() + "' WHERE ID = '" + Var1 + "'");
                            }
                            else if (action.Equals("D"))
                            {
                                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_PROG_PAL WHERE ID = '" + Var1 + "' ");
                            }
                        }
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Value = "OK";
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Success;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Success;
                    }
                    else
                    {
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Value = "NO";
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.TitleInformation;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.TitleInformation;
                    }
                }
                catch (Exception ex)
                {
                    if (_cells.GetCell(StartColTable + 3, i).Value == "   ")
                    {
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Error;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Error;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Value = "Error Save";
                    }
                    else
                    {
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Error;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Style = StyleConstants.Error;
                        _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:AccionesPers()", ex.Message);
                    }
                }
                finally
                {
                    _cells.GetRange(Table.Columns.Count + 1, StartRowTable + i, Table.Columns.Count + 1, StartRowTable + i).Select();
                    i++;
                }
            }
            //}
            if (_cells != null) _cells.SetCursorDefault();
        }



        private void ProcesoHoja1(string action)
        {
            var SW = 0;
            ConexionDataBase("SIGMAN");
            try
            {
                if (action == "GUARDAR")
                {
                    
                    /*string[] Nombre = new string[Table.Columns.Count];
                    for (Int32 H = 0; H < Table.Columns.Count; H++)
                    {
                        Nombre[H] = "Var" + H;
                    }*/
                    Excel.Worksheet SheetOne = _excelApp.ActiveWorkbook.Sheets[SheetName01];
                    string Var1 = "";
                    string Var2 = "";
                    string Var3 = "";
                    string Var4 = "";
                    string Var5 = "";
                    string Var6 = "";
                    string[,] DatosWoG;
                    //string action = _cells.GetCell("A1").Value;
                    Int32 TamVar;
                    var FilaInicial = 6;
                    Int32 Filas;
                    Int32 Columnas;
                    Var1 = _cells.GetEmptyIfNull(_cells.GetRange(4, 4, 4, 4).Value);
                    Var2 = _cells.GetEmptyIfNull(_cells.GetRange(2, 3, 2, 3).Value);
                    Var3 = _cells.GetEmptyIfNull(_cells.GetRange(7, 3, 7, 3).Value);
                    Var4 = _cells.GetEmptyIfNull(_cells.GetRange(9, 3, 9, 3).Value);
                    Var5 = _cells.GetEmptyIfNull(_cells.GetRange(15, 3, 15, 3).Value);
                    Var6 = _cells.GetEmptyIfNull(_cells.GetRange(15, 4, 15, 4).Value);
                    var F1 = Var2.Substring(0, 4);
                    var F2 = Var2.Substring(5, 2);
                    var F3 = Var2.Substring(8, 2);
                    Var2 = F1 + F2 + F3;

                    TamVar = Tam1;
                    DatosWoG = GetStringArray(SheetOne.Range["A" + FilaInicial, "F" + (TamVar + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam2;
                    DatosWoG = GetStringArray(SheetOne.Range["I" + FilaInicial, "N" + (TamVar + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam3;
                    DatosWoG = GetStringArray(SheetOne.Range["Q" + FilaInicial, "V" + (TamVar + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    if (Tam4 != 0)
                    {
                        TamVar = Tam4;
                        DatosWoG = GetStringArray(SheetOne.Range["Y" + FilaInicial, "AD" + (TamVar + (FilaInicial - 1))].Cells.Value2);
                        //string[] DatosWo = DatosWoG.Distinct().ToArray();
                        Filas = DatosWoG.GetLength(0);
                        Columnas = DatosWoG.GetUpperBound(1);
                        SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);
                    }

                    TamVar = Tam5;
                    DatosWoG = GetStringArray(SheetOne.Range["AG" + FilaInicial, "AN" + (TamVar + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 2, Var5, Var6);

                    //Segundo Nivel
                    /*TamVar = Tam14 + Tam1;
                    DatosWoG = GetStringArray(SheetOne.Range["A" + (FilaInicial + Tam1 + 3), "F" + (TamVar + 4 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);
                    */
                    TamVar = Tam6 + Tam2;
                    DatosWoG = GetStringArray(SheetOne.Range["I" + (FilaInicial + Tam2 + 3), "N" + (TamVar + 3 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam7 + Tam3;
                    DatosWoG = GetStringArray(SheetOne.Range["Q" + (FilaInicial + Tam3 + 3), "V" + (TamVar + 3 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam8 + Tam4;
                    DatosWoG = GetStringArray(SheetOne.Range["Y" + (FilaInicial + Tam4 + 3), "AD" + (TamVar + 3 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);


                    TamVar = Tam9 + Tam5;
                    DatosWoG = GetStringArray(SheetOne.Range["AG" + (FilaInicial + Tam5 + 3), "AN" + (TamVar + 3 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 2, Var5, Var6);




                    //Tercer Nivel
                    TamVar = Tam1 + Tam10;
                    DatosWoG = GetStringArray(SheetOne.Range["A" + (FilaInicial + (Tam1) + 3), "F" + (TamVar + 3 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam2 + Tam6 + Tam15;
                    DatosWoG = GetStringArray(SheetOne.Range["I" + (FilaInicial + (Tam2 + Tam6) + 6), "N" + (TamVar + 6 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam4 + Tam8 + Tam11;
                    DatosWoG = GetStringArray(SheetOne.Range["Y" + (FilaInicial + (Tam4 + Tam8) + 6), "AD" + (TamVar + 6 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);

                    TamVar = Tam5 + Tam9 + Tam12;
                    DatosWoG = GetStringArray(SheetOne.Range["AG" + (FilaInicial + (Tam5 + Tam9) + 6), "AN" + (TamVar + 6 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 2, Var5, Var6);

                    TamVar = Tam5 + Tam9 + Tam12 + Tam13;
                    DatosWoG = GetStringArray(SheetOne.Range["AG" + (FilaInicial + (Tam5 + Tam9 + Tam12) + 9), "AL" + (TamVar + 9 + (FilaInicial - 1))].Cells.Value2);
                    //string[] DatosWo = DatosWoG.Distinct().ToArray();
                    Filas = DatosWoG.GetLength(0);
                    Columnas = DatosWoG.GetUpperBound(1);
                    SaveOneSheet(action, Filas, Columnas, Var1, Var2, Var3, Var4, DatosWoG, 1, Var5, Var6);
                }
                else if(action.Equals("ELIMINAR"))
                {
                    SaveOneSheet(action, 0, 0, _cells.GetRange(4, 4, 4, 4).Value2, null, null, null, null, 1, null, null);
                }
            }
            catch (Exception ex)
            {
                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_COMBUSTIBLE WHERE SHIFTINDEX = " + _cells.GetRange(4, 4, 4, 4).Value + " ");
                //Debugger.LogError("RibbonEllipse.cs:AccionesPers()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
                SW = 1;
            }
            finally
            {
                if(SW == 0)
                {
                    MessageBox.Show(_cells.GetCell("A1").Value2 + @" Exitoso. ");
                }
            }

        }

        private void SaveOneSheet(string action, Int32 Filas, Int32 Columnas, string Var1, string Var2, string Var3, string Var4, string[,] DatosWoG, Int32 tipe = 1, string Var5 = "", string Var6 = "")
        {
            if (action.Equals("GUARDAR"))
            {
                if (tipe == 1)
                {
                    for (int i = 0; i < Filas; i++)
                    {
                        for (int j = 0; j < 1; j++)
                        {
                            if (DatosWoG[i, j] != null)
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_COMBUSTIBLE (SHIFTINDEX, FECHA, TURNO, GRUPO, EQUIP_NO, OPERADOR, RUTA, CUMPLIMIENTO, GALONES, HORA, DESPACHADOR, SUPERVISOR) VALUES ('" + Var1 + "', '" + Var2.Trim().ToUpper() + "', '" + Var3.Trim() + "', '" + Var4.Trim() + "', '" + DatosWoG[i, j] + "', '" + DatosWoG[i, j + 1] + "', '" + DatosWoG[i, j + 2] + "', '" + DatosWoG[i, j + 3] + "', '" + DatosWoG[i, j + 4] + "', '" + DatosWoG[i, j + 5] + "', '" + Var5 + "', '" + Var6 + "')");

                            }
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < Filas; i++)
                    {
                        for (int j = 0; j < 1; j++)
                        {
                            if (DatosWoG[i, j] != null)
                            {
                                EjecutarSql("INSERT INTO SIGMAN.APP_PTC_COMBUSTIBLE (SHIFTINDEX, FECHA, TURNO, GRUPO, EQUIP_NO, OPERADOR, RUTA, CUMPLIMIENTO, GALONES, HORA, HORA_PROG, DEMORA, DESPACHADOR, SUPERVISOR) VALUES ('" + Var1 + "', '" + Var2.Trim().ToUpper() + "', '" + Var3.Trim() + "', '" + Var4.Trim() + "', '" + DatosWoG[i, j] + "', '" + DatosWoG[i, j + 1] + "', '" + DatosWoG[i, j + 2] + "', '" + DatosWoG[i, j + 3] + "', '" + DatosWoG[i, j + 4] + "', '" + DatosWoG[i, j + 5] + "', '" + DatosWoG[i, j + 6] + "', '" + DatosWoG[i, j + 7] + "', '" + Var5 + "', '" + Var6 + "')");

                            }
                        }
                    }
                }
            }
            else if (action.Equals("ACTUALIZAR"))
            {
                for (int i = 0; i < Filas; i++)
                {
                    for (int j = 0; j < 1; j++)
                    {
                        EjecutarSql("UPDATE SIGMAN.APP_PTC_COMBUSTIBLE SET NAME = '" + Var2.Trim().ToUpper() + "', TYPE = '" + Var3.Trim().ToUpper() + "', CEDULA = '" + Var4.Trim() + "' WHERE ID = '" + Var1 + "'");
                    }
                }
            }
            else if (action.Equals("ELIMINAR"))
            {
                EjecutarSql("DELETE FROM SIGMAN.APP_PTC_COMBUSTIBLE WHERE SHIFTINDEX = '" + Var1 + "' ");
            }
        }

        private void bLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            Limpieza();
        }


        private void Limpieza()
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _cells.DeleteTableRange(tableName02);
            }
        }


        public void borrarTabla()
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //Excel._Worksheet _cells2 = _excelApp.ActiveWorkbook.ActiveSheet;
                //Excel._Worksheet Hoja = _excelApp.ActiveWorkbook.Sheets[Name_Hoja];
                //Hoja.ListObjects(tableName01);
                //_cells.DeleteTableRange(_excelApp.ActiveWorkbook.Sheets[Name_Hoja].Table.Name);
                _cells.DeleteTableRange(tableName01);
                return;
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _cells.GetRange("E" + 6, "E1005").Clear();
                _cells.DeleteTableRange(tableName02);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                _cells.GetRange("C" + 6, "C1005").Clear();
                _cells.DeleteTableRange(tableName03);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
            {
                _cells.GetRange("C" + 6, "C1005").Clear();
                _cells.DeleteTableRange(tableName04);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
            {
                _cells.GetRange("C" + 6, "C1005").Clear();
                _cells.DeleteTableRange(tableName05);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
            {
                _cells.GetRange("F" + 6, "F1005").Clear();
                _cells.DeleteTableRange(tableName06);
            }

        }




        /// <summary>
        /// Establece el resultado de búsqueda de la descripción de un equipo después de que este es escrita
        /// </summary>
        /// <param name="target"></param>
        /// <param name="changedRanges"></param>
        void GetTableChangedValue(Excel.Range target, ListRanges changedRanges)//Excel.Range target)
        {
            //_eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            switch (target.Column)
            {
                case 3://Equipo
                    try
                    {
                        if (string.IsNullOrWhiteSpace("" + target.Value))
                        {
                            _cells.GetCell(target.Column + 1, target.Row).Value = "";
                            break;
                        }

                        _cells.GetCell(target.Column + 1, target.Row).Value = "Buscando Equipo...";
                        string description = "Culo";

                        _cells.GetCell(target.Column + 1, target.Row).Value = !string.IsNullOrWhiteSpace(description) ? description.Trim() : "Equipo no encontrado";
                        _cells.GetCell(target.Column + 1, target.Row).Columns.AutoFit();
                    }
                    catch (NullReferenceException ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(target.Column + 1, target.Row).Value = "No fue Posible Obtener Informacion!";
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(target.Column + 1, target.Row).Value = "No fue Posible Obtener Informacion!";
                    }
                    break;
                case 5://Estadística
                    try
                    {
                        var equipNo = "" + _cells.GetCell(3, target.Row).Value;
                        var statType = "" + "Chucha";

                        var statDate = "" + _cells.GetCell(1, target.Row).Value;

                        if (string.IsNullOrWhiteSpace(equipNo) || string.IsNullOrWhiteSpace(statType) || string.IsNullOrWhiteSpace(statDate))
                        {
                            _cells.GetCell(7, target.Row).Value = "No fue Posible Obtener Información";
                            _cells.GetCell(8, target.Row).Value = "No fue Posible Obtener Información";
                        }
                        else
                        {
                            var lastStatReg = "PERRA";

                            _cells.GetCell(7, target.Row).Value = !string.IsNullOrWhiteSpace(lastStatReg) ? lastStatReg.Trim() : "";
                            _cells.GetCell(8, target.Row).Value = !string.IsNullOrWhiteSpace(lastStatReg) ? lastStatReg.Trim() : "";

                        }
                        _cells.GetCell(7, target.Row).Columns.AutoFit();
                        _cells.GetCell(8, target.Row).Columns.AutoFit();
                    }
                    catch (NullReferenceException ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(7, target.Row).Value = "";
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:GetTableChangedValue()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(7, target.Row).Value = "Se ha producido un error";
                    }
                    break;
            }
        }


        private void btnRestoreEvents_Click(object sender, RibbonControlEventArgs e)
        {
            RestoreEvents();
        }

        public void RestoreEvents()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            var table = _cells.GetRange(RangeOne).Worksheet.ListObjects[RangeOne];
            var tableObject = Globals.Factory.GetVstoObject(table);
            try
            {
                tableObject.Change -= GetTableChangedValue;
            }
            catch
            {
                //ignored
            }
            tableObject.Change += GetTableChangedValue;

        }



        private void NotifyChanges()
        {
            _excelApp.Visible = false;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;
          
            //_excelApp.ActiveSheet.Names.Add(Name: "compositeRange", RefersToR1C1: _cells.GetRange("O3", "S3"));
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[SheetName01]);
            //changesRange = worksheet.Controls.AddNamedRange(_cells.GetRange(StartColTable + 1, StartRowTable + 1, StartColTable + 15, (StartRowTable + Tam2)), "RangoTaladros");
            changesRange = worksheet.Controls.AddNamedRange(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[SheetName01].Columns, "RangoTaladros");
            //changesRange = worksheet.Application.Worksheets[SheetName01];
            changesRange.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change);
            _excelApp.ScreenUpdating = true;
        }

        private void NotifyChanges2()
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;

            //_excelApp.ActiveSheet.Names.Add(Name: "compositeRange", RefersToR1C1: _cells.GetRange("O3", "S3"));
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[SheetName01]);
            
            //Palas
            changesRange = worksheet.Controls.AddNamedRange(_cells.GetCell("AG" + (6 + Tam5)), "RangoTaladros2");
            changesRange2 = worksheet.Controls.AddNamedRange(_cells.GetCell("AG" + (6 + Tam5 + Tam9 + 3)), "RangoTaladros3");
            changesRange3 = worksheet.Controls.AddNamedRange(_cells.GetCell("AG" + (6 + Tam5 + Tam9 + Tam12 + 6)), "RangoTaladros4");

            //Equipos Aux
            changesRange4 = worksheet.Controls.AddNamedRange(_cells.GetCell("A" + (6 + Tam1)), "RangoTaladros5");
            changesRange5 = worksheet.Controls.AddNamedRange(_cells.GetCell("A" + (6 + Tam1 + Tam10 + 3)), "RangoTaladros6");
            changesRange6 = worksheet.Controls.AddNamedRange(_cells.GetCell("I" + (6 + Tam2)), "RangoTaladros7");
            changesRange7 = worksheet.Controls.AddNamedRange(_cells.GetCell("I" + (6 + Tam2 + Tam6 + 3)), "RangoTaladros8");
            changesRange8 = worksheet.Controls.AddNamedRange(_cells.GetCell("Q" + (6 + Tam3)), "RangoTaladros9");
            changesRange9 = worksheet.Controls.AddNamedRange(_cells.GetCell("Q" + (6 + Tam3 + Tam7 + 3)), "RangoTaladros10");
            changesRange10 = worksheet.Controls.AddNamedRange(_cells.GetCell("Y" + (6 + Tam4)), "RangoTaladros11");
            changesRange11 = worksheet.Controls.AddNamedRange(_cells.GetCell("Y" + (6 + Tam4 + Tam8 + 3)), "RangoTaladros12");
            changesRange12 = worksheet.Controls.AddNamedRange(_cells.GetCell("AG" + (6 + Tam5 + Tam9 + Tam12 + 6)), "RangoTaladros13");

            changesRange13 = worksheet.Controls.AddNamedRange(_cells.GetRange("A" + (6), "A" + (5 + Tam1)), "RangoTaladros14");


            //Palas
            changesRange.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange2.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange3.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            //Equipos Aux
            changesRange4.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange5.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange6.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange7.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange8.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange9.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange10.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange11.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange12.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);
            changesRange13.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change2);


            _excelApp.ScreenUpdating = true;
        }

        void changesRange_Change2(Excel.Range Target)
        {
            //string cellAddress = Target.get_Address(Excel.XlReferenceStyle.xlA1);
            //MessageBox.Show("Cell " + cellAddress + " changed.");
            Recargar();
        }


        void changesRange_Change(Excel.Range Target)
        {
            //string cellAddress = Target.get_Address(Excel.XlReferenceStyle.xlA1);
            //MessageBox.Show("Cell " + cellAddress + " changed.");
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();
        }



        private void Recargar()
        {
            //Palas
            Int32 RowInicial = 6;
            _cells.GetCell("Y1").NumberFormat = "0";
            _cells.GetCell("Y1").FormulaLocal = "=COUNTA(AG" + (RowInicial) + ":AG" + (RowInicial + (Tam5 - 1)) + ") + " + "COUNTA(AG" + (RowInicial + Tam5 + 3) + ":AG" + (RowInicial + (Tam5 + Tam9 + 2)) + ") +" + "COUNTA(AG" + (RowInicial + Tam5 + Tam9 + 6) + ":AG" + (RowInicial + Tam5 + Tam9 + Tam12 + 5) + ")";


            _cells.GetCell("AA1").NumberFormat = "0";
            _cells.GetCell("AA1").FormulaLocal = "=COUNTIF(AJ" + (RowInicial) + ":AJ" + (RowInicial + (Tam5 - 1)) + "," + @"""SIN COMB""" + ") + " + "COUNTIF(AJ" + (RowInicial + Tam5 + 3) + ":AJ" + (RowInicial + (Tam5 + Tam9 + 2)) + ", " + @"""SIN COMB""" + ") +" + "COUNTIF(AJ" + (RowInicial + Tam5 + Tam9 + 6) + ":AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 5) + ", " + @"""SIN COMB""" + ")";

            _cells.GetCell("AA2").NumberFormat = "0";
            _cells.GetCell("AA2").FormulaLocal = "=COUNTIF(D" + (RowInicial) + ":D" + (RowInicial + (Tam1 - 1)) + "," + @"""SIN COMB""" + ") + " + "COUNTIF(L" + (RowInicial) + ":L" + (RowInicial + (Tam2 - 1)) + ", " + @"""SIN COMB""" + ") +" + "COUNTIF(T" + (RowInicial) + ":T" + (RowInicial + (Tam3 - 1)) + ", " + @"""SIN COMB""" + ")+ " + "COUNTIF(AB" + (RowInicial) + ":AB" + (RowInicial + (Tam4)) + ", " + @"""SIN COMB""" + ")+" + "COUNTIF(L" + (RowInicial + Tam2 + 3) + ":L" + (RowInicial + (Tam2 + Tam6 + 2)) + ", " + @"""SIN COMB""" + ") + " + "COUNTIF(T" + (RowInicial + Tam3 + 3) + ":T" + (RowInicial + (Tam3 + Tam7 + 2)) + ", " + @"""SIN COMB""" + ") +" + "COUNTIF(AB" + (RowInicial + Tam4 + 3) + ":AB" + (RowInicial + (Tam4 + Tam8 + 2)) + ", " + @"""SIN COMB""" + ") +" + "COUNTIF(AB" + (RowInicial + Tam4 + Tam8 + 6) + ":AB" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ", " + @"""SIN COMB""" + ") +" + "COUNTIF(AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 9) + ":AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 8) + ", " + @"""SIN COMB""" + ")";

            _cells.GetCell("AA3").NumberFormat = "0";
            _cells.GetCell("AA3").FormulaLocal = "=COUNTIF(D" + (RowInicial + Tam1 + 3) + ":D" + (RowInicial + (Tam1 + Tam10 + 2)) + "," + @"""SIN COMB""" + ")";

            _cells.GetCell("AA4").NumberFormat = "0";
            _cells.GetCell("AA4").FormulaLocal = "=COUNTIF(AB" + (RowInicial + Tam4 + Tam8 + 6) + ":AB" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ", " + @"""SIN COMB""" + ")";

            //Cumplidoss
            _cells.GetCell("AC1").NumberFormat = "0";
            _cells.GetCell("AC1").FormulaLocal = "=COUNTA(AK" + (RowInicial) + ":AK" + (RowInicial + (Tam5 - 1)) + ") + " + "COUNTA(AK" + (RowInicial + Tam5 + 3) + ":AK" + (RowInicial + (Tam5 + Tam9 + 2)) + ") +" + "COUNTA(AK" + (RowInicial + Tam5 + Tam9 + 6) + ":AK" + (RowInicial + Tam5 + Tam9 + Tam12 + 5) + ") - AA1";

            _cells.GetCell("AC2").NumberFormat = "0";
            _cells.GetCell("AC2").FormulaLocal = "=COUNTA(E" + (RowInicial) + ":E" + (RowInicial + (Tam1 - 1)) + ") + " + "COUNTA(M" + (RowInicial) + ":M" + (RowInicial + (Tam2 - 1)) + ") +" + "COUNTA(U" + (RowInicial) + ":U" + (RowInicial + (Tam3 - 1)) + ")+ " + "COUNTA(AC" + (RowInicial) + ":AC" + (RowInicial + (Tam4)) + ")+" + "COUNTA(M" + (RowInicial + Tam2 + 3) + ":M" + (RowInicial + (Tam2 + Tam6 + 2)) + ") + " + "COUNTA(U" + (RowInicial + Tam3 + 3) + ":U" + (RowInicial + (Tam3 + Tam7 + 2)) + ") +" + "COUNTA(AC" + (RowInicial + Tam4 + 3) + ":AC" + (RowInicial + (Tam4 + Tam8 + 2)) + ") +" + "COUNTA(AC" + (RowInicial + Tam4 + Tam8 + 6) + ":AC" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ") +" + "COUNTA(AK" + (RowInicial + Tam5 + Tam9 + Tam12 + 9) + ":AK" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 8) + ") - AA2";

            _cells.GetCell("AC3").NumberFormat = "0";
            _cells.GetCell("AC3").FormulaLocal = "=COUNTA(E" + (RowInicial + Tam1 + 3) + ":E" + (RowInicial + (Tam1 + Tam10 + 2)) + ") - AA3";

            _cells.GetCell("AC4").NumberFormat = "0";
            _cells.GetCell("AC4").FormulaLocal = "=COUNTA(AC" + (RowInicial + Tam4 + Tam8 + 6) + ":AC" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ") - AA4";

            //Equipos Aux
            _cells.GetCell("Y2").NumberFormat = "0";
            _cells.GetCell("Y2").FormulaLocal = "=COUNTA(A" + (RowInicial) + ":A" + (RowInicial + (Tam1 - 1)) + ") + " + "COUNTA(I" + (RowInicial) + ":I" + (RowInicial + (Tam2 - 1)) + ") +" + "COUNTA(Q" + (RowInicial) + ":Q" + (RowInicial + (Tam3 - 1)) + ")+ " + "COUNTA(Y" + (RowInicial) + ":Y" + (RowInicial + (Tam4)) + ")+" + "COUNTA(I" + (RowInicial + Tam2 + 3) + ":I" + (RowInicial + (Tam2 + Tam6 + 2)) + ") + " + "COUNTA(Q" + (RowInicial + Tam3 + 3) + ":Q" + (RowInicial + (Tam3 + Tam7 + 2)) + ") +" + "COUNTA(Y" + (RowInicial + Tam4 + 3) + ":Y" + (RowInicial + (Tam4 + Tam8 + 2)) + ") +" + "COUNTA(Y" + (RowInicial + Tam4 + Tam8 + 6) + ":Y" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ") +" + "COUNTA(AG" + (RowInicial + Tam5 + Tam9 + Tam12 + 9) + ":AG" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 8) + ")";

            //Equipos Cargadores
            //Int32 P13 = Convert.ToInt32(string.IsNullOrWhiteSpace(_cells.GetCell("A" + (RowInicial + Tam1 + Tam10 + 3)).Value) ? "0" : _cells.GetCell("A" + (RowInicial + Tam1 + Tam10 + 3)).Value);
            _cells.GetCell("Y3").NumberFormat = "0";
            _cells.GetCell("Y3").FormulaLocal = "=COUNTA(A" + (RowInicial + Tam1 + 3) + ":A" + (RowInicial + (Tam1 + Tam10 + 2)) + ")";

            //Equipos Gaico
            //Int32 P14 = Convert.ToInt32(string.IsNullOrWhiteSpace(_cells.GetCell("Y" + (RowInicial + Tam1 + Tam10 + 3)).Value) ? "0" : _cells.GetCell("Y" + (RowInicial + Tam1 + Tam10 + 3)).Value);
            _cells.GetCell("Y4").NumberFormat = "0";
            _cells.GetCell("Y4").FormulaLocal = "=COUNTA(Y" + (RowInicial + Tam4 + Tam8 + 6) + ":Y" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ")";


            //Cumplimiento Cuenta
            //PALAS
            _cells.GetCell("AJ" + (RowInicial + Tam5)).NumberFormat = "0";
            _cells.GetCell("AJ" + (RowInicial + Tam5)).FormulaLocal = "=COUNTIF(AJ" + (RowInicial) + ":AJ" + (RowInicial + (Tam5 - 1)) + "," + @"""SIN COMB""" + ")";

            _cells.GetCell("AJ" + (RowInicial + (Tam5 + Tam9 + 3))).NumberFormat = "0";
            _cells.GetCell("AJ" + (RowInicial + (Tam5 + Tam9 + 3))).FormulaLocal = "=COUNTIF(AJ" + (RowInicial + Tam5 + 3) + ":AJ" + (RowInicial + (Tam5 + Tam9 + 2)) + "," + @"""SIN COMB""" + ")";

            _cells.GetCell("AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 6)).NumberFormat = "0";
            _cells.GetCell("AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 6)).FormulaLocal = "=COUNTIF(AJ" + (RowInicial + Tam5 + Tam9 + 6) + ":AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 5) + "," + @"""SIN COMB""" + ")";
            //AUX
            _cells.GetCell("D" + (RowInicial + Tam1)).NumberFormat = "0";
            _cells.GetCell("D" + (RowInicial + Tam1)).FormulaLocal = "=COUNTIF(D" + (RowInicial) + ":D" + (RowInicial + (Tam1 - 1)) + "," + @"""SIN COMB""" + ")";
            _cells.GetCell("L" + (RowInicial + Tam2)).NumberFormat = "0";
            _cells.GetCell("L" + (RowInicial + Tam2)).FormulaLocal = "=COUNTIF(L" + (RowInicial) + ":L" + (RowInicial + (Tam2 - 1)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("T" + (RowInicial + Tam3)).NumberFormat = "0";
            _cells.GetCell("T" + (RowInicial + Tam3)).FormulaLocal = "=COUNTIF(T" + (RowInicial) + ":T" + (RowInicial + (Tam3 - 1)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("AB" + (RowInicial + Tam4)).NumberFormat = "0";
            _cells.GetCell("AB" + (RowInicial + Tam4)).FormulaLocal = "=COUNTIF(AB" + (RowInicial) + ":AB" + (RowInicial + (Tam4)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("L" + (RowInicial + Tam2 + Tam6 + 3)).NumberFormat = "0";
            _cells.GetCell("L" + (RowInicial + Tam2 + Tam6 + 3)).FormulaLocal = "=COUNTIF(L" + (RowInicial + Tam2 + 3) + ":L" + (RowInicial + (Tam2 + Tam6 + 2)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("T" + (RowInicial + (Tam3 + Tam7 + 3))).NumberFormat = "0";
            _cells.GetCell("T" + (RowInicial + (Tam3 + Tam7 + 3))).FormulaLocal = "=COUNTIF(T" + (RowInicial + Tam3 + 3) + ":T" + (RowInicial + (Tam3 + Tam7 + 2)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("AB" + (RowInicial + (Tam4 + Tam8 + 3))).NumberFormat = "0";
            _cells.GetCell("AB" + (RowInicial + (Tam4 + Tam8 + 3))).FormulaLocal = "=COUNTIF(AB" + (RowInicial + Tam4 + 3) + ":AB" + (RowInicial + (Tam4 + Tam8 + 2)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("AB" + (RowInicial + (Tam4 + Tam8 + Tam11 + 6))).NumberFormat = "0";
            _cells.GetCell("AB" + (RowInicial + (Tam4 + Tam8 + Tam11 + 6))).FormulaLocal = "=COUNTIF(AB" + (RowInicial + Tam4 + Tam8 + 6) + ":AB" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ", " + @"""SIN COMB""" + ")";
            _cells.GetCell("AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 9)).NumberFormat = "0";
            _cells.GetCell("AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 9)).FormulaLocal = "=COUNTIF(AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + 9) + ":AJ" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 8) + ", " + @"""SIN COMB""" + ")";
            //Cargador
            _cells.GetCell("D" + (RowInicial + Tam1 + Tam10 + 3)).NumberFormat = "0";
            _cells.GetCell("D" + (RowInicial + Tam1 + Tam10 + 3)).FormulaLocal = "=COUNTIF(D" + (RowInicial + Tam1 + 3) + ":D" + (RowInicial + Tam1 + Tam10 + 2) + "," + @"""SIN COMB""" + ")";















            _cells.GetCell("AC2").NumberFormat = "0";
            _cells.GetCell("AC2").FormulaLocal = "=COUNTA(E" + (RowInicial) + ":E" + (RowInicial + (Tam1 - 1)) + ") + " + "COUNTA(M" + (RowInicial) + ":M" + (RowInicial + (Tam2 - 1)) + ") +" + "COUNTA(U" + (RowInicial) + ":U" + (RowInicial + (Tam3 - 1)) + ")+ " + "COUNTA(AC" + (RowInicial) + ":AC" + (RowInicial + (Tam4)) + ")+" + "COUNTA(M" + (RowInicial + Tam2 + 3) + ":M" + (RowInicial + (Tam2 + Tam6 + 2)) + ") + " + "COUNTA(U" + (RowInicial + Tam3 + 3) + ":U" + (RowInicial + (Tam3 + Tam7 + 2)) + ") +" + "COUNTA(AC" + (RowInicial + Tam4 + 3) + ":AC" + (RowInicial + (Tam4 + Tam8 + 2)) + ") +" + "COUNTA(AC" + (RowInicial + Tam4 + Tam8 + 6) + ":AC" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ") +" + "COUNTA(AK" + (RowInicial + Tam5 + Tam9 + Tam12 + 9) + ":AK" + (RowInicial + Tam5 + Tam9 + Tam12 + Tam13 + 8) + ") - AA2";

            _cells.GetCell("AC3").NumberFormat = "0";
            _cells.GetCell("AC3").FormulaLocal = "=COUNTA(E" + (RowInicial + Tam1 + 3) + ":E" + (RowInicial + (Tam1 + Tam10 + 2)) + ") - AA3";

            _cells.GetCell("AC4").NumberFormat = "0";
            _cells.GetCell("AC4").FormulaLocal = "=COUNTA(AC" + (RowInicial + Tam4 + Tam8 + 6) + ":AC" + (RowInicial + (Tam4 + Tam8 + Tam11 + 5)) + ") - AA4";




            //% Cumplimiento
            _cells.GetCell("AF1").NumberFormat = "0.00%";
            _cells.GetCell("AF1").FormulaLocal = "=AC1/Y1";
            _cells.GetCell("AF2").NumberFormat = "0.00%";
            _cells.GetCell("AF2").FormulaLocal = "=AC2/Y2";
            _cells.GetCell("AF3").NumberFormat = "0.00%";
            _cells.GetCell("AF3").FormulaLocal = "=AC3/Y3";
            _cells.GetCell("AF4").NumberFormat = "0.00%";
            _cells.GetCell("AF4").FormulaLocal = "=AC4/Y4";




        }

        private void Reload_Info_Select(Excel.Range Rango, List<string> Lista, Int32 TamLetr = 8, Int32 Rf = 166, Int32 Gf = 166, Int32 Bf = 166, Int32 Rl = 0, Int32 Gl = 0, Int32 Bl = 0)
        {
            Rango.Clear();
            FormatCamposMenu(Rango, true, true, true, "", "", TamLetr, Rf: Rf, Gf: Gf, Bf: Bf, Rl: Rl, Gl: Gl, Bl: Bl, Aline: "L");
            Rango.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), Lista), Type.Missing);
            FormatBordes(Rango);
        }

        //TRANFORMAR RANGO EN UNA MATRIZ----- FUNCIONA OK
        static string[,] GetStringArray(Object rangeValues)
        {
            string[,] stringArray = null;
            Array array = rangeValues as Array;
            if (null != array)
            {
                int rank = array.Rank;
                if (rank > 1)
                {
                    int rowCount = array.GetLength(0);
                    int columnCount = array.GetUpperBound(1);

                    stringArray = new string[rowCount, columnCount];
                    //stringArray[0, 0] = "H"; 
                    for (int Col = 0; Col < columnCount; Col++)
                    {
                        for (int Row = 0; Row < rowCount; Row++)
                        {
                            //stringArray[Col] = new string[columnCount - 1];
                            //stringArray = new string[Row, Col];
                            Object obj = array.GetValue(Row + 1, Col + 1);
                            if (null != obj)
                            {
                                //string value = obj.ToString();
                                stringArray[Row, Col] = obj.ToString();
                            }
                        }
                    }
                }
            }

            return stringArray;
        }


    }
}

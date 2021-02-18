using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;


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
using Debugger = SharedClassLibrary.Utilities.Debugger;
using SharedClassLibrary.Classes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;



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
        public Int32 StartColTable = 2;
        public Int32 StartRowTable = 1;
        //UTILIDADES PARA LOS MOVIMIENTOS DINAMICOS
        public Int32 Mayor = 0;
        public Int32 CntIndicador = 0;
        //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //public SqlConnection cnnx;
        public String Sql = "";

        //private const string ValidationSheetName01 = "ValidationSheetEventos";
        //private const string ValidationSheetName02 = "ValidationSheetCargas";
        private const string SheetName01 = "Eventos";
        private const string SheetName02 = "Cargas";
        private const string SheetName03 = "Eventos_Pivot";

        private const string tableName01 = "_01Eventos";
        private const string tableName02 = "_01Cargas";
        private const string tableName03 = "Pivot_Eventos";
        //private const int titleRow = 8;

        //Variables de Conexion 
        private string SQL;
        private string DataBase;
        private string User; //Ej. SIGCON, CONSULBO
        private string Pw;
        // ReSharper disable once InconsistentNaming
        public string DbLink; //Ej. @DBLMIMS

        //OracleConnection Conexion;
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = new string[] { "lmnsql01", "lmnsql03" };
            //var enviroments = {Productivo};
            //var enviroments = Environments.GetEnviromentList();
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

        private void btnFormatear_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = true;
            try
            {
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                Formatear("GANTT DE PARADA - ELLIPSE 8", SheetName01, true);
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



            FormatCamposMenu(_cells.GetRange("A1", "T2"), true, true, true, "PLAN DE TANQUEO DE COMBUSTIBLE", "", 22, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("A1", "T2"));
            FormatCamposMenu(_cells.GetRange("U1", "W4"), true, true, true, "CUMPLIMIENTO PLANES", "", 11, Rf: 255, Gf: 217, Bf: 102, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("U1", "W4"));

            //Primera fila Palas Indicadores
            FormatCamposMenu(_cells.GetRange("X1", "Y1"), true, true, true, "PALAS:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("X1", "Y1"));
            FormatCamposMenu(_cells.GetCell("Z1"), true, false, false, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("Z1"));
            FormatCamposMenu(_cells.GetCell("AA1"), true, false, true, "SIN COMB:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AA1"));
            FormatCamposMenu(_cells.GetCell("AB1"), true, false, false, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AB1"));
            FormatCamposMenu(_cells.GetCell("AC1"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AC1"));
            FormatCamposMenu(_cells.GetRange("AD1","AE1"), true, true, true, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("AD1", "AE1"));
            FormatCamposMenu(_cells.GetCell("AF1"), true, false, true, "% de cump.", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AF1"));
            FormatCamposMenu(_cells.GetCell("AG1"), true, false, true, "", "", 11, Rf: 255, Gf: 192, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AG1"));

            //Segunda fila Auxiliares Indicadores
            FormatCamposMenu(_cells.GetRange("X2", "Y2"), true, true, true, "AUXILIARES:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("X2", "Y2"));
            FormatCamposMenu(_cells.GetCell("Z2"), true, false, false, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("Z2"));
            FormatCamposMenu(_cells.GetCell("AA2"), true, false, true, "SIN COMB:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AA2"));
            FormatCamposMenu(_cells.GetCell("AB2"), true, false, false, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AB2"));
            FormatCamposMenu(_cells.GetCell("AC2"), true, false, true, "CUMPLIDOS:", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AC2"));
            FormatCamposMenu(_cells.GetRange("AD2", "AE2"), true, true, true, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetRange("AD2", "AE2"));
            FormatCamposMenu(_cells.GetCell("AF2"), true, false, true, "% de cump.", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AF2"));
            FormatCamposMenu(_cells.GetCell("AG2"), true, false, true, "", "", 11, Rf: 169, Gf: 208, Bf: 142, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell("AG2"));










            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();

            #endregion
            //_excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }

        private void FormatCamposMenu(Excel.Range Celda, bool Col, bool Merge, bool Negrita, String Texto = "", String Comentario = "", /*bool Bords, */Int32 TamLetra = 9, Int32 Rf = 91, Int32 Gf = 155, Int32 Bf = 213, Int32 Rl = 255, Int32 Gl = 255, Int32 Bl = 255)
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
                Celda.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Celda.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
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




    }
}

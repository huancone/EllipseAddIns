using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Services.Ellipse;
using System.Web.Services;
using FormularioAutenticacion;
using FormAutetication;
using System.Data.SqlClient;
using System.Drawing;
using System.Data;
using data = System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using Application = Microsoft.Office.Interop.Excel.Application;
using FormAuthenticate = EllipseCommonsClassLibrary.FormAuthenticate;






namespace EllipseAddInInfoPm
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;

        private const string SheetName01 = "INFORMACION_PM";
        private const string TableName01 = "_01INFORMACION_PM";
        private const int titleRow = 8;
        private Thread _thread;
        private bool _progressUpdate = true;
        public String Sql = "";
        static object useDefault = Type.Missing;
        private const Int32 StartColHrs = 19;
        private const Int32 DatosAgregados = 3;
        public SqlConnection cnnx;
        FormularioConfirmar Confir = new FormularioConfirmar();
        FormularioAutenticacionType _AuthG = new FormularioAutenticacionType();

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
        public Int32 StartRowTable = 3;
        //UTILIDADES PARA LOS MOVIMIENTOS DINAMICOS
        public Int32 Mayor = 0;
        public Int32 CntIndicador = 0;
        //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //FINALES FILAS Y COLUMNAS DE TABLA DE GANTT PARADA
        public Int32 FinColTablaOneSheet = 0;
        public Int32 FinRowTablaOneSheet = 0;
        public string[] NombreColumnas;
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            _excelApp.EnableEvents = true;
            //var tableObject = Globals.Factory.GetVstoObject(_excelApp.ActiveWorkbook.Sheets[SheetName01].Name);
            //tableObject.Change = GetTableChangedValue;
            var enviroments = Environments.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }
        //Ejecutar PL Sql
        public data.DataTable getdata(string SQL, string DataBase = "SIGCOPRD", string User = "consulbo", string Pw = "consulbo", string DbLink = "@DBLELLIPSE8")
        {

            _eFunctions.SetDBSettings(DataBase, User, Pw, DbLink);
            var dat = _eFunctions.GetQueryResult(SQL);
            data.DataTable DATA = new data.DataTable();
            DATA.Load(dat);
            return DATA;
        }
        //Ejecutar Sql Server
        public data.DataTable getdataSql(string SQL)
        {
            cnnx = new SqlConnection(Conexion("lmnsql01"));
            data.DataTable DATA = new data.DataTable();
            SqlDataAdapter dat = new SqlDataAdapter(SQL, cnnx);
            dat.Fill(DATA);
            return DATA;
        }
        private string Conexion(String Servidor)
        {
            String connectionString = "server=" + Servidor + ";database= PowerView; uid=xblmnsql01; pwd =Pview012019;Connection Timeout=0";
            return connectionString;
        }
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBoxExcelAddIn About = new AboutBoxExcelAddIn("Gustavo Vargas", "");
            About.ShowDialog();
        }
        public void Formatear(string Titulo, string NombreHoja, bool SubEncab = false)
        {
            //String Titulo = "";
            CntIndicador = CntIndicador + 1;
            //_eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            #region CONSTRUYO LA HOJA 1 y 2
            //while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
            /*if(CntIndicador == 1 )
            {
                _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.ActiveWorkbook.ActiveSheet.Name);
            }
            else
            {
                _excelApp.ActiveWorkbook.Worksheets.Add();
            }*/
            _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.ActiveWorkbook.ActiveSheet);
            _excelApp.ActiveWorkbook.ActiveSheet.Name = NombreHoja;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            //Excel.Worksheet _cells = (Excel.Worksheet)_excelApp.Worksheets.Add();

            if (CntIndicador == 1)
            {
                if (EndRowTitulo >= EndRowImg)
                {
                    Mayor = EndRowTitulo;
                }
                else
                {
                    Mayor = EndRowImg;
                }
                StartRowInputMenu = StartRowInputMenu + (Mayor + 1);
                StartRowTable = StartRowTable + (StartRowInputMenu + 2);
            }
            /*
                _cells.GetRange("M1", "S1").NumberFormat = "@";
                _cells.GetCell("M1").Value = (StartColInputMenu + 1) + "---" + (StartRowInputMenu + 1);
                _cells.GetCell("N1").Value = (StartColInputMenu + 1) + "---" + (StartRowInputMenu);
                _cells.GetCell("O1").Value = (StartColInputMenu + 4) + "---" + (StartRowInputMenu);//FLOTA StartColInputMenu + 4, StartRowInputMenu
                _cells.GetCell("P1").Value = (StartColInputMenu + 4) + "---" + (StartRowInputMenu + 1);//EQUIPO
                _cells.GetCell("Q1").Value = (StartColInputMenu + 6) + "---" + (StartRowInputMenu + 1);
                //rango de tabla
                _cells.GetCell("R1").Value = StartRowTable;
                _cells.GetCell("S1").Value = StartColTable;
                //Color blanco para la letra de la prueba de escritorio
                _cells.GetRange("M1", "S1").Font.Color = System.Drawing.Color.White;
            */

            TituloAndLogo(@"..\Resources\Cerrejon.png", _cells.GetRange(StartColImg, StartColImg, EndColImg, EndRowImg), Titulo, _cells.GetRange(StartColTitulo, StartRowTitulo, EndColTitulo, EndRowTitulo));
            if (SubEncab)
            {
                SubEncabezado();
            }


            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();

            #endregion
            //_excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
        }
        private void TituloAndLogo(String Ruta, Excel.Range RngImg, String Titulo, Excel.Range RngTitulo)
        {
            //FORMAT IMAGEN
            RngImg.Select();
            RngImg.Merge();
            RngImg.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            RngImg.Borders.Weight = "2";
            float Left = (float)((double)RngImg.Left);
            float Top = (float)((double)RngImg.Top);
            const float ImageSize = 23;
            //_excelApp.ActiveSheet.Shapes.AddPicture(Ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue,/*IZQUIERDA, ARRIBA, ANCHO, ALTO*/ Left + 1, Top + 1, ImageSize + 80, ImageSize);
            //RngImg.Style = _cells.GetStyle(StyleConstants.HeaderDefault);
            RngTitulo.Select();
            RngTitulo.Merge();
            RngTitulo.Value = Titulo;
            RngTitulo.Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(91, 155, 213)));
            RngTitulo.Font.Color = System.Drawing.ColorTranslator.ToOle((Color.White));
            RngTitulo.Font.Size = 20;
            RngTitulo.Font.Bold = true;
            RngTitulo.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            RngTitulo.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RngTitulo.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            RngTitulo.Borders.Weight = "2";
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
        private void SubEncabezado()
        {
            //_cells.GetCell("A1").Value = "CERREJÓN";
            //Excel.Range IMG = (Excel.Range)RngImg;
            //FORMAT TITULO
            //FECHAS DE LA HOJA 
            FormatCamposMenu(_cells.GetCell(StartColInputMenu, StartRowInputMenu), true, "EQUIPO INI");
            FormatBordes(_cells.GetCell(StartColInputMenu, StartRowInputMenu));
            FormatCamposMenu(_cells.GetCell(StartColInputMenu, StartRowInputMenu + 1), true, "EQUIPO FIN");
            FormatBordes(_cells.GetCell(StartColInputMenu, StartRowInputMenu + 1));
            // AGRGADO DE LISTAS DESPLEGABLES DE LAS FECHAS
            //var List_2 = string.Join(Separador(), ListaDatos(2));
            Excel.Range Fecha1 = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu);
            ///Fecha1.Validation.Delete();
            Fecha1.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertWarning, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), ListaDatos(2, "ASC")), Type.Missing);
            Fecha1.Validation.IgnoreBlank = true;
            Fecha1.Validation.ShowError = false;
            Fecha1.Copy();
            //Fecha1.Value = ListaDatos(2)[0];
            //Fecha1.Value = "";
            //DateTime dateToDisplay = DateTime.Now;
            _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).PasteSpecial();
            //_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value = "'" + DateTime.Now.ToString("yyyyMMdd");
            //_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value = "'20200226";
            //FORMATOS A CAMPOS FECHAS
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu), false, "", "0220251");
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1), false, "", "0220654");
            FormatBordes(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu));
            FormatBordes(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1));


            // CAMPOS DE FILTROS DE EQUIPOS FLOTAS Y TYPE CONSULTA


            //EQUIPOS Y FLOTAS DE LA HOJA 
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu), true, "DIA PROM");
            FormatBordes(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu));
            //FormatCamposMenu(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu + 1), true, "EQUIPO ELLIPSE");
            //FormatBordes(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu + 1));
            //IMPUT CAMPOS EQUIPOS y FLOTAS
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu), false, "", "EJ: 20 -- VALOR PROMEDIO PARA PROYECCION DE HORAS DE OPERACION DIARIO");
            //FormatCamposMenu(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1), false, "", "EQUIPO FORMATO ELLIPSE - [0220906] O [0050025]");
            FormatBordes(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu));
            //FormatBordes(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1));
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value = 20;
            // AGREGADO DE LISTAS DESPLEGABLES PARA FLOTAS
            //var List = string.Join(Separador(), ListaDatos(1));
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.Delete();
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, List, Type.Missing);
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.IgnoreBlank = true;
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.ShowError = true;


            /*
            //TYPE DE CONSULTA
            List<string> listRange2 = new List<string>();
            listRange2.Add("WORK_ORDER");
            listRange2.Add("TASK");
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu), true, "TYPE SQL");
            FormatBordes(_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu));
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu + 1), false, "WORK_ORDER", "WORK_ORDER - VER POR ENCABEZADO DE ORDENES O TASK - VER POR TAREAS");
            FormatBordes(_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu + 1));
            //AGREGADO DE LISTA PARA TYPE CONSULTA O SQL
            _cells.SetValidationList(_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu + 1), listRange2);


            //STATUS DE ORDENES
            List<string> listRange3 = new List<string>();
            listRange3.Add("Uncompleted");
            listRange3.Add("Closed");
            listRange3.Add("All");
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu), true, "WO STATUS");
            FormatBordes(_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu));
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu + 1), false, "Uncompleted", "BUSCAR POR ESTADO DE LAS WORK_ORDER");
            FormatBordes(_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu + 1));
            //AGREGADO DE LISTA PARA TYPE CONSULTA O SQL
            _cells.SetValidationList(_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu + 1), listRange3);
            */

        }
        List<string> ListaDatos(Int32 Tipo, String ORDEN = "DESC")
        {
            List<string> listRange = new List<string>();
            data.DataTable table = null;
            if (Tipo == 1)
            {
                Sql = (@"SELECT DISTINCT
                            CTD.FLOTA_ELLIPSE
                        FROM
                          SIGMAN.EQMTLIST CTD
                        WHERE
                          FLOTA_ELLIPSE IS NOT NULL
                        ORDER BY
                          1 " + ORDEN);
                table = getdata(Sql);
            }
            else if (Tipo == 2)
            {
                Sql = (@"SELECT
                      EQMTLIST.EQU
                    FROM
                      SIGMAN.EQMTLIST
                    ORDER BY
                      1  " + ORDEN);
                table = getdata(Sql);
            }
            else if (Tipo == 3)
            {
                Sql = (@"SELECT 
                            TT.TABLE_CODE AS RES_CODE--,
                            --TT.TABLE_DESC RES_DESC,
                        FROM 
                          ELLIPSE.MSF010 TT 
                        WHERE
                          TT.TABLE_TYPE = 'TT'
                        ORDER BY
                          1 " + ORDEN);
                table = getdata(Sql, "EL8PROD", "SIGCON", "ventyx", "");
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
        private void FormatCamposMenu(Excel.Range Celda, bool Col, String Texto = "", String Comentario = "", /*bool Bords, */Int32 TamLetra = 9, Int32 Rf = 91, Int32 Gf = 155, Int32 Bf = 213, Int32 Rl = 255, Int32 Gl = 255, Int32 Bl = 255)
        {

            Celda.NumberFormat = "@";
            Celda.Font.Bold = true;
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
            if (Texto != "")
            {
                Celda.Value = Texto;
            }
            if (Comentario != "")
            {
                Celda.AddComment(Comentario);
            }

        }
        private void btnFormatear_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = false;
            //búsquedas especiales de tabla
            //_cells.SetCursorWait();
            /*_AuthG.StartPosition = FormStartPosition.CenterScreen;
            if (_AuthG.ShowDialog() == DialogResult.OK)
            {*/
                try
                {
                    _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                    Formatear("INFORMACION DE PMS - ELLIPSE 9", SheetName01, true);
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
            /*}
            else
            {
                MessageBox.Show(@"Regrese cuando tenga autorizacion del administrador del sistema.");
                _excelApp.Application.Quit();
            }*/
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
        private void CentrarRango(Excel.Range Rango)
        {
            Rango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Rango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }
        private void btnConsultar_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                try
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ExecuteQuery);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
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
        public void ExecuteQuery()
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                //_cells.SetCursorWait();
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                string NameHoja = _excelApp.ActiveWorkbook.ActiveSheet.Name;
                borrarTabla(NameHoja);
                data.DataTable table;

                String Param1 = "";
                String Param2 = "";
                Int64 Param3 = 20;
                var sqlQuery = "";

                if (_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu).Value == null || _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value == null)
                {
                    MessageBox.Show("DEBE INGRESAR RANGO DE EQUIPOS");
                    return;
                }
                else
                {
                    Param1 = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu).Value;
                    Param2 = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value;
                    if(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value != null)
                    {
                        Param3 = Convert.ToInt64(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value);
                    }        
                    sqlQuery = Consulta(1, 1, Param1, Param2, Param3);
                    table = getdata(sqlQuery);
                }
    

                    
                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("NO SE ENCONTRO INFORMACION");
                    return;
                }
                data.DataTable table2;
                //DateTime date = DateTime.Now;
                double Dias = 0;
                //string[] format = { "yyyyMMdd" };
                int i = 0;
                string[,] data = new string[table.Rows.Count, table.Columns.Count];
                foreach (data.DataRow row in table.Rows)
                {
                    int j = 0;
                    //string P1 = "";
                    //string P2 = "";
                    //double P3 = 0;
                    //double P4 = 0;
                    //Columnas de la consulta
                    foreach (data.DataColumn col in table.Columns)
                    {
                        data[i, j] = row[col].ToString();
                        /*if (j == 14)
                        {
                            P3 = Convert.ToDouble(row[col].ToString());
                        }
                        if (j == 23)
                        {
                            P2 = row[col].ToString();
                        }
                        if (j == 24)
                        {
                            P1 = row[col].ToString();
                        }*/
                        j++;
                    }
                    if(row[23].ToString() != "")
                    {
                        sqlQuery = Consulta(1, 2, row[23].ToString(), row[22].ToString());
                        table2 = getdataSql(sqlQuery);
                        //HRS OPERACIONS DESPUES DEL PM
                        if (table2.Rows.Count != 0)
                        {
                            if(table2.Rows[0][0].ToString() == "")
                            {
                                data[i, 16] = "0";
                            }
                            else
                            {
                                data[i, 16] = table2.Rows[0][0].ToString();
                            }
                            data[i, 15] = (Convert.ToDouble(row[13].ToString()) - Convert.ToDouble(data[i, 16].ToString())).ToString() ;
                        }
                        table2 = null;
                        //ULTIMO EVENTO CTD Y ULTIMA FECHA TOMA DE CTD
                        sqlQuery = Consulta(1, 3, "", row[22].ToString());
                        table2 = getdataSql(sqlQuery);
                        if (table2.Rows.Count != 0)
                        {
                            data[i, 17] = table2.Rows[0][0].ToString();
                            data[i, 24] = table2.Rows[0][1].ToString();
                            if(table2.Rows[0][2].ToString() != "999")
                            {
                                //Dias = 0;
                                Dias = ((Convert.ToDouble(row[13].ToString()) - Convert.ToDouble(data[i, 16].ToString())) / Param3);
                                DateTime date = DateTime.Now;
                                date = date.AddDays(Dias);
                                data[i, 18] = date.ToString(/*"yyyy/mm/dd hh:mm:ss"*/);
                            }
                        }
                        table2 = null;
                        //ULTIMO TAJO DE TABLA CARGAS
                        sqlQuery = Consulta(1, 4, "", row[22].ToString());
                        table2 = getdataSql(sqlQuery);
                        if (table2.Rows.Count != 0)
                        {
                            data[i, 20] = table2.Rows[0][0].ToString();
                        }
                        else
                        {
                            //ULTIMO TAJO DE TABLA EVENTOS EQ AUX
                            table2 = null;
                            sqlQuery = Consulta(1, 5, "", row[22].ToString());
                            table2 = getdataSql(sqlQuery);
                            if (table2.Rows.Count != 0)
                            {
                                data[i, 20] = table2.Rows[0][0].ToString();
                            }
                        }
                        table2 = null;
                        //

                    }
                    
                    
                    i++;
                    //format row
                    //*if (i % 2 == 0)
                    //{
                      //  _cells.GetRange(StartColTable, (StartRowTable + i), table.Columns.Count + DatosAgregados, (StartRowTable + i)).Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(221, 235, 247)));
                    //}
                }


                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).NumberFormat = "@";
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = data;
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value;
                //CentrarRango(_cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable));
                Encabezado(table, _excelApp.ActiveWorkbook.ActiveSheet.Name);
                FormatTable(_cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable), NameHoja, 1, 1);
                table = null;

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //Convertir a Numero
                    ChangeFormatColumn(NameHoja, "#,##0.00", "P:Q");
                    ChangeFormatColumn(NameHoja, "mm/dd/yyyy hh:mm:ss", "S:S");
                }

                //hacemos estatica la primer fila y aplicamos filtros asi,
                _excelApp.Application.ActiveWindow.SplitRow = StartRowTable;
                _excelApp.Application.ActiveWindow.FreezePanes = true;

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
        private void FormatTable(Excel.Range Rango, string HojaName, Int32 StyleText = 0, Int32 TypeTable = 0)
        {
            //_excelApp.ActiveWorkbook.ActiveSheet.Select();
            //Rango.Select();
            String NameTable = "01";
            NameTable = NameTable + Convert.ToString(_excelApp.ActiveWorkbook.ActiveSheet.Name);
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
                //if (HojaName == SheetName01)
                //{
                    //TableFiltro.ShowHeaders = false;
               // }
                //Rango.AutoFilter(StartRowTable, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            }
            else if (TypeTable == 2)
            {

            }
            FormatBordes(Rango);

        }
        private void Encabezado(data.DataTable table, String Hoja)
        {
            //Formateando columnas de encabezado
            //_excelApp.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, _cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, StartRowTable), Type.Missing, Excel.XlYesNoGuess.xlNo, Type.Missing).Name = "TiTul01";
            int cont = StartColTable;
            for (var i = StartColTable; i <= table.Columns.Count; i++)
            {

                _cells.GetCell(cont, StartRowTable).Value = table.Columns[i- StartColTable].ColumnName.Trim();
                cont++;
            }

        }
        public void borrarTabla(String Name_Hoja)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //Excel._Worksheet _cells2 = _excelApp.ActiveWorkbook.ActiveSheet;
                //Excel._Worksheet Hoja = _excelApp.ActiveWorkbook.Sheets[Name_Hoja];
                //Hoja.ListObjects(tableName01);
                //_cells.DeleteTableRange(_excelApp.ActiveWorkbook.Sheets[Name_Hoja].Table.Name);
                _cells.DeleteTableRange(TableName01);
                return;
            }

        }
        public string Consulta(Int32 Hoja, Int32 Tipe, string Param1 = "", string Param2 = "", Int64 Param3 = 20)
        {
            string sqlQuery = "";
            if (Hoja == 1)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"WITH DATOS_ELLIPSE AS
                                (
                                    SELECT
                                    PM.WORK_GROUP,
                                    TRIM(PM.EQUIP_NO) AS EQUIP_NO,
                                    PM.WORK_ORDER AS WORK_ORDER_PM,
                                    EV.WORK_ORDER AS WORK_ORDER_EV,
                                    CASE
                                        WHEN EV.CLOSED_DT = '        ' OR EV.CLOSED_DT IS NULL THEN PM.CLOSED_DT
                                        ELSE EV.CLOSED_DT 
                                    END AS CLOSEDT_DT_SELECT,
                                    CASE
                                        WHEN EV.CLOSED_TIME = '        ' OR EV.CLOSED_TIME IS NULL THEN PM.CLOSED_TIME
                                        ELSE EV.CLOSED_TIME 
                                    END AS CLOSEDT_TIME_SELECT,
                                    PM.MAINT_SCH_TASK,
                                    PM.CLOSED_DT AS CLOSED_DT_PM,
                                    PM.CLOSED_TIME AS CLOSED_TIME_PM,
                                    EV.CLOSED_DT AS CLOSED_DT_EV,
                                    EV.CLOSED_TIME AS CLOSED_TIME_EV,  
                                    ROW_NUMBER() OVER(PARTITION BY PM.EQUIP_NO ORDER BY (PM.CLOSED_DT||PM.CLOSED_TIME) DESC NULLS LAST  )  AS N_ITEM,
                                    MST.SCHED_FREQ_1,
                                    MST.LAST_PERF_DATE,
                                    MST.SCHED_DESC_1,
                                    MST.SCHED_DESC_2
                                    FROM  
                                    ELLIPSE.MSF620@DBLELLIPSE8 PM
                                    LEFT JOIN ELLIPSE.MSF620@DBLELLIPSE8 EV ON (PM.DSTRCT_CODE=EV.DSTRCT_CODE AND PM.RELATED_WO=EV.WORK_ORDER)
                                    INNER JOIN ELLIPSE.MSF600@DBLELLIPSE8 EQ ON(PM.EQUIP_NO=EQ.EQUIP_NO)
                                    LEFT JOIN ELLIPSE.MSF700@DBLELLIPSE8 MST ON 
                                                                (
                                                                  ELLIPSE.PM.DSTRCT_CODE = ELLIPSE.MST.DSTRCT_CODE
                                                                  AND ELLIPSE.PM.EQUIP_NO = ELLIPSE.MST.EQUIP_NO 
                                                                  AND ELLIPSE.PM.MAINT_SCH_TASK = ELLIPSE.MST.MAINT_SCH_TASK 
                                                                  AND ELLIPSE.PM.COMP_CODE = ELLIPSE.MST.COMP_CODE
                                                                  AND ELLIPSE.PM.COMP_MOD_CODE = ELLIPSE.MST.COMP_MOD_CODE 
                                                                )
                                    WHERE
                                    EQ.DSTRCT_CODE = 'ICOR'
                                    --AND PM.EQUIP_NO = '0220438     '--'0220290     '
                                    --AND EQ.EQUIP_NO BETWEEN '0220251' AND '0220654'
                                    AND EQ.EQUIP_NO BETWEEN '" + Param1 + "' AND '" + Param2 + @"'
                                    AND EQ.ACTIVE_FLG = 'Y'
                                    AND PM.WO_STATUS_M = 'C'
                                    AND MST.JOB_DESC_CODE = 'Z9'
                                ),
                                PRIMERA AS
                                (
                                    SELECT
                                    --DATOS_ELLIPSE.*,
                                    --('20200812') AS MIERCOLES,
                                    --('7') AS D_PromAutomaticas,
                                    --0 AS HrsPromedioManuales,
                                    --(@Prompt('Fecha Inicio a Programar', 'A',, mono, free, persistent)) AS MIERCOLES,
                                    --(@Prompt('Dias Para Promedio', 'A',, mono, free, persistent)) AS D_PromAutomaticas,
                                    --To_Number((@Prompt('Hrs Promedio Manual', 'A',, mono, free, persistent))) AS HrsPromedioManuales,
                                    DATOS_ELLIPSE.WORK_GROUP,
                                    DATOS_ELLIPSE.EQUIP_NO,
                                    DATOS_ELLIPSE.WORK_ORDER_PM,
                                    DATOS_ELLIPSE.MAINT_SCH_TASK,
                                    DATOS_ELLIPSE.CLOSED_DT_PM,
                                    DATOS_ELLIPSE.CLOSED_TIME_PM,
                                    DATOS_ELLIPSE.WORK_ORDER_EV,
                                    DATOS_ELLIPSE.CLOSED_DT_EV,
                                    DATOS_ELLIPSE.CLOSED_TIME_EV,
                                    DATOS_ELLIPSE.CLOSEDT_DT_SELECT,
                                    DATOS_ELLIPSE.CLOSEDT_TIME_SELECT,
                                    SIGMAN.FNU_B_MST_DESC(DATOS_ELLIPSE.EQUIP_NO, DATOS_ELLIPSE.MAINT_SCH_TASK) AS ANTERIOR_MST,
                                    DATOS_ELLIPSE.MAINT_SCH_TASK AS ULT_MST,
                                    SIGMAN.FNU_FIRST_LAST_MST1(DATOS_ELLIPSE.EQUIP_NO, DATOS_ELLIPSE.MAINT_SCH_TASK) AS PROX_MST,
                                    (
                                        SELECT
                                            INTER.SCHED_DESC_1
                                        FROM
                                            ELLIPSE.MSF700@DBLELLIPSE8 INTER
                                        WHERE
                                            INTER.REC_700_TYPE = 'ES'
                                            AND INTER.EQUIP_NO = RPAD(DATOS_ELLIPSE.EQUIP_NO,12,' ')
                                            AND INTER.MAINT_SCH_TASK = SIGMAN.FNU_FIRST_LAST_MST1(DATOS_ELLIPSE.EQUIP_NO, DATOS_ELLIPSE.MAINT_SCH_TASK)
                                            AND INTER.SCHED_IND_700 <> '9'
                                            AND ROWNUM = 1
                                    ) AS DESC_PROX_MST,
                                    DATOS_ELLIPSE.SCHED_FREQ_1 AS TIPO_PM,
                                    DATOS_ELLIPSE.LAST_PERF_DATE AS FECHA_ULT_PM,
                                    DATOS_ELLIPSE.SCHED_DESC_1,
                                    DATOS_ELLIPSE.SCHED_DESC_2,
                                    CASE
                                        WHEN DATOS_ELLIPSE.CLOSEDT_DT_SELECT <> '        ' THEN TO_CHAR(TO_DATE(DATOS_ELLIPSE.CLOSEDT_DT_SELECT || DATOS_ELLIPSE.CLOSEDT_TIME_SELECT, 'YYYY-MM-DDHH24MISS'), 'YYYY-MM-DD HH24:MI:SS') 
                                        ELSE NULL
                                    END AS DATE_HRS
                                    FROM
                                    DATOS_ELLIPSE
                                    WHERE
                                    DATOS_ELLIPSE.N_ITEM = 1
                                )
                                SELECT
                                    CTD_ELLIPSE.UNIT,
                                    PRIMERA.WORK_GROUP,
                                    PRIMERA.EQUIP_NO,
                                    PRIMERA.WORK_ORDER_PM,
                                    PRIMERA.WORK_ORDER_EV,
                                    PRIMERA.CLOSEDT_DT_SELECT,
                                    PRIMERA.ANTERIOR_MST,
                                    PRIMERA.ULT_MST,
                                    PRIMERA.SCHED_DESC_1 AS DESC_ULT_MST,
                                    PRIMERA.PROX_MST,
                                    PRIMERA.DESC_PROX_MST,
                                    PRIMERA.FECHA_ULT_PM,
                                    PRIMERA.SCHED_DESC_2 AS ATENCION,
                                    PRIMERA.TIPO_PM,
                                    " + Param3 + @" AS PRO_HR,
                                    '' AS HRS_RESTANTES,
                                    '' AS HRS_DESP_PM,
                                    '' AS ULTIMO_ESTADO_CTD,
                                    '' AS PROX_FECHA_PM,
                                    TRIM(CTD_ELLIPSE.FLOTA_ELLIPSE) AS FLOTA_ELLIPSE,
                                    '' AS LAST_PIT,
                                    TRIM(CTD_ELLIPSE.EQMTTYPE) AS EQMTTYPE,
                                    CTD_ELLIPSE.eqmt EQ_CTD,
                                    PRIMERA.DATE_HRS,
                                    '' AS F_TOMA_CTD
                                FROM
                                    PRIMERA
                                    INNER JOIN SIGMAN.EQMTLIST CTD_ELLIPSE ON (PRIMERA.EQUIP_NO = CTD_ELLIPSE.EQU AND CTD_ELLIPSE.ACTIVE_FLG = 'Y')
                                    --ORDER BY
                                    --PRIMERA.EQUIP_NO";
                }
                else if(Tipe == 2)
                {
                    sqlQuery = @"SELECT
                                    ROUND(SUM(hist_statusevents.duration)/3600,2) AS HRS_M_PM      
                                FROM
                                    dbo.hist_statusevents
                                    LEFT JOIN dbo.hist_exproot hist_turnos ON (hist_statusevents.shiftindex=hist_turnos.shiftindex )
                                WHERE
                                    DATEADD(SECOND, (hist_statusevents.starttime + hist_turnos.start), Convert(datetime, hist_turnos.shiftdate, 112)) >= '" + Param1 + @"'
                                    AND hist_statusevents.eqmt = '" + Param2 + @"'
                                    AND hist_statusevents.category IN('2', '5')";
                }
                else if (Tipe == 3)
                {
                    sqlQuery = @"SELECT TOP 1
                                  (icr_codigoscategoria_200502.status) + ' - ' + (icr_codigoscategoria_200502.descripcion) AS RESULTADO1,
                                  DATEADD(SECOND,(hist_statusevents.endtime+hist_turnos.start),Convert(datetime,hist_turnos.shiftdate,112)) AS RESULTADO2,
                                  hist_statusevents.reason 
                                FROM
                                      dbo.hist_statusevents hist_statusevents
                                      LEFT OUTER JOIN  dbo.icr_codigoscategoria_200502 icr_codigoscategoria_200502 ON hist_statusevents.reason = icr_codigoscategoria_200502.codigo   AND hist_statusevents.status = icr_codigoscategoria_200502.statusnum
                                    INNER JOIN dbo.hist_exproot hist_turnos ON(hist_statusevents.shiftindex=hist_turnos.shiftindex)
                                WHERE
                                      hist_statusevents.shiftindex = 
                                                                      (
                                                                              SELECT 
                                                                              MAX(hist_statusevents.shiftindex) 
                                                                              FROM 
                                                                              dbo.hist_statusevents hist_statusevents
                                                                              WHERE
                                                                              hist_statusevents.eqmt = '" + Param2 + @"'
                                                                      ) 
                                      AND hist_statusevents.eqmt = '" + Param2 + @"'
                                ORDER BY
                                  hist_statusevents.endtime DESC";
                }
                else if (Tipe == 4)
                {
                    sqlQuery = @"SELECT TOP(1)
                                  icr_tajos.Tajo AS Tajo_One
                                FROM
                                  dbo.hist_cargas
                                  LEFT JOIN dbo.icr_tajos ON(hist_cargas.Tajo=icr_tajos.Inicial)
                                WHERE
                                  hist_cargas.shiftindex = 
                                                          (
                                                            SELECT
                                                              MAX(hist_cargas.shiftindex) shiftindex
                                                            FROM
                                                              PowerView.dbo.hist_cargas
                                                            WHERE
                                                              (hist_cargas.Truck = '" + Param2 + @"' OR hist_cargas.excav = '" + Param2 + @"')
                                                               AND Tajo <> 'REM'
                                                          )
                                  AND (hist_cargas.Truck = '" + Param2 + @"' OR hist_cargas.excav = '" + Param2 + @"')
                                  AND hist_cargas.Tajo <> 'REM'
                                ORDER BY
                                  hist_cargas.timefull DESC";
                }
                else if (Tipe == 5)
                {
                    sqlQuery = @"SELECT
                                    TOP (1)
                                    TAJOS.TAJO AS Tajo_Two
                                FROM
                                    (
                                    SELECT
                                        CASE
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (50)
                                        THEN 'Patilla'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (51)
                                        THEN 'EWP'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (52,54)
                                        THEN 'La Puente'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (53,55,56)
                                        THEN 'Tabaco'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (57,64)
                                        THEN 'Oreganal 1'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (58, 61)
                                        THEN 'Tajo 100'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (65)
                                        THEN 'Comuneros'
                                        WHEN TAJO.status =4
                                        AND TAJO.reason IN (66)
                                        THEN 'Annex'
                                        END TAJO,
                                        shiftindex,
                                        endtime
                                    FROM
                                        dbo.hist_statusevents TAJO
                                    WHERE
                                        TAJO.shiftindex >= 
                                                            (
                                                            SELECT
                                                                MAX(hist_statusevents.shiftindex) shiftindex
                                                            FROM
                                                                PowerView.dbo.hist_statusevents
                                                            WHERE
                                                                (hist_statusevents.eqmt = '" + Param2 + @"')
                                                            ) - 765
                                    AND TAJO.eqmt = '" + Param2 + @"'
                                    AND TAJO.status  = 4
                                    AND TAJO.reason IN(50, 51, 52, 54, 53, 55, 56, 57, 64, 58, 61, 62, 63, 65, 66, 201, 202, 204, 205, 207, 208, 310, 312, 320)
                                    ) TAJOS
                                WHERE
                                    TAJO IS NOT NULL
                                ORDER BY
                                    shiftindex DESC,
                                    endtime DESC";
                }
            }
            return sqlQuery;
        }
        private void Stop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
        private void btnLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = true;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _cells.DeleteTableRange(TableName01);

            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            _excelApp.ScreenUpdating = true;
            //_excelApp.DisplayAlerts = true;
        }
        private bool ComprobarHoja(String NameHoja)
        {
            //_excelApp.Windows.Application.Sheets
            //_excelApp.ActiveWorkbook.Worksheets.Count;
            bool Exist = false;
            for (int v = 1; v <= _excelApp.Windows.Application.Sheets.Count; v++)
            {
                var wkSheet = _excelApp.Windows.Application.Sheets[v];
                if (wkSheet.Name == NameHoja)
                {
                    Exist = true;
                    break;
                }
            }
            return Exist;
        }
        private void CrearCampCalPivot(Excel.PivotTable pivotTable, string NameCol, string Formula)
        {
            Excel.CalculatedFields CampCalculado = pivotTable.CalculatedFields();
            Excel.PivotField DatoCalculado = CampCalculado.Add(NameCol, Formula, true);
            //DatoCalculado.NumberFormat = "#,##0.00";
            //PivotField NuevoCampo = pivotTable.AddDataField(DatoCalculado);
            //NuevoCampo.NumberFormat = "#,##0.00";
            //NuevoCampo.Name = NameCol;
            //pivotTable.AddDataField(DatoCalculado);
        }
        //TRANFORMAR RANGO EN UN VECTOR----- FUNCIONA OK
        static string[] GetStringArray(Object rangeValues)
        {
            string[] stringArray = null;
            Array array = rangeValues as Array;
            if (null != array)
            {
                int rank = array.Rank;
                if (rank > 1)
                {
                    int rowCount = array.GetLength(0);
                    int columnCount = array.GetUpperBound(1);

                    stringArray = new string[rowCount];
                    //stringArray[0] = "SS271";
                    //stringArray[index] = new string[columnCount - 1];
                    for (int Col = 0; Col < columnCount; Col++)
                    {
                        for (int Row = 0; Row < rowCount; Row++)
                        {
                            Object obj = array.GetValue(Row + 1, Col + 1);
                            if (null != obj)
                            {
                                string value = obj.ToString();
                                stringArray[Row] = value;
                            }
                            //stringArray[Row,Col] = new string[columnCount - 1];
                        }
                    }
                }
            }
            return stringArray;
        }
        private void TablaDinamica(string Nombre_Hoja)
        {
            //_excelApp.ScreenUpdating = true;
            // Declaracion de variables a utilizar
            Excel.Range Encabezado = null;
            Excel.PivotTable pivotTable = null;
            Excel.PivotCaches pivotCaches = null;
            Excel.PivotCache pivotCache = null;
            //Excel.PivotFields pivotFields = null;
            Excel.PivotField OcultSubt = null;
            Excel.PivotField OcultBlank = null;
            Excel.Range pivotDestination = null;
            //Excel.PivotField DatoCalculado = null;
            Excel.Range pivotData = null;


            //Convertir Columna en Numero
            ChangeFormatColumn(Nombre_Hoja, "#,##0.00", "P:Q");
            InsertarColumn(Nombre_Hoja, "T:T", "T9", "TURNO");
            ChangeFormatColumn(Nombre_Hoja, "mm/dd/yyyy hh:mm:ss", "T:T");
            InsertarColumn(Nombre_Hoja, "S:S", "S9", "F(YYYYMMDD)");

            //InsertarColumn(fromwrksh, fromwrksh.get_Range("V:V"), fromwrksh.Range[9, 20], "FECHA_PM");

            Excel.Worksheet fromwrksh = _excelApp.ActiveWorkbook.Sheets[Nombre_Hoja];
            //Excel.Range last = fromwrksh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            // Find the last real row
            //var LastRow = fromwrksh.Cells[10, 3].Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            //var Last = fromwrksh.Cells[10,21].SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var LastRow = fromwrksh.Range[fromwrksh.Cells[10, 2], fromwrksh.Cells[10, 2].End[Excel.XlDirection.xlDown]].Count -1;

            string[] DatosFechas = GetStringArray(fromwrksh.Range[fromwrksh.Cells[10, 20], fromwrksh.Cells[10 + LastRow, 21]].Value);
            //Excel.Range x = fromwrksh.Range[fromwrksh.Cells[10, 21], fromwrksh.Cells[10 + LastRow, 21]];
            //Excel.Range y = fromwrksh.Range[fromwrksh.Cells[10, 20], fromwrksh.Cells[10, 20].End[Excel.XlDirection.xlDown]];
            DateTime Fecha;
            for (int i = 0; i < DatosFechas.Length; i++)
            {
                //fromwrksh.Cells[i + 10, 20].NumberFormat = "yyyy/mm/dd";
                if (DatosFechas[i] != null)
                {
                    Fecha = Convert.ToDateTime(DatosFechas[i]);
                    if(Fecha.Hour >= 06 && Fecha.Hour < 18)
                    {
                        fromwrksh.Cells[i + 10, 21] = "A1";
                    }
                    else
                    {
                        if (Fecha.Hour >= 00 && Fecha.Hour < 6)
                        {
                            Fecha = Fecha.AddDays(-1);
                        }
                        fromwrksh.Cells[i + 10, 21] = "A2";
                    }
                    //fromwrksh.Cells[i + 10, 20].NumberFormat = "@";
                    fromwrksh.Cells[i + 10, 19] = Fecha.ToShortDateString();
                    fromwrksh.Cells[i + 10, 19].NumberFormat = "yyyy/mm/dd";
                }
            }
                //x.NumberFormat = "yyyy/mm/dd";
                //var vvv = x.Cells.Value2;
                //y.Value = x.Cells.Value;
            //y.Value = y.Value;
            //_cells.GetRange(fromwrksh.Range[fromwrksh.Cells[10, 21], fromwrksh.Cells[10, 21].End[Excel.XlDirection.xlDown]]).NumberFormat = "mm/dd/yyyy";
            //_cells.GetRange(fromwrksh.Range[fromwrksh.Cells[10, 21], fromwrksh.Cells[10, 21].End[Excel.XlDirection.xlDown]]).Value = DatosFechas;
            //_cells.GetRange(fromwrksh.Range[fromwrksh.Cells[10, 21], fromwrksh.Cells[10, 21].End[Excel.XlDirection.xlDown]]).Value = _cells.GetRange(fromwrksh.Range[fromwrksh.Cells[10, 21], fromwrksh.Cells[10, 21].End[Excel.XlDirection.xlDown]]).Value;



            //Creamos el object Range donde inicia nuestra tabla
            Excel.Range InicioTable = _cells.GetCell(StartColTable, StartRowTable);
            pivotData = fromwrksh.Range[fromwrksh.Range[InicioTable, InicioTable.End[Excel.XlDirection.xlToRight]], fromwrksh.Range[InicioTable, InicioTable.End[Excel.XlDirection.xlDown]]];
            if (pivotData.ListObject == null)
            {
                MessageBox.Show(@"Tiene que existir Informacion, para poder generar la Pivot Table");
                return;
            }
            //y le decimos donde empezara nuestra tabla dinamica asi
            Excel.Worksheet wrksh = _excelApp.Worksheets.Add(After: _excelApp.ActiveWorkbook.Sheets[_excelApp.ActiveWorkbook.Sheets.Count]);
            pivotCaches = _excelApp.ActiveWorkbook.PivotCaches();
            pivotDestination = wrksh.Cells[1, 1];
            wrksh.Name = Nombre_Hoja + "_Pivot";
            //Excel.PivotCache oPivotCache = _excelApp.ActiveWorkbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, pivotData);
            pivotCache = pivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, pivotData);
            //Excel.PivotTable pivotTable = wrksh.PivotTables().Add(PivotCache: oPivotCache, TableDestination: pivotDestination, TableName: "Pivot_" + Nombre_Hoja);
            pivotTable = pivotCache.CreatePivotTable(pivotDestination, "Pivot_" + Nombre_Hoja, Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent);

            //damos un poco de formato a nuestra tabla dinamica..
            pivotTable.Format(Excel.XlPivotFormatType.xlTable9);//xlPTClassic, xlReport10
            pivotTable.InGridDropZones = true;
            pivotTable.TableStyle2 = "PivotStyleDark6";
            FormatBordes(pivotTable.TableRange2);
            pivotTable.ShowTableStyleRowHeaders = true;
            pivotTable.ShowTableStyleColumnStripes = true;
            pivotTable.ShowTableStyleRowStripes = true;
            pivotTable.MergeLabels = true;
            pivotTable.DisplayNullString = true;
            pivotTable.NullString = " -- ";
            pivotTable.CompactLayoutColumnHeader = "Proyecion PM";
            pivotTable.CompactLayoutRowHeader = "EQUIPO";
            pivotTable.RowGrand = false;
            pivotTable.ColumnGrand = false;



            string[] GroupFieldsCol;
            string[] GroupFieldsTyp;
            string[] GroupFieldsTypAgrup;
            string[] SliceFilter;
            if (Nombre_Hoja == SheetName01)
            {
                //Excel.CalculatedFields ViajeVacioAjDistancia = pivotTable.CalculatedFields();
                //ColumnCalculada = ViajeVacioAjDistancia.Add("DUR_HR.", "=DUR_HR/3600", true);
                GroupFieldsCol = new string[] { "EQUIP_NO", "F(YYYYMMDD)", "TURNO", "HRS_DESP_PM" };
                GroupFieldsTyp = new string[] { "Fila", "Column", "Fila", "Agrupado" };
                GroupFieldsTypAgrup = new string[] { null, null, null, "Sum" };
                SliceFilter = new string[] { "LAST_PIT", "FLOTA_ELLIPSE", "WORK_GROUP" };
                TypeColum(GroupFieldsCol, pivotTable, GroupFieldsTyp, GroupFieldsTypAgrup, SliceFilter, 2, Nombre_Hoja, 3);
            }
            Encabezado = wrksh.Range[pivotDestination, pivotDestination.End[Excel.XlDirection.xlToRight]];
            //GetPropertyValues(Encabezado);
            CentrarRango(Encabezado);
            //Ocultar Equipos en StandBy
            OcultBlank = pivotTable.PivotFields("F(YYYYMMDD).");
            OcultBlank.PivotItems("(blank)").Visible = false;

            //OcultBlank = null;
            //OcultBlank = pivotTable.PivotFields("TURNO.");

            OcultSubt = pivotTable.PivotFields("EQUIP_NO.");
            OcultSubt.Subtotals[1] = false;
            //OcultBlank.Subtotals[1] = false;


            _excelApp.ActiveWindow.DisplayGridlines = false;
            _excelApp.ActiveWorkbook.ShowPivotTableFieldList = false;
            _excelApp.DisplayAlerts = true;
            _excelApp.ScreenUpdating = true;

        }
        private void TypeColum(string[] GroupFieldsCol, Excel.PivotTable pivotTable, string[] GroupFieldsTyp, string[] GroupFieldsTypAgrup, string[] SliceFilter, Int32 N, String Nombre_Hoja, Int32 CambiFila = 3)
        {
            Excel.PivotField Campo = null;
            string NameCampo = "";
            //string FormatoNumero = "#,##0.00";
            for (int i = 0; i < GroupFieldsCol.Length; i++)
            {
                Campo = pivotTable.PivotFields(GroupFieldsCol[i]);
                if (Campo.IsCalculated)
                {
                    NameCampo = GroupFieldsCol[i];
                    Campo.Caption = NameCampo;
                    //Campo.Name = NameCampo;
                }
                else
                {
                    NameCampo = GroupFieldsCol[i] + ".";
                    Campo.Caption = NameCampo;
                    Campo.Name = NameCampo;
                }
                if (GroupFieldsTyp[i] == "Fila")
                {
                    Campo.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }
                else if (GroupFieldsTyp[i] == "Column")
                {
                    Campo.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                }
                else if (GroupFieldsTyp[i] == "Agrupado")
                {
                    Campo.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    if (GroupFieldsTypAgrup[i] == "Sum")
                    {
                        Campo.Function = Excel.XlConsolidationFunction.xlSum;
                        if (GroupFieldsCol[i] == "HRS_DESP_PM")
                        {
                            Campo.NumberFormat = "#,##0";
                        }
                        else
                        {
                            Campo.NumberFormat = "#,##0.00";
                        }
                    }
                    else if (GroupFieldsTypAgrup[i] == "Count")
                    {
                        Campo.Function = Excel.XlConsolidationFunction.xlCount;
                    }
                }
                else if (GroupFieldsTyp[i] == "Filtro")
                {
                    Campo.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                }
                else
                {
                    Campo.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }
                Campo = null;
            }

           
            float Top = (float)((double)pivotTable.TableRange2.Top);
            Int32 Tam = SliceFilter.Length;
            var Width = 120;
            var Height = 150;
            float Left;
            //float Top;
            //Int32 Cambio = 3;
            Int32 TopCambio = 0;
            Int32 x = 0;
            for (int i = 0; i < Tam; i++)
            {
                Left = (float)((double)pivotTable.TableRange2.Width);
                //Top = (float)((double)pivotTable.TableRange2.Top);
                if (i == CambiFila)
                {
                    x = 0;
                    TopCambio++;
                    //Left = (float)((double)pivotTable.TableRange2.Width);
                    //Top = (float)((double)pivotTable.TableRange2.Top);
                    Top = Top + (Height * TopCambio);
                }
                //+"."
                Left = Left + (Width * x);
                Campo = pivotTable.PivotFields(SliceFilter[i]);
                _excelApp.ActiveWorkbook.SlicerCaches.Add2(Source: pivotTable, SourceField: Campo, Name: Campo.Caption + Campo.SourceName, SlicerCacheType: Excel.XlSlicerCacheType.xlSlicer).Slicers.Add(SlicerDestination: Nombre_Hoja + "_Pivot", Name: Campo.SourceName, Caption: Campo.SourceName, Top: Top, Left: Left, Width: Width, Height: Height);
                x++;
            }


            _excelApp.Columns.AutoFit();
            _excelApp.Rows.AutoFit();
        }
        private void InsertarColumn(string Nombre_Hoja, string Ubc, string Ubct, string Name)
        {
            Excel.Worksheet fromwrksh = _excelApp.ActiveWorkbook.Sheets[Nombre_Hoja];
            Excel.Range oRng = fromwrksh.Range[Ubc];
            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oRng = fromwrksh.Range[Ubct];
            oRng.Value2 = Name;
        }
        private void ChangeFormatColumn(string Nombre_Hoja, string Format, string Ubic)
        {
            Excel.Worksheet fromwrksh = _excelApp.ActiveWorkbook.Sheets[Nombre_Hoja];
            fromwrksh.get_Range(Ubic).NumberFormat = Format;
            fromwrksh.get_Range(Ubic).Value = fromwrksh.get_Range(Ubic).Value;
        }
        private void btnProy_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;
            //_excelApp.Calculation = XlCalculation.xlCalculationManual;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                try
                {
                    string Nombre_Hoja = _excelApp.ActiveWorkbook.ActiveSheet.Name;
                    //Nombre_Hoja = Nombre_Hoja + "_Pivot";
                    if (ComprobarHoja(Nombre_Hoja + "_Pivot"))
                    {
                        if (Confir.ShowDialog() == DialogResult.OK)
                        {
                            BorrarSheets(Nombre_Hoja + "_Pivot");
                        }
                        else //if(Confir.ShowDialog() == DialogResult.Cancel)
                        {
                            _excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                            return;
                        }
                    }
                    TablaDinamica(Nombre_Hoja);
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse:btnProy_Click()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                    MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
                }
                finally
                {
                    if (_cells != null)
                    _cells.SetCursorDefault();
                    //_excelApp.ActiveWorkbook.Sheets[SheetName01].Select();
                    //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                    _excelApp.ScreenUpdating = true;
                    _excelApp.DisplayAlerts = true;
                }
            }
            else
            {
                MessageBox.Show(@"debe formatear o cambiar a la Hoja " + SheetName01 + " e intente nuevamente.");
            }
        }
        private void Prueba()
        {
            String Nombre_Hoja = _excelApp.ActiveWorkbook.ActiveSheet.Name;
            //Nombre_Hoja = Nombre_Hoja + "_Pivot";
            if ("Sheet1" == Nombre_Hoja)
            {
                if (Confir.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show(@"Aceptastes");
                    return;
                }
                else
                {
                    //MessageBox.Show(@"Cancelastes");
                    return;
                }

            }
        }
        private void Prueba_Click(object sender, RibbonControlEventArgs e)
        {
            Prueba();
        }
    }
}

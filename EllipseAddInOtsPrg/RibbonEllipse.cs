using System;
//using System.Data; //NUEVA
//using System.Data.SqlClient; //NUEVA
using System.Data;
using data = System.Data;
//using System.Data.SqlClient;
//using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseWorkOrdersClassLibrary;
using Authenticator = EllipseCommonsClassLibrary.AuthenticatorService;
//using Authenticator = EllipseCommonsClassLibrary.AuthenticatorService;
using WorkOrderTaskService = EllipseWorkOrdersClassLibrary.WorkOrderTaskService;
using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using ResourceReqmntsService = EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using MaterialReqmntsService = EllipseWorkOrdersClassLibrary.MaterialReqmntsService;
using EquipmentReqmntsService = EllipseWorkOrdersClassLibrary.EquipmentReqmntsService;
//using System.Data.Odbc;
//using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse;
using System.Web.Services;
using Screen = EllipseCommonsClassLibrary.ScreenService; //si es screen service
using System.IO;
using System.Runtime.InteropServices;
using VarEncript = SharedClassLibrary.Utilities.Encryption;

namespace EllipseAddInOtsPrg
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        FormAutetication.FormularioAutenticacionType _AuthG = new FormAutetication.FormularioAutenticacionType();
        private Excel.Application _excelApp;

        private const string SheetName01 = "Gantt Parada Equipo";
        private const string SheetName02 = "Labor OT";
        private const string SheetName03 = "DurationWorkOrders";
        private const string SheetName04 = "Vales OT";
        //private const string ValidationSheetName = "ValidationSheetEqGantt";
        private const string tableName01 = "Gantt";
        private const string tableName02 = "_01Labor_OT";
        private const string TableName03 = "_01DurationWorkOrders";
        private const string TableName04 = "_01Vales_OT";
        private const int titleRow = 8;
        private Thread _thread;
        private bool _progressUpdate = true;
        public String Sql = "";
        static object useDefault = Type.Missing;
        private const Int32 StartColHrs = 22;
        private const Int32 DatosAgregados = 1;

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


        private void Encabezado(data.DataTable table, String Hoja)
        {
            //Formateando columnas de encabezado
            //_excelApp.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, _cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, StartRowTable), Type.Missing, Excel.XlYesNoGuess.xlNo, Type.Missing).Name = "TiTul01";
            int cont = StartColTable;
            if (Hoja == "Gantt Parada Equipo")
            {
                Excel.Range Rango_Two = _cells.GetRange(StartColTable, StartRowTable, table.Columns.Count, StartRowTable);//SOLO LOS DATOS

                // RANGO DE DATOS
                Rango_Two.Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.Yellow));
                Rango_Two.Font.FontStyle = "Negrita";
                Rango_Two.Font.Color = System.Drawing.ColorTranslator.ToOle((Color.Black));
                //_cells.GetRange(StartColTable, StartRowTable, table.Columns.Count, StartRowTable).Font.Size = 20;
                Rango_Two.Font.Bold = true;
                Rango_Two.AutoFilter(StartRowTable, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                Excel.Range Rango_Four = _cells.GetRange(table.Columns.Count, StartRowTable, table.Columns.Count + DatosAgregados, StartRowTable + table.Rows.Count);//ULTIMOS DATOS


                //RANGO DE DATOS CALCULADOS
                _cells.GetCell(table.Columns.Count + DatosAgregados, StartRowTable).Value = "CAMBIO ?";
                FormatBordes(Rango_Four, Excel.XlBorderWeight.xlMedium);
                CentrarRango(Rango_Four);
                _cells.GetRange(table.Columns.Count, StartRowTable, table.Columns.Count + DatosAgregados, StartRowTable).Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.Yellow));
                Rango_Four.Font.FontStyle = "Negrita";
                Rango_Four.Font.Color = System.Drawing.ColorTranslator.ToOle((Color.Black));
                Rango_Four.Font.Bold = true;



                NombreColumnas = new string[(table.Columns.Count)];
                for (var i = StartColTable; i <= table.Columns.Count; i++)
                {
                    _cells.GetCell(cont, StartRowTable).Value = table.Columns[i - 1].ColumnName.Trim();
                    NombreColumnas[i - 1] = table.Columns[i - 1].ColumnName.ToString();
                    cont++;
                }

            }
            else
            {
                for (var i = StartColTable; i <= table.Columns.Count; i++)
                {
                    _cells.GetCell(cont, StartRowTable).Value = table.Columns[i].ColumnName.Trim();
                    cont++;
                }
            }

        }



        public void ExecuteQuery()
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = true;
            //Excel.Application NombreExcel = _excelApp.Application;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {

                //Excel.Names IsExist =  _excelApp.ActiveWorkbook.ActiveSheet.Names;
                //var xxxx = _excelApp.ActiveWorkbook.ActiveSheet.Names.count;//IsExist.Count;
                //Excel.Name Nombre = _excelApp.ActiveWorkbook.ActiveSheet.Names(tableName01);
                if (_excelApp.ActiveWorkbook.ActiveSheet.Names.count > 0)
                {
                    _excelApp.Application.Goto(tableName01);
                    _excelApp.Application.Selection.EntireRow.Delete();
                }

                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Select();
            }
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                //_cells.SetCursorWait();
                _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                string NameHoja = _excelApp.ActiveWorkbook.ActiveSheet.Name;
                //borrarTabla(NameHoja);
                data.DataTable table;

                String FechaFinal = "";
                Int32 HR_ADD = 0;
                String ESTADO = "";
                var sqlQuery = "";

                if (_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value != null)
                {
                    FechaFinal = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu).Value;
                }
                else
                {
                    FechaFinal = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu).Value;
                    HR_ADD = 6;
                }
                if (_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu + 1).Value == "Uncompleted")
                {
                    ESTADO = "A','O";
                }
                else if (_cells.GetCell(StartColInputMenu + 8, StartRowInputMenu + 1).Value == "All")
                {
                    ESTADO = "A','O','C";
                }
                else
                {
                    ESTADO = "C";
                }

                if (_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu + 1).Value == "TASK")
                {
                    sqlQuery = Consulta(1, 1, FechaFinal, HR_ADD, ESTADO);
                    table = getdata(sqlQuery);
                }
                else
                {
                    sqlQuery = Consulta(2, 1, FechaFinal, HR_ADD, ESTADO);
                    table = getdata(sqlQuery);
                }
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
                    if (i % 2 == 0)
                    {
                        _cells.GetRange(StartColTable, (StartRowTable + i), table.Columns.Count + DatosAgregados, (StartRowTable + i)).Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(221, 235, 247)));
                    }
                }
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).NumberFormat = "@";
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = data;
                _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value = _cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable).Value;
                //CentrarRango(_cells.GetRange(StartColTable, StartRowTable + 1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable));
                if (NameHoja == SheetName01)
                {
                    Encabezado(table, _excelApp.ActiveWorkbook.ActiveSheet.Name);
                    CentrarRango(_cells.GetRange(StartColTable, (StartRowTable + 1), ((8 + StartColTable) - 1), (table.Rows.Count + StartRowTable)));
                    CentrarRango(_cells.GetRange(10, (StartRowTable + 1), ((table.Columns.Count + StartColTable) - 2), (table.Rows.Count + StartRowTable)));
                    Excel.Range FormatTextos = _cells.GetRange(StartColTable, StartRowTable - 1, ((table.Columns.Count - 1) + StartColTable) - 1, table.Rows.Count + StartRowTable);
                    FormatTextos.Font.FontStyle = "Negrita";
                    FormatTextos.Font.ColorIndex = ColorTranslator.ToOle(Color.Black);
                    FormatTextos.Font.Size = 10;
                    FormatTextos.Font.TintAndShade = 0;
                    //FormatTable(_cells.GetRange(StartColTable, StartRowTable-1, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable), NameHoja, 1, 1);
                    FormatBordes(_cells.GetRange(StartColTable, StartRowTable - 1, (table.Columns.Count + StartColTable) - 1, (table.Rows.Count + StartRowTable)));
                    SobreEncabezado(table.Columns.Count);
                    FinColTablaOneSheet = table.Columns.Count;
                    FinRowTablaOneSheet = table.Rows.Count;
                    _cells.GetRange(StartColTable, StartRowTable - 1, (table.Columns.Count + StartColTable) - 1, (table.Rows.Count + StartRowTable)).Select();
                    _excelApp.ActiveWorkbook.Names.Add("Gantt", _cells.GetRange(StartColTable, StartRowTable - 1, (table.Columns.Count + StartColTable) - 1, (table.Rows.Count + StartRowTable)));
                }
                else
                {
                    Encabezado(table, _excelApp.ActiveWorkbook.ActiveSheet.Name);
                    FormatTable(_cells.GetRange(StartColTable, StartRowTable, (table.Columns.Count + StartColTable) - 1, table.Rows.Count + StartRowTable), NameHoja, 1, 1);
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
        private void btnFormatear_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = true;
            //búsquedas especiales de tabla
            //_cells.SetCursorWait();
            _AuthG.StartPosition = FormStartPosition.CenterScreen;
            if (_AuthG.ShowDialog() == DialogResult.OK)
            {
                if (_AuthG.Permiso == "2")
                {
                    menuAcciones.Items[3].Visible = false;
                    menuAcciones.Items[4].Visible = false;
                    menuAcciones.Items[5].Visible = false;
                }
                try
                {
                    _excelApp.Cursor = Excel.XlMousePointer.xlWait;
                    Formatear("GANTT DE PARADA - ELLIPSE 9", SheetName01, true);
                    Formatear("CARGAR LABOR - ELLIPSE 9", SheetName02);
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
            else
            {
                MessageBox.Show(@"Regrese cuando tenga autorizacion del administrador del sistema.");
                _excelApp.Application.Quit();
            }
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

        private void SubEncabezado()
        {
            //_cells.GetCell("A1").Value = "CERREJÓN";
            //Excel.Range IMG = (Excel.Range)RngImg;
            //FORMAT TITULO
            //FECHAS DE LA HOJA 
            FormatCamposMenu(_cells.GetCell(StartColInputMenu, StartRowInputMenu), true, "FECHA DESDE");
            FormatBordes(_cells.GetCell(StartColInputMenu, StartRowInputMenu));
            FormatCamposMenu(_cells.GetCell(StartColInputMenu, StartRowInputMenu + 1), true, "FECHA HASTA");
            FormatBordes(_cells.GetCell(StartColInputMenu, StartRowInputMenu + 1));
            // AGRGADO DE LISTAS DESPLEGABLES DE LAS FECHAS
            //var List_2 = string.Join(Separador(), ListaDatos(2));
            Excel.Range Fecha1 = _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu);
            ///Fecha1.Validation.Delete();
            Fecha1.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertWarning, Excel.XlFormatConditionOperator.xlBetween, string.Join(Separador(), ListaDatos(2)), Type.Missing);
            Fecha1.Validation.IgnoreBlank = true;
            Fecha1.Validation.ShowError = false;
            Fecha1.Copy();
            //Fecha1.Value = ListaDatos(2)[0];
            //Fecha1.Value = "";
            //DateTime dateToDisplay = DateTime.Now;
            _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).PasteSpecial();
            _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value = "'" + DateTime.Now.ToString("yyyyMMdd");
            //_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value = "'20200226";
            //FORMATOS A CAMPOS FECHAS
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu), false, "", "AAAAMMDD");
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1), false, "", "AAAAMMDD");
            FormatBordes(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu));
            FormatBordes(_cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1));


            // CAMPOS DE FILTROS DE EQUIPOS FLOTAS Y TYPE CONSULTA


            //EQUIPOS Y FLOTAS DE LA HOJA 
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu), true, "EQUIPO DESDE");
            FormatBordes(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu));
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu + 1), true, "EQUIPO HASTA");
            FormatBordes(_cells.GetCell(StartColInputMenu + 3, StartRowInputMenu + 1));
            //IMPUT CAMPOS EQUIPOS y FLOTAS
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu), false, "", "EJ: 0220251");
            FormatCamposMenu(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1), false, "", "EQUIPO FORMATO ELLIPSE - [0220906] O [0050025]");
            FormatBordes(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu));
            FormatBordes(_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1));
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value = "0050025";
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value = "0050025";
            // AGREGADO DE LISTAS DESPLEGABLES PARA FLOTAS
            //var List = string.Join(Separador(), ListaDatos(1));
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.Delete();
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, List, Type.Missing);
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.IgnoreBlank = true;
            //_cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Validation.ShowError = true;



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
                Sql = (@"SELECT DISTINCT
                          FECHA112
                        FROM
                          SIGMAN.HIST_TURNOS
                        WHERE
                          FECHA112 >= TO_CHAR(ADD_MONTHS(SYSDATE,-1),'YYYYMMDD')
                        ORDER BY
                          1 " + ORDEN);
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
                table = getdata(Sql, "54uYIi2N+rjk+uTlaspHWP2OjjE3SPTT3frFZk+NV5vxGT04vVZyxLGeZyiiWk01uzrCQEo3OgZFg2YPMVksCaVAPS8w6T4zT0ZQZ6DDdbgtdRbST3obu0g8Y0KX1oSI", "zmkMKxigvBzWvgW6j2/Mrj49SchBcuaP1Hah7cW+5Sx0OYUdaJ0QNuzIMAwCFaVsPPruPt9tT9mcFsGx4JotWlKjZySpjLUN8YPmW0uhkWjSB57XVA2bOSuWAEf2jRF+", "6EoC7MZbMxosgzMyRQj2+7tDUxwofJn/gb2Vz+JbR4uSrw5wWLNy47zpLZ8SuRaxJE5qpK2mRIODKUgkrpCZ7Y6c8ucDbITmpgh62KJ/PwQ2ozIBh7VhGp6XK8CX5cRo", "");
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

        public data.DataTable getdata(string SQL, string DataBase = "CrOkubls0sZ8lj8iUOR+QY18P9jBSp7MV17Q1hMCt0zpW2WGmMHYV5XXc8j/FdQQNSMJhAHs3GXzbxU0zB+CNt5K1PIiJBvP7RlVJqPn+vHh1mLdhaACGMniPn234d2s", string User = "x4yNNf5qsgLpNdA1xUaBM1GaKhwrINqfzNsmDA7rZmZWVx8308y12p1zvsIuEzx+yszVVnhqhQ1cFWL+lBB8yYb53Yx1kBkvdWcXspKfG8buz4RuwCjtXcXkvGOQwdzw", string Pw = "M8/fjRkEAGaWFKtzyECz8mlJluF8xZevecMTrJ8tf0uboneZPAzICqYYB1WWx23w6sF5AXHDY3MtMZNJVGJ1ALO2D76lFq0M9fLmnU8Q8aOYcANWnlQCQzpX/EqnO8Ow", string DbLink = "@DBLELLIPSE8")
        {
            _eFunctions.SetDBSettings(VarEncript.Encryption.Decrypt(DataBase), VarEncript.Encryption.Decrypt(User), VarEncript.Encryption.Decrypt(Pw), DbLink);
            var dat = _eFunctions.GetQueryResult(SQL);
            data.DataTable DATA = new data.DataTable();
            DATA.Load(dat);
            return DATA;
        }

        public string Consulta(Int32 Tipe, Int32 Hoja, string FechaFinal = "", Int32 HR_ADD = 6, string ESTADO = "A','O")
        {
            string sqlQuery = "";
            if (Hoja == 1)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT

                                      DATOS.*
                                    FROM
                                    (
                                      WITH PRIMERA AS
                                      (
                                        SELECT

                                        TR.PLAN_STR_DATE,
                                        TR.PLAN_STR_TIME,
                                        TR.PLAN_FIN_DATE,
                                        TR.PLAN_FIN_TIME,
                                        TRIM(OT.EQUIP_NO) AS EQUIP_NO,
                                        TR.WORK_ORDER,
                                        OT.RELATED_WO,
                                        TR.WO_TASK_NO,
                                        TR.WO_TASK_DESC AS WO_DESC,
                                        --EQ.EQUIP_LOCATION,
                                        OT.MAINT_SCH_TASK,
                                        TR.TASK_STATUS_M AS WO_STATUS_M,
                                        TRIM(EQ.EQUIP_GRP_ID) AS FLOTA,
                                        OT.STD_JOB_NO,
                                        TR.COMP_CODE,
                                        OT.MAINT_TYPE,
                                        OT.WO_TYPE,
                                        TR.WORK_GROUP,
                                        TR.TSK_DUR_HOURS AS EST_DUR_HRS,
                                        /*CASE 
										  WHEN TR.WO_TASK_NO = '001' THEN SIGMAN.FNU_INDICADORES_PROGRAM(TR.WORK_ORDER,7)
										END AS LAB_EST,*/
                                        CASE

                                          WHEN TR.WO_TASK_NO = '001' THEN SIGMAN.FNU_INDICADORES_PROGRAM(TR.WORK_ORDER, 4)

                                        END AS DUR_REAL,
                                        TR.EST_LAB_COST AS LABOR_EST,
                                        CASE

                                          WHEN TR.WO_TASK_NO = '001' THEN SIGMAN.FNU_INDICADORES_PROGRAM(TR.WORK_ORDER, 5)

                                        END AS LAB_REAL,
                                        TR.TASK_PRIORITY AS ORIG_PRIORITY,
                                        --COT.CALC_LAB_HRS,
                                        LOCATION_TO.REF_CODE,
                                        COLOR.REF_CODE_C,
                                        FIRST_VALUE(TR.PLAN_STR_DATE) OVER(ORDER BY TR.PLAN_STR_DATE ASC) AS F_Min,
                                        SECUENCIA.REF_CODE_SEC,
                                          FIRST_VALUE(TR.PLAN_STR_DATE || TR.PLAN_STR_TIME) OVER(ORDER BY TR.PLAN_STR_DATE || TR.PLAN_STR_TIME ASC NULLS LAST) AS F_HR_PSTART_MIN,
                                          FIRST_VALUE(TR.PLAN_FIN_DATE || TR.PLAN_FIN_TIME) OVER(ORDER BY TR.PLAN_FIN_DATE || TR.PLAN_FIN_TIME DESC NULLS FIRST) AS F_HR_PFIN_MAX/*, 
										  CASE
										WHEN 
										  FIRST_VALUE(TR.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_STR_DATE ASC)  = '        ' THEN TR.PLAN_STR_DATE
										ELSE
										  FIRST_VALUE(TR.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_STR_DATE ASC)
										END AS PLAN_STR_DATE_MIN,
										CASE
										WHEN 
										  FIRST_VALUE(TR.PLAN_FIN_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_FIN_DATE DESC)= '        ' THEN TR.PLAN_STR_DATE
										ELSE
										  FIRST_VALUE(TR.PLAN_FIN_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_FIN_DATE DESC)
										END AS PLAN_FIN_DATE_MAX,
										CASE 
										WHEN 
										  FIRST_VALUE(TR.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO  ORDER BY TR.PLAN_STR_DATE, TR.PLAN_STR_TIME  ASC)= '      ' THEN TR.PLAN_STR_TIME
										ELSE
										  FIRST_VALUE(TR.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO  ORDER BY TR.PLAN_STR_DATE, TR.PLAN_STR_TIME  ASC)
										END AS PLAN_STR_TIME_MIN,
										CASE
										WHEN
										  FIRST_VALUE(TR.PLAN_FIN_TIME) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_FIN_DATE DESC, TR.PLAN_FIN_TIME DESC)= '      ' THEN TR.PLAN_FIN_TIME
										ELSE
										  FIRST_VALUE(TR.PLAN_FIN_TIME) OVER(PARTITION BY OT.EQUIP_NO ORDER BY TR.PLAN_FIN_DATE DESC, TR.PLAN_FIN_TIME DESC)
										END AS PLAN_FIN_TIME_MAX,
										TO_DATE(FIRST_VALUE(TR.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO,COLOR.REF_CODE_C ORDER BY TR.PLAN_STR_DATE ASC ) || FIRST_VALUE(TR.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO,COLOR.REF_CODE_C ORDER BY TR.PLAN_STR_TIME ASC ),'YYYY/MM/DD hh24:mi:ss') AS F_R
										*/

                                        FROM

                                        ELLIPSE.MSF620@DBLELLIPSE8 OT

                                        INNER JOIN ELLIPSE.MSF600@DBLELLIPSE8 EQ ON EQ.EQUIP_NO = OT.EQUIP_NO

                                        INNER JOIN ELLIPSE.MSF623@DBLELLIPSE8 TR ON OT.WORK_ORDER = TR.WORK_ORDER
                                        --INNER JOIN ELLIPSE.MSF621@DBLELLIPSE8 COT ON COT.WORK_ORDER = OT.WORK_ORDER

                                        LEFT JOIN
                                        (
                                          SELECT

                                          RC.REF_CODE AS REF_CODE,
                                          SUBSTR(RC.ENTITY_VALUE, 6, 8) AS NO_OT

                                          FROM

                                          ELLIPSE.MSF071@DBLELLIPSE8 RC,
                                          ELLIPSE.MSF070@DBLELLIPSE8 RCE

                                          WHERE

                                          RC.ENTITY_TYPE = RCE.ENTITY_TYPE

                                          AND RC.REF_NO = RCE.REF_NO

                                          AND RCE.ENTITY_TYPE = 'WKO'

                                          AND RC.REF_NO = '031'

                                          AND RC.SEQ_NUM = '001'
                                        )LOCATION_TO ON OT.WORK_ORDER = LOCATION_TO.NO_OT

                                        LEFT JOIN
                                        (
                                            SELECT

                                            RC.REF_CODE AS REF_CODE_C,
                                            SUBSTR(RC.ENTITY_VALUE, 6, 8) AS NO_OT_C

                                            FROM

                                            ELLIPSE.MSF071@DBLELLIPSE8 RC,
                                            ELLIPSE.MSF070@DBLELLIPSE8 RCE

                                            WHERE

                                            RC.ENTITY_TYPE = RCE.ENTITY_TYPE

                                            AND RC.REF_CODE IS NOT NULL

                                            AND RC.REF_NO = RCE.REF_NO

                                            AND RCE.ENTITY_TYPE = 'WKO'

                                            AND RC.REF_NO = '025'

                                        )COLOR ON OT.WORK_ORDER = COLOR.NO_OT_C

                                        LEFT JOIN
                                        (
                                            SELECT

                                            TRIM(RC.REF_CODE) AS REF_CODE_SEC,
                                            SUBSTR(RC.ENTITY_VALUE, 6, 8) AS SEC_OT

                                            FROM

                                            ELLIPSE.MSF071@DBLELLIPSE8 RC,
                                            ELLIPSE.MSF070@DBLELLIPSE8 RCE

                                            WHERE

                                            RC.ENTITY_TYPE = RCE.ENTITY_TYPE

                                            AND RC.REF_NO = RCE.REF_NO

                                            AND RCE.ENTITY_TYPE = 'WKO'

                                            AND RC.REF_NO = '036'
                                        )SECUENCIA ON OT.WORK_ORDER = SECUENCIA.SEC_OT

                                        WHERE

                                        OT.DSTRCT_CODE = 'ICOR'
                                        --AND COT.DSTRCT_CODE = 'ICOR'

                                        AND EQ.DSTRCT_CODE = 'ICOR'
                                        --AND COLOR.REF_CODE_C IS NOT NULL

                                        AND TR.TASK_STATUS_M IN('" + ESTADO + @"')--AND--OT.PLAN_STR_DATE > TO_CHAR(SYSDATE - 15, 'YYYYMMDD')--ADD_MONTHS(TO_CHAR(SYSDATE, 'DD/MM/YYYY'), -1)

                                        AND TR.PLAN_STR_DATE BETWEEN TO_CHAR(TO_DATE('" + FechaFinal + @"', 'YYYYMMDD') + " + HR_ADD + @", 'YYYYMMDD') AND '" + _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value + @"'

                                        AND TRIM(OT.EQUIP_NO) >= '" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value + @"'

                                        AND TRIM(OT.EQUIP_NO) <= '" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value + @"'
                                        /*AND
                                        (
                                            LOCATION_TO.REF_CODE IS NULL

                                          OR

                                            LOCATION_TO.REF_CODE = 'TL                                      '
                                        )*/
                                          --AND EQ.EQUIP_GRP_ID = 'EH5000'
                                        --ORDER BY
                                          --OT.EQUIP_NO,
                                          --TO_NUMBER(TRIM(COLOR.REF_CODE_C)),
                                          --SECUENCIA.REF_CODE_SEC
                                      )
                                      SELECT
                                        PRIMERA.FLOTA,
                                        PRIMERA.EQUIP_NO,
                                        PRIMERA.PLAN_STR_DATE,
                                        PRIMERA.PLAN_STR_TIME,
                                        PRIMERA.PLAN_FIN_DATE,
                                        PRIMERA.PLAN_FIN_TIME,
                                        PRIMERA.WORK_ORDER,
                                        PRIMERA.RELATED_WO,
                                        PRIMERA.WO_TASK_NO AS TASK,
                                        PRIMERA.WO_STATUS_M AS STUS,
                                        PRIMERA.WO_DESC DESCRIPCION,
                                        PRIMERA.EST_DUR_HRS AS DUR_EST, --OK

                                        PRIMERA.DUR_REAL, --OK

                                        PRIMERA.LABOR_EST AS LAB_EST, --X

                                        PRIMERA.LAB_REAL, --X

                                        PRIMERA.ORIG_PRIORITY AS PRI,
                                        PRIMERA.REF_CODE UBIC,
                                        PRIMERA.REF_CODE_C AS COD,
                                        PRIMERA.REF_CODE_SEC SEC,
                                        TRUNC((((TO_DATE(PRIMERA.F_HR_PFIN_MAX, 'YYYYMMDD HH24MISS')) - (TO_DATE(PRIMERA.F_HR_PSTART_MIN, 'YYYYMMDD HH24MISS'))) * 24), 1) AS PARADA_EQUIPO

                                        FROM

                                          PRIMERA
                                    ) DATOS
                                    ORDER BY

                                        DATOS.EQUIP_NO,DATOS.PLAN_STR_DATE||DATOS.PLAN_STR_TIME ASC, TRIM(DATOS.SEC) ASC";
                    //TRIM(DATOS.COD),WORK_ORDER,TASK,DATOS.PLAN_STR_DATE || DATOS.PLAN_STR_TIME ASC, TRIM(DATOS.SEC) ASC
                }
                else if (Tipe == 2)
                {
                    sqlQuery = @"SELECT
									  DATOS.*
									FROM
									( 
									  WITH PRIMERA AS 
									  (
										SELECT
										  OT.PLAN_STR_DATE,
										  OT.PLAN_STR_TIME,
										  OT.PLAN_FIN_DATE,
										  OT.PLAN_FIN_TIME,
										  TRIM(OT.EQUIP_NO) AS EQUIP_NO,
										  OT.WORK_ORDER,
										  OT.RELATED_WO,
										  OT.WO_DESC,
										  --EQ.EQUIP_LOCATION,
										  OT.MAINT_SCH_TASK,
										  OT.WO_STATUS_M,
										  TRIM(EQ.EQUIP_GRP_ID) AS FLOTA,
										  OT.STD_JOB_NO,
										  OT.COMP_CODE,
										  OT.MAINT_TYPE,
										  OT.WO_TYPE,
										  OT.WORK_GROUP,
										  COT.EST_DUR_HRS,
										  SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,4) AS DUR_REAL,
										  NVL(SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,1,1),0) AS LABOR_EST,
										  --COT.EST_LAB_HRS AS LABOR_EST,
										  SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,5) AS LAB_REAL,
										  OT.ORIG_PRIORITY,
										  --COT.CALC_LAB_HRS,
										  LOCATION_TO.REF_CODE,
										  COLOR.REF_CODE_C,
										  FIRST_VALUE(OT.PLAN_STR_DATE) OVER(ORDER BY OT.PLAN_STR_DATE ASC) AS F_Min, 
										  SECUENCIA.REF_CODE_SEC,
										  FIRST_VALUE(OT.PLAN_STR_DATE||OT.PLAN_STR_TIME) OVER(ORDER BY OT.PLAN_STR_DATE||OT.PLAN_STR_TIME ASC NULLS LAST) AS F_HR_PSTART_MIN, 
										  FIRST_VALUE(OT.PLAN_FIN_DATE||OT.PLAN_FIN_TIME) OVER(ORDER BY OT.PLAN_FIN_DATE||OT.PLAN_FIN_TIME DESC NULLS FIRST) AS F_HR_PFIN_MAX/*, 
										  CASE
											WHEN 
											  FIRST_VALUE(OT.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_STR_DATE ASC)  = '        ' THEN OT.PLAN_STR_DATE
											ELSE
											  FIRST_VALUE(OT.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_STR_DATE ASC)
										  END AS PLAN_STR_DATE_MIN,
										  CASE
											WHEN 
											  FIRST_VALUE(OT.PLAN_FIN_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_FIN_DATE DESC)= '        ' THEN OT.PLAN_STR_DATE
											ELSE
											  FIRST_VALUE(OT.PLAN_FIN_DATE) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_FIN_DATE DESC)
										  END AS PLAN_FIN_DATE_MAX,
										  CASE 
											WHEN 
											  FIRST_VALUE(OT.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO  ORDER BY OT.PLAN_STR_DATE, OT.PLAN_STR_TIME  ASC)= '      ' THEN OT.PLAN_STR_TIME
											ELSE
											  FIRST_VALUE(OT.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO  ORDER BY OT.PLAN_STR_DATE, OT.PLAN_STR_TIME  ASC)
										  END AS PLAN_STR_TIME_MIN,
										  CASE
											WHEN
											  FIRST_VALUE(OT.PLAN_FIN_TIME) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_FIN_DATE DESC, OT.PLAN_FIN_TIME DESC)= '      ' THEN OT.PLAN_FIN_TIME
											ELSE
											  FIRST_VALUE(OT.PLAN_FIN_TIME) OVER(PARTITION BY OT.EQUIP_NO ORDER BY OT.PLAN_FIN_DATE DESC, OT.PLAN_FIN_TIME DESC)
										  END AS PLAN_FIN_TIME_MAX,
										  TO_DATE(FIRST_VALUE(OT.PLAN_STR_DATE) OVER(PARTITION BY OT.EQUIP_NO,COLOR.REF_CODE_C ORDER BY OT.PLAN_STR_DATE ASC ) || FIRST_VALUE(OT.PLAN_STR_TIME) OVER(PARTITION BY OT.EQUIP_NO,COLOR.REF_CODE_C ORDER BY OT.PLAN_STR_TIME ASC ),'YYYY/MM/DD hh24:mi:ss') AS F_R
										*/
										FROM
										  ELLIPSE.MSF620@DBLELLIPSE8 OT
										  INNER JOIN ELLIPSE.MSF600@DBLELLIPSE8 EQ ON EQ.EQUIP_NO=OT.EQUIP_NO
										  INNER JOIN ELLIPSE.MSF621@DBLELLIPSE8 COT ON COT.WORK_ORDER=OT.WORK_ORDER
										  LEFT JOIN 
										  (
											  SELECT  
												RC.REF_CODE AS REF_CODE,
												SUBSTR(RC.ENTITY_VALUE,6,8) AS NO_OT
											  FROM  
												ELLIPSE.MSF071@DBLELLIPSE8 RC, 
												ELLIPSE.MSF070@DBLELLIPSE8 RCE 
											  WHERE  
												RC.ENTITY_TYPE = RCE.ENTITY_TYPE  
												AND RC.REF_NO = RCE.REF_NO 
												AND RCE.ENTITY_TYPE = 'WKO'  
												AND RC.REF_NO = '031'  
												AND RC.SEQ_NUM = '001' 
										  )LOCATION_TO ON OT.WORK_ORDER=LOCATION_TO.NO_OT
										  LEFT JOIN
										  (
												  SELECT  
													RC.REF_CODE AS REF_CODE_C,
													SUBSTR(RC.ENTITY_VALUE,6,8) AS NO_OT_C
												  FROM  
													ELLIPSE.MSF071@DBLELLIPSE8 RC, 
													ELLIPSE.MSF070@DBLELLIPSE8 RCE 
												  WHERE  
													RC.ENTITY_TYPE = RCE.ENTITY_TYPE 
													AND RC.REF_CODE IS NOT NULL 
													AND RC.REF_NO = RCE.REF_NO 
													AND RCE.ENTITY_TYPE = 'WKO'  
													AND RC.REF_NO = '025'

										  )COLOR ON OT.WORK_ORDER=COLOR.NO_OT_C
										  LEFT JOIN
										  (
												  SELECT  
													TRIM(RC.REF_CODE) AS REF_CODE_SEC,
													SUBSTR(RC.ENTITY_VALUE,6,8) AS SEC_OT
												  FROM  
													ELLIPSE.MSF071@DBLELLIPSE8 RC, 
													ELLIPSE.MSF070@DBLELLIPSE8 RCE 
												  WHERE  
													RC.ENTITY_TYPE = RCE.ENTITY_TYPE  
													AND RC.REF_NO = RCE.REF_NO 
													AND RCE.ENTITY_TYPE = 'WKO'  
													AND RC.REF_NO = '036' 
										  )SECUENCIA ON OT.WORK_ORDER=SECUENCIA.SEC_OT
										WHERE
										  OT.DSTRCT_CODE='ICOR' 
										  AND COT.DSTRCT_CODE='ICOR'
										  AND EQ.DSTRCT_CODE='ICOR'
										  --AND COLOR.REF_CODE_C IS NOT NULL
										  AND OT.WO_STATUS_M IN  ('" + ESTADO + @"') --AND --OT.PLAN_STR_DATE > TO_CHAR(SYSDATE-15,'YYYYMMDD')--ADD_MONTHS(TO_CHAR(SYSDATE,'DD/MM/YYYY'),-1)
										  AND OT.PLAN_STR_DATE BETWEEN TO_CHAR(TO_DATE('" + FechaFinal + @"', 'YYYYMMDD') + " + HR_ADD + @", 'YYYYMMDD') AND '" + _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value + @"'
										  AND TRIM(OT.EQUIP_NO) >= '" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value + @"'
                                          AND TRIM(OT.EQUIP_NO) <= '" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value + @"'
										  /*AND 
										  (
											  LOCATION_TO.REF_CODE IS NULL
											OR
											  LOCATION_TO.REF_CODE = 'TL                                      '
										  )*/
										  --AND EQ.EQUIP_GRP_ID = 'EH5000'	  
										--ORDER BY
										  --OT.EQUIP_NO,
										  --TO_NUMBER(TRIM(COLOR.REF_CODE_C)),
										  --SECUENCIA.REF_CODE_SEC
									  )
									  SELECT
                                        PRIMERA.FLOTA,
                                        PRIMERA.EQUIP_NO,
										PRIMERA.PLAN_STR_DATE,
										PRIMERA.PLAN_STR_TIME,
										PRIMERA.PLAN_FIN_DATE,
										PRIMERA.PLAN_FIN_TIME,
										PRIMERA.WORK_ORDER,
										PRIMERA.RELATED_WO,
										'' AS TASK,
										PRIMERA.WO_STATUS_M AS STUS,
										PRIMERA.WO_DESC AS DESCRIPCION,
										PRIMERA.EST_DUR_HRS AS DUR_EST,--OK
										PRIMERA.DUR_REAL,--OK
										PRIMERA.LABOR_EST AS LAB_EST,--X
										PRIMERA.LAB_REAL,--X
										PRIMERA.ORIG_PRIORITY AS PRI,
										PRIMERA.REF_CODE AS UBIC,
										PRIMERA.REF_CODE_C AS COD,
										PRIMERA.REF_CODE_SEC AS SEC,
										TRUNC( ( ( ( TO_DATE(PRIMERA.F_HR_PFIN_MAX,'YYYYMMDD HH24MISS') )-( TO_DATE(PRIMERA.F_HR_PSTART_MIN,'YYYYMMDD HH24MISS') ) )*24) ,1) AS PARADA_EQUIPO
										FROM
										  PRIMERA
									) DATOS
									ORDER BY
										DATOS.EQUIP_NO,DATOS.PLAN_STR_DATE||DATOS.PLAN_STR_TIME ASC, TRIM(DATOS.SEC) ASC ";
                }
                else if (Tipe == 3)
                {
                    sqlQuery = (@"WITH PRIMERA AS
                                (
                                    SELECT
                                        --OT.WORK_ORDER,
                                        SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,2) AS NUMER_CUMP_PROGF,
                                        SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,1,1) AS EST_LAB,
                                        SIGMAN.FNU_INDICADORES_PROGRAM(OT.WORK_ORDER,1,2) AS REAL_LAB
                                        --OT.WO_TYPE,
                                        --OT.MAINT_TYPE,
                                        --TRIM(LOCATION_TO.REF_CODE) AS REF_CODE,
                                        --OT.RELATED_WO
                                    FROM
                                        ELLIPSE.MSF620@DBLELLIPSE8 OT
                                        INNER JOIN
                                        (
                                                SELECT  
                                                    RC.REF_CODE AS REF_CODE_C,
                                                    SUBSTR(RC.ENTITY_VALUE,6,8) AS NO_OT_C
                                                FROM  
                                                    ELLIPSE.MSF071@DBLELLIPSE8 RC, 
                                                    ELLIPSE.MSF070@DBLELLIPSE8 RCE 
                                                WHERE  
                                                    RC.ENTITY_TYPE = RCE.ENTITY_TYPE  
                                                    AND RC.REF_NO = RCE.REF_NO 
                                                    AND RCE.ENTITY_TYPE = 'WKO'  
                                                    AND RC.REF_NO = '025'
                                        )COLOR ON OT.WORK_ORDER=COLOR.NO_OT_C
                                        /*LEFT JOIN 
                                        (
                                            SELECT  
                                            RC.REF_CODE AS REF_CODE,
                                            SUBSTR(RC.ENTITY_VALUE,6,8) AS NO_OT
                                            FROM  
                                            ELLIPSE.MSF071@DBLELLIPSE8 RC, 
                                            ELLIPSE.MSF070@DBLELLIPSE8 RCE 
                                            WHERE  
                                            RC.ENTITY_TYPE = RCE.ENTITY_TYPE  
                                            AND RC.REF_NO = RCE.REF_NO 
                                            AND RCE.ENTITY_TYPE = 'WKO'  
                                            AND RC.REF_NO = '031'  
                                            AND RC.SEQ_NUM = '001' 
                                            --AND RC.REF_CODE IS NOT NULL
                                        )LOCATION_TO ON OT.WORK_ORDER=LOCATION_TO.NO_OT*/
                                    WHERE
                                        OT.DSTRCT_CODE='ICOR'
                                        AND OT.PLAN_STR_DATE BETWEEN '" + _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value + "' AND TO_CHAR(TO_DATE('" + FechaFinal + @"','YYYYMMDD')+" + HR_ADD + @",'YYYYMMDD')
                                        --AND OT.WORK_ORDER = '00818402'
                                        AND TRIM(OT.EQUIP_NO) = '" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value + @"'
                                ),
                                SEGUND AS
                                (
                                    SELECT
                                    ROUND(coalesce(SUM(PRIMERA.NUMER_CUMP_PROGF) / nullif(COUNT(PRIMERA.NUMER_CUMP_PROGF),0),0)*100,2) AS CUMP_PROGF,
                                    ROUND(coalesce(SUM(PRIMERA.REAL_LAB) / nullif(SUM(PRIMERA.EST_LAB),0),0)*100,2) AS CUMP_LABF
                                    FROM
                                    PRIMERA
                                )
                                SELECT
                                    SEGUND.*,
                                    SIGMAN.FNU_INDICADORES_PROGRAM(SIGMAN.SHARE_OT_REL('" + _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value + @"','" + _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value + "','" + HR_ADD + @"','" + FechaFinal + @"'),6) AS CUMP_DURF
                                FROM
                                    SEGUND");
                }

            }
            else if (Hoja == 2)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"WITH RES_REAL AS
                                (
                                    SELECT 
                                    TR.DSTRCT_CODE,
                                    WT.WORK_GROUP,
                                    TR.WORK_ORDER,
                                    TR.WO_TASK_NO,
                                    WT.WO_TASK_DESC,
                                    TR.RESOURCE_TYPE RES_CODE,
                                    TT.TABLE_DESC RES_DESC,
                                    SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS
                                    FROM 
                                    ELLIPSE.MSFX99 TX 
                                    INNER JOIN ELLIPSE.MSF900 TR ON(
                                                                    TR.FULL_PERIOD = TX.FULL_PERIOD 
                                                                    AND TR.WORK_ORDER = TX.WORK_ORDER 
                                                                    AND TR.USERNO = TX.USERNO
                                                                    AND TR.TRANSACTION_NO = TX.TRANSACTION_NO
                                                                    AND TR.ACCOUNT_CODE   = TX.ACCOUNT_CODE
                                                                    AND TR.REC900_TYPE    = TX.REC900_TYPE
                                                                    AND TR.PROCESS_DATE   = TX.PROCESS_DATE
                                                                    AND TR.DSTRCT_CODE    = TX.DSTRCT_CODE 
                                                                    )
                                    INNER JOIN ELLIPSE.MSF010 TT
                                    ON TT.TABLE_CODE  = TR.RESOURCE_TYPE
                                    AND TT.TABLE_TYPE = 'TT'
                                    LEFT JOIN ELLIPSE.MSF623 WT
                                    ON WT.DSTRCT_CODE    = TR.DSTRCT_CODE
                                    AND WT.WORK_ORDER    = TR.WORK_ORDER
                                    AND WT.WO_TASK_NO    = TR.WO_TASK_NO
                                    WHERE 
                                    TR.DSTRCT_CODE = 'ICOR'
                                    AND TR.WORK_ORDER IN ('" + FechaFinal + @"')
                                    --AND TR.WORK_ORDER IN ('SS033423','SS033570','SS033575','SS033416')
                                    --AND TR.WO_TASK_NO    = 'SS033423','SS033570','SS033575','SS033416' 
                                    AND POSTED_STATUS = 'B'
                                    GROUP BY 
                                    TR.DSTRCT_CODE,
                                    WT.WORK_GROUP,
                                    TR.WORK_ORDER,
                                    TR.WO_TASK_NO,
                                    WT.WO_TASK_DESC,
                                    TR.RESOURCE_TYPE,
                                    TT.TABLE_DESC
                                    ),
                                    RES_EST AS
                                    (
                                    SELECT 
                                        TSK.DSTRCT_CODE,
                                        TSK.WORK_GROUP,
                                        TSK.WORK_ORDER,
                                        TSK.WO_TASK_NO,
                                        TSK.WO_TASK_DESC,
                                        RS.RESOURCE_TYPE RES_CODE,
                                        TT.TABLE_DESC RES_DESC,
                                        TO_NUMBER(RS.CREW_SIZE) QTY_REQ,
                                        RS.EST_RESRCE_HRS
                                    FROM 
                                    ELLIPSE.MSF623 TSK
                                    INNER JOIN ELLIPSE.MSF735 RS ON(RS.KEY_735_ID = TSK.DSTRCT_CODE||TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.REC_735_TYPE = 'WT')
                                    INNER JOIN ELLIPSE.MSF010 TT ON(TT.TABLE_CODE      = RS.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT')
                                    WHERE 
                                    TSK.DSTRCT_CODE = 'ICOR'
                                    AND TSK.WORK_ORDER IN ('" + FechaFinal + @"')
                                    --AND TSK.WORK_ORDER IN ('SS033423','SS033570','SS033575','SS033416')
                                    --AND TSK.WO_TASK_NO    = 'SS033423','SS033570','SS033575','SS033416'
                                    ),
                                    TABLA_REC AS
                                    (
                                    SELECT 
                                        DECODE(RES_EST.DSTRCT_CODE,NULL,RES_REAL.DSTRCT_CODE,RES_EST.DSTRCT_CODE) DSTRCT_CODE,
                                        DECODE(RES_EST.WORK_GROUP,NULL,RES_REAL.WORK_GROUP,RES_EST.WORK_GROUP) WORK_GROUP,
                                        DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER,
                                        DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO,
                                        DECODE(RES_EST.WO_TASK_DESC,NULL,RES_REAL.WO_TASK_DESC,RES_EST.WO_TASK_DESC) WO_TASK_DESC,
                                        DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE,
                                        DECODE(RES_EST.RES_DESC,NULL,RES_REAL.RES_DESC,RES_EST.RES_DESC) RES_DESC,
                                        RES_EST.QTY_REQ,
                                        RES_REAL.ACT_RESRCE_HRS,
                                        RES_EST.EST_RESRCE_HRS
                                    FROM RES_REAL
                                    FULL JOIN RES_EST ON(RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO AND RES_REAL.RES_CODE   = RES_EST.RES_CODE)
                                    )
                                SELECT 
                                    TABLA_REC.DSTRCT_CODE,
                                    TABLA_REC.WORK_GROUP,
                                    TABLA_REC.WORK_ORDER,
                                    TABLA_REC.WO_TASK_NO,
                                    TABLA_REC.WO_TASK_DESC,
                                    'M' AS ACCIÓN,
                                    'LAB' REQ_TYPE,
                                    '' SEQ_NO,
                                    TABLA_REC.RES_CODE,
                                    TRIM(TABLA_REC.RES_DESC) AS RES_DESC,
                                    '' UNITS,
                                    TABLA_REC.QTY_REQ,
                                    NULL QTY_ISS,
                                    DECODE(TABLA_REC.EST_RESRCE_HRS, NULL, 0, TABLA_REC.EST_RESRCE_HRS) EST_RESRCE_HRS,
                                    DECODE(TABLA_REC.ACT_RESRCE_HRS, NULL, 0, TABLA_REC.ACT_RESRCE_HRS) ACT_RESRCE_HRS
                                FROM 
                                    TABLA_REC";
                }
            }
            else if (Hoja == 3)
            {
                if (Tipe == 1)
                {
                    sqlQuery = @"SELECT
                                  ELLIPSE.MSF232.EQUIP_NO AS Equipo,
                                  ELLIPSE.MSF232.WORK_ORDER AS OT,
                                  ELLIPSE.MSF620.WO_DESC AS DESCRIPCION_OT,
                                  MSF140.IREQ_NO,
                                  ELLIPSE.MSF141.STOCK_CODE,
                                  (SELECT trim(part_no) FROM ellipse.msf110 where stock_code = ELLIPSE.MSF141.STOCK_CODE and pref_part_ind = '01' and rownum = 1) as PART_NO,
                                  ELLIPSE.MSF100.STK_DESC AS DESCRIPCION_PART_NO,
                                  ELLIPSE.MSF141.QTY_REQ AS Cant_Requerida,
                                  ELLIPSE.MSF141.QTY_ISSUED AS Cant_Despachada,
                                  (SELECT trim(MNEMONIC) FROM ellipse.msf110 where stock_code = ELLIPSE.MSF141.STOCK_CODE and pref_part_ind = '01' and rownum = 1) AS MNEMONIC,
                                  ELLIPSE.MSF140.CREATION_DATE AS FECHA,
                                  --ELLIPSE.MSF141.WHOUSE_ID,
                                  --ELLIPSE.MSF140.REQUESTED_BY,
                                  --ELLIPSE.MSF620.ORIG_PRIORITY,
                                  --MSF100.DESC_LINEX1 AS ITEM_DESC,
                                  --ELLIPSE.MSF140.REQUESTED_BY AS SOLICITADO_PR,
                                  ELLIPSE.MSF140.AUTHSD_BY AS Autor,
                                  ELLIPSE.MSF140.REQ_BY_DATE,
                                  ELLIPSE.MSF140.CREATION_DATE AS Date_Entered,
                                  --(ellipse.get_soh('ICOR',ELLIPSE.MSF141.stock_code) + (ellipse.get_consign_soh('ICOR',ELLIPSE.MSF141.stock_code)) - ellipse.get_soh('ICOR',ELLIPSE.MSF141.stock_code, 'OS'||'&'||'D') - ellipse.get_soh('ICOR',ELLIPSE.MSF141.stock_code, 'DISC')) as soh_real,
                                  ellipse.get_soh('ICOR',ELLIPSE.MSF141.stock_code) as soh_total
                                  --ELLIPSE.MSF170.rop, 
                                  --ELLIPSE.MSF170.roq AS RDC,
                                  --ELLIPSE.MSF170.reorder_qty AS ROQ,
                                  --ELLIPSE.MSF170.dues_in,
                                  --(SELECT E.OFF_IN_TRANSIT FROM ELLIPSE.MSF175 E WHERE E.Stock_Code = ELLIPSE.MSF141.stock_code and full_acct_per between (select to_char(add_months(sysdate,-24),'yyyymm') from dual) and (select to_char(sysdate,'yyyymm') from dual) AND ROWNUM = 1)+(select SUM(WH_XFER_ITRANS) FROM ELLIPSE.MSF180 WHERE dstrct_code=ELLIPSE.MSF170.dstrct_code AND stock_code = ELLIPSE.MSF141.stock_code)+ELLIPSE.MSF170.CONSIGN_ITRANS AS in_transit
                                  --ELLIPSE.MSF140.AUTHSD_DATE || ELLIPSE.MSF140.AUTHSD_TIME AS FECHA_HR_VALE_AUT
                                  --CAST(SIGMAN.FNU_ANT_VALE(ELLIPSE.MSF141.STOCK_CODE,ELLIPSE.MSF232.EQUIP_NO,ELLIPSE.MSF140.CREATION_DATE || ELLIPSE.MSF140.CREATION_TIME,1) AS INT) DIAS
                                FROM
                                  ELLIPSE.MSF140
                                  INNER JOIN ELLIPSE.MSF141 ON(MSF140.DSTRCT_CODE=MSF141.DSTRCT_CODE AND MSF140.IREQ_NO=MSF141.IREQ_NO)
                                  LEFT JOIN ELLIPSE.MSF232 ON(MSF141.DSTRCT_CODE=MSF232.DSTRCT_CODE AND MSF232.REQUISITION_NO = MSF141.IREQ_NO || '  ' || '0000'  /*AND ALLOC_COUNT = '01'*/)--AGREGAR EL OTRO INDEX ALLOC
                                  LEFT JOIN ELLIPSE.MSF620 ON (MSF232.DSTRCT_CODE=MSF620.DSTRCT_CODE AND MSF232.WORK_ORDER=MSF620.WORK_ORDER)
                                  LEFT JOIN ELLIPSE.MSF100 ON(MSF141.STOCK_CODE=ELLIPSE.MSF100.STOCK_CODE)
                                  LEFT JOIN ELLIPSE.MSF170 ON(MSF141.STOCK_CODE=ELLIPSE.MSF170.STOCK_CODE)
                                  --LEFT JOIN ELLIPSE.MSFX96 MSFX96 ON(MSF141.STOCK_CODE=MSF141.STOCK_CODE AND MSFX96.IREQ_NO=MSF141.IREQ_NO)
                                WHERE
                                  (
                                    MSF140.DSTRCT_CODE='ICOR'
                                    AND MSF170.DSTRCT_CODE='ICOR'
                                    --AND ELLIPSE.MSF141.STOCK_CODE IN('003236056','000425188','000425141','000427891','003236007','000425208','000425171','000427910','003236049','003881406','003881398','000427892','003235611','000425190','000425144','000427893','000425170','000425207','003236072','000425143','000425189','003236064','003238987','003862968','003862976','000427894','003353299','000425211','000425175','000427913','003239027','000425192','000425147','000427896','003917416','000425209','000425172','000427911','003544103','000425210','000425174','000427912','003239381','000425194','000425149','000427897','003239019','000425193','000425148','003917424','000425191','000425146','000427895','002875177','003136173','003136181','000427900','002875193','003227675','000425164','000427907','002875201','000425204','000425166','000427908','002875227','003881448','003881430','000427898','002875219','003136165','003136157','000427899','002959716','000425195','000425152','000417393','000425200','000425158','000427901','000417258','000425196','000425154','000417259','000425197','000425155','000417392','000425199','000425157','000417260','000425198','000425156','000417261','000425205','000425167','003415296','000425202','000425161','000427902','002177756','002724300','002724292','003581014','002613586','002791861','000427909','003415270','003833118','003833126','000427904','000427905','001423169','000425203','000425162','000427903','000428147','000425206','003692993','000425159','003881414','003881422','003692985','003790953','003353448','003917671','002877322','002998474','003357225','003521853','003439635','003680568','003896107','003896115','003896123','003897501','003897519','003897527','000403814','000408048','000408047','002178846','003439627','003258241','003256161','003439643','003258134','003680550','002875011','003775582','002955599','003268869','002693794','002833838','003258159','003268851','003521671','002326577','003443280','001361880','003470044','002178838','003516465','003470010','003470036','003470028','003516473','003470002','002178820','003680576','002875003','003258142','002157923','003775608','003268844','003439650','003258225','002875029','003439601','003775590','003268836','003257870','003439668','002342392','002955631','002157865','002157857','003257896','003196003','002178853','003439619','003257904','002178887','002157899','002157881','003257912','003257888','003326659','002955615','003268752','003268760','003268802','003268810','003268828','003258209','003258191','003258167','003258175','002723138','002724318','003268901','003268919','003268877','003268885','003258225','003680618','000417239')
                                    --AND ELLIPSE.MSF170.EXP_ELEMENT  =  '525'
                                    --AND ELLIPSE.MSF141.STOCK_CODE IN('003353299','003239027','000427912','003239019','000427894','003470028','003258142','003258159','000424478','003470036','003258225','002342392','002157865')
                                    and MSF232.ACCT_CODE_TYPE='1'
                                    --AND MSF140.IREQ_NO = 'B29468'
                                    AND MSF232.REQ_232_TYPE='I'
                                    AND ELLIPSE.MSF232.WORK_ORDER = '" + FechaFinal + @"'
                                  )";
                }
            }

            return sqlQuery;
        }

        private void CentrarRango(Excel.Range Rango)
        {
            Rango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Rango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        private void SobreEncabezado(Int32 FinCol)
        {
            //Formateando columnas de encabezado
            //var Prueba = _cells.GetCell(FinCol, StartRowTable + 1).Value;

            Excel.Range RangoFechaTitulo = _cells.GetRange(StartColTable, StartRowTable - 1, FinCol, StartRowTable - 1);
            //_cells.GetRange(StartColTable, StartRowTable - 1, StartColHrs - 1, StartRowTable - 1).Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(91, 155, 213)));
            FormatBordes(RangoFechaTitulo, Excel.XlBorderWeight.xlMedium);
            CentrarRango(RangoFechaTitulo);
            RangoFechaTitulo.Interior.Color = System.Drawing.ColorTranslator.ToOle((Color.FromArgb(91, 155, 213)));
            RangoFechaTitulo.Font.FontStyle = "Negrita";
            RangoFechaTitulo.Font.Color = System.Drawing.ColorTranslator.ToOle((Color.White));
            RangoFechaTitulo.Font.Size = 20;
            RangoFechaTitulo.Font.Bold = true;
            RangoFechaTitulo.Font.TintAndShade = 0;
            //Rango para Merge de Titulo de Datos
            _cells.GetCell(StartColTable, StartRowTable - 1).Value = "Datos";
            CentrarRango(_cells.GetCell(StartColTable, StartRowTable - 1));
            _cells.GetRange(StartColTable, StartRowTable - 1, FinCol, StartRowTable - 1).Merge();

           
        }

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

        private void ConsultLabor()
        {
            _excelApp.Visible = true;
            _cells.SetCursorWait();
            _excelApp.ScreenUpdating = false;
            //var taskCells = new ExcelStyleCells(_excelApp, SheetName01);
            Excel.Worksheet SheetGantt = _excelApp.ActiveWorkbook.Sheets[SheetName01];
            Excel.Worksheet SheetLabor = _excelApp.ActiveWorkbook.Sheets[SheetName02];
            var WoCol = FindColumna("WORK_ORDER");
            //Excel.Range DatosOts = SheetGantt.Range[SheetGantt.Cells[StartRowTable + 1, WoCol], SheetGantt.Cells[StartRowTable + FinRowTablaOneSheet, WoCol]];
            string[] DatosWoG = GetStringArray(SheetGantt.Range[SheetGantt.Cells[StartRowTable + 1, WoCol], SheetGantt.Cells[StartRowTable + FinRowTablaOneSheet, WoCol]].Cells.Value2);
            string[] DatosWo = DatosWoG.Distinct().ToArray();
            //string[,] DatosWo = GetStringArray(SheetGantt.Range[SheetGantt.Cells[StartRowTable + 1, WoCol], SheetGantt.Cells[StartRowTable + FinRowTablaOneSheet, WoCol]].Cells.Value);
            //string[][] DatosWo = GetStringArray(SheetGantt.Range[SheetGantt.Cells[StartRowTable + 1, WoCol], SheetGantt.Cells[StartRowTable + FinRowTablaOneSheet, WoCol+1]].Cells.Value2);
            /*if (DatosWo.Length == 0)
            {
                MessageBox.Show(@"Debe existir ordenes en la pestaña del Gantt para poder consultar esta informacion.");
                return;
            }*/


            //string ListWo = string.Join("','", DatosWo);
            //_cells.GetCell(StartColTable, StartRowTable).Value2 = LIST_WORK_ORDER;
            //Int32 RowMtz = 0;
            //Int32 ColMtz = 0;
            //List<string> Encabezados = new <string> ArrayList();
            IList<string> Encabezados = new List<string>();
            List<string> Acciones = new List<string>();
            List<string> ReqType = new List<string>();
            Acciones.Add("C");
            Acciones.Add("M");
            Acciones.Add("D");
            ReqType.Add("LAB");
            ReqType.Add("MAT");
            ReqType.Add("EQU");
            Int32 FinCol = 0;
            Int32 FinRowForFormat = 0;
            var StrCol = StartColTable;
            var StrRow = StartRowTable + 1;
            var FinRow = StartRowTable + 1;
            //string ListAcciones = string.Join(Separador(), Acciones);
            //Int32 UltRow = 0;
            //Excel.Range ColOt = SheetLabor.Range[SheetLabor.Cells[StrRow, 3], SheetLabor.Cells[StrRow, 3].End[Excel.XlDirection.xlDown]];   //, SheetLabor.Cells[WoCol, StrRow].End[Excel.XlDirection.xlDown]];//SheetLabor.Cells[WoCol, StrRow]
            //ColOt.NumberFormat = "@";
            //Excel.Range ColOt = _cells.GetRange(_cells.GetCell(WoCol, StrRow), _cells.GetCell(WoCol, StrRow).End[Excel.XlDirection.xlDown]);
            for (Int32 w = 0; w < DatosWo.Length; w++)
            {
                string sqlQuery = Consulta(1, 2, DatosWo[w]);
                data.DataTable table = getdata(sqlQuery, "brw6hTk7tyzbWMnkgOAGm7T5ISbOxIDZzSuf/5nvKn94VsLindO9npazUR8CDo7/5YX0KUYHtN+VxayBURC3BPWpjIhFlX+hVWYxVGV3FBoO5gv6XYTiHcXupsZ5bm5S", "x4yNNf5qsgLpNdA1xUaBM1GaKhwrINqfzNsmDA7rZmZWVx8308y12p1zvsIuEzx+yszVVnhqhQ1cFWL+lBB8yYb53Yx1kBkvdWcXspKfG8buz4RuwCjtXcXkvGOQwdzw", "Td/V9ZKxqcRFLUfFZD15bv4qZwZIHI0IhNQjdK3EoZQL+8ZJb0vhv5x/XhxtfrN6TxiMJud/+TWSgU6GOTq5YiKRDVJMlSV+f8dswzHxZJ7xjfL8fjyYpd0rFQRMCK41", "");
                if (w == 0)
                {
                    foreach (data.DataColumn Col in table.Columns)
                    {
                        Encabezados.Add(Col.ColumnName.ToString());
                    }
                }
                string[,] data = new string[table.Rows.Count, table.Columns.Count];
                for (int Row = 0; Row < table.Rows.Count; Row++)
                {
                    for (int Col = 0; Col < table.Columns.Count; Col++)
                    {
                        //if(table.Columns[Col].ColumnName.ToString() == "ACCIÓN")
                        //{

                        //}
                        data[Row, Col] = table.Rows[Row][Col].ToString();
                    }
                }
                FinCol = table.Columns.Count;
                _cells.GetRange(StrCol, StrRow, FinCol - 5, (FinRow + (table.Rows.Count - 1))).NumberFormat = "@";
                if (table.Rows.Count == 0)
                {
                    //WoSinLabor.Add(DatosWo[w].ToString());
                    _cells.GetCell(StrCol, StrRow).Value = "'ICOR";
                    _cells.GetCell(StrCol + 2, StrRow).Value = DatosWo[w];
                    _cells.GetCell(StrCol + 6, StrRow).Value = "LAB";
                    StrRow = StrRow + 1;
                    FinRow = FinRow + 1;
                }
                else
                {
                    _cells.GetRange(StrCol, StrRow, FinCol, (FinRow + (table.Rows.Count - 1))).Value = data;
                    _cells.GetRange(StrCol, StrRow, FinCol, (FinRow + (table.Rows.Count - 1))).Value = _cells.GetRange(StrCol, StrRow, FinCol, (FinRow + (table.Rows.Count - 1))).Value;
                    StrRow = (StrRow + table.Rows.Count);
                    FinRow = (FinRow + table.Rows.Count);
                }
                //UltRow = table.Rows.Count;
            }
            FinRowForFormat = FinRow - 1;
            //Eliminar elementos repetidos de una lista
            //IEnumerable<string> ElementosDistinct = Encabezados.Distinct();
            //IEnumerable<string> ListObject = Encabezados;
            //Convertir List en Array
            string[] ArrayEncabezadoLab = new string[Encabezados.Count];
            Encabezados.CopyTo(ArrayEncabezadoLab, 0);
            Int32 ColAccion = FindColumna("ACCIÓN", ArrayEncabezadoLab);
            Int32 ColReqType = FindColumna("REQ_TYPE", ArrayEncabezadoLab);
            Int32 ColResCode = FindColumna("RES_CODE", ArrayEncabezadoLab);
            _cells.SetValidationList(_cells.GetRange(ColAccion, StartRowTable + 1, ColAccion, FinRowForFormat), Acciones);
            _cells.SetValidationList(_cells.GetRange(ColReqType, StartRowTable + 1, ColReqType, FinRowForFormat), ReqType);
            _cells.SetValidationList(_cells.GetRange(ColResCode, StartRowTable + 1, ColResCode, FinRowForFormat), ListaDatos(3, "ASC"));
            Int32 ColFormat = FindColumna("UNITS", ArrayEncabezadoLab);
            //var UnitCol = FindColumna("UNITS", Encabezados);
            //Agregando de titulos
            var ColumName = StartColTable;
            var RowName = StartRowTable;
            foreach (var ColName in Encabezados)
            {
                _cells.GetCell(ColumName, RowName).Value = ColName;
                ColumName++;
            }
            //Resultado.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
            //Resultado.Font.Bold = true;
            //FormatBordes(Resultado);
            Excel.Range Format = _cells.GetRange(StartColTable, StartRowTable, FinCol + 1, FinRowForFormat);

            FormatTable(Format, _excelApp.ActiveWorkbook.ActiveSheet.Name, 1, 1);
            //Centrar titulo del encabezado de la tabla
            CentrarRango(_cells.GetRange(StartColTable, StartRowTable, FinCol, StartRowTable));
            //Centrar valores despues de la columna UNITS
            CentrarRango(_cells.GetRange(ColFormat + 1, StartRowTable + 1, FinCol + 1, FinRowForFormat));
            //_cells.GetRange(ColFormat + 1, StartRowTable + 1, FinCol, FinRowForFormat).NumberFormat = "#,##0";
            //Centrar de las primeras tres columnas
            CentrarRango(_cells.GetRange(StartColTable, StartRowTable + 1, 4, FinRowForFormat));
            //Centrar de la Quinta hasta la ocho
            CentrarRango(_cells.GetRange(6, StartRowTable + 1, 8, FinRowForFormat));


            //FORMATEANDO ULTIMA COLUMNA
            FormatCamposMenu(Celda: _cells.GetCell(FinCol + 1, StartRowTable), Col: true, Texto: "RESULTADO", Comentario: "SI LA ACCIÓN SE EJECUTO CORRETAMENTE.", Rf: 255, Gf: 243, Bf: 0, Rl: 0, Gl: 0, Bl: 0);
            FormatBordes(_cells.GetCell(FinCol + 1, StartRowTable));

            _excelApp.ActiveWindow.Zoom = 80;
            //CentrarRango(Format);
            _excelApp.Columns.AutoFit();
            _excelApp.Rows.AutoFit();
            _excelApp.ScreenUpdating = true;
            _excelApp.DisplayAlerts = true;
            _cells.SetCursorDefault();
        }

        private Int32 FindColumna(string ColName, string[] VectorCol = null)
        {
            var Encontrado = 0;
            if (VectorCol == null)
            {
                VectorCol = NombreColumnas;
            }
            for (var i = 0; i < VectorCol.Length; i++)
            {
                if (VectorCol[i] == ColName)
                {
                    Encontrado = i + 1;
                    break;
                }
            }
            return Encontrado;
        }

        private void btnConsultarLab_Click(object sender, RibbonControlEventArgs e)
        {
            //_excelApp.DisplayAlerts = true;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                Excel.Worksheet SheetGantt = _excelApp.ActiveWorkbook.Sheets[SheetName01];
                var WoCol = FindColumna("WORK_ORDER");

                if (SheetGantt.Cells[StartRowTable + 1, StartColTable + (WoCol - 1)].Value == null)
                {
                    MessageBox.Show(@"Debe existir ordenes en la pestaña del Gantt para poder consultar esta informacion.");
                    return;
                }
                try
                {
                    _cells.DeleteTableRange(tableName02);
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ConsultLabor);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ConsultLabor()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show(@"Se ha producido un error: " + ex.Message);
                }
                finally
                {
                    if (_cells != null)
                        _cells.SetCursorDefault();
                    _eFunctions.CloseConnection();
                    _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Select();
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnActualizaDatos_Click(object sender, RibbonControlEventArgs e)
        {
            //_eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            //var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            //CalcularFechaHr();
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                Excel.Worksheet SheetGantt = _excelApp.ActiveWorkbook.Sheets[SheetName01];
                if (SheetGantt.Cells[StartRowTable + 1, StartColTable + 4].Value == null)
                {
                    MessageBox.Show(@"Debe existir ordenes en la pestaña del Gantt para poder realizar esta Acción.");
                    return;
                }
                _cells.GetCell(StartColTable + FinColTablaOneSheet, StartRowTable - 1).Select();
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {


                    if (_cells.GetCell(StartColInputMenu + 6, StartRowInputMenu + 1).Value == "TASK")
                    {

                        ActualizarGanttTaskOt(2);
                    }
                    else
                    {
                        ActualizarGanttEncabezadoOt(2);
                    }

                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ActualizarGanttTaskOt(int tipo)
        {
            int filas = StartRowTable + 1;
            String PLAN_STR_DATE = "";
            Int32 PlanStrDate = FindColumna("PLAN_STR_DATE");
            String PLAN_STR_TIME = "";
            Int32 PlanStrTime = FindColumna("PLAN_STR_TIME");
            String PLAN_FIN_DATE = "";
            Int32 PlanFinDate = FindColumna("PLAN_FIN_DATE");
            String PLAN_FIN_TIME = "";
            Int32 PlanFinTime = FindColumna("PLAN_FIN_TIME");
            String WORK_ORDER = "";
            Int32 Wo = FindColumna("WORK_ORDER");
            String WO_TASK_NO = "";
            Int32 WoTask = FindColumna("TASK");
            String WO_DESC = "";
            Int32 WoDesc = FindColumna("DESCRIPCION");
            String TSK_DUR_HOURS = "";
            Int32 TskDurHr = FindColumna("DUR_EST");
            String TASK_PRIORITY = "";
            Int32 TskPriori = FindColumna("PRI");
            String UBIC = "";
            Int32 Ubic = FindColumna("UBIC");
            String COL = "";
            Int32 Cod = FindColumna("COD");
            String SEC = "";
            Int32 Sec = FindColumna("SEC");
            while (_cells.GetCell(StartColTable, filas).Value != null)
            {
                try
                {
                    if (tipo == 1)
                    {
                        PLAN_STR_DATE = "" + _cells.GetCell(FinColTablaOneSheet + 1, filas).Value;
                        PLAN_STR_TIME = "" + _cells.GetCell(FinColTablaOneSheet + 2, filas).Value;
                        PLAN_FIN_DATE = "";
                        PLAN_FIN_TIME = "";
                    }
                    else
                    {
                        PLAN_STR_DATE = "" + _cells.GetCell(PlanStrDate, filas).Value;
                        PLAN_STR_TIME = "" + _cells.GetCell(PlanStrTime, filas).Value;
                        PLAN_FIN_DATE = "";
                        PLAN_FIN_TIME = "";
                    }
                    WORK_ORDER = "" + _cells.GetCell(Wo, filas).Value;
                    WO_TASK_NO = "" + _cells.GetCell(WoTask, filas).Value;
                    WO_DESC = "" + _cells.GetCell(WoDesc, filas).Value;
                    TSK_DUR_HOURS = "" + _cells.GetCell(TskDurHr, filas).Value;
                    TASK_PRIORITY = "" + _cells.GetCell(TskPriori, filas).Value;
                    //ReferentCode
                    UBIC = "" + _cells.GetCell(Ubic, filas).Value;
                    COL = "" + _cells.GetCell(Cod, filas).Value;
                    SEC = "" + _cells.GetCell(Sec, filas).Value;

                    var distrito = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? _frmAuth.EllipseDsct : "ICOR";
                    var userName = _frmAuth.EllipseUser.ToUpper();


                    WorkOrderTaskService.WorkOrderTaskService proxySheet_t = new WorkOrderTaskService.WorkOrderTaskService();


                    WorkOrderTaskService.WorkOrderTaskServiceModifyRequestDTO requestParamsSheet_t = new WorkOrderTaskService.WorkOrderTaskServiceModifyRequestDTO();
                    WorkOrderTaskService.WorkOrderTaskServiceModifyReplyDTO replySheet_t = new WorkOrderTaskService.WorkOrderTaskServiceModifyReplyDTO();

                    var workOrderA_t = new WorkOrderTaskService.WorkOrderDTO();

                    workOrderA_t.no = WORK_ORDER.Substring(2, 6);
                    workOrderA_t.prefix = WORK_ORDER.Substring(0, 2);

                    proxySheet_t.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/WorkOrderTaskService";

                    var opSheet_t = new WorkOrderTaskService.OperationContext
                    {
                        district = _frmAuth.EllipseDsct,
                        position = _frmAuth.EllipsePost,
                        maxInstances = 100,
                        maxInstancesSpecified = true,
                        returnWarnings = Debugger.DebugWarnings,
                        returnWarningsSpecified = true,
                    };

                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                    requestParamsSheet_t.districtCode = distrito;
                    requestParamsSheet_t.planStrDate = PLAN_STR_DATE;
                    requestParamsSheet_t.planStrTime = PLAN_STR_TIME;
                    requestParamsSheet_t.planFinDate = PLAN_FIN_DATE;
                    requestParamsSheet_t.planFinTime = PLAN_FIN_TIME;
                    requestParamsSheet_t.workOrder = workOrderA_t;
                    requestParamsSheet_t.WOTaskNo = WO_TASK_NO;
                    requestParamsSheet_t.WOTaskDesc = WO_DESC;
                    requestParamsSheet_t.priority = TASK_PRIORITY;

                    var woTask = new WorkOrderTask
                    {
                        DistrictCode = distrito,
                        WorkOrder = WORK_ORDER,
                        WoTaskNo = WO_TASK_NO,
                        EstimatedDurationsHrs = TSK_DUR_HOURS
                    };

                    woTask.SetWorkOrderDto(woTask.WorkOrder);



                    ReplyMessage replyMsg = null;



                    string messageResult = replyMsg == null ? "OK" : replyMsg.Message;

                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Value = messageResult;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Style = StyleConstants.Success;

                    replySheet_t = proxySheet_t.modify(opSheet_t, requestParamsSheet_t);
                    var reply = WorkOrderTaskActions.ModifyWorkOrderTask(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet_t, woTask);
                    WorkOrderTaskActions.SetWorkOrderTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, woTask);

                    if (_cells.GetCell(WoTask, filas).Value == "001" || _cells.GetCell(WoTask, filas).Value == "")
                    {
                        ActualizarRefCodes(filas, distrito, UBIC, COL, SEC, WORK_ORDER);
                    }

                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Value = messageResult;
                    //_cells.GetCell("GD" + filas).Style = _cells.GetStyle(StyleConstants.ItalicSmall);
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Style = StyleConstants.Success;
                    //_cells.GetCell("GD").Borders.Weight = "2";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Style = StyleConstants.Error;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ActualizarGanttTaskOt()", ex.Message);
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Select();

                }
                finally
                {
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Font.Size = 10;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Borders.Weight = 3d;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Font.Bold = true;
                    filas++;
                    _cells.GetCell(FinColTablaOneSheet + 3, filas).Select();
                    //_cells.GetCell("GC" + filas).Value = "OK";
                    //_cells.GetCell("GC" + filas).Style = StyleConstants.Success;
                }
            }
            _cells.SetCursorDefault();
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Select();
        }

        private void ActualizarRefCodes(int fila, string distrit, string UBIC, string COLOR, string SEC, string WORKORDER)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var distrito = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";

            var district = distrito;
            var workOrder = WORKORDER;
            var localiza = UBIC;
            var colores = COLOR;
            var secuencia = SEC;


            var woRefCodes = new WorkOrderReferenceCodes
            {
                Localizacion = localiza,
                CodigoCertificacion = colores,
                SecuenciaOt = SEC
                //secuencia
            };

            var replyRefCode = WorkOrderActions.UpdateWorkOrderReferenceCodes(_eFunctions, urlService, opSheet, district, workOrder, woRefCodes);


            if (replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
            {
                var errorList = "";
                // ReSharper disable once LoopCanBeConvertedToQuery
                foreach (var error in replyRefCode.Errors)
                    errorList = errorList + "\nError: " + error;
            }
        }

        private void ActualizarGanttEncabezadoOt(int tipo)
        {

            int filas = StartRowTable + 1;
            String PLAN_STR_DATE = "";
            Int32 PlanStrDate = FindColumna("PLAN_STR_DATE");
            String PLAN_STR_TIME = "";
            Int32 PlanStrTime = FindColumna("PLAN_STR_TIME");
            String PLAN_FIN_DATE = "";
            Int32 PlanFinDate = FindColumna("PLAN_FIN_DATE");
            String PLAN_FIN_TIME = "";
            Int32 PlanFinTime = FindColumna("PLAN_FIN_TIME");
            String WORK_ORDER = "";
            Int32 Wo = FindColumna("WORK_ORDER");
            String RELATED_WO = "";
            Int32 RltWo = FindColumna("RELATED_WO");
            //String WO_TASK_NO = "";
            //Int32 WoTask = FindColumna("TASK");
            String WO_DESC = "";
            Int32 WoDesc = FindColumna("DESCRIPCION");
            String DUR_HOURS = "";
            Int32 EstDurHr = FindColumna("DUR_EST");
            String PRIORITY = "";
            Int32 Priori = FindColumna("PRI");
            String UBIC = "";
            Int32 Ubic = FindColumna("UBIC");
            String COL = "";
            Int32 Cod = FindColumna("COD");
            String SEC = "";
            Int32 Sec = FindColumna("SEC");



            while (_cells.GetCell(StartColTable, filas).Value != null)
            {
                try
                {
                    if (tipo == 1)
                    {
                        PLAN_STR_DATE = "" + _cells.GetCell(FinColTablaOneSheet + 1, filas).Value;
                        PLAN_STR_TIME = "" + _cells.GetCell(FinColTablaOneSheet + 2, filas).Value;
                        PLAN_FIN_DATE = "";
                        PLAN_FIN_TIME = "";
                    }
                    else
                    {
                        PLAN_STR_DATE = "" + _cells.GetCell(PlanStrDate, filas).Value;
                        PLAN_STR_TIME = "" + _cells.GetCell(PlanStrTime, filas).Value;
                        PLAN_FIN_DATE = "";
                        PLAN_FIN_TIME = "";
                    }
                    WORK_ORDER = "" + _cells.GetCell(Wo, filas).Value;
                    RELATED_WO = "" + _cells.GetCell(RltWo, filas).Value;
                    //WO_TASK_NO = "" + _cells.GetCell(WoTask, filas).Value;
                    WO_DESC = "" + _cells.GetCell(WoDesc, filas).Value;
                    DUR_HOURS = "" + _cells.GetCell(EstDurHr, filas).Value;
                    PRIORITY = "" + _cells.GetCell(Priori, filas).Value;
                    //ReferentCode
                    UBIC = "" + _cells.GetCell(Ubic, filas).Value;
                    COL = "" + _cells.GetCell(Cod, filas).Value;
                    SEC = "" + _cells.GetCell(Sec, filas).Value;

                    var distrito = string.IsNullOrWhiteSpace(_frmAuth.EllipseDsct) ? _frmAuth.EllipseDsct : "ICOR";
                    var userName = _frmAuth.EllipseUser.ToUpper();

                    WorkOrderService.WorkOrderService proxySheet = new WorkOrderService.WorkOrderService();

                    WorkOrderService.WorkOrderServiceModifyRequestDTO requestParamsSheet = new WorkOrderService.WorkOrderServiceModifyRequestDTO();
                    WorkOrderService.WorkOrderServiceModifyReplyDTO replySheet = new WorkOrderService.WorkOrderServiceModifyReplyDTO();

                    var workOrderA = new WorkOrderService.WorkOrderDTO();
                    var workOrderB = new WorkOrderService.WorkOrderDTO();

                    workOrderA.no = WORK_ORDER.Substring(2, 6);
                    workOrderA.prefix = WORK_ORDER.Substring(0, 2);
                    /*Int32 sw = 0;
                    if (_cells.GetNullIfTrimmedEmpty(RELATED_WO) != "")
                    {
                        workOrderB.no = RELATED_WO.Substring(2, 6);
                        workOrderB.prefix = RELATED_WO.Substring(0, 2);
                        sw = 1;
                    }*/
                    workOrderB.prefix = "  ";
                    workOrderB.no = "      ";
                    if (RELATED_WO != "")
                    {
                        workOrderB.no = RELATED_WO.Substring(2, 6);
                        workOrderB.prefix = RELATED_WO.Substring(0, 2);
                    }

                    proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/WorkOrderService";

                    var opSheet = new WorkOrderService.OperationContext
                    {
                        district = _frmAuth.EllipseDsct,
                        position = _frmAuth.EllipsePost,
                        maxInstances = 100,
                        maxInstancesSpecified = true,
                        returnWarnings = Debugger.DebugWarnings,
                        returnWarningsSpecified = true,
                    };

                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                    requestParamsSheet.districtCode = distrito;
                    requestParamsSheet.workOrder = workOrderA;
                    /*if (sw == 1)
                    {
                        requestParamsSheet.relatedWo = workOrderB;
                    }*/
                    requestParamsSheet.relatedWo = workOrderB;
                    requestParamsSheet.planStrDate = PLAN_STR_DATE;
                    requestParamsSheet.planStrTime = PLAN_STR_TIME;
                    requestParamsSheet.planFinDate = PLAN_FIN_DATE;
                    requestParamsSheet.planFinTime = PLAN_FIN_TIME;
                    requestParamsSheet.workOrderDesc = WO_DESC;
                    requestParamsSheet.origPriority = PRIORITY;

                    replySheet = proxySheet.modify(opSheet, requestParamsSheet);
                    ActualizarRefCodes(filas, distrito, UBIC, COL, SEC, WORK_ORDER);

                    ReplyMessage replyMsg = null;
                    string messageResult = replyMsg == null ? "OK" : replyMsg.Message;



                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Value = messageResult;
                    //_cells.GetCell("GD" + filas).Style = _cells.GetStyle(StyleConstants.ItalicSmall);
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Style = StyleConstants.Success;
                    //_cells.GetCell("GD").Borders.Weight = "2";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Style = StyleConstants.Error;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ActualizarGanttTaskOt()", ex.Message);
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Select();

                }
                finally
                {
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Font.Size = 10;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Borders.Weight = 3d;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Font.Bold = true;
                    //_cells.GetCell("GD").Borders.Weight = "2";
                    filas++;
                    _cells.GetCell(FinColTablaOneSheet + DatosAgregados, filas).Select();
                }
            }
            _cells.SetCursorDefault();
            _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Select();
        }

        private void btnActualizarDurLab_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = true;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                try
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ExecuteTaskActions);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ExecuteTaskActions()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
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
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ExecuteTaskActions()
        {
            _cells.GetCell(StartColTable + 15, StartRowTable).Select();
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _excelApp.Visible = true;
                _excelApp.ScreenUpdating = false;

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                // if(true)
                {
                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    _cells.SetCursorWait();

                    _cells.ClearTableRangeColumn(tableName02, 16);
                    var i = StartRowTable + 1;

                    var opSheetResource = new ResourceReqmntsService.OperationContext
                    {
                        district = _frmAuth.EllipseDsct,
                        position = _frmAuth.EllipsePost,
                        maxInstances = 100,
                        returnWarnings = Debugger.DebugWarnings,
                        returnWarningsSpecified = true,
                        maxInstancesSpecified = true
                    };
                    var opSheetMaterial = new MaterialReqmntsService.OperationContext
                    {
                        district = _frmAuth.EllipseDsct,
                        position = _frmAuth.EllipsePost,
                        maxInstances = 100,
                        returnWarnings = Debugger.DebugWarnings,
                        returnWarningsSpecified = true,
                        maxInstancesSpecified = true
                    };
                    var opSheetEquipment = new EquipmentReqmntsService.OperationContext()
                    {
                        district = _frmAuth.EllipseDsct,
                        position = _frmAuth.EllipsePost,
                        maxInstances = 100,
                        returnWarnings = Debugger.DebugWarnings,
                        returnWarningsSpecified = true,
                        maxInstancesSpecified = true
                    };


                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    //Int32 Wo = FindColumna("WORK_ORDER");
                    while (!string.IsNullOrEmpty("" + _cells.GetCell(StartColTable + 2, i).Value) /*&& !string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value)*/)
                    {
                        if (_cells.GetCell(StartColTable + 3, i).Value != "" && _cells.GetCell(StartColTable + 5, i).Value != "")
                        {
                            try
                            {
                                // ReSharper disable once UseObjectOrCollectionInitializer
                                var taskReq = new TaskRequirement();
                                string action = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 5, i).Value);                         //_cells.GetCell(6, i).Value = "M";

                                taskReq.DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable, i).Value);                  //_cells.GetCell(1, i).Value = "" + req.DistrictCode; 
                                taskReq.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 1, i).Value);                     //_cells.GetCell(2, i).Value = "" + req.WorkGroup;    
                                taskReq.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 2, i).Value);                     //_cells.GetCell(3, i).Value = "" + req.WorkOrder;     
                                taskReq.WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 3, i).Value);                      //_cells.GetCell(4, i).Value = "" + req.WoTaskNo;      
                                taskReq.WoTaskNo = string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? "001" : taskReq.WoTaskNo;
                                taskReq.WoTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 4, i).Value);                    //_cells.GetCell(5, i).Value = "" + req.WoTaskDesc;
                                taskReq.ReqType = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 6, i).Value);                       //_cells.GetCell(7, i).Value = "" + req.ReqType;       
                                taskReq.SeqNo = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 7, i).Value);                         //_cells.GetCell(8, i).Value = "" + req.SeqNo;         
                                taskReq.ReqCode = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 8, i).Value);                       //_cells.GetCell(9, i).Value = "" + req.ReqCode;      
                                taskReq.ReqDesc = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 9, i).Value);                      //_cells.GetCell(10, i).Value = "" + req.ReqDesc;
                                taskReq.UoM = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 10, i).Value);                          //_cells.GetCell(11, i).Value = "" + req.UoM;
                                taskReq.QtyReq = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 11, i).Value);                       //_cells.GetCell(12, i).Value = "" + req.QtyReq;       
                                taskReq.QtyIss = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 12, i).Value);                       //_cells.GetCell(13, i).Value = "" + req.QtyIss;      
                                taskReq.HrsReq = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 13, i).Value);                       //_cells.GetCell(14, i).Value = "" + req.HrsReq;      
                                taskReq.HrsReal = _cells.GetEmptyIfNull(_cells.GetCell(StartColTable + 14, i).Value);                      //_cells.GetCell(15, i).Value = "" + req.HrsReal;     


                                if (string.IsNullOrWhiteSpace(action))
                                    continue;
                                else if (action.Equals("C"))
                                {
                                    if (taskReq.ReqType.Equals("LAB"))
                                        WorkOrderTaskActions.CreateTaskResource(urlService, opSheetResource, taskReq);
                                    else if (taskReq.ReqType.Equals("MAT"))
                                        WorkOrderTaskActions.CreateTaskMaterial(urlService, opSheetMaterial, taskReq);
                                    else if (taskReq.ReqType.Equals("EQU"))
                                        WorkOrderTaskActions.CreateTaskEquipment(urlService, opSheetEquipment, taskReq);
                                }
                                else if (action.Equals("M"))
                                {
                                    if (taskReq.ReqType.Equals("LAB"))
                                        WorkOrderTaskActions.ModifyTaskResource(urlService, opSheetResource, taskReq);
                                    else if (taskReq.ReqType.Equals("MAT"))
                                        WorkOrderTaskActions.ModifyTaskMaterial(urlService, opSheetMaterial, taskReq);
                                    else if (taskReq.ReqType.Equals("EQU"))
                                        WorkOrderTaskActions.ModifyTaskEquipment(urlService, opSheetEquipment, taskReq);
                                }
                                else if (action.Equals("D"))
                                {
                                    if (taskReq.ReqType.Equals("LAB"))
                                        WorkOrderTaskActions.DeleteTaskResource(urlService, opSheetResource, taskReq);
                                    else if (taskReq.ReqType.Equals("MAT"))
                                        WorkOrderTaskActions.DeleteTaskMaterial(urlService, opSheetMaterial, taskReq);
                                    else if (taskReq.ReqType.Equals("EQU"))
                                        WorkOrderTaskActions.DeleteTaskEquipment(urlService, opSheetEquipment, taskReq);
                                }
                                _cells.GetCell(StartColTable + 15, i).Value = "OK";
                                _cells.GetCell(StartColTable, i).Style = StyleConstants.Success;
                                _cells.GetCell(StartColTable + 15, i).Style = StyleConstants.Success;
                            }
                            catch (Exception ex)
                            {
                                if (_cells.GetCell(StartColTable + 3, i).Value == "   ")
                                {
                                    _cells.GetCell(StartColTable, i).Style = StyleConstants.Error;
                                    _cells.GetCell(StartColTable + 15, i).Style = StyleConstants.Error;
                                    _cells.GetCell(StartColTable + 15, i).Value = "Lab Save Sn Task_NO";
                                }
                                else
                                {
                                    _cells.GetCell(StartColTable, i).Style = StyleConstants.Error;
                                    _cells.GetCell(StartColTable + 15, i).Style = StyleConstants.Error;
                                    _cells.GetCell(StartColTable + 15, i).Value = "ERROR: " + ex.Message;
                                    Debugger.LogError("RibbonEllipse.cs:ExecuteTaskActions()", ex.Message);
                                }
                            }
                            finally
                            {
                                _cells.GetCell(StartColTable + 15, i).Select();
                                i++;
                            }
                        }
                        else
                        {
                            i++;
                        }
                    }

                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }


            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            _excelApp.Visible = true;
            _excelApp.ScreenUpdating = false;
            //_excelApp.DisplayAlerts = true;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu).Value = null;
                _cells.GetCell(StartColInputMenu + 1, StartRowInputMenu + 1).Value = null;
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu).Value = null;
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Value = null;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Names.count > 0)
                {
                    _excelApp.Application.Goto(tableName01);
                    _excelApp.Application.Selection.EntireRow.Delete();
                }
                _cells.GetCell(StartColInputMenu + 4, StartRowInputMenu + 1).Select();

            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _cells.DeleteTableRange(tableName02);
            }
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                _cells.DeleteTableRange(TableName03);
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            _excelApp.ScreenUpdating = true;
            //_excelApp.DisplayAlerts = true;
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
    }
}

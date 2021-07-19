using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Connections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using SharedClassLibrary.Connections.Oracle;
using Excel = Microsoft.Office.Interop.Excel;
using InvestigacionesIcam;

namespace InvestigacionesICAM
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private Excel.Application _excelApp;
        private SharedClassLibrary.Connections.Oracle.OracleConnector _conn;
        private Thread _thread;

        private int TitleRow01 = 9;
        private string TableName01 = "InvestigacionesIcam";
        private string SheetName01 = "ICAM";

        private string ValidationSheetName = "Validaciones";

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            try
            {
                Settings.Initiate();
                _excelApp = Globals.ThisAddIn.Application;

                var environments = new List<string> { "Productivo", "Test" };
                foreach (var env in environments)
                {
                    var item = Factory.CreateRibbonDropDownItem();
                    item.Label = env;
                    drpEnvironment.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #region Button Controls
        private void btnExecution_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => ExecuteQuery("All"));

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteQuery(All)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }
        private void btnReviewAccidents_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ExecuteQuery("Accidents"));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteQuery(All)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReviewRecomendations_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ExecuteQuery("Recomendations"));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteQuery(All)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReviewPlans_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ExecuteQuery("Plans"));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteQuery(All)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        #endregion
        private void StartConnection(string environment)
        {
            string connection;
            string username;
            string password;

            if (environment.ToUpper().Equals("PRODUCTIVO"))
            {
                connection = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=lmndbs03.cerrejon.com)(PORT=1521))(CONNECT_DATA=(SID=webprd10)))";
                username = "consulbo";
                password = "pwinicial";
            }
            else
            {
                connection = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=lmndbs05.cerrejon.com)(PORT=1521))(CONNECT_DATA=(SID=webtst11)))";
                username = "adminsiio";
                password = "adminsiio";
            }

            _conn = new OracleConnector(connection, username, password);
        }
        private void ExecuteQuery(string reviewType)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var cells = new ExcelStyleCells(_excelApp, true);

                var paramList = new List<KeyValuePair<int?, string>>();
                //codigo
                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("A3").Value2), "" + cells.GetCell("B3").Value2));
                //responsable
                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("A4").Value2), "" + cells.GetCell("B4").Value2));
                //potencial
                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("A5").Value2), "" + cells.GetCell("B5").Value2));
                //estado
                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("A6").Value2), "" + cells.GetCell("B6").Value2));

                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("C3").Value2), "" + cells.GetCell("D3").Value2));

                paramList.Add(new KeyValuePair<int?, string>((int?)SearchParameters.GetValueFromDescription(cells.GetCell("C4").Value2), "" + cells.GetCell("D4").Value2));


                StartConnection(drpEnvironment.SelectedItem.Label);

                var titleRow = TitleRow01;
                var tableName = TableName01;

                string sqlQuery = "";

                if (reviewType.Equals("All"))
                    sqlQuery = Queries.GetAccidentsAllQuery(paramList);
                else if (reviewType.Equals("Accidents"))
                    sqlQuery = Queries.GetAccidentsQuery(paramList);
                else if (reviewType.Equals("Recomendations"))
                    sqlQuery = Queries.GetRecomendationsQuery(paramList);
                else if (reviewType.Equals("Plans"))
                    sqlQuery = Queries.GetPlansQuery(paramList);




                var dataReader = _conn.GetQueryResult(sqlQuery);

                if (dataReader == null)
                    return;

                cells.DeleteTableRange(tableName);

                if (reviewType.Equals("All"))
                    ListAllFromDataReader(cells, dataReader, tableName, titleRow);
                else if (reviewType.Equals("Accidents"))
                    ListAccidentesFromDataReader(cells, dataReader, tableName, titleRow);
                else if (reviewType.Equals("Recomendations"))
                    ListRecomendacionesFromDataReader(cells, dataReader, tableName, titleRow);
                else if (reviewType.Equals("Plans"))
                    ListPlanFromDataReader(cells, dataReader, tableName, titleRow);
                

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _conn.CloseConnection();
            }
        }


        private void ListAccidentesFromDataReader(ExcelStyleCells cells, IDataReader dataReader, string tableName, int titleRow)
        {

            //Creo el encabezado de la tabla y doy formato
            cells.GetCell(1, titleRow).Value2 = "Código";
            cells.GetCell(2, titleRow).Value2 = "Fecha";
            cells.GetCell(3, titleRow).Value2 = "Hora";
            cells.GetCell(4, titleRow).Value2 = "Id Responsable";
            cells.GetCell(5, titleRow).Value2 = "Responsable";
            cells.GetCell(6, titleRow).Value2 = "Comentarios";
            cells.GetCell(7, titleRow).Value2 = "Potencial";
            cells.GetCell(8, titleRow).Value2 = "Cód SuperIntendencia";
            cells.GetCell(9, titleRow).Value2 = "SuperIntendencia";

            var columnQty = 9;
            var selectedRange = cells.GetRange(1, titleRow, columnQty, titleRow + 1);

            var tableRange = cells.FormatAsTable(selectedRange, tableName);

            //cargo los datos 
            if (dataReader.IsClosed) return;


            var currentRow = titleRow + 1;
            while (dataReader.Read())
            {
                var item = new Accidente(dataReader);

                cells.GetCell(1, currentRow).Value2 = "'" + item.CodigoAccidente;
                cells.GetCell(2, currentRow).Value2 = "'" + item.FechaCaso;
                cells.GetCell(3, currentRow).Value2 = "'" + item.HoraCaso;
                cells.GetCell(4, currentRow).Value2 = "'" + item.IdResponsable;
                cells.GetCell(5, currentRow).Value2 = "'" + item.NombreResponsable;
                cells.GetCell(6, currentRow).Value2 = "'" + item.Comentarios;
                cells.GetCell(7, currentRow).Value2 = "'" + item.Potencial;
                cells.GetCell(8, currentRow).Value2 = "'" + item.CodigoSuperIntendencia;
                cells.GetCell(9, currentRow).Value2 = "'" + item.SuperIntendencia;

                currentRow++;
            }

            selectedRange = cells.GetRange(1, titleRow, columnQty, currentRow);
            selectedRange.WrapText = false;
        }

        private void ListRecomendacionesFromDataReader(ExcelStyleCells cells, IDataReader dataReader, string tableName, int titleRow)
        {

            //Creo el encabezado de la tabla y doy formato
            cells.GetCell(1, titleRow).Value2 = "Código";
            cells.GetCell(2, titleRow).Value2 = "Nro. Rec.";
            cells.GetCell(3, titleRow).Value2 = "Estado";
            cells.GetCell(4, titleRow).Value2 = "Id Responsable";
            cells.GetCell(5, titleRow).Value2 = "Responsable";
            cells.GetCell(6, titleRow).Value2 = "Fecha Creacion";
            cells.GetCell(7, titleRow).Value2 = "Fecha Plan";
            cells.GetCell(8, titleRow).Value2 = "Fecha Cierre";
            cells.GetCell(9, titleRow).Value2 = "Cód. Departamento";
            cells.GetCell(10, titleRow).Value2 = "Departamento";
            cells.GetCell(11, titleRow).Value2 = "Descripción";

            var columnQty = 11;
            var selectedRange = cells.GetRange(1, titleRow, columnQty, titleRow + 1);

            var tableRange = cells.FormatAsTable(selectedRange, tableName);

            //cargo los datos 
            if (dataReader.IsClosed) return;


            var currentRow = titleRow + 1;
            while (dataReader.Read())
            {
                var item = new Recomendacion(dataReader);

                cells.GetCell(1, currentRow).Value2 = "'" + item.CodigoAccidente;
                cells.GetCell(2, currentRow).Value2 = "'" + item.CodigoRecomendacion;
                cells.GetCell(3, currentRow).Value2 = "'" + item.Estado;
                cells.GetCell(4, currentRow).Value2 = "'" + item.IdResponsable;
                cells.GetCell(5, currentRow).Value2 = "'" + item.NombreResponsable;
                cells.GetCell(6, currentRow).Value2 = "'" + item.FechaCreacion;
                cells.GetCell(7, currentRow).Value2 = "'" + item.FechaPlaneada;
                cells.GetCell(8, currentRow).Value2 = "'" + item.FechaCierre;
                cells.GetCell(9, currentRow).Value2 = "'" + item.CodigoDepartamento;
                cells.GetCell(10, currentRow).Value2 = "'" + item.Departamento;
                cells.GetCell(11, currentRow).Value2 = "'" + item.Descripcion;

                if (!string.IsNullOrWhiteSpace(item.FechaPlaneada))
                {
                    var planDate = MyUtilities.ToDate(item.FechaPlaneada, MyUtilities.DateTime.DateYYYYMMDD);
                    var closedDate = DateTime.Today;
                    if(!string.IsNullOrWhiteSpace(item.FechaCierre))
                        closedDate = MyUtilities.ToDate(item.FechaCierre, MyUtilities.DateTime.DateYYYYMMDD);
                    if (planDate < closedDate)
                        cells.GetCell(8, currentRow).Style = StyleConstants.Error;
                    else
                        cells.GetCell(8, currentRow).Style = StyleConstants.Normal;
                }

                currentRow++;
            }

            selectedRange = cells.GetRange(1, titleRow, columnQty, currentRow);
            selectedRange.WrapText = false;
        }

        private void ListPlanFromDataReader(ExcelStyleCells cells, IDataReader dataReader, string tableName, int titleRow)
        {

            //Creo el encabezado de la tabla y doy formato
            cells.GetCell(1, titleRow).Value2 = "Código";
            cells.GetCell(2, titleRow).Value2 = "Nro. Rec.";
            cells.GetCell(3, titleRow).Value2 = "Nro. Plan.";
            cells.GetCell(4, titleRow).Value2 = "Estado";
            cells.GetCell(5, titleRow).Value2 = "Id Responsable";
            cells.GetCell(6, titleRow).Value2 = "Responsable";
            cells.GetCell(7, titleRow).Value2 = "Fecha Plan";
            cells.GetCell(8, titleRow).Value2 = "% Avance";
            cells.GetCell(9, titleRow).Value2 = "Descripción";
            cells.GetCell(10, titleRow).Value2 = "Avance";

            var columnQty = 10;
            var selectedRange = cells.GetRange(1, titleRow, columnQty, titleRow + 1);

            var tableRange = cells.FormatAsTable(selectedRange, tableName);

            //cargo los datos 
            if (dataReader.IsClosed) return;


            var currentRow = titleRow + 1;
            while (dataReader.Read())
            {
                var item = new PlanAccion(dataReader);

                cells.GetCell(1, currentRow).Value2 = "'" + item.CodigoAccidente;
                cells.GetCell(2, currentRow).Value2 = "'" + item.CodigoRecomendacion;
                cells.GetCell(3, currentRow).Value2 = "'" + item.CodigoPlan;
                cells.GetCell(4, currentRow).Value2 = "'" + item.Estado;
                cells.GetCell(5, currentRow).Value2 = "'" + item.IdResponsable;
                cells.GetCell(6, currentRow).Value2 = "'" + item.NombreResponsable;
                cells.GetCell(7, currentRow).Value2 = "'" + item.FechaPlaneada;
                cells.GetCell(8, currentRow).Value2 = "'" + item.PorcentajeAvance;
                cells.GetCell(9, currentRow).Value2 = "'" + item.Descripcion;
                cells.GetCell(10, currentRow).Value2 = "'" + item.Avance;
                

                if (!string.IsNullOrWhiteSpace(item.FechaPlaneada))
                {
                    var planDate = MyUtilities.ToDate(item.FechaPlaneada, MyUtilities.DateTime.DateYYYYMMDD);
                    var closedDate = DateTime.Today;
                    if (planDate < closedDate && MyUtilities.ToInteger16(item.PorcentajeAvance) < 100)
                        cells.GetCell(7, currentRow).Style = StyleConstants.Error;
                    else
                        cells.GetCell(7, currentRow).Style = StyleConstants.Normal;
                }

                currentRow++;
            }

            selectedRange = cells.GetRange(1, titleRow, columnQty, currentRow);
            selectedRange.WrapText = false;
        }

        private void ListAllFromDataReader(ExcelStyleCells cells, IDataReader dataReader, string tableName, int titleRow)
        {

            //Creo el encabezado de la tabla y doy formato
            cells.GetCell(1, titleRow).Value2 = "Código";
            cells.GetCell(2, titleRow).Value2 = "Nro. Rec.";
            cells.GetCell(3, titleRow).Value2 = "Estado Rec.";
            cells.GetCell(4, titleRow).Value2 = "Nro. Plan.";
            cells.GetCell(5, titleRow).Value2 = "Estado Plan";
            cells.GetCell(6, titleRow).Value2 = "% Plan";
            cells.GetCell(7, titleRow).Value2 = "Id Responsable";
            cells.GetCell(8, titleRow).Value2 = "Responsable";
            cells.GetCell(9, titleRow).Value2 = "Fecha Caso";
            cells.GetCell(10, titleRow).Value2 = "Fecha Creacion";
            cells.GetCell(11, titleRow).Value2 = "Fecha Plan";
            cells.GetCell(12, titleRow).Value2 = "Fecha Cierre";
            cells.GetCell(13, titleRow).Value2 = "Cód. Departamento";
            cells.GetCell(14, titleRow).Value2 = "Departamento";
            cells.GetCell(15, titleRow).Value2 = "Descripción Accidente";
            cells.GetCell(16, titleRow).Value2 = "Descripción Recomendación";
            cells.GetCell(17, titleRow).Value2 = "Descripción Plan";
            cells.GetCell(18, titleRow).Value2 = "Avance Plan";

            var columnQty = 18;
            var selectedRange = cells.GetRange(1, titleRow, columnQty, titleRow + 1);

            var tableRange = cells.FormatAsTable(selectedRange, tableName);

            //cargo los datos 
            if (dataReader.IsClosed) return;


            var currentRow = titleRow + 1;
            while (dataReader.Read())
            {
                var item = new AccRecPlan(dataReader);

                cells.GetCell(1, currentRow).Value2 = "'" + item.CodigoAccidente;
                cells.GetCell(2, currentRow).Value2 = "'" + item.CodigoRecomendacion;
                cells.GetCell(3, currentRow).Value2 = "'" + item.EstadoRecomendacion;
                cells.GetCell(4, currentRow).Value2 = "'" + item.CodigoPlan;
                cells.GetCell(5, currentRow).Value2 = "'" + item.EstadoPlan;
                cells.GetCell(6, currentRow).Value2 = "'" + item.PorcentajeAvancePlan;
                cells.GetCell(7, currentRow).Value2 = "'" + item.IdResponsableRecomendacion;
                cells.GetCell(8, currentRow).Value2 = "'" + item.NombreResponsableRecomendacion;
                cells.GetCell(9, currentRow).Value2 = "'" + item.FechaCasoAccidente;
                cells.GetCell(10, currentRow).Value2 = "'" + item.FechaCreacionRecomendacion;
                cells.GetCell(11, currentRow).Value2 = "'" + item.FechaPlaneadaRecomendacion;
                cells.GetCell(12, currentRow).Value2 = "'" + item.FechaCierreRecomendacion;
                cells.GetCell(13, currentRow).Value2 = "'" + item.CodigoDepartamentoRecomendacion;
                cells.GetCell(14, currentRow).Value2 = "'" + item.DepartamentoRecomendacion;
                cells.GetCell(15, currentRow).Value2 = "'" + item.ComentariosAccidente;
                cells.GetCell(16, currentRow).Value2 = "'" + item.DescripcionRecomendacion;
                cells.GetCell(17, currentRow).Value2 = "'" + item.DescripcionPlan;
                cells.GetCell(18, currentRow).Value2 = "'" + item.AvancePlan;



                if (!string.IsNullOrWhiteSpace(item.FechaPlaneadaRecomendacion))
                {
                    var planDate = MyUtilities.ToDate(item.FechaPlaneadaRecomendacion, MyUtilities.DateTime.DateYYYYMMDD);
                    var closedDate = DateTime.Today;
                    if (!string.IsNullOrWhiteSpace(item.FechaCierreRecomendacion))
                        closedDate = MyUtilities.ToDate(item.FechaCierreRecomendacion, MyUtilities.DateTime.DateYYYYMMDD);
                    if (planDate < closedDate)
                        cells.GetCell(12, currentRow).Style = StyleConstants.Error;
                    else
                        cells.GetCell(12, currentRow).Style = StyleConstants.Normal;
                }

                if (!string.IsNullOrWhiteSpace(item.FechaPlaneadaPlan))
                {
                    var planDate = MyUtilities.ToDate(item.FechaPlaneadaPlan, MyUtilities.DateTime.DateYYYYMMDD);
                    var closedDate = DateTime.Today;
                    if (planDate < closedDate && MyUtilities.ToInteger16(item.PorcentajeAvancePlan) < 100)
                        cells.GetCell(6, currentRow).Style = StyleConstants.Error;
                    else
                        cells.GetCell(6, currentRow).Style = StyleConstants.Normal;

                }

                currentRow++;
            }

            selectedRange = cells.GetRange(1, titleRow, columnQty, currentRow);
            selectedRange.WrapText = false;
        }


        private void ListDefaultFromDataReader(ExcelStyleCells cells, IDataReader dataReader, string tableName, int titleRow)
        {

            //Cargo el encabezado de la tabla y doy formato
            for (var i = 0; i < dataReader.FieldCount; i++)
                cells.GetCell(i + 1, titleRow).Value2 = "'" + dataReader.GetName(i);

            var selectedRange = cells.GetRange(1, titleRow, dataReader.FieldCount, titleRow + 1);

            var tableRange = cells.FormatAsTable(selectedRange, tableName);

            //cargo los datos 
            if (dataReader.IsClosed) return;


            var currentRow = titleRow + 1;
            while (dataReader.Read())
            {
                for (var i = 0; i < dataReader.FieldCount; i++)
                    cells.GetCell(i + 1, currentRow).Value2 = "'" + dataReader[i].ToString().Trim();
                currentRow++;
            }

            selectedRange = cells.GetRange(1, titleRow, dataReader.FieldCount, currentRow);
            selectedRange.WrapText = false;
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _conn?.Dispose();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FormatSheet();
                //si ya hay un thread corriendo que no se ha detenido
                /*
                 * if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(FormatSheet);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
                */
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatSheet()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                StartConnection(drpEnvironment.SelectedItem.Label);

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells?.SetCursorWait();

                var cells = new ExcelStyleCells(_excelApp, true);

                #region CONSTRUYO LA HOJA 1
                var titleRow = TitleRow01;
                var tableName = TableName01;
                var sheetName = SheetName01;
                var validationSheetName = ValidationSheetName;

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;
                cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                cells.GetCell("A1").Value = "CERREJÓN";
                cells.GetCell("A1").Style = cells.GetStyle(StyleConstants.HeaderDefault);
                cells.MergeCells("A1", "B2");

                cells.GetCell("C1").Value = "INVESTIGACIONES ICAM";
                cells.GetCell("C1").Style = cells.GetStyle(StyleConstants.HeaderDefault);
                cells.MergeCells("C1", "J2");

                /*
                cells.GetCell("K1").Value = "OBLIGATORIO";
                cells.GetCell("K1").Style = cells.GetStyle(StyleConstants.TitleRequired);
                cells.GetCell("K2").Value = "OPCIONAL";
                cells.GetCell("K2").Style = cells.GetStyle(StyleConstants.TitleOptional);
                cells.GetCell("K3").Value = "INFORMATIVO";
                cells.GetCell("K3").Style = cells.GetStyle(StyleConstants.TitleInformation);
                cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                cells.GetCell("K4").Style = cells.GetStyle(StyleConstants.TitleAction);
                cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                cells.GetCell("K5").Style = cells.GetStyle(StyleConstants.TitleAdditional);
                */

                var paramList = new List<string>();
                var paramDateList = new List<string>();
                var paramPotencialList = new List<string> { "BAJO", "MEDIO", "ALTO", "ICAM" };
                var paramEstados = new List<string> { "1 - ABIERTA", "2 - EN PROCESO", "3 - CERRADA", "4 - CANCELADA" };
                foreach (Enum i in Enum.GetValues(typeof(IxSearchParameters)))
                {
                    var paramDescription = SearchParameters.GetDescription(i);
                    if (paramDescription.ToUpper().Contains("FECHA"))
                        paramDateList.Add(paramDescription);
                    else
                        paramList.Add(paramDescription);
                }


                _cells.SetValidationList(_cells.GetRange("A3", "A6"), paramList, validationSheetName, 1, false);
                cells.GetCell("A3").Value = SearchParameters.GetDescription(IxSearchParameters.CodigoAccidente);
                cells.GetCell("A4").Value = SearchParameters.GetDescription(IxSearchParameters.IdResponsableRecomendacion);
                cells.GetCell("A5").Value = SearchParameters.GetDescription(IxSearchParameters.PotencialAccidente);
                cells.GetCell("A6").Value = SearchParameters.GetDescription(IxSearchParameters.CodigoEstadoRecomendacion);
                cells.GetCell("B5").Value = "ICAM";
                _cells.SetValidationList(_cells.GetRange("B6"), paramEstados, validationSheetName, 3, false);
                cells.GetCell("B6").Value = paramEstados[0];
                cells.GetRange("A3", "A6").Style = cells.GetStyle(StyleConstants.Option);
                cells.GetRange("B3", "B6").Style = cells.GetStyle(StyleConstants.Select);

                _cells.SetValidationList(_cells.GetRange("C3", "C4"), paramDateList, validationSheetName, 2, false);
                cells.GetCell("C3").Value = SearchParameters.GetDescription(IxSearchParameters.FechaInicialAccidente);
                cells.GetCell("C4").Value = SearchParameters.GetDescription(IxSearchParameters.FechaFinalAccidente);
                cells.GetRange("C3", "C4").Style = cells.GetStyle(StyleConstants.Option);
                cells.GetRange("D3", "D4").Style = cells.GetStyle(StyleConstants.Select);

                cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                cells.GetCell("D3").AddComment("YYYYMMDD");
                cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                cells.GetCell("D4").AddComment("YYYYMMDD");


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }


    }
}
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using SharedClassLibrary.Cority;
using SharedClassLibrary.Cority.MGIPService;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace TamizajeExcelAddIn
{
    public partial class RibbonTamizaje
    {
        private ExcelStyleCells _cells;
        private Application _excelApp;

        public int TitleRowQrh = 1;
        public int ResultColumnQrh = 18;
        public string SheetNameQrh = "Cuestionarios";
        public string TableNameQrh = "QRH_Loader_Table";

        public string ValidationSheetName = "Validaciones";

        private Thread _thread;

        private void RibbonTamizaje_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                var environments = new List<string>()
                    {"Productivo", "Test"};

                foreach (var env in environments)
                {
                    var item = Factory.CreateRibbonDropDownItem();
                    item.Label = env;
                    drpEnvironment.Items.Add(item);
                }

                Debugger.LocalDataPath = "c:\\cority\\";
                Debugger.DebugginMode = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatLoadQuestionary();
        }

        public void FormatLoadQuestionary()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                #region CONSTRUYO LA HOJA DE CARGUE QRH
                var titleRow = TitleRowQrh;
                var resultColumn = ResultColumnQrh;
                var tableName = TableNameQrh;
                var sheetName = SheetNameQrh;

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
                
                //GENERAL
                _cells.GetCell(1, titleRow).Value = "Código de Barras";
                _cells.GetCell(2, titleRow).Value = "Nombre 1";
                _cells.GetCell(3, titleRow).Value = "Nombre 2";
                _cells.GetCell(4, titleRow).Value = "Apellido 1";
                _cells.GetCell(5, titleRow).Value = "Apellido 2";
                _cells.GetCell(6, titleRow).Value = "Tipo de Documento";
                _cells.GetCell(7, titleRow).Value = "Número de Identificación";
                _cells.GetCell(8, titleRow).Value = "Sexo";
                _cells.GetCell(9, titleRow).Value = "Teléfono";
                _cells.GetCell(10, titleRow).Value = "Ciudad";
                _cells.GetCell(11, titleRow).Value = "Fecha Solicitud";
                _cells.GetCell(12, titleRow).Value = "Fecha Validación";
                _cells.GetCell(13, titleRow).Value = "Laboratorio";
                _cells.GetCell(14, titleRow).Value = "Área";
                _cells.GetCell(15, titleRow).Value = "Resultado";
                _cells.GetCell(16, titleRow).Value = "QRH";
                _cells.GetCell(17, titleRow).Value = "A revisar por";
                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleRequired;

                _cells.GetCell(resultColumn, titleRow).Value = "MENSAJE";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion
                

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("FormatLoadQuestionary()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void btnLoadQuestionary_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameQrh)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(LoadQuestionaryList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void LoadQuestionaryList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var cells = cbAllowBackgroundWork.Checked ? new ExcelStyleCells(_excelApp, SheetNameQrh) : _cells;
            cells.ClearTableRangeColumn(TableNameQrh, ResultColumnQrh);

            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var titleRow = TitleRowQrh;
            var resultColumn = ResultColumnQrh;
            var tableName = TableNameQrh;
            var sheetName = SheetNameQrh;

            var service = new MGIPService();
            service.Url = urlService;
            Authentication.Authenticate("hector.hernandez.ext", "Hha6jhr8");
            var i = titleRow + 1;
            var idColumn = 7;
            try
            {
                while (!string.IsNullOrEmpty("" + cells.GetCell(idColumn, i).Value))
                {
                    try
                    {
                        var quest = new CovidQuestionary();
                        quest.CodigoBarras.Value = "" + cells.GetCell(1, i).Value;
                        quest.FirstName = "" + cells.GetCell(2, i).Value;
                        quest.SecondName = "" + cells.GetCell(3, i).Value;
                        quest.FirstLastName = "" + cells.GetCell(4, i).Value;
                        quest.SecondLastName = "" + cells.GetCell(5, i).Value;
                        quest.DocumentType = "" + cells.GetCell(6, i).Value;
                        quest.EmployeeId = "" + cells.GetCell(7, i).Value;
                        quest.Sex = "" + cells.GetCell(8, i).Value;
                        quest.Telefono.Value = "" + cells.GetCell(9, i).Value;
                        quest.Ciudad.Value = "" + cells.GetCell(10, i).Value;
                        quest.FechaRespuesta.Value = "" + cells.GetCell(11, i).Value;
                        quest.FechaResultado.Value = "" + cells.GetCell(12, i).Value;
                        quest.Laboratorio.Value = "" + cells.GetCell(13, i).Value;
                        quest.Area.Value = "" + cells.GetCell(14, i).Value;
                        quest.EstadoPrueba.Value = "" + cells.GetCell(15, i).Value;
                        
                        quest.Header.Qrh = "" + cells.GetCell(16, i).Value;
                        quest.Header.ToBeReviewed = "" + cells.GetCell(17, i).Value;
                        quest.Header.EmployeeId = "" + quest.EmployeeId;
                        quest.Header.QuestionaryCode = "" + quest.QuestionaryCode;
                        quest.Header.DateOfResponse = "" + quest.FechaRespuesta.Value;

                        //
                        if (quest.EstadoPrueba.Value.ToUpperInvariant().Contains("POSITIV"))
                        {
                            if(quest.EstadoPrueba.Value.ToUpperInvariant().Contains("ANT"))//antígeno
                                quest.EstadoPrueba.Value = "6";
                            else
                                quest.EstadoPrueba.Value = "1";
                            quest.Conducta.Value = "2";
                            quest.TipoCaso.Value = "1";
                        }
                        else if (quest.EstadoPrueba.Value.ToUpperInvariant().Contains("REPROCESAD"))
                        {
                            quest.EstadoPrueba.Value = "8";
                            quest.Conducta.Value = "5";
                            quest.TipoCaso.Value = "4";
                        }
                        else if(quest.EstadoPrueba.Value.ToUpperInvariant().Contains("NEGATIV"))
                        {
                            quest.EstadoPrueba.Value = "2";
                            quest.Conducta.Value = "6";
                            quest.TipoCaso.Value = "2";
                        }

                        quest.Site.Value = null;
                        quest.Ubicacion.Value = null;
                        quest.Evolucion.Value = null;
                        quest.Severidad.Value = null;
                        quest.FuenteCaso.Value = "5"; //Reincorporación a la operación
                        quest.EtapaPrueba.Value = "8";//Tamizaje
                        //

                        cells.GetCell(ResultColumnQrh, i).Value  = CovidActions.CreateCovidQuestionary(service, quest);

                    }
                    catch (Exception ex)
                    {
                        cells.GetCell(idColumn, i).Style = StyleConstants.Error;
                        cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ReReviewWODetailedList()", ex.Message);
                    }
                    finally
                    {
                        if (!cbAllowBackgroundWork.Checked)
                            cells.GetCell(2, i).Select();
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("LoadQuestionaryList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                cells?.ActiveSheet.Cells.Columns.AutoFit();
                cells?.SetCursorDefault();
            }


        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
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

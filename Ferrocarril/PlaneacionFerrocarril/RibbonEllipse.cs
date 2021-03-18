using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Settings = SharedClassLibrary.Ellipse.Settings;
using System.Text;
using EllipseJobsClassLibrary;
using EllipseJobsClassLibrary.WorkOrderTaskMWPService;
using EllipseStandardJobsClassLibrary;
using Microsoft.Office.Tools.Excel;
using PlaneacionFerrocarril.TemperatureWagon;

namespace PlaneacionFerrocarril
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;

        private const string SheetNameWkP = "Programación";
        private const int TitleRowWkP = 22;
        private const int ResourceRowWkP = 7;
        private const int ResultColumnWkP = 11;
        private const string TableNameWkP = "WeekPlanning";

        private const string SheetNameFcEq = "Forecast";
        private const int TitleRowFcEq = 13;
        private const int ResultColumnFcEq = 11;
        private const string TableNameFcEq = "ForecastTable";

        private const string SheetNameFcTk = "ForecastTask";
        private const int TitleRowFcTk = 7;
        private const int ResultColumnFcTk = 27;
        private const string TableNameFcTk = "ForecastTaskTable";

        private const string SheetNameFcRq = "ForecastRequirements";
        private const int TitleRowFcRq = 7;
        private const int ResultColumnFcRq = 33;
        private const string TableNameFcRq = "ForecastRequirementsTable";

        private const string SheetNameMse345 = "MSE345";
        private const int TitleRowMse345 = 13;
        private const int ResultColumnMse345 = 11;
        private const string TableNameMse345 = "Mse345Table";

        private const string SheetNamePlain = "Plano";
        private const int TitleRowPlain = 1;
        private const string TableNamePlain = "PlainTable";


        private const string SheetNamePlanHist = "BD";
        private const int TitleRowPlanHist = 1;
        private const int ResultColumnPlanHist = 7;
        private const string TableNamePlanHist = "PlanHistoryTable";
        private string SelectedStartDatePlanHist = "";
        private const string ValidationSheetName = "ValidationSheet";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        #region General
        private void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }

            settings.SetDefaultCustomSettingValue(TempWagonConstants.IgnoreLocomotives, "true");
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
            //TempWagon
            var ignoreLocomotives = MyUtilities.IsTrue(settings.GetCustomSettingValue(TempWagonConstants.IgnoreLocomotives));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            cbTempWagIgnoreLocomotives.Checked = ignoreLocomotives;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatWeeklyPlanning();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(@"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            const string developerName1 = "Héctor Hernandez <hernandezrhectorj@gmail.com>";
            const string developerName2 = "Hugo Mendoza <huancone@gmail.com>";

            new AboutBoxExcelAddIn(developerName1, developerName2).ShowDialog();
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }

        

        

        #endregion

        #region Forecast
        private void btnForecastFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatForecast();
        }

        private void btnReviewForecastTask_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameFcTk)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewForecastTasks);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnForecastReviewRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameFcRq)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewForecastRequirements);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
    
        private void FormatForecast()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                var sheetName = SheetNameFcEq;
                var titleRow = TitleRowFcEq;
                var resultColumn = ResultColumnFcEq;
                var tableName = TableNameFcEq;

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;



                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "EQUIPMENT FORECAST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");
                _cells.MergeCells("C6", "L11");
                _cells.MergeCells("A12", "L12");

                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("A6", "A11").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B6", "B11").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleInformation;


                _cells.GetCell(1, titleRow).Value = "WORKGROUP";
                _cells.GetCell(2, titleRow).Value = "EQUIPO";
                _cells.GetCell(3, titleRow).Value = "EQ DESC";
                _cells.GetCell(4, titleRow).Value = "COMP CODE";
                _cells.GetCell(5, titleRow).Value = "COMP MOD CODE";
                _cells.GetCell(6, titleRow).Value = "WO/FORECAST";
                _cells.GetCell(7, titleRow).Value = "MST";
                _cells.GetCell(8, titleRow).Value = "STD JOB NO";
                _cells.GetCell(9, titleRow).Value = "WO DESC";
                _cells.GetCell(10, titleRow).Value = "PLAN STR DATE";


                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                #region Hoja 2

                sheetName = SheetNameFcTk;
                titleRow = TitleRowFcTk;
                resultColumn = ResultColumnFcTk;
                tableName = TableNameFcTk;

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;



                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "TASK FORECAST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");


                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleInformation;


                _cells.GetCell(1, titleRow).Value = "WORKGROUP";
                _cells.GetCell(2, titleRow).Value = "EQUIPO";
                _cells.GetCell(3, titleRow).Value = "EQ DESC";
                _cells.GetCell(4, titleRow).Value = "COMP CODE";
                _cells.GetCell(5, titleRow).Value = "COMP MOD CODE";
                _cells.GetCell(6, titleRow).Value = "WO/FORECAST";
                _cells.GetCell(7, titleRow).Value = "MST";
                _cells.GetCell(8, titleRow).Value = "STD JOB NO";
                _cells.GetCell(9, titleRow).Value = "WO DESC";
                _cells.GetCell(10, titleRow).Value = "PLAN STR DATE";

                _cells.GetCell(11, titleRow).Value = "TASK NO";
                _cells.GetCell(12, titleRow).Value = "TASK DESC";
                _cells.GetCell(13, titleRow).Value = "JOB DESC CODE";
                _cells.GetCell(14, titleRow).Value = "SAFETY INSTR";
                _cells.GetCell(15, titleRow).Value = "COMPLETE INSTR";
                _cells.GetCell(16, titleRow).Value = "COMP TEXT CODE";
                _cells.GetCell(17, titleRow).Value = "ASSIGN PERSON";
                _cells.GetCell(18, titleRow).Value = "EST MACH HOURS";
                _cells.GetCell(19, titleRow).Value = "EST DUR HRS";
                _cells.GetCell(20, titleRow).Value = "NO LABOR";
                _cells.GetCell(21, titleRow).Value = "NO MATERIAL";
                _cells.GetCell(22, titleRow).Value = "APL EGI";
                _cells.GetCell(23, titleRow).Value = "APL TYPE";
                _cells.GetCell(24, titleRow).Value = "APL COMP CODE";
                _cells.GetCell(25, titleRow).Value = "APL COMP MOD CODE";
                _cells.GetCell(26, titleRow).Value = "APL SEQ NO";

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region Hoja 3

                sheetName = SheetNameFcRq;
                titleRow = TitleRowFcRq;
                resultColumn = ResultColumnFcRq;
                tableName = TableNameFcRq;

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;



                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "REQUIREMENTS FORECAST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");


                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleInformation;


                _cells.GetCell(1, titleRow).Value = "WORKGROUP";
                _cells.GetCell(2, titleRow).Value = "EQUIPO";
                _cells.GetCell(3, titleRow).Value = "EQ DESC";
                _cells.GetCell(4, titleRow).Value = "COMP CODE";
                _cells.GetCell(5, titleRow).Value = "COMP MOD CODE";
                _cells.GetCell(6, titleRow).Value = "WO/FORECAST";
                _cells.GetCell(7, titleRow).Value = "MST";
                _cells.GetCell(8, titleRow).Value = "STD JOB NO";
                _cells.GetCell(9, titleRow).Value = "WO DESC";
                _cells.GetCell(10, titleRow).Value = "PLAN STR DATE";

                _cells.GetCell(11, titleRow).Value = "TASK NO";
                _cells.GetCell(12, titleRow).Value = "TASK DESC";
                _cells.GetCell(13, titleRow).Value = "JOB DESC CODE";
                _cells.GetCell(14, titleRow).Value = "SAFETY INSTR";
                _cells.GetCell(15, titleRow).Value = "COMPLETE INSTR";
                _cells.GetCell(16, titleRow).Value = "COMP TEXT CODE";
                _cells.GetCell(17, titleRow).Value = "ASSIGN PERSON";
                _cells.GetCell(18, titleRow).Value = "EST MACH HOURS";
                _cells.GetCell(19, titleRow).Value = "EST DUR HRS";
                _cells.GetCell(20, titleRow).Value = "NO LABOR";
                _cells.GetCell(21, titleRow).Value = "NO MATERIAL";
                _cells.GetCell(22, titleRow).Value = "APL EGI";
                _cells.GetCell(23, titleRow).Value = "APL TYPE";
                _cells.GetCell(24, titleRow).Value = "APL COMP CODE";
                _cells.GetCell(25, titleRow).Value = "APL COMP MOD CODE";
                _cells.GetCell(26, titleRow).Value = "APL SEQ NO";
                _cells.GetCell(27, titleRow).Value = "REQ TYPE";
                _cells.GetCell(28, titleRow).Value = "REQ SEQ NO";
                _cells.GetCell(29, titleRow).Value = "REQ CODE";
                _cells.GetCell(30, titleRow).Value = "REQ DESC";
                _cells.GetCell(31, titleRow).Value = "QTY REQ";
                _cells.GetCell(32, titleRow).Value = "HRS REQ";
                
                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatEquipmentForecast()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                throw new Exception(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ReviewForecastTasks()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();

            const int titleRow01 = TitleRowFcEq;
            const int titleRow02 = TitleRowFcTk;
            const string sheetName01 = SheetNameFcEq;
            const string sheetName02 = SheetNameFcTk;
            const string tableName01 = TableNameFcEq;
            const string tableName02 = TableNameFcTk;
            const int resultColumn01 = ResultColumnFcEq;
            const int resultColumn02 = ResultColumnFcTk;

            var stdCells = new ExcelStyleCells(_excelApp, sheetName01);
            var taskCells = new ExcelStyleCells(_excelApp, sheetName02);

            taskCells.ClearTableRange(tableName02);
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);





            var j = titleRow01 + 1;//itera según cada estándar
            var i = titleRow02 + 1;//itera la celda para cada tarea

            while (!string.IsNullOrEmpty("" + stdCells.GetCell(6, j).Value))
            {
                try
                {
                    var stdJob = new StandardJob();
                    stdJob.DistrictCode = "ICOR";
                    stdJob.WorkGroup = _cells.GetEmptyIfNull(stdCells.GetCell(1, j).Value2);
                    var equipmentNo = _cells.GetEmptyIfNull(stdCells.GetCell(2, j).Value2);
                    var equipmentDesc = _cells.GetEmptyIfNull(stdCells.GetCell(3, j).Value2);
                    stdJob.CompCode = _cells.GetEmptyIfNull(stdCells.GetCell(4, j).Value2);
                    stdJob.CompModCode = _cells.GetEmptyIfNull(stdCells.GetCell(5, j).Value2);
                    var woForecastId = _cells.GetEmptyIfNull(stdCells.GetCell(6, j).Value2);
                    stdJob.NoMsts = _cells.GetEmptyIfNull(stdCells.GetCell(7, j).Value2);
                    stdJob.StandardJobNo = _cells.GetEmptyIfNull(stdCells.GetCell(8, j).Value2);
                    stdJob.StandardJobDescription = _cells.GetEmptyIfNull(stdCells.GetCell(9, j).Value2);
                    var planStartDate = _cells.GetEmptyIfNull(stdCells.GetCell(10, j).Value2);

                    var taskList = StandardJobActions.FetchStandardJobTask(_eFunctions, stdJob.DistrictCode, stdJob.WorkGroup, stdJob.StandardJobNo);


                    foreach (var task in taskList)
                    {
                        //Para resetear el estilo
                        taskCells.GetRange(1, i, resultColumn02, i).Style = StyleConstants.Normal;
                        //GENERAL
                        taskCells.GetCell(1, i).Value = "" + stdJob.WorkGroup;
                        taskCells.GetCell(2, i).Value = "'" + equipmentNo;
                        taskCells.GetCell(3, i).Value = "'" + equipmentDesc;
                        taskCells.GetCell(4, i).Value = "'" + stdJob.CompCode;
                        taskCells.GetCell(5, i).Value = "'" + stdJob.CompModCode;
                        taskCells.GetCell(6, i).Value = "'" + woForecastId;
                        taskCells.GetCell(7, i).Value = "'" + stdJob.NoMsts;
                        taskCells.GetCell(8, i).Value = "'" + task.StandardJob;
                        taskCells.GetCell(9, i).Value = "" + task.StandardJobDescription;
                        taskCells.GetCell(10, i).Value = "" + planStartDate;
                        //GENERAL
                        taskCells.GetCell(11, i).Value = "'" + task.SjTaskNo;
                        taskCells.GetCell(12, i).Value = "" + task.SjTaskDesc;
                        taskCells.GetCell(13, i).Value = "'" + task.JobDescCode;
                        taskCells.GetCell(14, i).Value = "'" + task.SafetyInstr;
                        taskCells.GetCell(15, i).Value = "'" + task.CompleteInstr;
                        taskCells.GetCell(16, i).Value = "'" + task.ComplTextCode;

                        taskCells.GetCell(17, i).Value = "" + task.AssignPerson;
                        taskCells.GetCell(18, i).Value = "'" + task.EstimatedMachHrs;
                        taskCells.GetCell(19, i).Value = "" + task.EstimatedDurationsHrs;
                        taskCells.GetCell(20, i).Value = "" + task.NoLabor;
                        taskCells.GetCell(21, i).Value = "" + task.NoMaterial;
                        //APL
                        taskCells.GetCell(22, i).Value = "'" + task.AplEquipmentGrpId;
                        taskCells.GetCell(23, i).Value = "'" + task.AplType;
                        taskCells.GetCell(24, i).Value = "'" + task.AplCompCode;
                        taskCells.GetCell(25, i).Value = "'" + task.AplCompModCode;
                        taskCells.GetCell(26, i).Value = "'" + task.AplSeqNo;


                        i++;//aumenta tarea
                    }
                }
                catch (Exception ex)
                {
                    taskCells.GetCell(1, i).Style = StyleConstants.Error;
                    taskCells.GetCell(1, i).Value = _cells.GetEmptyIfNull(stdCells.GetCell(1, j).Value2);
                    taskCells.GetCell(2, i).Value = _cells.GetEmptyIfNull(stdCells.GetCell(2, j).Value2);
                    taskCells.GetCell(3, i).Value = _cells.GetEmptyIfNull(stdCells.GetCell(3, j).Value2);
                    taskCells.GetCell(resultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewForecastTasks()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta std
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
            _eFunctions.CloseConnection();
        }

        private void ReviewForecastRequirements()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            const string sheetName02 = SheetNameFcTk;
            const string sheetName03 = SheetNameFcRq;
            const string tableName02 = TableNameFcTk;
            const string tableName03 = TableNameFcRq;
            const int resultColumn02 = ResultColumnFcTk;
            const int resultColumn03 = ResultColumnFcRq;
            const int titleRow02 = TitleRowFcTk;
            const int titleRow03 = TitleRowFcRq;

            var taskCells = new ExcelStyleCells(_excelApp, sheetName02);
            var reqCells = new ExcelStyleCells(_excelApp, sheetName03);

            reqCells.ClearTableRange(tableName03);
            var j = titleRow02 + 1;//itera según cada tarea
            var i = titleRow03 + 1;//itera la celda para cada requerimiento

            while (!string.IsNullOrEmpty("" + taskCells.GetCell(6, j).Value) && !string.IsNullOrEmpty("" + taskCells.GetCell(11, j).Value))
            {
                try
                {
                    var task = new StandardJobTask();
                    var stdJob = new StandardJob();
                    stdJob.DistrictCode = "ICOR";
                    stdJob.WorkGroup = _cells.GetEmptyIfNull(taskCells.GetCell(1, j).Value2);
                    var equipmentNo = _cells.GetEmptyIfNull(taskCells.GetCell(2, j).Value2);
                    var equipmentDesc = _cells.GetEmptyIfNull(taskCells.GetCell(3, j).Value2);
                    stdJob.CompCode = _cells.GetEmptyIfNull(taskCells.GetCell(4, j).Value2);
                    stdJob.CompModCode = _cells.GetEmptyIfNull(taskCells.GetCell(5, j).Value2);
                    var woForecastId = _cells.GetEmptyIfNull(taskCells.GetCell(6, j).Value2);
                    stdJob.NoMsts = _cells.GetEmptyIfNull(taskCells.GetCell(7, j).Value2);
                    stdJob.StandardJobNo = _cells.GetEmptyIfNull(taskCells.GetCell(8, j).Value2);
                    stdJob.StandardJobDescription = _cells.GetEmptyIfNull(taskCells.GetCell(9, j).Value2);
                    var planStartDate = _cells.GetEmptyIfNull(taskCells.GetCell(10, j).Value2);

                    task.SjTaskNo = _cells.GetEmptyIfNull(taskCells.GetCell(11, j).Value2);
                    task.SjTaskDesc = _cells.GetEmptyIfNull(taskCells.GetCell(12, j).Value2);
                    task.JobDescCode = _cells.GetEmptyIfNull(taskCells.GetCell(13, j).Value2);
                    task.SafetyInstr = _cells.GetEmptyIfNull(taskCells.GetCell(14, j).Value2);
                    task.CompleteInstr = _cells.GetEmptyIfNull(taskCells.GetCell(15, j).Value2);
                    task.ComplTextCode = _cells.GetEmptyIfNull(taskCells.GetCell(16, j).Value2);
                    task.AssignPerson = _cells.GetEmptyIfNull(taskCells.GetCell(17, j).Value2);
                    task.EstimatedMachHrs = _cells.GetEmptyIfNull(taskCells.GetCell(18, j).Value2);
                    task.EstimatedDurationsHrs = _cells.GetEmptyIfNull(taskCells.GetCell(19, j).Value2);
                    task.NoLabor = _cells.GetEmptyIfNull(taskCells.GetCell(20, j).Value2);
                    task.NoMaterial = _cells.GetEmptyIfNull(taskCells.GetCell(21, j).Value2);
                    task.AplEquipmentGrpId = _cells.GetEmptyIfNull(taskCells.GetCell(22, j).Value2);
                    task.AplType = _cells.GetEmptyIfNull(taskCells.GetCell(23, j).Value2);
                    task.AplCompCode = _cells.GetEmptyIfNull(taskCells.GetCell(24, j).Value2);
                    task.AplCompModCode = _cells.GetEmptyIfNull(taskCells.GetCell(25, j).Value2);
                    task.AplSeqNo = _cells.GetEmptyIfNull(taskCells.GetCell(26, j).Value2);


                    var reqList = StandardJobActions.FetchTaskRequirements(_eFunctions, stdJob.DistrictCode, stdJob.WorkGroup, stdJob.StandardJobNo, task.SjTaskNo);

                    foreach (var req in reqList)
                    {
                        //Para resetear el estilo
                        reqCells.GetRange(1, i, resultColumn02, i).Style = StyleConstants.Normal;
                        //GENERAL
                        reqCells.GetCell(1, i).Value = "" + task.WorkGroup;
                        reqCells.GetCell(2, i).Value = "'" + equipmentNo;
                        reqCells.GetCell(3, i).Value = "'" + equipmentDesc;
                        reqCells.GetCell(4, i).Value = "'" + stdJob.CompCode;
                        reqCells.GetCell(5, i).Value = "'" + stdJob.CompModCode;
                        reqCells.GetCell(6, i).Value = "'" + woForecastId;
                        reqCells.GetCell(7, i).Value = "'" + stdJob.NoMsts;
                        reqCells.GetCell(8, i).Value = "'" + task.StandardJob;
                        reqCells.GetCell(9, i).Value = "" + task.StandardJobDescription;
                        reqCells.GetCell(10, i).Value = "" + planStartDate;
                        //GENERAL
                        reqCells.GetCell(11, i).Value = "'" + task.SjTaskNo;
                        reqCells.GetCell(12, i).Value = "" + task.SjTaskDesc;
                        reqCells.GetCell(13, i).Value = "'" + task.JobDescCode;
                        reqCells.GetCell(14, i).Value = "'" + task.SafetyInstr;
                        reqCells.GetCell(15, i).Value = "'" + task.CompleteInstr;
                        reqCells.GetCell(16, i).Value = "'" + task.ComplTextCode;

                        reqCells.GetCell(17, i).Value = "" + task.AssignPerson;
                        reqCells.GetCell(18, i).Value = "'" + task.EstimatedMachHrs;
                        reqCells.GetCell(19, i).Value = "" + task.EstimatedDurationsHrs;
                        reqCells.GetCell(20, i).Value = "" + task.NoLabor;
                        reqCells.GetCell(21, i).Value = "" + task.NoMaterial;
                        //APL
                        reqCells.GetCell(22, i).Value = "'" + task.AplEquipmentGrpId;
                        reqCells.GetCell(23, i).Value = "'" + task.AplType;
                        reqCells.GetCell(24, i).Value = "'" + task.AplCompCode;
                        reqCells.GetCell(25, i).Value = "'" + task.AplCompModCode;
                        reqCells.GetCell(26, i).Value = "'" + task.AplSeqNo;

                        reqCells.GetCell(27, i).Value = "" + req.ReqType;
                        reqCells.GetCell(28, i).Value = "" + req.SeqNo;
                        reqCells.GetCell(29, i).Value = "" + req.ReqCode;
                        reqCells.GetCell(30, i).Value = "" + req.ReqDesc;
                        reqCells.GetCell(31, i).Value = "" + req.QtyReq;
                        reqCells.GetCell(32, i).Value = "" + req.HrsReq;
                        i++;//aumenta req
                    }
                }
                catch (Exception ex)
                {
                    reqCells.GetCell(1, i).Style = StyleConstants.Error;
                    reqCells.GetCell(1, i).Value = _cells.GetEmptyIfNull(taskCells.GetCell(1, j).Value2);
                    reqCells.GetCell(2, i).Value = _cells.GetEmptyIfNull(taskCells.GetCell(2, j).Value2);
                    reqCells.GetCell(3, i).Value = _cells.GetEmptyIfNull(taskCells.GetCell(3, j).Value2);
                    reqCells.GetCell(4, i).Value = _cells.GetEmptyIfNull(taskCells.GetCell(6, j).Value2);
                    reqCells.GetCell(resultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewForecastRequirements()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta task
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }

        #endregion

        #region  WeekPlanning
        private void FormatWeeklyPlanning()
        {
            try
            {
                _excelApp.Workbooks.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.SetCursorWait();

                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                #region CONSTRUYO LA HOJA 1
                var titleRow = TitleRowWkP;
                var resultColumn = ResultColumnWkP;
                var tableName = TableNameWkP;
                var sheetName = SheetNameWkP;

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "PROGRAMACIÓN SEMANAL DE TRABAJO";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();

                _cells.GetCell("A3").Value = "GRUPO";
                _cells.SetValidationList(_cells.GetCell("B3"), workGroupList, ValidationSheetName, 2, false);
                _cells.GetCell("A4").Value = "ADICIONAL";
                _cells.GetCell("B4").Value = "BACKLOG";
                _cells.GetCell("A5").Value = "BUSQUEDA";

                var searchType = new List<string>
                {
                    SearchType.WorkOrderOnly,
                    //SearchType.MstForecastOnly,
                    SearchType.WorkOrderAndMstForecast
                };
                _cells.SetValidationList(_cells.GetCell("B5"), searchType, ValidationSheetName, 3);
                _cells.GetCell("B5").Value = SearchType.WorkOrderOnly;

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                var fromDate = DateTime.Today;
                var toDate = DateTime.Today.AddDays(7);
                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("D3").Value = string.Format("{0:0000}", fromDate.Year) + string.Format("{0:00}", fromDate.Month) + string.Format("{0:00}", fromDate.Day + 1);
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", toDate.Year) + string.Format("{0:00}", toDate.Month) + string.Format("{0:00}", toDate.Day);
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);




                //Task Table
                _cells.GetCell(1, titleRow).Value = "EQUIPO";
                _cells.GetCell(2, titleRow).Value = "OT/STAD";
                _cells.GetCell(3, titleRow).Value = "TAREA";
                _cells.GetCell(4, titleRow).Value = "DESCRIPCIÓN";
                _cells.GetCell(5, titleRow).Value = "ESTADO";
                _cells.GetCell(6, titleRow).Value = "FECHA PLAN";
                _cells.GetCell(7, titleRow).Value = "RECURSO";
                _cells.GetCell(8, titleRow).Value = "HH ACTUAL";
                _cells.GetCell(9, titleRow).Value = "HH ESTIMADAS";
                _cells.GetCell(10, titleRow).Value = "PENDIENTE (EST-ACT)";

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;

                _cells.GetRange(1, titleRow + 1, resultColumn - 1, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn - 1, titleRow + 1), tableName);
                //

                //Resource Table
                var resourceRow = ResourceRowWkP;

                _cells.GetCell(7, resourceRow).Value = "HORAS HOMBRE";
                _cells.GetCell(7, resourceRow).Style = StyleConstants.TitleRequired;
                _cells.MergeCells(7, resourceRow, 9, resourceRow);

                _cells.GetCell(7, resourceRow + 1).Value = "Tipo Recurso";
                _cells.GetCell(8, resourceRow + 1).Value = "Disponible";
                _cells.GetCell(9, resourceRow + 1).Value = "Programadas";
                _cells.GetRange(7, resourceRow + 1, 9, resourceRow + 1).Style = StyleConstants.TitleOptional;

                _cells.GetCell(7, resourceRow + 12).Value = "TOTAL";
                _cells.GetCell(7, resourceRow + 12).Style = StyleConstants.Option;
                _cells.GetCell(8, resourceRow + 12).Formula = "=SUM(H9:H18)";
                _cells.GetCell(9, resourceRow + 12).Formula = "=SUM(I9:I18)";
                _cells.GetCell(8, resourceRow + 12).Style = StyleConstants.Select;
                _cells.GetCell(9, resourceRow + 12).Style = StyleConstants.Select;


                for (var i = resourceRow + 2; i < resourceRow + 12; i++)
                {
                    _cells.GetCell(7, i).Style = StyleConstants.Option;
                    _cells.GetCell(8, i).Style = StyleConstants.Select;
                    _cells.GetCell(9, i).Style = StyleConstants.Select;
                    _cells.GetCell(9, i).Formula = "=SUMIF(WeekPlanning[RECURSO],G" + i + ",WeekPlanning[PENDIENTE (EST-ACT)])";

                }

                //Chart
                // Add chart
                var charts = _cells.ActiveSheet.ChartObjects() as Excel.ChartObjects;
                var chartRange = _cells.GetRange(1, resourceRow, 5, resourceRow + 12);
                if (charts != null)
                {
                    var chartObject = charts.Add(chartRange.Left, chartRange.Top, chartRange.Width, chartRange.Height);

                    var chart = chartObject.Chart;

                    // Set chart range and data
                    chart.SetSourceData(_cells.GetRange(8, resourceRow + 1, 9, resourceRow + 11));
                    var seriesRange = _cells.GetRange(7, resourceRow + 2, 7, resourceRow + 11);
                    var xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
                    xAxis.CategoryNames = seriesRange;
                }

                //Total Availability
                _cells.GetCell(10, 5).Value = "HH DISP/HH PRO";
                _cells.GetCell(10, 5).Style = StyleConstants.TitleOptional;

                var totalAvRange = _cells.GetCell(8, resourceRow + 12);
                var totalPrRange = _cells.GetCell(9, resourceRow + 12);
                _cells.GetCell(10, 6).Formula = "=" + totalPrRange.Address + "/" + totalAvRange.Address;
                _cells.GetCell(10, 6).Style = StyleConstants.Select;
                _cells.GetCell(10, 6).NumberFormat = NumberFormatConstants.Percentage;

                Excel.FormatCondition cond1 = _cells.GetCell(10, 6).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween, "=0.9", "=1.1");
                cond1.Font.Bold = true;
                cond1.Interior.Color = 13434828;
                Excel.FormatCondition cond2 = _cells.GetCell(10, 6).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0.9", "=1.1");
                cond2.Font.Bold = true;
                cond2.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                cond2.Interior.Color = 192;
                Excel.FormatCondition cond3 = _cells.GetCell(10, 6).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0.9", "=1.1");
                cond3.Font.Bold = true;
                cond3.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                cond3.Interior.Color = 192;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatWeeklyPlanning()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_SheetHeaderError} . {ex.Message}");
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }




        private void btnReviewWeekPlanning_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameWkP)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    //_frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    //_frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    //if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(() => ReviewWeekPlanningAndResources());

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWeekPlanningAndResources()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void ReviewWeekPlanning()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var workGroup = "" + _cells.GetCell(2, 3).Value;
                var additional = "" + _cells.GetCell(2, 4).Value;
                var searchType = "" + _cells.GetCell(2, 5).Value;
                var startDate = "" + _cells.GetCell(4, 3).Value;
                var finishDate = "" + _cells.GetCell(4, 4).Value;

                if (searchType.Equals(SearchType.MstForecastOnly) || searchType.Equals(SearchType.WorkOrderAndMstForecast))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                }

                _cells?.SetCursorWait();

                var tableName = TableNameWkP;
                _cells.ClearTableRange(tableName);

                if (searchType.Equals(SearchType.WorkOrderOnly))
                    ReviewWeekPlanningTasks(_eFunctions, workGroup, startDate, finishDate, additional);
                if (searchType.Equals(SearchType.MstForecastOnly) || searchType.Equals(SearchType.WorkOrderAndMstForecast))
                    ReviewWeekPlanningTasksServices(_eFunctions, workGroup, startDate, finishDate, additional);
                UpdateResourceRequiredTable();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWeekPlanning()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void ReviewWeekPlanningAndResources()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var workGroup = "" + _cells.GetCell(2, 3).Value;
                var additional = "" + _cells.GetCell(2, 4).Value;
                var searchType = "" + _cells.GetCell(2, 5).Value;
                var startDate = "" + _cells.GetCell(4, 3).Value;
                var finishDate = "" + _cells.GetCell(4, 4).Value;

                if (searchType.Equals(SearchType.MstForecastOnly) || searchType.Equals(SearchType.WorkOrderAndMstForecast))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                }

                _cells?.SetCursorWait();

                var tableName = TableNameWkP;
                _cells.ClearTableRange(tableName);

                ReviewWeekPlanningAvailableResources(_eFunctions, workGroup);
                if (searchType.Equals(SearchType.WorkOrderOnly))
                    ReviewWeekPlanningTasks(_eFunctions, workGroup, startDate, finishDate, additional);
                if (searchType.Equals(SearchType.MstForecastOnly) || searchType.Equals(SearchType.WorkOrderAndMstForecast))
                    ReviewWeekPlanningTasksServices(_eFunctions, workGroup, startDate, finishDate, additional);
                UpdateResourceRequiredTable();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWeekPlanning()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ReviewWeekPlanningAvailableResources(EllipseFunctions eFunctions, string workGroup)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var availableResources = WeekPlanning.GetWorkGroupAvailableResources(eFunctions, workGroup);
            if (availableResources.Count > 10)
                throw new Exception("Error en Recursos de Grupo de Trabajo. No se pueden analizar más de 10 tipos de recursos diferentes");

            const int resourceRow = ResourceRowWkP;

            for (var i = resourceRow + 2; i < resourceRow + 12; i++)
            {
                _cells.GetCell(7, i).ClearComments();
                _cells.GetCell(7, i).Value = "";
                _cells.GetCell(8, i).Value = "";
            }

            var currentRow = resourceRow + 2;
            foreach (var res in availableResources)
            {
                _cells.GetCell(7, currentRow).AddComment(res.Description);
                _cells.GetCell(7, currentRow).Value = ("" + res.Type).Trim();
                _cells.GetCell(8, currentRow).Value = res.EstimatedHours;
                currentRow++;
            }
        }

        private void ReviewWeekPlanningTasks(EllipseFunctions eFunctions, string workGroup, string startDate, string finishDate, string additional)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            const int titleRow = TitleRowWkP;

            var currentRow = titleRow + 1;

            var taskList = WeekPlanning.GetWorkGroupTaskItems(eFunctions, workGroup, startDate, finishDate, additional);
            foreach (var task in taskList)
            {
                _cells.GetCell(1, currentRow).ClearComments();
                _cells.GetCell(1, currentRow).AddComment(task.EquipDesc);
                _cells.GetCell(1, currentRow).Value = task.EquipNo;
                _cells.GetCell(2, currentRow).Value = task.StdWo;
                _cells.GetCell(3, currentRow).Value = task.TaskNo;
                _cells.GetCell(4, currentRow).Value = task.TaskDescription;
                _cells.GetCell(5, currentRow).Value = task.TaskStatus;
                _cells.GetCell(6, currentRow).Value = task.NextSchedule;
                _cells.GetCell(7, currentRow).Value = task.ResType;
                _cells.GetCell(8, currentRow).Value = MyUtilities.ToDecimal(task.ActResHours, IxConversionConstant.DefaultNullAndEmpty);
                _cells.GetCell(9, currentRow).Value = MyUtilities.ToDecimal(task.EstResHours, IxConversionConstant.DefaultNullAndEmpty);
                var resPending = MyUtilities.ToDecimal(task.EstResHours, IxConversionConstant.DefaultNullAndEmpty) - MyUtilities.ToDecimal(task.ActResHours, IxConversionConstant.DefaultNullAndEmpty);
                _cells.GetCell(10, currentRow).Value = resPending > 0 ? resPending : 0;
                if (resPending < 0)
                {
                    _cells.GetCell(10, currentRow).Style = StyleConstants.Error;
                    _cells.GetCell(10, currentRow).ClearComments();
                    _cells.GetCell(10, currentRow).AddComment("Valor menor a Cero = " + resPending);
                }

                currentRow++;
            }
        }
        private void btnFormatWeekPlanning_Click(object sender, RibbonControlEventArgs e)
        {
            FormatWeeklyPlanning();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(@"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnReviewWeekPlanning_Click_1(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameWkP)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    //_frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    //_frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    //if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(() => ReviewWeekPlanning());

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWeekPlanning()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }


        private void ReviewWeekPlanningTasksServices(EllipseFunctions eFunctions, string workGroup, string startDate, string finishDate, string additional, string searchType = null)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var district = _frmAuth.EllipseDstrct;

            const int titleRow = TitleRowWkP;
            

            var currentRow = titleRow + 1;

            var taskOperationContext = new OperationContext
            {
                district = district,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            var workGroupList = new List<string> { workGroup };
            var taskSearchParam = new TaskSearchParam();
            taskSearchParam.AdditionalInformation = false;
            taskSearchParam.IncludeMst = true;
            taskSearchParam.OverlappingDates = true;
            taskSearchParam.StartDate = startDate;
            taskSearchParam.FinishDate = finishDate;
            taskSearchParam.DateInclude = additional;
            taskSearchParam.District = district;
            taskSearchParam.WorkGroups = workGroupList;
            taskSearchParam.SearchEntity = searchType;
            List<JobTask> ellipseJobTasks = null;
            try
            {
                ellipseJobTasks = JobActions.FetchJobsTasks(_eFunctions, urlService, taskOperationContext, taskSearchParam);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            if (ellipseJobTasks == null)
                return;
            foreach (var task in ellipseJobTasks)
            {
                try
                {
                    if (task.LabourResourcesList.Count > 0)
                    {
                        foreach (var r in task.LabourResourcesList)
                        {
                            _cells.GetCell(1, currentRow).ClearComments();
                            _cells.GetCell(1, currentRow).AddComment("" + task.ItemName1);
                            _cells.GetCell(1, currentRow).Value = task.EquipNo;
                            _cells.GetCell(2, currentRow).Value = task.WorkOrder ?? task.MaintSchTask + " " + task.StdJobNo;
                            if (string.IsNullOrWhiteSpace(task.WorkOrder))
                                _cells.GetCell(2, currentRow).Style = StyleConstants.Warning;
                            _cells.GetCell(3, currentRow).Value = task.WoTaskNo ?? task.StdJobTask;
                            _cells.GetCell(4, currentRow).Value = task.WoTaskDesc ?? task.WoDesc;
                            _cells.GetCell(5, currentRow).Value = task.WoStatusUDescription;
                            _cells.GetCell(6, currentRow).Value = task.PlanStrDate;
                            _cells.GetCell(7, currentRow).Value = r.ResourceCode;
                            _cells.GetCell(8, currentRow).Value = r.RealLabourHours;
                            _cells.GetCell(9, currentRow).Value = r.EstimatedLabourHours;
                            var resPending = r.EstimatedLabourHours - r.RealLabourHours;
                            _cells.GetCell(10, currentRow).Value = resPending > 0 ? resPending : 0;
                            if (resPending < 0)
                            {
                                _cells.GetCell(10, currentRow).Style = StyleConstants.Error;
                                _cells.GetCell(10, currentRow).ClearComments();
                                _cells.GetCell(10, currentRow).AddComment("Valor menor a Cero = " + resPending);
                            }

                            currentRow++;
                        }
                    }
                    else
                    {
                        double estimatedLabHours = 0;
                        double realLabHours = 0;
                        string resourceCode = "";
                        foreach (var r in task.LabourResourcesList)
                        {
                            resourceCode += " " + r.ResourceCode;
                            estimatedLabHours += r.EstimatedLabourHours;
                            realLabHours += r.RealLabourHours;
                        }

                        resourceCode = resourceCode.Trim();

                        _cells.GetCell(1, currentRow).ClearComments();
                        _cells.GetCell(1, currentRow).AddComment(task.ItemName1);
                        _cells.GetCell(1, currentRow).Value = task.EquipNo;
                        _cells.GetCell(2, currentRow).Value = task.WorkOrder ?? task.MaintSchTask + " " + task.StdJobNo;
                        if (string.IsNullOrWhiteSpace(task.WorkOrder))
                            _cells.GetCell(2, currentRow).Style = StyleConstants.Warning;
                        _cells.GetCell(3, currentRow).Value = task.WoTaskNo ?? task.StdJobTask;
                        _cells.GetCell(4, currentRow).Value = task.WoTaskDesc ?? task.WoDesc;
                        _cells.GetCell(5, currentRow).Value = "";
                        _cells.GetCell(6, currentRow).Value = task.PlanStrDate;
                        _cells.GetCell(7, currentRow).Value = resourceCode;
                        _cells.GetCell(8, currentRow).Value = realLabHours;
                        _cells.GetCell(9, currentRow).Value = estimatedLabHours;
                        var resPending = estimatedLabHours - realLabHours;
                        _cells.GetCell(10, currentRow).Value = resPending > 0 ? resPending : 0;
                        if (resPending < 0)
                        {
                            _cells.GetCell(10, currentRow).Style = StyleConstants.Error;
                            _cells.GetCell(10, currentRow).ClearComments();
                            _cells.GetCell(10, currentRow).AddComment("Valor menor a Cero = " + resPending);
                        }

                        currentRow++;
                    }
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse.cs:ReviewWeekPlanning()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                    MessageBox.Show(@"Se ha producido un error: " + ex.Message);
                    currentRow++;
                }
                finally
                {
                    _cells?.SetCursorDefault();
                }
            }
        }

        private void UpdateResourceRequiredTable()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            var resourceRow = ResourceRowWkP;
            var titleRow = TitleRowWkP;
            var currentRow = titleRow + 1;
            var schedResList = new List<string>();
            
            while(!string.IsNullOrWhiteSpace(_cells.GetCell(2, currentRow).Value) || !string.IsNullOrWhiteSpace(_cells.GetCell(6, currentRow).Value))
            {
                string resType = ("" + _cells.GetCell(7, currentRow).Value).Trim();
                
                if(!string.IsNullOrWhiteSpace(resType) && !schedResList.Contains(resType))
                    schedResList.Add(resType);
                currentRow++;
            }

            currentRow = resourceRow + 2;
            while (!string.IsNullOrWhiteSpace(_cells.GetCell(7, currentRow).Value))
            {
                string resType = "" + _cells.GetCell(7, currentRow).Value;
                if (schedResList.Contains(resType))
                    schedResList.Remove(resType);
                currentRow++;
            }
            foreach(var res in schedResList)
            {
                _cells.GetCell(7, currentRow).Value = res;
                currentRow++;
            }
        }

        private void btnUpdateReqResourceTable_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateResourceRequiredTable();
        }

        private void btnUpdateAvaResourceTable_Click(object sender, RibbonControlEventArgs e)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var workGroup = "" + _cells.GetCell(2, 3).Value;

            ReviewWeekPlanningAvailableResources(_eFunctions, workGroup);
            UpdateResourceRequiredTable();
        }
        #endregion

        #region Temperature Wagons
        
        private void TransformPlainLogsToMse345()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                FormatTempLogsMse345();

                _cells.SetCursorWait();
                var cellp = new ExcelStyleCells(_excelApp, SheetNamePlain);
                var cellm = new ExcelStyleCells(_excelApp, SheetNameMse345);

                var tableName = TableNameMse345;
                cellm.ClearTableRange(tableName);

                var errorsFound = 0;

                var monSet = "AT - ANALISIS TEMPERATURA";
                var component = "EJE  - EJE";

                var plainRow = TitleRowPlain + 1;
                var currentRow = TitleRowMse345 + 1;
                var currentAxis = 1;
                var previousCarNumber = "";
                while (!string.IsNullOrWhiteSpace("" + cellp.GetCell(1, plainRow).Value))
                {
                    try
                    {
                        string carNumber = "" + cellp.GetCell(4, plainRow).Value;
                        if (string.IsNullOrWhiteSpace(carNumber))
                            continue;
                        if (cbTempWagIgnoreLocomotives.Checked && carNumber.Length == 7 && carNumber.StartsWith("10000"))
                            continue;
                        //reinicio el eje si hay cambio de carro
                        if (!previousCarNumber.Equals("" + carNumber, StringComparison.InvariantCultureIgnoreCase))
                            currentAxis = 1;

                        //Campos Comunes
                        cellm.GetRange(1, currentRow, 1, currentRow + 4).Value = monSet;
                        cellm.GetRange(2, currentRow, 2, currentRow + 4).Value = carNumber;
                        cellm.GetRange(3, currentRow, 3, currentRow + 4).Value = cellp.GetCell(1, plainRow).Value;
                        cellm.GetRange(4, currentRow, 4, currentRow + 4).Value = component;
                        //Codigo
                        cellm.GetCell(7, currentRow).Value = "TEJE" + currentAxis + "DIF";
                        cellm.GetCell(7, currentRow + 1).Value = "TEJE" + currentAxis + "ROD";
                        cellm.GetCell(7, currentRow + 2).Value = "TEJE" + currentAxis + "ROI";
                        cellm.GetCell(7, currentRow + 3).Value = "TEJE" + currentAxis + "RUD";
                        cellm.GetCell(7, currentRow + 4).Value = "TEJE" + currentAxis + "RUI";
                        //Valores
                        cellm.GetCell(9, currentRow).Value = cellp.GetCell(7, plainRow).Value;
                        cellm.GetCell(9, currentRow + 1).Value = cellp.GetCell(8, plainRow).Value;
                        cellm.GetCell(9, currentRow + 2).Value = cellp.GetCell(9, plainRow).Value;
                        cellm.GetCell(9, currentRow + 3).Value = cellp.GetCell(10, plainRow).Value;
                        cellm.GetCell(9, currentRow + 4).Value = cellp.GetCell(11, plainRow).Value;

                        previousCarNumber = carNumber;
                        currentAxis++;
                        currentRow += 5;
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:TransformPlainLogsToMse345()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        cellp.GetCell(1, plainRow).Style = StyleConstants.Error;
                        cellp.GetCell(1, plainRow).ClearComments();
                        cellp.GetCell(1, plainRow).AddComment(ex.Message);

                        errorsFound++;
                    }
                    finally
                    {
                        plainRow++;
                    }
                }

                if (errorsFound > 0)
                    MessageBox.Show("Se encontraron " + errorsFound + " errores. Verifique la hoja de format plano para más detalles");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:TransformPlainLogsToMse345()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show("Error. " + ex.Message, "Transform Plain Formt to MSE345", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void LoadLocationLogsToPlain()
        {
            try
            {


                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                var defaultPath = Settings.CurrentSettings.GetCustomSettingValue(CustomVariables.WagonTemperatureLogPath);
                var fbd = new FolderBrowserDialog { SelectedPath = defaultPath };
                var result = fbd.ShowDialog();

                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(fbd.SelectedPath)) return;

                _cells.SetCursorWait();
                Settings.CurrentSettings.SetCustomSettingValue(CustomVariables.WagonTemperatureLogPath, fbd.SelectedPath);
                Settings.CurrentSettings.SaveCustomSettings();

                if (!_cells.ActiveSheet.Name.Equals("Plano"))
                    FormatTempLogsPlain();

                const string tableName = TableNamePlain;

                //_cells.ClearTableRange(tableName);

                var files = Directory.GetFiles(fbd.SelectedPath);
                var filesList = new List<string>();


                foreach (var file in files)
                {
                    if (!file.EndsWith(".log"))
                        continue;

                    var startingRow = _cells.GetTableLastRowIndex(tableName) + 1;
                    TempWagonActions.TransformLogToPlain(file, _cells, startingRow);
                    filesList.Add(Path.GetFileName(file));
                }

                var filesLoaded = filesList.Aggregate("Archivos Cargados:", (current, file) => current + ("\n" + file));
                MessageBox.Show(filesLoaded, "Load Logs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.LoadLocationLogs()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.Message, "Load Logs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void FormatTempLogsMse345()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameMse345;

                _cells.SetCursorWait();

                const int titleRow = TitleRowMse345;
                const int resultColumn = ResultColumnMse345;
                const string tableName = TableNameMse345;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");
                _cells.MergeCells("C6", "L11");
                _cells.MergeCells("A12", "L12");

                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("M1", "XFD1048576").Columns.Hidden = true;


                _cells.GetRange("A6", "A11").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B6", "B11").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("A6").Value = "MONITOREO";

                _cells.GetCell("A7").Value = "EQUIPO";
                _cells.GetCell("B7").NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell("A8").Value = "FECHA";
                _cells.GetCell("B8").Value = DateTime.Now.ToString("yyyyMMdd");

                _cells.GetCell("A9").Value = "INSPECTOR 1";
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A10").Value = "INSPECTOR 2";
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A11").Value = "INSPECTOR 3";
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleInformation;


                _cells.GetCell(1, titleRow).Value = "MONITOREO";
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = "EQUIPO";
                _cells.GetCell(2, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, titleRow).Value = "FECHA";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, titleRow).Value = "COMPONENTE";
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, titleRow).Value = "MODIFICADOR";
                _cells.GetCell(5, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, titleRow).Value = "POSICION";
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).Value = "CODIGO";
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(8, titleRow).Value = "DESCRIPCION";
                _cells.GetCell(8, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetCell(9, titleRow).Value = "VALOR ENCONTRADO";
                _cells.GetCell(9, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(10, titleRow).Value = "COMENTARIO";
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatLogsMse345()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                throw new Exception(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void FormatTempLogsPlain()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNamePlain;
                _cells.SetCursorWait();

                const int titleRow = TitleRowPlain;
                const int resultColumn = 11;
                const string tableName = TableNamePlain;

                _cells.GetCell(1, titleRow).Value = "DATE";
                _cells.GetCell(2, titleRow).Value = "CAR ORDER";
                _cells.GetCell(3, titleRow).Value = "CAR OWNER";
                _cells.GetCell(4, titleRow).Value = "AXLE NUMBER";
                _cells.GetCell(5, titleRow).Value = "---";
                _cells.GetCell(6, titleRow).Value = "SPACING (M)";
                _cells.GetCell(7, titleRow).Value = "CH1 (C)";
                _cells.GetCell(8, titleRow).Value = "CH2 (C)";
                _cells.GetCell(9, titleRow).Value = "CH3 (C)";
                _cells.GetCell(10, titleRow).Value = "CH4 (C)";
                _cells.GetCell(11, titleRow).Value = "ALARMS";

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetRange(7, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = "0.0";

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatTempLogsPlain()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                throw new Exception(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }

            _cells.SetCursorDefault();
        }

        private void btnLoadTempLogMse345_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNamePlain)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(TransformPlainLogsToMse345);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                {
                    MessageBox.Show(
                        @"Debe seleccionar una hoja con el formato de Carga Plano");
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatTempLogsMse345()",
                    "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }

        private void btnLoadTempLogPlain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(LoadLocationLogsToPlain);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:LoadLocationLogsToPlain()",
                    "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }


        private void cbTempWagIgnoreLocomotives_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue(TempWagonConstants.IgnoreLocomotives, cbTempWagIgnoreLocomotives.Checked.ToString());
            Settings.CurrentSettings.SaveCustomSettings();
        }


        #endregion

        #region Plan History
        private void btnPlanHistoryFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatPlanHistory();
        }
        private void btnPlanHistoryReview_Click(object sender, RibbonControlEventArgs e)
        {
            //si ya hay un thread corriendo que no se ha detenido
            if (_thread != null && _thread.IsAlive) return;
            _thread = new Thread(ReviewPlanHistory);

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void btnPlanHistoryLoad_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNamePlanHist)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(LoadPlanHistory);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                {
                    MessageBox.Show(
                        @"Debe seleccionar una hoja con el formato de Cargue de Historia de Programación");
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:LoadPlanHistory()",
                    "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void FormatPlanHistory()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                var sheetName = SheetNamePlanHist;
                var titleRow = TitleRowPlanHist;
                var resultColumn = ResultColumnPlanHist;
                var tableName = TableNamePlanHist;

                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;


                _cells.GetCell(1, titleRow).Value = "FECHA";
                _cells.GetCell(2, titleRow).Value = "GRUPO";
                _cells.GetCell(3, titleRow).Value = "ID_CONCEPTO";
                _cells.GetCell(4, titleRow).Value = "CONCEPTO";
                _cells.GetCell(5, titleRow).Value = "VALOR1";
                _cells.GetCell(6, titleRow).Value = "VALOR2";

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleResult);

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatPlanHistory()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                throw new Exception(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ReviewPlanHistory()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();
            if (!_cells.ActiveSheet.Name.Equals(SheetNamePlanHist))
                FormatPlanHistory();

            const int titleRow = TitleRowPlanHist;
            const string sheetName = SheetNamePlanHist;
            const string tableName = TableNamePlanHist;
            const int resultColumn = ResultColumnPlanHist;


            _cells.ClearTableRange(tableName);
            _eFunctions.SetDBSettings(Environments.SigcorProductivo);

            var i = titleRow + 1;//itera la celda para cada tarea

            if (string.IsNullOrWhiteSpace(SelectedStartDatePlanHist))
                SelectedStartDatePlanHist = MyUtilities.ToString(DateTime.Today, MyUtilities.DateTime.DateYYYYMMDD);
            var startDate = InputBox.GetValue("CONSULTAR HISTORIA DE PLANEACIÓN", "Fecha de Pogramación a Consultar (YYYYMMDD):", SelectedStartDatePlanHist);
            var planItems = PlanHistory.PlanHistoryActions.ReviewPlanHistory(_eFunctions, startDate, null, null, null);

            foreach (var item in planItems)
            {
                //Para resetear el estilo
                _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                _cells.GetCell(1, i).Value = item.Fecha;
                _cells.GetCell(2, i).Value = item.Grupo;
                _cells.GetCell(3, i).Value = "'" + item.IdConcepto;
                _cells.GetCell(4, i).Value = item.Concepto;
                _cells.GetCell(5, i).Value = item.Valor1;
                _cells.GetCell(6, i).Value = item.Valor2;

                i++;//aumenta tarea
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
            _eFunctions.CloseConnection();
        }

        private void LoadPlanHistory()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();

            const int titleRow = TitleRowPlanHist;
            const string sheetName = SheetNamePlanHist;
            const string tableName = TableNamePlanHist;
            const int resultColumn = ResultColumnPlanHist;

            _eFunctions.SetDBSettings(Environments.SigcorProductivo);

            var i = titleRow + 1;//itera la celda para cada tarea

            while(!string.IsNullOrWhiteSpace(_cells.GetCell(1, i).Value))
            {
                try
                {
                    var item = new PlanHistory.PlanHistoryItem();
                    item.Fecha = _cells.GetCell(1, i).Value;
                    item.Grupo = _cells.GetCell(2, i).Value;
                    item.IdConcepto = _cells.GetCell(3, i).Value;
                    item.Concepto = _cells.GetCell(4, i).Value;
                    item.Valor1 = "" + _cells.GetCell(5, i).Value;
                    item.Valor2 = "" + _cells.GetCell(6, i).Value;

                    var result = PlanHistory.PlanHistoryActions.LoadPlanHistoryItem(_eFunctions, item);
                    if (result == 0)
                        throw new Exception("No se ha cargado ningún registro");

                    _cells.GetCell(resultColumn, i).Value = "CARGADO";
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse:LoadPlanHistory()",
                        "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                    _cells.GetCell(resultColumn, i).Value = ex.Message;
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                }
                finally
                {
                    i++;
                }
                
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
            _eFunctions.CloseConnection();
        }
        #endregion


    }
}

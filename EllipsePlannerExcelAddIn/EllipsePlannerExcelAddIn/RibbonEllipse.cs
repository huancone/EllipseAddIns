using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Utilities;
using EllipseJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using OperationContext = EllipseWorkOrdersClassLibrary.ResourceReqmntsService.OperationContext;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SearchFieldCriteriaType = EllipseJobsClassLibrary.SearchFieldCriteriaType;

//si es screen service

// ReSharper disable UseNullPropagation
// ReSharper disable UseStringInterpolation
// ReSharper disable UseIndexedProperty
// ReSharper disable SuggestVarOrType_Elsewhere

namespace EllipsePlannerExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private Application _excelApp;
        private FormAuthenticate _frmAuth;
        private Thread _thread;

        //Hojas
        private const string ValidationSheetName = "Validacion";
        private const string ResourcesSheetName = "Planeados";
        private const string EllipseResourcesSheetName = "Estimados";
        private const string PeopleSoftResourcesSheetName = "PeopleSoft";
        private const string DailySheetName = "Plan Diario";

        //Tablas
        private const string TableJobResources = "JobResources";
        private const string TableEllipseResources = "EllipseResources";
        private const string TableIndicator = "Indicator";
        private const string TableDailyEllipseResources = "DailyEllipseResources";
        private const string TablePsoftResources = "PsoftResources";
        private const string TableDaily = "DailyResources";

        //Titulos
        private const int TitleRowResources = 8;//Hoja 1
        private const int TitleRowEllipse = 6;//Hoja 2, 3 y 4

        //Columnas de Resultado
        private const int ResultColumnResources = 31;
        private const int ResultColumnEllipse = 7;
        private const int ActionColumn = 17;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            
            LoadSettings();
        }

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

            settings.SetDefaultCustomSettingValue("DeviationStats", "Y");
            settings.SetDefaultCustomSettingValue("SplitByResource", "Y");
            settings.SetDefaultCustomSettingValue("IncludeMsts", "Y");
            settings.SetDefaultCustomSettingValue("OverlappingDateSearch", "N");

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //Setting of Configuration Options from Config File (or default)
            cbDeviationStats.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("DeviationStats"));
            cbSplitTaskByResource.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("SplitByResource"));
            cbIncludeMsts.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IncludeMsts"));
            cbOverlappingDateSearch.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("OverlappingDateSearch"));

            settings.SaveCustomSettings();
            //
        }

        #region Buttons
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
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnReviewJobs_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewJobListPost);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnReviewJobs()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnLoadData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(LoadJobPlan);
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnLoadData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateEllipse_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(LoadEllipse);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnUpdateEllipse()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateOrder_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ExecuteRequirementActions);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:btnUpdateOrder_Click()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        #endregion
private void FormatSheet()
        {

            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 5)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                //hoja de validación
                _cells.CreateNewWorksheet(ValidationSheetName);


                #region hoja 1
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = ResourcesSheetName;

                var titleRow = TitleRowResources;
                var resultColumn = ResultColumnResources;
                var districtList = Districts.GetDistrictList();

                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;

                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);

                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.GetCell("A4").AddComment("--ÁREA GERENCIAL/SUPERINTENDENCIA--\n" +
                    "INST: IMIS, MINA\n" +
                    "" + ManagementArea.ManejoDeCarbon.Key + ": " + QuarterMasters.Ferrocarril.Key + ", " + QuarterMasters.PuertoBolivar.Key + ", " + QuarterMasters.PlantasDeCarbon.Key + "\n" +
                    "" + ManagementArea.Mantenimiento.Key + ": MINA\n" +
                    "" + ManagementArea.SoporteOperacion.Key + ": ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.GetCell("A5").Value = "Trabajos Adicionales";

                var aditionalJobsLis = new List<string> { "Backlog", "Unscheduled", "Backlog and Unscheduled", "Backlog Only", "Unscheduled Only", "Backlog and Unscheduled Only" };

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 3, false);
                _cells.SetValidationList(_cells.GetCell("B5"), aditionalJobsLis, ValidationSheetName, 4, false);

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);


                var nextWednesday = DateTime.Now;
                while (nextWednesday.DayOfWeek != DayOfWeek.Wednesday)
                    nextWednesday = nextWednesday.AddDays(1);
                var lastThursday = DateTime.Now;
                while (lastThursday.DayOfWeek != DayOfWeek.Thursday)
                    lastThursday = lastThursday.AddDays(-1);


                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = EllipseJobsClassLibrary.SearchDateCriteriaType.PlannedStart.Value;
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = lastThursday.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = nextWednesday.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RECURSOS NECESARIOS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");
                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetCell(1, titleRow).Value = "Grupo";
                _cells.GetCell(2, titleRow).Value = "Equipo";
                _cells.GetCell(3, titleRow).Value = "Componente";
                _cells.GetCell(4, titleRow).Value = "Mod";
                _cells.GetCell(5, titleRow).Value = "Eq Desc";
                _cells.GetCell(6, titleRow).Value = "MST";
                _cells.GetCell(7, titleRow).Value = "Referencia";
                _cells.GetCell(8, titleRow).Value = "Ref Desc";
                _cells.GetCell(9, titleRow).Value = "Tarea";
                _cells.GetCell(10, titleRow).Value = "Recurso";
                _cells.GetCell(11, titleRow).Value = "Horas Estimadas";
                _cells.GetCell(12, titleRow).Value = "Horas Reales";
                _cells.GetCell(13, titleRow).Value = "Horas restantes";
                _cells.GetCell(14, titleRow).Value = "Fecha Planeada";
                _cells.GetCell(15, titleRow).Value = "Hora Planeada";
                _cells.GetCell(16, titleRow).Value = "Duracion";
                _cells.GetCell(17, titleRow).Value = "Accion";
                _cells.GetCell(18, titleRow).Value = "Codigo de Cierre";
                _cells.GetCell(19, titleRow).Value = "Fecha de Cierre";
                _cells.GetCell(20, titleRow).Value = "Asignado";
                _cells.GetCell(21, titleRow).Value = "Tipo MT";
                _cells.GetCell(22, titleRow).Value = "Job Code";
                _cells.GetCell(23, titleRow).Value = "Estad. Pr.";
                _cells.GetCell(24, titleRow).Value = "Fecha Original";
                _cells.GetCell(25, titleRow).Value = "Mínima Fecha";
                _cells.GetCell(26, titleRow).Value = "Máxima Fecha";
                _cells.GetCell(27, titleRow).Value = "Estad. Original";
                _cells.GetCell(28, titleRow).Value = "Estad. Actual";
                _cells.GetCell(29, titleRow).Value = "Mínima Estad.";
                _cells.GetCell(30, titleRow).Value = "Máxima Estad.";
                _cells.GetCell(resultColumn, titleRow).Value = "Resultado";


                _cells.GetCell(ActionColumn, titleRow).Style = StyleConstants.TitleAction;
                _cells.SetValidationList(_cells.GetCell(ActionColumn, titleRow + 1), new List<string> { "M", "C", "D", "Close Task" });
                _cells.GetCell(ActionColumn, titleRow).AddComment("C: Crear Requerimiento \nM: Modificar Requerimiento \nD: Eliminar Requerimiento \nClose Task: Cerrar Tarea");


                var completeCodeList = _eFunctions.GetItemCodes("SC").Select(item => item.Code + " - " + item.Description).ToList();
                _cells.SetValidationList(_cells.GetCell(18, titleRow + 1), completeCodeList, ValidationSheetName, 10, false);

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), TableJobResources);
                _cells.GetRange(1, titleRow, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region hoja 2 - Estimados

                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = EllipseResourcesSheetName;
                titleRow = TitleRowEllipse;


                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RECURSOS ELLIPSE - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");
                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetCell(1, titleRow).Value = "Grupo";
                _cells.GetCell(2, titleRow).Value = "Fecha";
                _cells.GetCell(3, titleRow).Value = "Recurso";
                _cells.GetCell(4, titleRow).Value = "Planeadas";
                _cells.GetCell(5, titleRow).Value = "Disponibles";
                _cells.GetRange(1, titleRow, 5, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, 5, titleRow + 1), TableEllipseResources);
                _cells.GetRange(1, titleRow, 5, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(7, titleRow).Value = "Grupo";
                _cells.GetCell(8, titleRow).Value = "Fecha";
                _cells.GetCell(9, titleRow).Value = "Horas Planeadas";
                _cells.GetCell(10, titleRow).Value = "Horas Disponibles";
                _cells.GetCell(11, titleRow).Value = "Indicador de Planeacion";
                _cells.GetRange(7, titleRow, 11, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(7, titleRow, 11, titleRow + 1), TableIndicator);
                _cells.GetRange(7, titleRow, 10, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(11, titleRow + 1).NumberFormat = NumberFormatConstants.Percentage;

                _cells.GetCell(13, titleRow - 1).Value = "Datos Ellipse";
                _cells.MergeCells(13, titleRow - 1, 19, titleRow - 1);
                _cells.GetRange(13, titleRow - 1, 19, titleRow - 1).Style = StyleConstants.TitleInformation;

                _cells.GetCell(13, titleRow).Value = "Grupo";
                _cells.GetCell(14, titleRow).Value = "fecha";
                _cells.GetCell(15, titleRow).Value = "Recurso";
                _cells.GetCell(16, titleRow).Value = "Cantidad";
                _cells.GetCell(17, titleRow).Value = "Estimado";
                _cells.GetCell(18, titleRow).Value = "Disponible";
                _cells.GetCell(19, titleRow).Value = "resultado";
                _cells.GetRange(13, titleRow, 18, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(13, titleRow, 19, titleRow + 1), TableDailyEllipseResources);
                _cells.GetRange(13, titleRow, 19, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region hoja 3 - Recursos Peoplesoft
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = PeopleSoftResourcesSheetName;
                titleRow = TitleRowEllipse;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RECURSOS ELLIPSE - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");
                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetCell(1, titleRow).Value = "Grupo";
                _cells.GetCell(2, titleRow).Value = "Fecha";
                _cells.GetCell(3, titleRow).Value = "Recurso";
                _cells.GetCell(4, titleRow).Value = "Cedula";
                _cells.GetCell(5, titleRow).Value = "Nombre";
                _cells.GetCell(6, titleRow).Value = "Horas";
                _cells.GetRange(1, titleRow, 6, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, 6, titleRow + 1), TablePsoftResources);
                _cells.GetRange(1, titleRow, 6, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region hoja 4 - Plan Diario
                _excelApp.ActiveWorkbook.Sheets.get_Item(4).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = DailySheetName;
                titleRow = TitleRowEllipse;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RECURSOS DIARIOS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");
                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetCell(1, titleRow).Value = "Grupo";
                _cells.GetCell(2, titleRow).Value = "Equipo";
                _cells.GetCell(3, titleRow).Value = "Componente";
                _cells.GetCell(4, titleRow).Value = "Mod";
                _cells.GetCell(5, titleRow).Value = "Eq Desc";
                _cells.GetCell(6, titleRow).Value = "MST";
                _cells.GetCell(7, titleRow).Value = "Referencia";
                _cells.GetCell(8, titleRow).Value = "Ref Desc";
                _cells.GetCell(9, titleRow).Value = "Tarea";
                _cells.GetCell(10, titleRow).Value = "Turno";
                _cells.GetCell(11, titleRow).Value = "Hora Inicio";
                _cells.GetCell(12, titleRow).Value = "Hora Fin";
                _cells.GetCell(13, titleRow).Value = "Duracion Turno";
                _cells.GetCell(14, titleRow).Value = "Recurso";
                _cells.GetCell(15, titleRow).Value = "Horas Requeridas";

                _cells.GetRange(1, titleRow, 15 - 1, titleRow).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, 15, titleRow + 1), TableJobResources);
                _cells.GetRange(1, titleRow, 15, titleRow + 1).NumberFormat = NumberFormatConstants.Text;


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
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void ReviewJobListPost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();


                #region Hoja 1 - Tareas
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                var taskSearchParam = new TaskSearchParam();
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
                var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
                var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
                var dateInclude = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
                var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
                var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);
                var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
                var groupList = new List<string>();

                if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    groupList.Add(searchCriteriaValue1);
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    groupList = Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList();
                else
                    groupList = Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList();


                taskSearchParam.AdditionalInformation = cbDeviationStats.Checked;
                taskSearchParam.IncludeMst = cbIncludeMsts.Checked;
                taskSearchParam.OverlappingDates = cbOverlappingDateSearch.Checked;
                taskSearchParam.StartDate = startDate;
                taskSearchParam.FinishDate = endDate;
                taskSearchParam.DateInclude = dateInclude;
                taskSearchParam.District = district;
                taskSearchParam.WorkGroups = groupList;


                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var taskOperationContext = new EllipseJobsClassLibrary.WorkOrderTaskMWPService.OperationContext()
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                //consumo de servicio de msewts
                List<JobTask> ellipseJobTasks = JobActions.FetchJobsTasks(_eFunctions, urlService, taskOperationContext, taskSearchParam);

                //consulta sobre tabla de Ellipse mso720
                List<LabourResources> ellipseResources = JobActions.GetEllipseResources(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                //consulta sobre tabla de Ellipse mso720
                List<LabourResources> pSoftResources = JobActions.GetPsoftResources(district, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                //jobs sin repetirse
                //var singlejobs = ellipseJobs.GroupBy(a=> new { a.WorkGroup, a.MstReference, a.Reference })


                //recursos planeados ellipse agrupados por grupo/fecha/recurso
                var ellipseTotalresource = (from task in ellipseJobTasks from resources in task.LabourResourcesList select resources).GroupBy(l => new { l.WorkGroup, l.Date, l.ResourceCode })
                    .Select(cl => new LabourResources
                    {
                        WorkGroup = cl.First().WorkGroup,
                        ResourceCode = cl.First().ResourceCode,
                        Date = cl.First().Date,
                        EstimatedLabourHours = cl.Sum(c => c.EstimatedLabourHours),
                        RealLabourHours = cl.Sum(c => c.RealLabourHours)
                    }).ToList();

                //union de tareas planeadas y horas disponibles por grupo/recurso/fecha
                var totalDailyResources = ellipseTotalresource.Union(ellipseResources).GroupBy(a => new { a.WorkGroup, a.Date, a.ResourceCode }).Select(cl => new LabourResources
                {
                    WorkGroup = cl.First().WorkGroup,
                    ResourceCode = cl.First().ResourceCode,
                    Date = cl.First().Date,
                    Quantity = cl.Max(c => c.Quantity),
                    EstimatedLabourHours = cl.Sum(c => c.EstimatedLabourHours),
                    RealLabourHours = cl.Sum(c => c.RealLabourHours),
                    AvailableLabourHours = cl.Sum(c => c.AvailableLabourHours)
                }).ToList();

                //Recursos disponibles agrupados por grupo/recurso
                var totalWeeklyResources = (from r in totalDailyResources select r).GroupBy(l => new { l.WorkGroup, l.ResourceCode })
                    .Select(cl => new LabourResources
                    {
                        WorkGroup = cl.First().WorkGroup,
                        Date = startDate,
                        ResourceCode = cl.First().ResourceCode,
                        Quantity = cl.Max(c => c.Quantity),
                        AvailableLabourHours = cl.Sum(c => c.AvailableLabourHours),
                        EstimatedLabourHours = cl.Sum(c => c.EstimatedLabourHours),
                        RealLabourHours = cl.Sum(c => c.RealLabourHours)
                    }).ToList();

                //variable para el calculo del indicador de disponible/planeado
                var plannerIndicator = (from r in totalDailyResources select r).GroupBy(l => new { l.WorkGroup })
                    .Select(cl => new LabourResources
                    {
                        WorkGroup = cl.First().WorkGroup,
                        Date = startDate,
                        EstimatedLabourHours = cl.Sum(c => c.EstimatedLabourHours),
                        RealLabourHours = cl.Sum(c => c.RealLabourHours),
                        AvailableLabourHours = cl.Sum(c => c.AvailableLabourHours)
                    }).ToList();

                //LLenado de tablas
                var titleRow = TitleRowResources;
                var resultColumn = ResultColumnResources;

                _cells.ClearTableRange(TableJobResources);
                var i = titleRow + 1;
                foreach (var jt in ellipseJobTasks)
                {
                    if (jt.LabourResourcesList.Count > 0 && cbSplitTaskByResource.Checked)
                    {
                        foreach (var r in jt.LabourResourcesList)
                        {
                            _cells.GetCell(1, i).Value = jt.WorkGroup;
                            _cells.GetCell(2, i).Value = jt.EquipNo;
                            _cells.GetCell(3, i).Value = jt.CompCode;
                            _cells.GetCell(4, i).Value = jt.CompModCode;
                            _cells.GetCell(5, i).Value = jt.ItemName1;
                            _cells.GetCell(6, i).Value = jt.MaintSchTask;
                            _cells.GetCell(7, i).Value = jt.WorkOrder ?? jt.StdJobNo;
                            _cells.GetCell(8, i).Value = jt.WoTaskNo ?? jt.StdJobTask;
                            _cells.GetCell(9, i).Value = jt.WoTaskDesc ?? jt.WoDesc ;
                            _cells.GetCell(10, i).Value = r.ResourceCode;
                            _cells.GetCell(11, i).Value = r.EstimatedLabourHours;
                            _cells.GetCell(12, i).Value = r.RealLabourHours;
                            _cells.GetCell(13, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                            _cells.GetCell(14, i).Value = jt.PlanStrDate;
                            _cells.GetCell(15, i).Value = jt.PlanStrTime;
                            _cells.GetCell(16, i).Value = jt.EstimatedDurationsHrs;
                            _cells.GetCell(20, i).Value = jt.AssignPerson;
                            if (jt.Additional != null)
                            {
                                if(string.IsNullOrWhiteSpace(jt.AssignPerson))
                                    _cells.GetCell(20, i).Value = string.IsNullOrWhiteSpace(jt.Additional.AssignPerson) ? jt.Additional.WorkOrderAssignPerson : jt.Additional.AssignPerson;
                                _cells.GetCell(21, i).Value = "" + jt.Additional.WorkOrderType;
                                _cells.GetCell(22, i).Value = "" + jt.Additional.JobDescCode;
                                _cells.GetCell(23, i).Value = "" + jt.Additional.EquipPrimaryStatType;
                                _cells.GetCell(24, i).Value = "" + jt.Additional.OriginalSchedDate;
                                _cells.GetCell(25, i).Value = "" + jt.Additional.MinSchedDate;
                                _cells.GetCell(26, i).Value = "" + jt.Additional.MaxSchedDate;
                                _cells.GetCell(27, i).Value = "" + jt.Additional.ScheduleStatValue;
                                _cells.GetCell(28, i).Value = "" + jt.Additional.ActualStatValue;
                                _cells.GetCell(29, i).Value = "" + jt.Additional.MinSchedStat;
                                _cells.GetCell(30, i).Value = "" + jt.Additional.MaxSchedStat;
                            }
                            i++;
                        }
                    }
                    else
                    {
                        double estimatedLabHours = 0;
                        double realLabHours = 0;
                        string resourceCode = "";
                        foreach (var r in jt.LabourResourcesList)
                        {
                            resourceCode += " " + r.ResourceCode;
                            estimatedLabHours += r.EstimatedLabourHours;
                            realLabHours += r.RealLabourHours;
                        }

                        resourceCode = resourceCode.Trim();
                            
                        _cells.GetCell(1, i).Value = jt.WorkGroup;
                        _cells.GetCell(2, i).Value = jt.EquipNo;
                        _cells.GetCell(3, i).Value = jt.CompCode;
                        _cells.GetCell(4, i).Value = jt.CompModCode;
                        _cells.GetCell(5, i).Value = jt.ItemName1;
                        _cells.GetCell(6, i).Value = jt.MaintSchTask;
                        _cells.GetCell(7, i).Value = jt.WorkOrder ?? jt.StdJobNo;
                        _cells.GetCell(8, i).Value = jt.WoTaskNo ?? jt.StdJobTask;
                        _cells.GetCell(9, i).Value = jt.WoDesc;
                        _cells.GetCell(10, i).Value = resourceCode;
                        _cells.GetCell(11, i).Value = estimatedLabHours;
                        _cells.GetCell(12, i).Value = realLabHours;
                        _cells.GetCell(13, i).Value = estimatedLabHours - realLabHours;
                        _cells.GetCell(14, i).Value = jt.PlanStrDate;
                        _cells.GetCell(15, i).Value = jt.PlanStrTime;
                        _cells.GetCell(16, i).Value = jt.EstimatedDurationsHrs;
                        _cells.GetCell(20, i).Value = jt.AssignPerson;
                        if (jt.Additional != null)
                        {
                            if (string.IsNullOrWhiteSpace(jt.AssignPerson))
                                _cells.GetCell(20, i).Value = string.IsNullOrWhiteSpace(jt.Additional.AssignPerson) ? jt.Additional.WorkOrderAssignPerson : jt.Additional.AssignPerson;
                            _cells.GetCell(21, i).Value = "" + jt.Additional.WorkOrderType;
                            _cells.GetCell(22, i).Value = "" + jt.Additional.JobDescCode;
                            _cells.GetCell(23, i).Value = "" + jt.Additional.EquipPrimaryStatType;
                            _cells.GetCell(24, i).Value = "" + jt.Additional.OriginalSchedDate;
                            _cells.GetCell(25, i).Value = "" + jt.Additional.MinSchedDate;
                            _cells.GetCell(26, i).Value = "" + jt.Additional.MaxSchedDate;
                            _cells.GetCell(27, i).Value = "" + jt.Additional.ScheduleStatValue;
                            _cells.GetCell(28, i).Value = "" + jt.Additional.ActualStatValue;
                            _cells.GetCell(29, i).Value = "" + jt.Additional.MinSchedStat;
                            _cells.GetCell(30, i).Value = "" + jt.Additional.MaxSchedStat;
                        }
                        i++;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion
                #region Hoja 2 - Estimados

                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                titleRow = TitleRowEllipse;
                resultColumn = ResultColumnEllipse;
                _cells.ClearTableRange(TableEllipseResources);
                _cells.ClearTableRange(TableDailyEllipseResources);
                _cells.ClearTableRange(TableIndicator);

                i = titleRow + 1;
                foreach (var r in totalDailyResources)
                {
                    _cells.GetCell(1, i).Value = r.WorkGroup;
                    _cells.GetCell(2, i).Value = r.Date;
                    _cells.GetCell(3, i).Value = r.ResourceCode;
                    _cells.GetCell(4, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                    _cells.GetCell(5, i).Value = r.AvailableLabourHours;
                    i++;
                }

                i = titleRow + 1;
                foreach (var r in plannerIndicator)
                {
                    _cells.GetCell(7, i).Value = r.WorkGroup;
                    _cells.GetCell(8, i).Value = r.Date;
                    _cells.GetCell(9, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                    _cells.GetCell(10, i).Value = r.AvailableLabourHours;
                    _cells.GetCell(11, i).Value = (r.EstimatedLabourHours - r.RealLabourHours) / r.AvailableLabourHours;
                    i++;
                }

                i = titleRow + 1;
                foreach (var r in totalWeeklyResources)
                {
                    _cells.GetCell(13, i).Value = r.WorkGroup;
                    _cells.GetCell(14, i).Value = r.Date;
                    _cells.GetCell(15, i).Value = r.ResourceCode;
                    _cells.GetCell(16, i).Value = r.Quantity;
                    _cells.GetCell(17, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                    _cells.GetCell(18, i).Value = r.AvailableLabourHours;
                    i++;
                }

                // Add chart.
                var charts = _excelApp.ActiveWorkbook.ActiveSheet.ChartObjects() as ChartObjects;
                if (charts != null)
                {
                    var chartObject = charts.Add(60, 10, 300, 300);
                    var chart = chartObject.Chart;

                    // Set chart range.
                    var range = _cells.GetRange(1, titleRow, resultColumn, titleRow + totalDailyResources.Count);
                    chart.SetSourceData(range);

                    // Set chart properties.
                    chart.ChartType = XlChartType.xlColumnClustered;
                    chart.ChartWizard(range);
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region Hoja 3 - PeopleSoft
                //hoja de ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                titleRow = TitleRowEllipse;

                _cells.ClearTableRange(TablePsoftResources);

                i = titleRow + 1;
                foreach (var r in pSoftResources)
                {
                    _cells.GetCell(1, i).Value = r.WorkGroup;
                    _cells.GetCell(2, i).Value = r.Date;
                    _cells.GetCell(3, i).Value = r.ResourceCode;
                    _cells.GetCell(4, i).Value = r.EmployeeId;
                    _cells.GetCell(5, i).Value = r.EmployeeName;
                    _cells.GetCell(6, i).Value = r.AvailableLabourHours;
                    i++;
                }
                #endregion
                
                #region Hoja 4 - Plan Diario
                _excelApp.ActiveWorkbook.Sheets.get_Item(4).Activate();
                titleRow = TitleRowEllipse;

                i = titleRow + 1;
                foreach (var jt in ellipseJobTasks)
                {
                    if (jt.LabourResourcesList.Count <= 0) continue;
                    foreach (var r in jt.LabourResourcesList)
                    {
                        List<DailyJobs> singleTask = JobActions.GetEllipseSingleTask(_eFunctions, district, jt.WorkOrder ?? jt.StdJobNo, jt.WoTaskNo ?? jt.StdJobTask, jt.PlanStrDate, jt.PlanStrTime, jt.PlanFinDate, jt.PlanFinTime, startDate, endDate, r.ResourceCode);
                        foreach (var k in singleTask)
                        {
                            _cells.GetCell(1, i).Value = k.WorkGroup;                       //"Grupo"
                            _cells.GetCell(2, i).Value = jt.EquipNo;                         //"Equipo"
                            _cells.GetCell(3, i).Value = k.WorkOrder;                       //"Component"
                            _cells.GetCell(4, i).Value = k.WorkOrder;                       //"Modificador"
                            _cells.GetCell(5, i).Value = jt.ItemName1;                       //"Eq Desc"
                            _cells.GetCell(6, i).Value = jt.MaintSchTask;                    //"MST"
                            _cells.GetCell(7, i).Value = k.WorkOrder;                       //"Referencia"
                            _cells.GetCell(8, i).Value = k.WoTaskNo;                        //"Ref Desc"
                            _cells.GetCell(9, i).Value = k.WoTaskDesc;                      //"Tarea"
                            _cells.GetCell(10, i).Value = k.Shift;                           //"Turno"
                            _cells.GetCell(11, i).Value = k.PlanStrDate;                     //"Hora Inicio"
                            _cells.GetCell(12, i).Value = k.PlanFinDate;                    //"Hora Fin"
                            _cells.GetCell(13, i).Value = k.EstimatedShiftDurationsHrs;     //"Duracion Turno"
                            _cells.GetCell(14, i).Value = k.ResourceCode;                   //"Recurso"
                            _cells.GetCell(15, i).Value = k.ShiftLabourHours;               //"Horas Requeridas"
                            i++;
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewJobList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar ejecutar la funcion. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void LoadEllipse()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                //Hoja de Ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
                var opContext = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDsct,
                    _frmAuth.EllipsePost);
                var titleRow = TitleRowEllipse;

                var i = titleRow + 1;

                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value) != null)
                {
                    var resourcesToSave = new LabourResources
                    {
                        WorkGroup = _cells.GetCell(13, i).Value.ToString(),
                        Date = _cells.GetCell(14, i).Value.ToString(),
                        ResourceCode = _cells.GetCell(15, i).Value.ToString(),
                        Quantity = Convert.ToDouble(_cells.GetCell(16, i).Value),
                        EstimatedLabourHours = Convert.ToDouble(_cells.GetCell(17, i).Value),
                        AvailableLabourHours = Convert.ToDouble(_cells.GetCell(18, i).Value)
                    };
                    var reply = JobActions.UpdateEllipseResources(_eFunctions, urlService, opContext, resourcesToSave);
                    _cells.GetCell(19, i).Value = reply.Errors == null
                        ? reply.Message
                        : string.Join(",", reply.Errors.Select(p => p.ToString()).ToArray());
                    i += 1;
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewJobList()",
                    "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void LoadJobPlan()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();


                #region Hoja 1 - Tareas
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                var titleRow = TitleRowResources;

                var i = titleRow + 1;
                var tasksToSave = new List<JobTask>();
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value) != null)
                {
                    var j = new JobTask();
                    j.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString());
                    j.PlanStrDate = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value.ToString());
                    j.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value.ToString());
                    j.WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value.ToString());
                    tasksToSave.Add(j);
                    i += 1;
                }
                JobActions.SaveTasks(tasksToSave);
                #endregion

                #region Hoja 2 - Hoja de Ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                titleRow = TitleRowEllipse;

                i = titleRow + 1;
                var resourcesToSave = new List<LabourResources>();
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value) != null)
                {
                    var l = new LabourResources();
                    l.WorkGroup = _cells.GetCell(13, i).Value.ToString();
                    l.Date = _cells.GetCell(14, i).Value.ToString();
                    l.ResourceCode = _cells.GetCell(15, i).Value.ToString();
                    l.Quantity = Convert.ToDouble(_cells.GetCell(16, i).Value);
                    l.EstimatedLabourHours = Convert.ToDouble(_cells.GetCell(17, i).Value);
                    l.AvailableLabourHours = Convert.ToDouble(_cells.GetCell(18, i).Value);
                    resourcesToSave.Add(l);
                    i += 1;
                }
                JobActions.SaveResources(resourcesToSave);
                #endregion
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewJobList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar ejecutar la funcion. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }


        private void ExecuteRequirementActions()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            #region Hoja 1 - Tareas
            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDsct, _frmAuth.EllipsePost);

            var titleRow = TitleRowResources;
            var resultColumn = ResultColumnResources;

            var i = titleRow + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null)
            {

                try
                {
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(ActionColumn, i).Value);

                    var taskReq = new TaskRequirement
                    {
                        DistrictCode = _frmAuth.EllipseDsct,
                        WorkGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                        WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                        WoTaskNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                        ReqCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value),
                        UnitsQty = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value)
                    };

                    if (string.IsNullOrWhiteSpace(action))
                        continue;
                    else if (action.Equals("C"))
                    {
                        WorkOrderTaskActions.CreateTaskResource(urlService, opContext, taskReq);
                    }
                    else if (action.Equals("M"))
                    {
                        WorkOrderTaskActions.ModifyTaskResource(urlService, opContext, taskReq);
                    }
                    else if (action.Equals("D"))
                    {
                        WorkOrderTaskActions.DeleteTaskResource(urlService, opContext, taskReq);
                    }
                    else if (action.Equals("Close Task"))
                    {
                        var taskOpContext = new EllipseWorkOrdersClassLibrary.WorkOrderTaskService.OperationContext
                        {
                            district = opContext.district,
                            position = opContext.position,
                            maxInstances = opContext.maxInstances,
                            maxInstancesSpecified = opContext.maxInstancesSpecified,
                            returnWarnings = opContext.returnWarnings,
                            returnWarningsSpecified = opContext.returnWarningsSpecified
                        };
                        var woTask = new WorkOrderTask
                        {
                            DistrictCode = _frmAuth.EllipseDsct,
                            WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                            WoTaskNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                            CompletedBy = _frmAuth.EllipseUser,
                            CompletedCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value)),
                            ClosedDate = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value))
                        };

                        WorkOrderTaskActions.CompleteWorkOrderTask(urlService, taskOpContext, woTask);
                    }
                    _cells.GetCell(resultColumn, i).Value = "OK";
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;

                }
                catch (Exception ex)
                {
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteRequirementActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            #endregion
            _cells.SetCursorDefault();
        }

        private void cbDeviationStats_Click(object sender, RibbonControlEventArgs e)
        {

            Settings.CurrentSettings.SetCustomSettingValue("DeviationStats", MyUtilities.ToString(cbDeviationStats.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbSplitTaskByResource_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("SplitByResource", MyUtilities.ToString(cbSplitTaskByResource.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbOverlappingDateSearch_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("OverlappingDateSearch", MyUtilities.ToString(cbSplitTaskByResource.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }

        private void cbIncludeMsts_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.CurrentSettings.SetCustomSettingValue("IncludeMsts", MyUtilities.ToString(cbSplitTaskByResource.Checked));
            Settings.CurrentSettings.SaveCustomSettings();
        }
    }
}

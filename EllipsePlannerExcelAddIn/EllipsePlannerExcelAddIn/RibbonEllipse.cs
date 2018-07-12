using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

using Screen = EllipseCommonsClassLibrary.ScreenService;
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
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        private Application _excelApp;
        private readonly FormAuthenticate _frmAuth = new FormAuthenticate();
        private Thread _thread;

        //Hojas
        private const string ValidationSheetName = "Validacion";
        private const string ResourcesSheetName = "Planeados";
        private const string EllipseResourcesSheetName = "Estimados";
        private const string PeopleSoftResourcesSheetName = "PeopleSoft";

        //Tablas
        private const string TableJobResources = "JobResources";
        private const string TableEllipseResources = "EllipseResources";
        private const string TableIndicator = "Indicator";
        private const string TableDailyEllipseResources = "DailyEllipseResources";
        private const string TablePsoftResources = "PsoftResources";

        //Titulos
        private const int TitleRowResources = 8;
        private const int TitleRowEllipse = 6;

        //Columnas de Resultado
        private const int ResultColumnResources = 14;
        private const int ResultColumnEllipse = 5;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            var enviroments = Environments.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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

        private void FormatSheet()
        {

            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 4)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                //hoja de validación
                _cells.CreateNewWorksheet(ValidationSheetName);


                //hoja 1
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = ResourcesSheetName;
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

                _cells.GetCell(1, TitleRowResources).Value = "Grupo";
                _cells.GetCell(2, TitleRowResources).Value = "Equipo";
                _cells.GetCell(3, TitleRowResources).Value = "Eq Desc";
                _cells.GetCell(4, TitleRowResources).Value = "MST";
                _cells.GetCell(5, TitleRowResources).Value = "Referencia";
                _cells.GetCell(6, TitleRowResources).Value = "Ref Desc";
                _cells.GetCell(7, TitleRowResources).Value = "Tarea";
                _cells.GetCell(8, TitleRowResources).Value = "Recurso";
                _cells.GetCell(9, TitleRowResources).Value = "Horas Estimadas";
                _cells.GetCell(10, TitleRowResources).Value = "Horas Reales";
                _cells.GetCell(11, TitleRowResources).Value = "Horas restantes";
                _cells.GetCell(12, TitleRowResources).Value = "Fecha Planeada";
                _cells.GetCell(13, TitleRowResources).Value = "Accion";
                _cells.GetCell(14, TitleRowResources).Value = "Resultado";


                _cells.GetCell(13, TitleRowResources).Style = StyleConstants.TitleAction;
                _cells.SetValidationList(_cells.GetCell(13, TitleRowResources + 1), new List<string> { "M", "C", "D" });
                _cells.GetCell(13, TitleRowResources).AddComment("C: Crear Requerimiento \nM: Modificar Requerimiento \nD: Eliminar Requerimiento");

                _cells.GetRange(1, TitleRowResources, ResultColumnResources - 1, TitleRowResources).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowResources, ResultColumnResources, TitleRowResources + 1), TableJobResources);
                _cells.GetRange(1, TitleRowResources, ResultColumnResources, TitleRowResources + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(ResultColumnResources, TitleRowResources).Style = StyleConstants.TitleResult;


                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //hoja 2
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = EllipseResourcesSheetName;

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

                _cells.GetCell(1, TitleRowEllipse).Value = "Grupo";
                _cells.GetCell(2, TitleRowEllipse).Value = "Fecha";
                _cells.GetCell(3, TitleRowEllipse).Value = "Recurso";
                _cells.GetCell(4, TitleRowEllipse).Value = "Planeadas";
                _cells.GetCell(5, TitleRowEllipse).Value = "Disponibles";
                _cells.GetRange(1, TitleRowEllipse, 5, TitleRowEllipse).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowEllipse, 5, TitleRowEllipse + 1), TableEllipseResources);
                _cells.GetRange(1, TitleRowEllipse, 5, TitleRowEllipse + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(7, TitleRowEllipse).Value = "Grupo";
                _cells.GetCell(8, TitleRowEllipse).Value = "Fecha";
                _cells.GetCell(9, TitleRowEllipse).Value = "Horas Planeadas";
                _cells.GetCell(10, TitleRowEllipse).Value = "Horas Disponibles";
                _cells.GetCell(11, TitleRowEllipse).Value = "Indicador de Planeacion";
                _cells.GetRange(7, TitleRowEllipse, 11, TitleRowEllipse).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(7, TitleRowEllipse, 11, TitleRowEllipse + 1), TableIndicator);
                _cells.GetRange(7, TitleRowEllipse, 10, TitleRowEllipse + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(11, TitleRowEllipse + 1).NumberFormat = NumberFormatConstants.Percentage;

                _cells.GetCell(13, TitleRowEllipse - 1).Value = "Datos Ellipse";
                _cells.MergeCells(13, TitleRowEllipse - 1, 19, TitleRowEllipse - 1);
                _cells.GetRange(13, TitleRowEllipse - 1, 19, TitleRowEllipse - 1).Style = StyleConstants.TitleInformation;

                _cells.GetCell(13, TitleRowEllipse).Value = "Grupo";
                _cells.GetCell(14, TitleRowEllipse).Value = "fecha";
                _cells.GetCell(15, TitleRowEllipse).Value = "Recurso";
                _cells.GetCell(16, TitleRowEllipse).Value = "Cantidad";
                _cells.GetCell(17, TitleRowEllipse).Value = "Estimado";
                _cells.GetCell(18, TitleRowEllipse).Value = "Disponible";
                _cells.GetCell(19, TitleRowEllipse).Value = "resultado";
                _cells.GetRange(13, TitleRowEllipse, 18, TitleRowEllipse).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(13, TitleRowEllipse, 19, TitleRowEllipse + 1), TableDailyEllipseResources);
                _cells.GetRange(13, TitleRowEllipse, 19, TitleRowEllipse + 1).NumberFormat = NumberFormatConstants.Text;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //hoja 3
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = PeopleSoftResourcesSheetName;

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

                _cells.GetCell(1, TitleRowEllipse).Value = "Grupo";
                _cells.GetCell(2, TitleRowEllipse).Value = "Fecha";
                _cells.GetCell(3, TitleRowEllipse).Value = "Recurso";
                _cells.GetCell(4, TitleRowEllipse).Value = "Cedula";
                _cells.GetCell(5, TitleRowEllipse).Value = "Nombre";
                _cells.GetCell(6, TitleRowEllipse).Value = "Horas";
                _cells.GetRange(1, TitleRowEllipse, 6, TitleRowEllipse).Style = StyleConstants.TitleInformation;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowEllipse, 6, TitleRowEllipse + 1), TablePsoftResources);
                _cells.GetRange(1, TitleRowEllipse, 6, TitleRowEllipse + 1).NumberFormat = NumberFormatConstants.Text;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

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

                ////hoja de Tareas
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();

                var urlServicePost = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, ServiceType.PostService);
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
                var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
                var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
                var dateInclude = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
                var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
                var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);
                var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;

                _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlServicePost);
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                //consumo de servicio de msewts
                List<Jobs> ellipseJobs = JobActions.FetchJobsPost(_eFunctions, district, dateInclude, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                //consulta sobre tabla de Ellipse mso720
                List<LabourResources> ellipseResources = JobActions.GetEllipseResources(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                //consulta sobre tabla de Ellipse mso720
                List<LabourResources> pSoftResources = JobActions.GetPsoftResources(district, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                //recursos planeados ellipse agrupados por grupo/fecha/recurso
                var ellipseTotalresource = (from jobs in ellipseJobs from resources in jobs.LabourResourcesList select resources).GroupBy(l => new { l.WorkGroup, l.Date, l.ResourceCode })
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

                _cells.ClearTableRange(TableJobResources);
                var i = TitleRowResources + 1;
                foreach (var j in ellipseJobs)
                {
                    if (j.LabourResourcesList.Count > 0)
                    {
                        foreach (var r in j.LabourResourcesList)
                        {
                            _cells.GetCell(1, i).Value = j.WorkGroup;
                            _cells.GetCell(2, i).Value = j.EquipNo;
                            _cells.GetCell(3, i).Value = j.ItemName1;
                            _cells.GetCell(4, i).Value = j.MaintSchTask;
                            _cells.GetCell(5, i).Value = j.WorkOrder ?? j.StdJobNo;
                            _cells.GetCell(6, i).Value = j.WoTaskNo ?? j.StdJobTask;
                            _cells.GetCell(7, i).Value = j.WoDesc;
                            _cells.GetCell(8, i).Value = r.ResourceCode;
                            _cells.GetCell(9, i).Value = r.EstimatedLabourHours;
                            _cells.GetCell(10, i).Value = r.RealLabourHours;
                            _cells.GetCell(11, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                            _cells.GetCell(12, i).Value = j.PlanStrDate;
                            i++;
                        }
                    }
                    else
                    {
                        _cells.GetCell(1, i).Value = j.WorkGroup;
                        _cells.GetCell(2, i).Value = j.EquipNo;
                        _cells.GetCell(3, i).Value = j.ItemName1;
                        _cells.GetCell(4, i).Value = j.MaintSchTask;
                        _cells.GetCell(5, i).Value = j.WorkOrder ?? j.StdJobNo;
                        _cells.GetCell(6, i).Value = j.WoTaskNo ?? j.StdJobTask;
                        _cells.GetCell(7, i).Value = j.WoDesc;
                        _cells.GetCell(12, i).Value = j.PlanStrDate;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //hoja de ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();

                _cells.ClearTableRange(TableEllipseResources);
                _cells.ClearTableRange(TableDailyEllipseResources);
                _cells.ClearTableRange(TableIndicator);

                i = TitleRowEllipse + 1;
                foreach (var r in totalDailyResources)
                {
                    _cells.GetCell(1, i).Value = r.WorkGroup;
                    _cells.GetCell(2, i).Value = r.Date;
                    _cells.GetCell(3, i).Value = r.ResourceCode;
                    _cells.GetCell(4, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                    _cells.GetCell(5, i).Value = r.AvailableLabourHours;
                    i++;
                }

                i = TitleRowEllipse + 1;
                foreach (var r in plannerIndicator)
                {
                    _cells.GetCell(7, i).Value = r.WorkGroup;
                    _cells.GetCell(8, i).Value = r.Date;
                    _cells.GetCell(9, i).Value = r.EstimatedLabourHours - r.RealLabourHours;
                    _cells.GetCell(10, i).Value = r.AvailableLabourHours;
                    _cells.GetCell(11, i).Value = (r.EstimatedLabourHours - r.RealLabourHours) / r.AvailableLabourHours;
                    i++;
                }

                i = TitleRowEllipse + 1;
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
                    var range = _cells.GetRange(1, TitleRowEllipse, ResultColumnEllipse, TitleRowEllipse + totalDailyResources.Count);
                    chart.SetSourceData(range);

                    // Set chart properties.
                    chart.ChartType = XlChartType.xlColumnClustered;
                    chart.ChartWizard(range);
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //hoja de ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();

                _cells.ClearTableRange(TablePsoftResources);

                i = TitleRowEllipse + 1;
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
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewJobList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar ejecutar la funcion. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
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

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
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

                var i = TitleRowEllipse + 1;

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


                //Hoja de Tareas
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                var i = TitleRowResources + 1;
                var tasksToSave = new List<Jobs>();
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value.ToString()) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value.ToString()) != null &
                       _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value.ToString()) != null)
                {
                    tasksToSave.Add(new Jobs
                    {
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString()),
                        PlanStrDate = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value.ToString()),
                        WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value.ToString()),
                        WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value.ToString())
                    });
                    i += 1;
                }
                JobActions.SaveTasks(tasksToSave);


                //hoja de ellipse
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                i = TitleRowEllipse + 1;
                var resourcesToSave = new List<LabourResources>();
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value) != null)
                {
                    resourcesToSave.Add(new LabourResources
                    {
                        WorkGroup = _cells.GetCell(13, i).Value.ToString(),
                        Date = _cells.GetCell(14, i).Value.ToString(),
                        ResourceCode = _cells.GetCell(15, i).Value.ToString(),
                        Quantity = Convert.ToDouble(_cells.GetCell(16, i).Value),
                        EstimatedLabourHours = Convert.ToDouble(_cells.GetCell(17, i).Value),
                        AvailableLabourHours = Convert.ToDouble(_cells.GetCell(18, i).Value)
                    });
                    i += 1;
                }
                JobActions.SaveResources(resourcesToSave);
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

            //Hoja de jobs
            _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
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

            var i = TitleRowResources + 1;

            while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null)
            {

                try
                {
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value);
                    var taskReq = new TaskRequirement
                    {
                        DistrictCode = _frmAuth.EllipseDsct,
                        WorkGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                        WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                        WoTaskNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value),
                        ReqCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                        HrsReq = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value)
                    };

                    if (string.IsNullOrWhiteSpace(action))
                        continue;
                    else if (action.Equals("C"))
                    {
                        WorkOrderActions.CreateTaskResource(urlService, opContext, taskReq);
                    }
                    else if (action.Equals("M"))
                    {
                        WorkOrderActions.ModifyTaskResource(urlService, opContext, taskReq);
                    }
                    else if (action.Equals("D"))
                    {
                        WorkOrderActions.DeleteTaskResource(urlService, opContext, taskReq);
                    }
                    _cells.GetCell(ResultColumnResources, i).Value = "OK";
                    _cells.GetCell(ResultColumnResources, i).Style = StyleConstants.Success;

                }
                catch (Exception ex)
                {
                    _cells.GetCell(13, i).Style = StyleConstants.Error;
                    _cells.GetCell(13, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteRequirementActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            _cells.SetCursorDefault();
        }
    }
}

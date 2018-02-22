using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseJobsClassLibrary;
using EllipseJobsClassLibrary.JobsMWPService;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipsePlannerExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private Application _excelApp;
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Thread _thread;
        private const string SheetNameJobs = "SheetJobs";
        private const string ValidationSheetName = "ValidationSheetJobs";
        private const int TitleRowJobs = 8;
        private const string TableNameJobs = "Tablejobs";
        private const int ResultColumnJobs = 13;

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
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {

            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameJobs;


                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK ORDERS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                var districtList = Districts.GetDistrictList();

                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

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

                var aditionalJobsLis = new List<string> {"Backlog", "Unscheduled", "Backlog and Unscheduled", "Backlog Only", "Unscheduled Only", "Backlog and Unscheduled Only" };

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 3, false);
                _cells.SetValidationList(_cells.GetCell("B5"), aditionalJobsLis, ValidationSheetName, 4, false);

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.PlannedStart.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 5);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                //GENERAL
                _cells.GetCell(1, TitleRowJobs).Value = "workGroup";
                _cells.GetCell(2, TitleRowJobs).Value = "equipNo";
                _cells.GetCell(3, TitleRowJobs).Value = "Equip Description";
                _cells.GetCell(4, TitleRowJobs).Value = "maintSchTask";
                _cells.GetCell(5, TitleRowJobs).Value = "stdJobNo";
                _cells.GetCell(6, TitleRowJobs).Value = "workOrder";
                _cells.GetCell(7, TitleRowJobs).Value = "woDesc";
                _cells.GetCell(8, TitleRowJobs).Value = "maintType";
                _cells.GetCell(9, TitleRowJobs).Value = "woType";
                _cells.GetCell(10, TitleRowJobs).Value = "origPriority";
                _cells.GetCell(11, TitleRowJobs).Value = "originalPlannedStartDate";
                _cells.GetCell(12, TitleRowJobs).Value = "planStrDate";
                _cells.GetCell(ResultColumnJobs , TitleRowJobs).Value = "Resultado";
                _cells.GetRange(1, TitleRowJobs, 13, TitleRowJobs).Style = StyleConstants.TitleInformation;

                _cells.FormatAsTable(_cells.GetRange(1, TitleRowJobs, ResultColumnJobs, TitleRowJobs + 1), TableNameJobs);

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

        private void btnReviewJobs_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameJobs)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewJobList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewJobList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ReviewJobList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _cells.ClearTableRange(TableNameJobs);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);


                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

                var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);

                var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
                var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
                var dateInclude = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

                var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
                var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
                var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

                var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;

                List<Jobs> replySheet = JobActions.FetchJobs(_eFunctions, urlService, opSheet, district, dateInclude, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);

                var totalResource = (from jobs in replySheet from task in jobs.JobTaskList from resources in task.LabourResourcesList select resources).ToList();

                var result = totalResource
                    .GroupBy(l => l.ResourceCode)
                    .Select(cl => new LabourResources
                    {
                        ResourceCode = cl.First().ResourceCode,
                        RealLabourHours = cl.Sum(c => c.RealLabourHours),
                        EstimatedLabourHours = cl.Sum(c => c.EstimatedLabourHours)
                    }).ToList();

                var i = TitleRowJobs + 1;
                foreach (var job in replySheet)
                {
                    _cells.GetCell(1, i).Value = job.WorkGroup;
                    _cells.GetCell(2, i).Value = job.EquipNo;
                    _cells.GetCell(3, i).Value = job.ItemName1;
                    _cells.GetCell(4, i).Value = job.MaintSchTask;
                    _cells.GetCell(5, i).Value = job.StdJobNo;
                    _cells.GetCell(6, i).Value = job.WorkOrder;
                    _cells.GetCell(7, i).Value = job.WoDesc;
                    _cells.GetCell(8, i).Value = job.MaintType;
                    _cells.GetCell(9, i).Value =job.WoType;
                    _cells.GetCell(10, i).Value =job.OrigPriority;
                    _cells.GetCell(11, i).Value =job.OriginalPlannedStartDate;
                    _cells.GetCell(12, i).Value = job.PlanStrDate;
                    i++;
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewJobList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar ejecutar la funcion. " + ex.Message);
            }
        }

        private void ReviewJobListPost()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableNameJobs);


            var urlServicePost = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, ServiceType.PostService);
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlServicePost);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);

            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var dateInclude = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            var searchCriteriaKey1 =
                searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;

            var responsePost = JobActions.FetchJobsPost(_eFunctions, district, dateInclude, searchCriteriaKey1, searchCriteriaValue1, startDate, endDate);
        }
    }
}

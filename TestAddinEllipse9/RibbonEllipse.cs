using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse.Post;
using Microsoft.Office.Tools.Ribbon;
using TestAddinEllipse9.WorkOrderService;
using Application = Microsoft.Office.Interop.Excel.Application;
using FormAuthenticate = EllipseCommonsClassLibrary.FormAuthenticate;
using Debugger = EllipseCommonsClassLibrary.Debugger;
using TaskRequirement = EllipseStandardJobsClassLibrary.TaskRequirement;
using System.Web.Services.Description;
using System.Xml.Linq;
using EllipseJobsClassLibrary;
using EllipseStandardJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using SearchFieldCriteriaType = EllipseJobsClassLibrary.SearchFieldCriteriaType;

namespace TestAddinEllipse9
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string ValidationSheetName = "ValidationSheetWorkOrder";
        private Thread _thread;


        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        public void LoadSettings()
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

            var defaultConfig = new Settings.Options();
            //defaultConfig.SetOption("OptionName1", "OptionValue1");
            //defaultConfig.SetOption("OptionName2", "OptionValue2");
            //defaultConfig.SetOption("OptionName3", "OptionValue3");

            var options = settings.GetOptionsSettings(defaultConfig);

            //Setting of Configuration Options from Config File (or default)
            //var optionItem1Value = MyUtilities.IsTrue(options.GetOptionValue("OptionName1"));
            //var optionItem1Value = options.GetOptionValue("OptionName2");
            //var optionItem1Value = options.GetOptionValue("OptionName3");

            //optionItem1.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.UpdateOptionsSettings(options);
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ClientConversation.debuggingMode = true;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SetUserAuthentication("HMENDO4", "", "COMC0", "ICOR");
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                ConsultarOt();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ConsultarOt()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ClientConversation.debuggingMode = true;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SetUserAuthentication("HMENDO4", "", "COMC0", "ICOR");
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                ConsultarJobPost();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ConsultarJobPost()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void ConsultarOt()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);


            //Creación del Servicio
            var service = new WorkOrderService.WorkOrderService();
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            service.Url = urlService + "/WorkOrderService";

            //Instanciar el Contexto de Operación
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost
            };

            //Instanciar el SOAP
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Se cargan los parámetros de la solicitud
            string workOrder = "" + _cells.GetCell(1, 1).Value;

            if (string.IsNullOrWhiteSpace(workOrder))
            {
                workOrder = "KL045745";
                _cells.GetCell(1, 1).Value = workOrder;
            }

            _cells.GetCell(1, 2).Value2 = service.Url;

            var request = new WorkOrderServiceReadRequestDTO();
            request.districtCode = "ICOR";
            request.workOrder = new WorkOrderDTO();
            request.workOrder.prefix = workOrder.Substring(0, 2);
            request.workOrder.no = workOrder.Substring(2);
            request.includeTasks = false;
            request.includeTasksSpecified = true;

            //se envía la acción
            var reply = service.read(opContext, request);

            //se analiza la respuesta y se hacen las acciones pertinentes

            var i = 3;
            _cells.GetCell(1, i).Value2 = reply.workOrder.prefix + reply.workOrder.no;
            _cells.GetCell(2, i).Value2 = reply.workOrderDesc;
            _cells.GetCell(3, i).Value2 = reply.equipmentNo;
            _cells.GetCell(4, i).Value2 = reply.workOrderType;
            _cells.GetCell(5, i).Value2 = reply.maintenanceType;
            _cells.GetCell(6, i).Value2 = reply.raisedDate;
            _cells.GetCell(7, i).Value2 = reply.originatorId;



        }

        private void ConsultarJobPost(string version = "E8")
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            //Creación del Servicio
            var service = new WorkOrderService.WorkOrderService();
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            service.Url = urlService + "/WorkOrderService";

            //Instanciar el Contexto de Operación
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost
            };

            //Instanciar el SOAP
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Se cargan los parámetros de la solicitud
            var district = "ICOR";
            var dateInclude = "";//BI, UI, BU, BO, UO, UB
            var searchKey = SearchFieldCriteriaType.WorkGroup.Key;
            var searchValue = "MTOLOC";
            var startDate = "20200816";
            var endDate = "20200831";
            var searchParam = new TaskSearchParam();

            //se envía la acción
            List<JobTask> reply;
            if(version.Equals("E8"))
                reply = FetchJobsTasksPost(_eFunctions, district, dateInclude, searchKey, searchValue, startDate, endDate, searchParam);
            else
                reply = FetchJobsTasksPost9(_eFunctions, district, dateInclude, searchKey, searchValue, startDate, endDate, searchParam);
            //se analiza la respuesta y se hacen las acciones pertinentes
            var i = 3;

            foreach (var task in reply)
            {
                _cells.GetCell(1, i).Value2 = task.WorkOrder;
                _cells.GetCell(2, i).Value2 = task.WoDesc;
                _cells.GetCell(3, i).Value2 = task.EquipNo;
                _cells.GetCell(4, i).Value2 = task.WoType;
                _cells.GetCell(5, i).Value2 = task.MaintType;
                _cells.GetCell(6, i).Value2 = task.RaisedDate;
                _cells.GetCell(7, i).Value2 = task.OriginalPlannedStartDate;
                i = i + 1;
            }
            
        }

        public List<JobTask> FetchJobsTasksPost(EllipseFunctions ef, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate, TaskSearchParam searchParam)
        {
            var serviceUrl = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label, "POST");
            _cells.GetCell(1, 1).Value2 = searchCriteriaValue1;
            _cells.GetCell(2, 1).Value2 = startDate;
            _cells.GetCell(3, 1).Value2 = endDate;
            _cells.GetCell(1, 2).Value2 = serviceUrl;
            ef.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var proxy = new PostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDstrct, serviceUrl);
            var resp = proxy.InitConexion();

            var groupList = new List<string>();

            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList.Add(searchCriteriaValue1);
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList();

            switch (dateInclude)
            {
                case "Backlog":
                    dateInclude = "BI";
                    break;
                case "Unscheduled":
                    dateInclude = "UI";
                    break;
                case "Backlog and Unscheduled":
                    dateInclude = "BU";
                    break;
                case "Backlog Only":
                    dateInclude = "BO";
                    break;
                case "Unscheduled Only":
                    dateInclude = "UO";
                    break;
                case "Backlog and Unscheduled Only":
                    dateInclude = "UB";
                    break;
            }

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.workordertaskmwp.WorkOrderTaskMWPService</name>";
            requestXml = requestXml + "				<operation>tasksSearch</operation>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto uuid=\"" + proxy.ConnectionId + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<taskSearchType>T</taskSearchType>";
            requestXml = requestXml + "					<isTaskSearch>Y</isTaskSearch>";
            requestXml = requestXml + "                 <workOrderSearchMethod>EM</workOrderSearchMethod>";
            requestXml = requestXml + "					<taskDatePreset>N</taskDatePreset>";
            requestXml = requestXml + "					<taskDateIncrement>1</taskDateIncrement>";
            requestXml = requestXml + "					<taskDateIncrementUnit>D</taskDateIncrementUnit>";
            requestXml = requestXml + "					<startDate>" + startDate.Substring(4, 2) + "/" + startDate.Substring(6, 2) + "/" + startDate.Substring(0, 4) + "</startDate>";
            requestXml = requestXml + "					<finishDate>" + endDate.Substring(4, 2) + "/" + endDate.Substring(6, 2) + "/" + endDate.Substring(0, 4) + "</finishDate>";
            requestXml = requestXml + "					<allDistrictsForTasks>" + district + "</allDistrictsForTasks>";
            requestXml = requestXml + "					<workGroupsForTasks>";
            requestXml = groupList.Aggregate(requestXml, (current, @group) => current + "                        <item>" + @group + "</item>");
            requestXml = requestXml + "					</workGroupsForTasks>";
            requestXml = requestXml + "					<status>N</status>";
            requestXml = requestXml + "					<unassigned>N</unassigned>";
            requestXml = requestXml + "					<overlappingDateSearch>" + MyUtilities.ToString(searchParam.OverlappingDates, "Y") + "</overlappingDateSearch>";
            requestXml = requestXml + "					<datePreset>T</datePreset>";
            requestXml = requestXml + "					<dateIncrement>1</dateIncrement>";
            requestXml = requestXml + "					<dateIncrementUnit>D</dateIncrementUnit>";
            requestXml = requestXml + "					<dateIncludes>" + dateInclude + "</dateIncludes>";
            requestXml = requestXml + "					<allDistricts>N</allDistricts>";
            requestXml = requestXml + "					<matchOnChildren>N</matchOnChildren>";
            requestXml = requestXml + "					<includeProjectHierarchy>N</includeProjectHierarchy>";
            requestXml = requestXml + "					<includeMSTis>" + MyUtilities.ToString(searchParam.IncludeMst, "Y") + "</includeMSTis>";
            requestXml = requestXml + "					<displayMSTiTaskDetails>N</displayMSTiTaskDetails>";
            requestXml = requestXml + "					<includeEquipmentHierarchy>N</includeEquipmentHierarchy>";
            requestXml = requestXml + "					<includeSubLists>N</includeSubLists>";
            requestXml = requestXml + "					<woStatusMSearch>U</woStatusMSearch>";
            requestXml = requestXml + "					<excludeWorkOrderType>N</excludeWorkOrderType>";
            requestXml = requestXml + "					<excludeMaintenanceType>N</excludeMaintenanceType>";
            requestXml = requestXml + "					<attachedToOutage>N</attachedToOutage>";
            requestXml = requestXml + "					<includePreferedEGI>N</includePreferedEGI>";
            requestXml = requestXml + "					<resourceTotalsOnly>N</resourceTotalsOnly>";
            requestXml = requestXml + "					<resourceWorkGroupTotalsOnly>N</resourceWorkGroupTotalsOnly>";
            requestXml = requestXml + "					<resourceCrewTotalsOnly>N</resourceCrewTotalsOnly>";
            requestXml = requestXml + "					<resourceDisableAvailabilityCache>N</resourceDisableAvailabilityCache>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "				<maxInstances>1000</maxInstances>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + PostService.GetNewConnectionId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + proxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msewts</application>";
            requestXml = requestXml + "	<applicationPage>results</applicationPage>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");

            var responseDto = proxy.ExecutePostRequest(requestXml);

            if (responseDto.GotErrorMessages())
            {
                var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
                if (!errorMessage.Equals(""))
                    throw new Exception(errorMessage);
                return null;
            }

            var xElement = XDocument.Parse(responseDto.ResponseString).Root;
            if (xElement == null) return null;

            var jobTasks = xElement.Descendants("dto").Select(dto => new JobTask
            {
                AssignPerson = (string)dto.Element("assignPerson"),
                DstrctAcctCode = (string)dto.Element("dstrctAcctCode"),
                DstrctCode = (string)dto.Element("dstrctCode"),
                EquipNo = (string)dto.Element("equipNo"),
                CompCode = (string)dto.Element("compCode"),
                CompModCode = (string)dto.Element("compModCode"),
                ItemName1 = (string)dto.Element("itemName1"),
                ItemName2 = (string)dto.Element("itemName2"),
                JobId = (string)dto.Element("jobId"),
                JobParentId = (string)dto.Element("jobParentId"),
                JobType = (string)dto.Element("jobType"),
                MaintSchTask = (string)dto.Element("maintSchTask"),
                MaintType = (string)dto.Element("maintType"),
                MstReference = (string)dto.Element("mstReference"),
                OrigPriority = (string)dto.Element("origPriority"),
                OriginalPlannedStartDate = (string)dto.Element("originalPlannedStartDate"),
                PlanPriority = (string)dto.Element("planPriority"),
                PlanStrDate = (string)dto.Element("planStrDate"),
                PlanStrTime = (string)dto.Element("planStrTime"),
                PlanFinDate = (string)dto.Element("planFinDate"),
                PlanFinTime = (string)dto.Element("planFinTime"),
                EstimatedDurationsHrs = (string)dto.Element("estDurHrs"),
                RaisedDate = (string)dto.Element("raisedDate"),
                Reference = (string)dto.Element("reference"),
                StdJobNo = (string)dto.Element("stdJobNo"),
                StdJobTask = (string)dto.Element("wOTaskNo"),
                WoStatusM = (string)dto.Element("woStatusM"),
                WoStatusU = (string)dto.Element("woStatusU"),
                WoType = (string)dto.Element("woType"),
                WorkGroup = (string)dto.Element("workGroup"),
                WorkOrder = (string)dto.Element("workOrder"),
                WoDesc = (string)dto.Element("woDesc"),
                WoTaskNo = (string)dto.Element("wOTaskNo"),
                WoTaskDesc = (string)dto.Element("taskDescription")
            }).ToList();

            jobTasks = jobTasks.GroupBy(r => r.Reference).Select(f => f.First()).ToList();

            foreach (var task in jobTasks)
            {
                task.LabourResourcesList = new List<LabourResources>();
                //si es una orden de trabajo.
                if (task.WorkOrder != null)
                {
                    var reqList = WorkOrderTaskActions.FetchRequirements(ef, task.DstrctCode, task.WorkOrder, RequirementType.Labour.Key, task.WoTaskNo);

                    foreach (var req in reqList)
                    {
                        var requirement = new LabourResources
                        {
                            WorkGroup = req.WorkGroup,
                            ResourceCode = req.ReqCode,
                            Date = task.PlanStrDate,
                            EstimatedLabourHours = MyUtilities.ToDouble(req.UnitsQty, MyUtilities.ConversionConstants.DefaultNullAndEmpty),
                            RealLabourHours = MyUtilities.ToDouble(req.RealQty, MyUtilities.ConversionConstants.DefaultNullAndEmpty)
                        };
                        task.LabourResourcesList.Add(requirement);
                    }
                }
                else if (task.StdJobNo != null)
                {
                    //obtengo la lista de tareas de la orden de trabajo
                    var reqList = StandardJobActions.FetchTaskRequirements(ef, task.DstrctCode, task.WorkGroup, task.StdJobNo);

                    foreach (var req in reqList)
                    {
                        task.StdJobTask = req.SJTaskNo;
                        var requirement = new LabourResources
                        {
                            WorkGroup = req.WorkGroup,
                            ResourceCode = req.ReqCode,
                            Date = task.PlanStrDate,
                            EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDouble(req.HrsReq) : 0,
                            RealLabourHours = 0
                        };
                        if (req.ReqType == "LAB")
                            task.LabourResourcesList.Add(requirement);
                    }
                }
            }

            if (searchParam.AdditionalInformation)
            {
                foreach (var task in jobTasks)
                {
                    try
                    {
                        var taskAdd = JobActions.GetJobTaskAdditional(ef, task);
                        task.Additional = taskAdd;
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("JobActions.cs:GetJobTaskAdditional()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                        //ignored;
                    }
                }
            }
            return jobTasks;
        }

        public void CreateWoPostE9()
        {
            var serviceUrl = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label, "POST");
            var equipNo = "1000019";
            var woType = "CO";
            var mType = "CO";
            var workGroup = "MTOLOC";
            var woDesc = "delete order";

            var proxy = new PostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDstrct, serviceUrl);
            proxy.InitConexionE9();

            var requestXml = "";
            requestXml = requestXml + "{                                                      ";
            requestXml = requestXml + "  \"interaction\": {                                                 ";
            requestXml = requestXml + "    \"actions\": {                                                   ";
            requestXml = requestXml + "      \"action\": {                                                  ";
            requestXml = requestXml + "        \"name\": \"service\",                                         ";
            requestXml = requestXml + "        \"data\": {                                                  ";
            requestXml = requestXml + "          \"name\": \"com.mincom.ellipse.service.m3620.work.WorkService\",";
            requestXml = requestXml + "          \"operation\": \"create\",                                   ";
            requestXml = requestXml + "          \"className\": \"mfui.actions.tree.node::TreeNodeSubmitAction\",";
            requestXml = requestXml + "          \"returnWarnings\": \"true\",                                ";
            requestXml = requestXml + "          \"dto\": {                                                 ";
            requestXml = requestXml + "            \"displayRiskInd\": \"Y\",                                 ";
            requestXml = requestXml + "            \"copyEqpFlag\": \"Y\",                                    ";
            requestXml = requestXml + "            \"copyMatFlag\": \"Y\",                                    ";
            requestXml = requestXml + "            \"copyTasksFlg\": \"Y\",                                   ";
            requestXml = requestXml + "            \"copyResFlag\": \"Y\",                                    ";
            requestXml = requestXml + "            \"fieldReleaseEnabled\": \"N\",                            ";
            requestXml = requestXml + "            \"loggedOnEmployeeId\": \"HMENDO4\",                       ";
            requestXml = requestXml + "            \"loggedOnUser\": \"HMENDO4\",                             ";
            requestXml = requestXml + "            \"externalProcurementFlag\": \"N\",                        ";
            requestXml = requestXml + "            \"conAstSeg\": \"#\",                                      ";
            requestXml = requestXml + "            \"enabledModules\": \"$MODULE\",                           ";
            requestXml = requestXml + "            \"calculatedDurationsFlag\": \"Y\",                        ";
            requestXml = requestXml + "            \"employee\": \"HMENDO4\",                                 ";
            requestXml = requestXml + "            \"tranDate\": \"08 / 26 / 2020\",                          ";
            requestXml = requestXml + "            \"calculatedLabFlag\": \"Y\",                              ";
            requestXml = requestXml + "            \"calculatedMatFlag\": \"Y\",                              ";
            requestXml = requestXml + "            \"calculatedEquipmentFlag\": \"Y\",                        ";
            requestXml = requestXml + "            \"calculatedOtherFlag\": \"Y\",                            ";
            requestXml = requestXml + "            \"calculatedTotalFlag\": \"Y\",                            ";
            requestXml = requestXml + "            \"matRequisitionAllowed\": \"N\",                          ";
            requestXml = requestXml + "            \"accountCodeEnabled\": \"N\",                             ";
            requestXml = requestXml + "            \"unschedFlag\": \"N\",                                    ";
            requestXml = requestXml + "            \"earnCode\": \"001\",                                     ";
            requestXml = requestXml + "            \"noticeExistFlg\": \"N\",                                 ";
            requestXml = requestXml + "            \"bPIExistFlag\": \"N\",                                   ";
            requestXml = requestXml + "            \"districtCode\": \"ICOR\",                                ";
            requestXml = requestXml + "            \"equipmentRef\": \"" + equipNo + "\",                                 ";
            requestXml = requestXml + "            \"immdtInspFlg\": \"N\",                                   ";
            requestXml = requestXml + "            \"workOrderDesc\": \"" + woDesc + "\",                       ";
            requestXml = requestXml + "            \"mSSSEnabledInd\": \"N\",                                 ";
            requestXml = requestXml + "            \"workOrderType\": \"" +woType + "\",                                 ";
            requestXml = requestXml + "            \"maintenanceType\": \"" + mType+ "\",                               ";
            requestXml = requestXml + "            \"workGroup\": \"" + workGroup + "\",                                 ";
            requestXml = requestXml + "            \"originatorId\": \"HMENDO4\",                             ";
            requestXml = requestXml + "            \"msssStatusInd\": \"N\",                                  ";
            requestXml = requestXml + "            \"integUpdateSw\": \"N\",                                  ";
            requestXml = requestXml + "            \"protRequiredByDate\": \"N\",                             ";
            requestXml = requestXml + "            \"capitalWorkOrderSw\": \"N\",                             ";
            requestXml = requestXml + "            \"aFUDCSuspendFlag\": \"N\",                               ";
            requestXml = requestXml + "            \"aPTWExistsInd\": \"N\",                                  ";
            requestXml = requestXml + "            \"taskAPTWExistsInd\": \"N\",                              ";
            requestXml = requestXml + "            \"pTWExistsInd\": \"N\",                                   ";
            requestXml = requestXml + "            \"grantFunding\": \" \",                                   ";
            requestXml = requestXml + "            \"extendedText\": \"WO\",                                  ";
            requestXml = requestXml + "            \"linkType\": \" \",                                       ";
            requestXml = requestXml + "            \"offsetType\": \" \",                                     ";
            requestXml = requestXml + "            \"contractorCostsExists\": \"N\",                          ";
            requestXml = requestXml + "            \"jobW0CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW1CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW2CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW3CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW4CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW5CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW6CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW7CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW8CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"jobW9CommentExists\": \"N\",                             ";
            requestXml = requestXml + "            \"instWOFlag\": \"N\",                                     ";
            requestXml = requestXml + "            \"paperHist\": \"N\",                                      ";
            requestXml = requestXml + "            \"finalCosts\": \"N\",                                     ";
            requestXml = requestXml + "            \"completionText\": \"CW\"                                 ";
            requestXml = requestXml + "          }                                                        ";
            requestXml = requestXml + "        },                                                         ";
            requestXml = requestXml + "        \"id\": \"6AD31D10 - AB02 - 973D - 94D6 - 297D43C44DB6\"       ";
            requestXml = requestXml + "      }                                                            ";
            requestXml = requestXml + "    },                                                             ";
            requestXml = requestXml + "    \"connectionId\": \"a6057476 - fc7a - 442c - bca6 - 05b326b76262\",";
            requestXml = requestXml + "    \"application\": \"msewot\",                                       ";
            requestXml = requestXml + "    \"applicationPage\": \"create\",                                   ";
            requestXml = requestXml + "    \"transaction\": \"true\"                                          ";
            requestXml = requestXml + "  }                                                                ";
            requestXml = requestXml + "}                                                                  ";


            var response = proxy.ExecutePostRequestE9(requestXml);
        }
        public List<JobTask> FetchJobsTasksPost9(EllipseFunctions ef, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate, TaskSearchParam searchParam)
        {
            var serviceUrl = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label, "POST");
            _cells.GetCell(1, 1).Value2 = searchCriteriaValue1;
            _cells.GetCell(2, 1).Value2 = startDate;
            _cells.GetCell(3, 1).Value2 = endDate;
            _cells.GetCell(1, 2).Value2 = serviceUrl;
            ef.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var proxy = new PostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDstrct, serviceUrl);
            var resp = proxy.InitConexionE9();

            var groupList = new List<string>();

            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList.Add(searchCriteriaValue1);
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList();

            switch (dateInclude)
            {
                case "Backlog":
                    dateInclude = "BI";
                    break;
                case "Unscheduled":
                    dateInclude = "UI";
                    break;
                case "Backlog and Unscheduled":
                    dateInclude = "BU";
                    break;
                case "Backlog Only":
                    dateInclude = "BO";
                    break;
                case "Unscheduled Only":
                    dateInclude = "UO";
                    break;
                case "Backlog and Unscheduled Only":
                    dateInclude = "UB";
                    break;
            }

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.workordertaskmwp.WorkOrderTaskMWPService</name>";
            requestXml = requestXml + "				<operation>tasksSearch</operation>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto uuid=\"" + proxy.ConnectionId + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<taskSearchType>T</taskSearchType>";
            requestXml = requestXml + "					<isTaskSearch>Y</isTaskSearch>";
            requestXml = requestXml + "                 <workOrderSearchMethod>EM</workOrderSearchMethod>";
            requestXml = requestXml + "					<taskDatePreset>N</taskDatePreset>";
            requestXml = requestXml + "					<taskDateIncrement>1</taskDateIncrement>";
            requestXml = requestXml + "					<taskDateIncrementUnit>D</taskDateIncrementUnit>";
            requestXml = requestXml + "					<startDate>" + startDate.Substring(4, 2) + "/" + startDate.Substring(6, 2) + "/" + startDate.Substring(0, 4) + "</startDate>";
            requestXml = requestXml + "					<finishDate>" + endDate.Substring(4, 2) + "/" + endDate.Substring(6, 2) + "/" + endDate.Substring(0, 4) + "</finishDate>";
            requestXml = requestXml + "					<allDistrictsForTasks>" + district + "</allDistrictsForTasks>";
            requestXml = requestXml + "					<workGroupsForTasks>";
            requestXml = groupList.Aggregate(requestXml, (current, @group) => current + "                        <item>" + @group + "</item>");
            requestXml = requestXml + "					</workGroupsForTasks>";
            requestXml = requestXml + "					<status>N</status>";
            requestXml = requestXml + "					<unassigned>N</unassigned>";
            requestXml = requestXml + "					<overlappingDateSearch>" + MyUtilities.ToString(searchParam.OverlappingDates, "Y") + "</overlappingDateSearch>";
            requestXml = requestXml + "					<datePreset>T</datePreset>";
            requestXml = requestXml + "					<dateIncrement>1</dateIncrement>";
            requestXml = requestXml + "					<dateIncrementUnit>D</dateIncrementUnit>";
            requestXml = requestXml + "					<dateIncludes>" + dateInclude + "</dateIncludes>";
            requestXml = requestXml + "					<allDistricts>N</allDistricts>";
            requestXml = requestXml + "					<matchOnChildren>N</matchOnChildren>";
            requestXml = requestXml + "					<includeProjectHierarchy>N</includeProjectHierarchy>";
            requestXml = requestXml + "					<includeMSTis>" + MyUtilities.ToString(searchParam.IncludeMst, "Y") + "</includeMSTis>";
            requestXml = requestXml + "					<displayMSTiTaskDetails>N</displayMSTiTaskDetails>";
            requestXml = requestXml + "					<includeEquipmentHierarchy>N</includeEquipmentHierarchy>";
            requestXml = requestXml + "					<includeSubLists>N</includeSubLists>";
            requestXml = requestXml + "					<woStatusMSearch>U</woStatusMSearch>";
            requestXml = requestXml + "					<excludeWorkOrderType>N</excludeWorkOrderType>";
            requestXml = requestXml + "					<excludeMaintenanceType>N</excludeMaintenanceType>";
            requestXml = requestXml + "					<attachedToOutage>N</attachedToOutage>";
            requestXml = requestXml + "					<includePreferedEGI>N</includePreferedEGI>";
            requestXml = requestXml + "					<resourceTotalsOnly>N</resourceTotalsOnly>";
            requestXml = requestXml + "					<resourceWorkGroupTotalsOnly>N</resourceWorkGroupTotalsOnly>";
            requestXml = requestXml + "					<resourceCrewTotalsOnly>N</resourceCrewTotalsOnly>";
            requestXml = requestXml + "					<resourceDisableAvailabilityCache>N</resourceDisableAvailabilityCache>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "				<maxInstances>1000</maxInstances>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + PostService.GetNewConnectionId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + proxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msewts</application>";
            requestXml = requestXml + "	<applicationPage>results</applicationPage>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");

            var responseDto = proxy.ExecutePostRequest(requestXml);

            if (responseDto.GotErrorMessages())
            {
                var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
                if (!errorMessage.Equals(""))
                    throw new Exception(errorMessage);
                return null;
            }

            var xElement = XDocument.Parse(responseDto.ResponseString).Root;
            if (xElement == null) return null;

            var jobTasks = xElement.Descendants("dto").Select(dto => new JobTask
            {
                AssignPerson = (string)dto.Element("assignPerson"),
                DstrctAcctCode = (string)dto.Element("dstrctAcctCode"),
                DstrctCode = (string)dto.Element("dstrctCode"),
                EquipNo = (string)dto.Element("equipNo"),
                CompCode = (string)dto.Element("compCode"),
                CompModCode = (string)dto.Element("compModCode"),
                ItemName1 = (string)dto.Element("itemName1"),
                ItemName2 = (string)dto.Element("itemName2"),
                JobId = (string)dto.Element("jobId"),
                JobParentId = (string)dto.Element("jobParentId"),
                JobType = (string)dto.Element("jobType"),
                MaintSchTask = (string)dto.Element("maintSchTask"),
                MaintType = (string)dto.Element("maintType"),
                MstReference = (string)dto.Element("mstReference"),
                OrigPriority = (string)dto.Element("origPriority"),
                OriginalPlannedStartDate = (string)dto.Element("originalPlannedStartDate"),
                PlanPriority = (string)dto.Element("planPriority"),
                PlanStrDate = (string)dto.Element("planStrDate"),
                PlanStrTime = (string)dto.Element("planStrTime"),
                PlanFinDate = (string)dto.Element("planFinDate"),
                PlanFinTime = (string)dto.Element("planFinTime"),
                EstimatedDurationsHrs = (string)dto.Element("estDurHrs"),
                RaisedDate = (string)dto.Element("raisedDate"),
                Reference = (string)dto.Element("reference"),
                StdJobNo = (string)dto.Element("stdJobNo"),
                StdJobTask = (string)dto.Element("wOTaskNo"),
                WoStatusM = (string)dto.Element("woStatusM"),
                WoStatusU = (string)dto.Element("woStatusU"),
                WoType = (string)dto.Element("woType"),
                WorkGroup = (string)dto.Element("workGroup"),
                WorkOrder = (string)dto.Element("workOrder"),
                WoDesc = (string)dto.Element("woDesc"),
                WoTaskNo = (string)dto.Element("wOTaskNo"),
                WoTaskDesc = (string)dto.Element("taskDescription")
            }).ToList();

            jobTasks = jobTasks.GroupBy(r => r.Reference).Select(f => f.First()).ToList();

            foreach (var task in jobTasks)
            {
                task.LabourResourcesList = new List<LabourResources>();
                //si es una orden de trabajo.
                if (task.WorkOrder != null)
                {
                    var reqList = WorkOrderTaskActions.FetchRequirements(ef, task.DstrctCode, task.WorkOrder, RequirementType.Labour.Key, task.WoTaskNo);

                    foreach (var req in reqList)
                    {
                        var requirement = new LabourResources
                        {
                            WorkGroup = req.WorkGroup,
                            ResourceCode = req.ReqCode,
                            Date = task.PlanStrDate,
                            EstimatedLabourHours = MyUtilities.ToDouble(req.UnitsQty, MyUtilities.ConversionConstants.DefaultNullAndEmpty),
                            RealLabourHours = MyUtilities.ToDouble(req.RealQty, MyUtilities.ConversionConstants.DefaultNullAndEmpty)
                        };
                        task.LabourResourcesList.Add(requirement);
                    }
                }
                else if (task.StdJobNo != null)
                {
                    //obtengo la lista de tareas de la orden de trabajo
                    var reqList = StandardJobActions.FetchTaskRequirements(ef, task.DstrctCode, task.WorkGroup, task.StdJobNo);

                    foreach (var req in reqList)
                    {
                        task.StdJobTask = req.SJTaskNo;
                        var requirement = new LabourResources
                        {
                            WorkGroup = req.WorkGroup,
                            ResourceCode = req.ReqCode,
                            Date = task.PlanStrDate,
                            EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDouble(req.HrsReq) : 0,
                            RealLabourHours = 0
                        };
                        if (req.ReqType == "LAB")
                            task.LabourResourcesList.Add(requirement);
                    }
                }
            }

            if (searchParam.AdditionalInformation)
            {
                foreach (var task in jobTasks)
                {
                    try
                    {
                        var taskAdd = JobActions.GetJobTaskAdditional(ef, task);
                        task.Additional = taskAdd;
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("JobActions.cs:GetJobTaskAdditional()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                        //ignored;
                    }
                }
            }
            return jobTasks;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ClientConversation.debuggingMode = true;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SetUserAuthentication("HMENDO4", "", "COMC0", "ICOR");
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                ConsultarJobPost("E9");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ConsultarJobPost()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ClientConversation.debuggingMode = true;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SetUserAuthentication("HMENDO4", "", "COMC0", "ICOR");
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                CreateWoPostE9();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ConsultarJobPost()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
    }
}

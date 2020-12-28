using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using System.Xml.Linq;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;
using EllipseStandardJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using TaskRequirement = EllipseStandardJobsClassLibrary.TaskRequirement;
using SharedClassLibrary.Ellipse.Connections;
using Debugger = SharedClassLibrary.Utilities.Debugger;

// ReSharper disable LoopCanBeConvertedToQuery

namespace EllipseJobsClassLibrary
{
    public static class JobActions
    {
        public static List<Jobs> FetchJobs(string urlService, JobsMWPService.OperationContext opContext, JobSearchParam searchParam)
        {
            var jobList = new List<Jobs>();

            var service = new JobsMWPService.JobsMWPService();
            service.Url = urlService + "/JobsMWPService";
            var jobDto = new JobsMWPService.JobsMWPDTO();

            switch (searchParam.DateIncludes)
            {
                case "Backlog":
                    searchParam.DateIncludes = "BI";
                    break;
                case "Unscheduled":
                    searchParam.DateIncludes = "UI";
                    break;
                case "Backlog and Unscheduled":
                    searchParam.DateIncludes = "BU";
                    break;
                case "Backlog Only":
                    searchParam.DateIncludes = "BO";
                    break;
                case "Unscheduled Only":
                    searchParam.DateIncludes = "UO";
                    break;
                case "Backlog and Unscheduled Only":
                    searchParam.DateIncludes = "UB";
                    break;
            }

            switch (searchParam.SearchEntity)
            {
                case "Work Orders Only":
                    searchParam.SearchEntity = "W";
                    break;
                case "MST Forecast Only":
                    searchParam.SearchEntity = "M";
                    break;
                case "Work Orders and MST Forecast":
                    searchParam.SearchEntity = "A";
                    break;
            }

            var searchParamDto = searchParam.ToDto();
            var result = service.jobsSearch(opContext, searchParamDto, jobDto);
            
            foreach (var item in result)
            {
                if(item != null && item.jobsMWPDTO != null)
                    jobList.Add(new Jobs(item.jobsMWPDTO));
            }
            return jobList;
        }
        public static List<JobTask> FetchJobsTasksPost(EllipseFunctions ef, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate, TaskSearchParam searchParam)
        {

            ef.InitiatePostConnection();

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
            requestXml = requestXml + "				<dto uuid=\"" + PostService.GetNewConnectionId() + "\" deleted=\"true\" modified=\"false\">";
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
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msewts</application>";
            requestXml = requestXml + "	<applicationPage>results</applicationPage>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");

            var responseDto = ef.ExecutePostRequest(requestXml);

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
                        var taskAdd = GetJobTaskAdditional(ef, task);
                        task.Additional = taskAdd;
                    }
                    catch(Exception ex)
                    {
                        Debugger.LogError("JobActions.cs:GetJobTaskAdditional()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                        //ignored;
                    }
                }
            }
            return jobTasks;
        }

        public static List<LabourResources> GetEllipseResources(EllipseFunctions ef, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var sqlQuery = Queries.GetEllipseResourcesQuery(ef.DbReference, ef.DbLink, district, primakeryKey, primaryValue, startDate, endDate);
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<LabourResources>();

            if (drResources == null || drResources.IsClosed) return list;
            while (drResources.Read())
            {
                var res = new LabourResources
                {
                    WorkGroup = drResources["GRUPO"].ToString().Trim(),
                    ResourceCode = drResources["RECURSO"].ToString().Trim(),
                    Date = drResources["FECHA"].ToString().Trim(),
                    Quantity = !string.IsNullOrEmpty(drResources["CANTIDAD"].ToString().Trim()) ? Convert.ToDouble(drResources["CANTIDAD"].ToString().Trim()) : 0,
                    AvailableLabourHours = !string.IsNullOrEmpty(drResources["HORAS"].ToString().Trim()) ? Convert.ToDouble(drResources["HORAS"].ToString().Trim()) : 0
                };
                list.Add(res);
            }

            return list;
        }


        public static List<LabourResources> GetPsoftResources(string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var conn = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo));
            try
            {
                var sqlQuery = Queries.GetPsoftResourcesQuery(conn.DbReference, conn.DbLink, district, primakeryKey, primaryValue, startDate, endDate);
                var drResources = conn.GetQueryResult(sqlQuery);
                var list = new List<LabourResources>();

                if (drResources == null || drResources.IsClosed) return list;
                while (drResources.Read())
                {
                    var res = new LabourResources
                    {
                        WorkGroup = drResources["GRUPO"].ToString().Trim(),
                        Date = drResources["FECHA"].ToString().Trim(),
                        ResourceCode = drResources["RECURSO"].ToString().Trim(),
                        EmployeeId = drResources["CEDULA"].ToString().Trim(),
                        EmployeeName = drResources["NOMBRE"].ToString().Trim(),
                        AvailableLabourHours = !string.IsNullOrEmpty(drResources["HORAS"].ToString().Trim()) ? Convert.ToDouble(drResources["HORAS"].ToString().Trim()) : 0
                    };
                    list.Add(res);
                }
                conn.CloseConnection(true);

                return list;
            }
            catch
            {
                conn.CloseConnection(true);
                throw;
            }
        }

        public static JobTaskAdditional GetJobTaskAdditional(EllipseFunctions eFunctions, JobTask task)
        {
            var sqlQuery = Queries.GetJobTaskAdditionalQuery(eFunctions.DbReference, eFunctions.DbLink, task);
            var drConn = eFunctions.GetQueryResult(sqlQuery);


            if (drConn == null || drConn.IsClosed) return null;
            drConn.Read();

            
            var taskAdd = new JobTaskAdditional()
            {
                DistrictCode= drConn["DSTRCT_CODE"].ToString().Trim(),
                WorkOrder= drConn["WORK_GROUP"].ToString().Trim(),
                TaskNo= drConn["WO_TASK_NO"].ToString().Trim(),
                EquipNo= drConn["EQUIP_NO"].ToString().Trim(),
                CompCode= drConn["COMP_CODE"].ToString().Trim(),
                CompModCode= drConn["COMP_MOD_CODE"].ToString().Trim(),
                MaintScheduleTask= drConn["MAINT_SCH_TASK"].ToString().Trim(),
                StandardJobNo= drConn["STD_JOB_NO"].ToString().Trim(),
                PlanStartDate= drConn["PLAN_STR_DATE"].ToString().Trim(),
                OriginalSchedDate= drConn["ORIG_SCHED_DATE"].ToString().Trim(),
                RequiredStartDate= drConn["REQ_START_DATE"].ToString().Trim(),
                RequiredByDate= drConn["REQ_BY_DATE"].ToString().Trim(),
                CompletedCode= drConn["COMPLETED_CODE"].ToString().Trim(),
                WorkOrderAssignPerson = drConn["WO_ASSIGN_PERSON"].ToString().Trim(),
                AssignPerson = drConn["ASSIGN_PERSON"].ToString().Trim(),
                MaintenanceType= drConn["MAINT_TYPE"].ToString().Trim(),
                JobDescCode = drConn["JOB_DESC_CODE"].ToString().Trim(),
                WorkOrderType = drConn["WO_TYPE"].ToString().Trim(),
                EquipPrimaryStatType= drConn["EQ_STAT_TYPE_PR"].ToString().Trim(),
                ScheduleStatValue= drConn["SCHED_STAT_VALUE"].ToString().Trim(),
                ActualStatValue= drConn["ACTUAL_STAT_VALUE"].ToString().Trim(),
                MinSchedDate= drConn["MIN_SCHED_DT"].ToString().Trim(),
                MaxSchedDate= drConn["MAX_SCHED_DT"].ToString().Trim(),
                MinSchedStat= drConn["MIN_SCH_STAT"].ToString().Trim(),
                MaxSchedStat= drConn["MAX_SCH_STAT"].ToString().Trim(),
            };

            return taskAdd;
        }
        public static List<DailyJobs> GetEllipseSingleTask(EllipseFunctions ef, string district, string reference, string referenceTask, string referenceStartDate, string referenceStartHour, string referenceFinDate, string referenceFinHour, string startDate, string finDate, string resourceCode)
        {
            var sqlQuery = Queries.GetEllipseSingleTaskQuery(ef.DbReference, ef.DbLink, district, reference, referenceTask, referenceStartDate, referenceStartHour, referenceFinDate, referenceFinHour, startDate, finDate, resourceCode);
            
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<DailyJobs>();

            if (drResources == null || drResources.IsClosed) return list;
            while (drResources.Read())
            {
                var res = new DailyJobs()
                {
                    WorkGroup = drResources["WORK_GROUP"].ToString().Trim(),
                    WorkOrder = drResources["WORK_ORDER"].ToString().Trim(),
                    WoTaskNo = drResources["WO_TASK_NO"].ToString().Trim(),
                    WoTaskDesc = drResources["WO_TASK_DESC"].ToString().Trim(),
                    Shift = drResources["SHIFT"].ToString().Trim(),
                    PlanStrDate = drResources["PLAN_STR_DATE"].ToString().Trim(),
                    PlanFinDate = drResources["PLAN_FIN_DATE"].ToString().Trim(),
                    EstimatedDurationsHrs = drResources["TSK_DUR_HOURS"].ToString().Trim(),
                    EstimatedShiftDurationsHrs = drResources["SHIFT_TSK_DUR_HOURS"].ToString().Trim(),
                    ResourceCode = drResources["RES_CODE"].ToString().Trim(),
                    ShiftLabourHours = drResources["SHIFT_LAB_HOURS"].ToString().Trim()
                };
                list.Add(res);
            }

            return list;
        }

        public static void SaveResources(List<LabourResources> resourcesToSave)
        {
            var ef = new EllipseFunctions();
            ef.SetDBSettings(Environments.SigcorProductivo);
            foreach (var sqlQuery in resourcesToSave.Select(r => Queries.SaveResourcesQuery(ef.DbReference, r)))
            {
                ef.GetQueryResult(sqlQuery);
            }
        }

        public static void SaveTasks(List<JobTask> tasksToSave)
        {
            var ef = new EllipseFunctions();
            ef.SetDBSettings(Environments.SigcorProductivo);
            foreach (var sqlQuery in tasksToSave.Select(r => Queries.SaveTaskQuery(ef.DbReference, r)))
            {
                ef.GetQueryResult(sqlQuery);
            }
        }

        public static ReplyMessage UpdateEllipseResources(EllipseFunctions eFunctions, string urlService, Screen.OperationContext opContext, LabourResources resourcesToSave)
        {
            var proxySheet = new Screen.ScreenService { Url = urlService };
            var replyMessage = new ReplyMessage();
            var arrayFields = new ArrayScreenNameValue();


            eFunctions.RevertOperation(opContext, proxySheet);
            var replySheet = proxySheet.executeScreen(opContext, "MSO720");

            if (replySheet.mapName != "MSM720A")
                throw new Exception("NO SE PUEDE INGRESAR AL PROGRAMA MSO720");

            arrayFields.Add("OPTION1I", "3");
            arrayFields.Add("WORK_GROUP1I", resourcesToSave.WorkGroup);

            var requestSheet = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            replySheet = proxySheet.submit(opContext, requestSheet);

            if (replySheet == null)
                throw new Exception("No se pudo entrar al MSO720 Opcion 3");
            if (eFunctions.CheckReplyError(replySheet) || eFunctions.CheckReplyWarning(replySheet))
                throw new Exception(replySheet.message);
            if (replySheet.mapName != "MSM72AA")
                throw new Exception("No se pudo ingresar a la pantalla MSM72AA");

            var replyArrayFields = new ArrayScreenNameValue(replySheet.screenFields);

            var screenIndex = 1;
            while (!string.IsNullOrWhiteSpace(replyArrayFields.GetField("RES_CODE1I" + screenIndex).value))
            {
                if (screenIndex > 12)
                {
                    //enviar Screen
                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opContext, requestSheet);
                    arrayFields = new ArrayScreenNameValue();
                    //
                    if (replySheet != null && replySheet.mapName != "MSM72AA")
                        break;
                    screenIndex = 1;
                }
                if (resourcesToSave.ResourceCode == replyArrayFields.GetField("RES_CLASS1I" + screenIndex).value + replyArrayFields.GetField("RES_CODE1I" + screenIndex).value)
                {
                    break;
                }
                screenIndex++;
            }
            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("RES_CLASS1I" + screenIndex, resourcesToSave.ResourceCode.Substring(0, 1));
            arrayFields.Add("RES_CODE1I" + screenIndex, resourcesToSave.ResourceCode.Substring(1, 3));
            arrayFields.Add("MAND_IND1I" + screenIndex, "N");
            arrayFields.Add("REQMT_TYPE1I" + screenIndex, "E");
            arrayFields.Add("RESRC_NO1I" + screenIndex, "" + resourcesToSave.Quantity);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";
            replySheet = proxySheet.submit(opContext, requestSheet);

            eFunctions.CheckReplyWarning(replySheet);//si hay debug activo muestra el warning de lo contrario depende del proceso del OP

            if (replySheet != null && !eFunctions.CheckReplyError(replySheet) && replySheet.mapName == "MSM720A")
                replyMessage.Message = "Ok";
            if (replySheet != null && eFunctions.CheckReplyError(replySheet))
                replyMessage.Errors = new[] { replyMessage.Message };
            return replyMessage;
        }
    }

   
}

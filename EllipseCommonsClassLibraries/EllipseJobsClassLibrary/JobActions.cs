using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using System.Xml.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseStandardJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using TaskRequirement = EllipseStandardJobsClassLibrary.TaskRequirement;
using System.Diagnostics;
using System.Web.Services.Description;

// ReSharper disable LoopCanBeConvertedToQuery

namespace EllipseJobsClassLibrary
{
    public static class JobActions
    {
        public static List<Jobs> FetchJobs(string urlService, JobsMWPService.OperationContext opContext, JobSearchParam searchParam)
        {
            var jobList = new List<Jobs>();

            var service = new JobsMWPService.JobsMWPService();
            service.Url = urlService + "JobsMWPService";
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
                    searchParam.DateIncludes = "W";
                    break;
                case "MST Forecast Only":
                    searchParam.DateIncludes = "M";
                    break;
                case "Work Orders and MST Forecast":
                    searchParam.DateIncludes = "A";
                    break;
            }

            var result = service.jobsSearch(opContext, searchParam.ToDto(), jobDto);
            foreach (var item in result)
            {
                if(item != null && item.jobsMWPDTO != null)
                    jobList.Add(new Jobs(item.jobsMWPDTO));
            }
            return jobList;
        }
        public static List<JobTask> FetchJobsTasksPost(EllipseFunctions ef, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate)
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
            requestXml = requestXml + "				<dto uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
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
            requestXml = requestXml + "					<overlappingDateSearch>Y</overlappingDateSearch>";
            requestXml = requestXml + "					<datePreset>T</datePreset>";
            requestXml = requestXml + "					<dateIncrement>1</dateIncrement>";
            requestXml = requestXml + "					<dateIncrementUnit>D</dateIncrementUnit>";
            requestXml = requestXml + "					<dateIncludes>" + dateInclude + "</dateIncludes>";
            requestXml = requestXml + "					<allDistricts>N</allDistricts>";
            requestXml = requestXml + "					<matchOnChildren>N</matchOnChildren>";
            requestXml = requestXml + "					<includeProjectHierarchy>N</includeProjectHierarchy>";
            requestXml = requestXml + "					<includeMSTis>Y</includeMSTis>";
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
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
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

            var jobs = xElement.Descendants("dto").Select(dto => new JobTask
            {
                AssignPerson = (string)dto.Element("assignPerson"),
                DstrctAcctCode = (string)dto.Element("dstrctAcctCode"),
                DstrctCode = (string)dto.Element("dstrctCode"),
                EquipNo = (string)dto.Element("equipNo"),
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

            jobs = jobs.GroupBy(r => r.Reference).Select(f => f.First()).ToList();

            foreach (var job in jobs)
            {
                job.LabourResourcesList = new List<LabourResources>();
                //si es una orden de trabajo.
                if (job.WorkOrder != null)
                {
                    var reqList = WorkOrderActions.FetchTaskRequirements(ef, job.DstrctCode, job.WorkGroup, job.WorkOrder, job.WoTaskNo);

                    foreach (var requirement in from req in reqList
                                                let requirement = new LabourResources
                                                {
                                                    WorkGroup = req.WorkGroup,
                                                    ResourceCode = req.ReqCode,
                                                    Date = job.PlanStrDate,
                                                    EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDouble(req.HrsReq) : 0,
                                                    RealLabourHours = !string.IsNullOrEmpty(req.HrsReal) ? Convert.ToDouble(req.HrsReal) : 0
                                                }
                                                where req.ReqType == "LAB"
                                                select requirement)
                    {
                        job.LabourResourcesList.Add(requirement);
                    }

                }
                else if (job.StdJobNo != null)
                {
                    //obtengo la lista de tareas de la orden de trabajo
                    var reqList = StandardJobActions.FetchTaskRequirements(ef, job.DstrctCode, job.WorkGroup, job.StdJobNo);

                    foreach (var req in reqList)
                    {
                        job.StdJobTask = req.SJTaskNo;
                        var requirement = new LabourResources
                        {
                            WorkGroup = req.WorkGroup,
                            ResourceCode = req.ReqCode,
                            Date = job.PlanStrDate,
                            EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDouble(req.HrsReq) : 0,
                            RealLabourHours = 0
                        };
                        if (req.ReqType == "LAB")
                            job.LabourResourcesList.Add(requirement);
                    }
                }
            }

            return jobs;
        }

        public static List<LabourResources> GetEllipseResources(EllipseFunctions ef, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var sqlQuery = Queries.GetEllipseResourcesQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, startDate, endDate);
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<LabourResources>();

            if (drResources == null || drResources.IsClosed || !drResources.HasRows) return list;
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
            var ef = new EllipseFunctions();
            ef.SetDBSettings(Environments.SigcorProductivo);
            var sqlQuery = Queries.GetPsoftResourcesQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, startDate, endDate);
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<LabourResources>();

            if (drResources == null || drResources.IsClosed || !drResources.HasRows) return list;
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
            return list;
        }

        public static List<DailyJobs> GetEllipseSingleTask(EllipseFunctions ef, string district, string reference, string referenceTask, string referenceStartDate, string referenceStartHour, string referenceFinDate, string referenceFinHour, string startDate, string finDate, string resourceCode)
        {
            var sqlQuery = Queries.GetEllipseSingleTaskQuery(ef.dbReference, ef.dbLink, district, reference, referenceTask, referenceStartDate, referenceStartHour, referenceFinDate, referenceFinHour, startDate, finDate, resourceCode);
            
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<DailyJobs>();

            if (drResources == null || drResources.IsClosed || !drResources.HasRows) return list;
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
            foreach (var sqlQuery in resourcesToSave.Select(r => Queries.SaveResourcesQuery(ef.dbReference, r)))
            {
                ef.GetQueryResult(sqlQuery);
            }
        }

        public static void SaveTasks(List<JobTask> tasksToSave)
        {
            var ef = new EllipseFunctions();
            ef.SetDBSettings(Environments.SigcorProductivo);
            foreach (var sqlQuery in tasksToSave.Select(r => Queries.SaveTaskQuery(ef.dbReference, r)))
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

    public static class Queries
    {
        public static string GetEllipseResourcesQuery(string dbReference, string dbLink, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var groupList = new List<string>();

            if (primakeryKey == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList.Add(primaryValue);
            else if (primakeryKey == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == primaryValue).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == primaryValue).Select(g => g.Name).ToList();

            var query = "WITH CTE_DATES ( CTE_DATE ) AS ( " +
                        "    SELECT CAST(TO_DATE('" + startDate + "','YYYYMMDD') AS DATE) CTE_DATE FROM DUAL " +
                        "    UNION ALL " +
                        "    SELECT CAST( (CTE_DATE + 1) AS DATE) CTE_DATE FROM CTE_DATES WHERE TRUNC(CTE_DATE) + 1 <= TO_DATE('" + endDate + "','YYYYMMDD') " +
                        "),FECHAS AS ( " +
                        "    SELECT TO_CHAR(CTE_DATE,'YYYYMMDD') FECHA FROM CTE_DATES " +
                        ") SELECT " +
                        "    ELL.WORK_GROUP GRUPO, " +
                        "    FECHAS.FECHA FECHA, " +
                        "    ELL.RESOURCE_TYPE RECURSO, " +
                        "    ELL.REQ_RESRC_NO CANTIDAD, " +
                        "    ROUND( (TO_DATE(FECHAS.FECHA || ' ' || DEF_STOP_TIME,'YYYYMMDD HH24MISS') - TO_DATE(FECHAS.FECHA || ' ' || DEF_STR_TIME,'YYYYMMDD HH24MISS') ) * 24 * ELL.REQ_RESRC_NO * (1 - ( (WG.BDOWN_ALLOW_PC + ASSIGN_OTH_PC) / 100) ),2) HORAS " +
                        "  FROM " +
                        "    " + dbReference + ".MSF730_RESRC_REQ" + dbLink + " ELL " +
                        "    INNER JOIN " + dbReference + ".MSF720" + dbLink + " WG " +
                        "    ON ELL.WORK_GROUP = WG.WORK_GROUP, " +
                        "    FECHAS " +
                        "  WHERE " +
                        "    ELL.WORK_GROUP IN (" + groupList.Aggregate("", (current, g) => current + "'" + g + "'") + ") ";

            return query;
        }

        public static string GetPsoftResourcesQuery(string dbReference, string dbLink, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var groupList = new List<string>();

            if (primakeryKey == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList.Add(primaryValue);
            else if (primakeryKey == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == primaryValue).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == primaryValue).Select(g => g.Name).ToList();

            var query = "WITH CTE_DATES ( CTE_DATE ) AS ( " +
                        "    SELECT CAST(TO_DATE('" + startDate + "','YYYYMMDD') AS DATE) CTE_DATE FROM DUAL " +
                        "    UNION ALL " +
                        "    SELECT CAST( (CTE_DATE + 1) AS DATE) CTE_DATE FROM CTE_DATES WHERE TRUNC(CTE_DATE) + 1 <= TO_DATE('" + endDate + "','YYYYMMDD') " +
                        "),FECHAS AS ( " +
                        "    SELECT TO_CHAR(CTE_DATE,'YYYYMMDD') FECHA FROM CTE_DATES " +
                        ") SELECT " +
                        "    WE.WORK_GROUP GRUPO, " +
                        "    FECHAS.FECHA, " +
                        "    EMP.RESOURCE_TYPE RECURSO, " +
                        "    TURNOS.CEDULA, " +
                        "    TRIM(EMP.FIRST_NAME) || ' ' || TRIM(EMP.SURNAME) NOMBRE, " +
                        "    ROUND(TURNOS.HORAS,2) HORAS " +
                        "  FROM " +
                        "    " + dbReference + ".MSF810" + dbLink + " EMP " +
                        "    INNER JOIN " + dbReference + ".MSF723" + dbLink + " WE " +
                        "    ON EMP.EMPLOYEE_ID = WE.EMPLOYEE_ID " +
                        "    AND   WE.STOP_DT_REVSD = '00000000' " +
                        "    AND WE.WORK_GROUP IN (" + groupList.Aggregate("", (current, g) => current + "'" + g + "'") + ") " +
                        "    LEFT JOIN SIGMAN.ASISTENCIA TURNOS " +
                        "    ON LPAD(EMP.EMPLOYEE_ID,11,'0') = LPAD(TURNOS.CEDULA,11,'0') " +
                        "    INNER JOIN FECHAS " +
                        "    ON   TURNOS.FECHAP = FECHAS.FECHA " +
                        "  WHERE " +
                        "    TRIM(EMP.RESOURCE_TYPE) IS NOT NULL  " +
                        "    AND   TRIM(EMP.RESOURCE_TYPE) NOT IN ('SMPT','SSUP') " +
                        "ORDER BY WE.WORK_GROUP, " +
                        "    FECHAS.FECHA, " +
                        "    EMP.RESOURCE_TYPE, " +
                        "    TURNOS.CEDULA ";
            return query;
        }

        public static string SaveResourcesQuery(string dbReference, LabourResources l)
        {
            var query = "MERGE INTO SIGMDC.RECURSOS_PROGRAMACION T USING " +
                         "(SELECT " +
                         " '" + l.WorkGroup + "' GRUPO, " +
                         " '" + l.ResourceCode + "' RECURSO, " +
                         " '" + l.Date + "' FECHA, " +
                         " '" + l.EstimatedLabourHours + "' HORAS_PRO, " +
                         " '" + l.AvailableLabourHours + "' HORAS_DISPO " +
                         " FROM DUAL)S ON ( " +
                         " T.GRUPO = S.GRUPO " +
                         " AND T.RECURSO = S.RECURSO " +
                         " AND T.FECHA = S.FECHA " +
                         ") " +
                         "WHEN MATCHED THEN UPDATE SET T.HORAS_PRO = S.HORAS_PRO, T.HORAS_DISPO = S.HORAS_DISPO " +
                         "WHEN NOT MATCHED THEN INSERT(GRUPO, RECURSO, FECHA, HORAS_PRO, HORAS_DISPO) " +
                         "VALUES(S.GRUPO, S.RECURSO, S.FECHA, S.HORAS_PRO, S.HORAS_DISPO) ";

            return query;
        }

        public static string SaveTaskQuery(string dbReference, JobTask t)
        {
            var query = "MERGE INTO SIGMDC.SEG_PROGRAMACION T USING " +
                         "(SELECT " +
                         " '" + t.WorkGroup + "' WORK_GROUP, " +
                         " '" + t.PlanStrDate + "' FECHA, " +
                         " '" + t.WorkOrder + "' WORK_ORDER, " +
                         " '" + t.WoTaskNo + "' WO_TASK_NO " +
                         " FROM DUAL)S ON ( " +
                         " T.WORK_GROUP = S.WORK_GROUP " +
                         " AND T.FECHA = S.FECHA " +
                         " AND T.WORK_ORDER = S.WORK_ORDER " +
                         " AND T.WO_TASK_NO = S.WO_TASK_NO " +
                         ") " +
                         "WHEN NOT MATCHED THEN INSERT(WORK_GROUP, FECHA, WORK_ORDER, WO_TASK_NO) " +
                         "VALUES(S.WORK_GROUP, S.FECHA, S.WORK_ORDER, S.WO_TASK_NO) ";
            return query;
        }

        public static string GetEllipseSingleTaskQuery(string dbReference, string dbLink, string district, string reference, string referenceTask, string referenceStartDate, string referenceStartHour, string referenceFinDate, string referenceFinHour, string startDate, string finDate, string resourceCode)
        {
            var query = "WITH CTE_DATES ( " +
                        "     STARTDATE, " +
                        "     ENDDATE " +
                        " ) AS ( " +
                        "     SELECT " +
                        "         CAST(TO_DATE('" + startDate + " 060000','YYYYMMDD HH24MISS') AS DATE) STARTDATE, " +
                        "         CAST(TO_DATE('" + startDate + " 180000','YYYYMMDD HH24MISS') AS DATE) ENDDATE " +
                        "     FROM " +
                        "         DUAL " +
                        "     UNION ALL " +
                        "     SELECT " +
                        "         CAST( (CTE_DATES.STARTDATE + 0.5) AS DATE) STARTDATE, " +
                        "         CAST( (CTE_DATES.ENDDATE + 0.5) AS DATE) ENDDATE " +
                        "     FROM " +
                        "         CTE_DATES " +
                        "     WHERE " +
                        "         TRUNC(CTE_DATES.ENDDATE) + 0.5 <= TO_DATE('" + finDate + " 180000','YYYYMMDD HH24MISS') " +
                        " ),TASKS AS ( " +
                        "     SELECT " +
                        "         'WT' TASK_TYPE, " +
                        "         WT.DSTRCT_CODE, " +
                        "         WT.WORK_GROUP, " +
                        "         WT.WORK_ORDER, " +
                        "         WT.WO_TASK_NO, " +
                        "         WT.WO_TASK_DESC, " +
                        "         TO_DATE(WT.PLAN_STR_DATE || WT.PLAN_STR_TIME,'YYYYMMDD HH24MISS') PLAN_STR_DATE, " +
                        "         TO_DATE(WT.PLAN_STR_DATE || WT.PLAN_STR_TIME,'YYYYMMDD HH24MISS') + WT.TSK_DUR_HOURS / 24 PLAN_FIN_DATE, " +
                        "         WT.TSK_DUR_HOURS, " +
                        "         WT.CALC_LAB_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSF623 WT " +
                        "     WHERE " +
                        "         WT.DSTRCT_CODE = 'ICOR' " +
                        "         AND WT.WORK_ORDER = '" + reference + "' " +
                        "         AND WT.WO_TASK_NO = '" + referenceTask + "' " +

                        "         AND TRIM(WT.TSK_DUR_HOURS) IS NOT NULL " +
                        "         AND TRIM(WT.PLAN_STR_DATE) IS NOT NULL " +
                        "         AND TRIM(WT.PLAN_STR_DATE) <> '00000000' " +
                        "     UNION ALL " +
                        "     SELECT " +
                        "         'ST' TASK_TYPE, " +
                        "         ST.DSTRCT_CODE, " +
                        "         ST.WORK_GROUP, " +
                        "         ST.STD_JOB_NO, " +
                        "         ST.STD_JOB_TASK, " +
                        "         ST.SJ_TASK_DESC, " +
                        "         TO_DATE('" + referenceStartDate + "' || '" + referenceStartHour + "','YYYYMMDD HH24MISS') PLAN_STR_DATE, " +
                        "         TO_DATE('" + referenceStartDate + "' || '" + referenceStartHour + "','YYYYMMDD HH24MISS') + ST.TSK_DUR_HOURS / 24 PLAN_FIN_DATE, " +
                        "         ST.TSK_DUR_HOURS, " +
                        "         ST.CALC_LAB_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSF693 ST " +
                        "     WHERE " +
                        "         ST.DSTRCT_CODE = 'ICOR' " +
                        "         AND ST.STD_JOB_NO = '" + reference + "' " +
                        " ),SHIFT_TASKS AS ( " +
                        "     SELECT " +
                        "         TASKS.DSTRCT_CODE, " +
                        "         TASKS.WORK_GROUP, " +
                        "         TASKS.WORK_ORDER, " +
                        "         TASKS.WO_TASK_NO, " +
                        "         TASKS.WO_TASK_DESC, " +
                        "         CTE_DATES.STARTDATE   SHIFT, " +
                        "         CASE " +
                        "             WHEN TASKS.PLAN_STR_DATE >= CTE_DATES.STARTDATE THEN " +
                        "                 TASKS.PLAN_STR_DATE " +
                        "             ELSE " +
                        "                 CTE_DATES.STARTDATE " +
                        "         END PLAN_STR_DATE, " +
                        "         CASE " +
                        "             WHEN TASKS.PLAN_FIN_DATE <= CTE_DATES.ENDDATE THEN " +
                        "                 TASKS.PLAN_FIN_DATE " +
                        "             ELSE " +
                        "                 CTE_DATES.ENDDATE " +
                        "         END PLAN_FIN_DATE, " +
                        "         TASKS.TSK_DUR_HOURS " +
                        "     FROM " +
                        "         TASKS " +
                        "         INNER JOIN CTE_DATES " +
                        "         ON TASKS.PLAN_STR_DATE < CTE_DATES.ENDDATE " +
                        "            AND TASKS.PLAN_FIN_DATE > CTE_DATES.STARTDATE " +
                        " ),RES_REAL AS ( " +
                        "     SELECT " +
                        "         TR.DSTRCT_CODE, " +
                        "         TR.WORK_ORDER, " +
                        "         TR.WO_TASK_NO, " +
                        "         TR.RESOURCE_TYPE   RES_CODE, " +
                        "         SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSFX99 TX " +
                        "         INNER JOIN ELLIPSE.MSF900 TR " +
                        "         ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                        "            AND TR.WORK_ORDER = TX.WORK_ORDER " +
                        "            AND TR.USERNO = TX.USERNO " +
                        "            AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                        "            AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                        "            AND TR.REC900_TYPE = TX.REC900_TYPE " +
                        "            AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                        "            AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                        "            AND TR.DSTRCT_CODE = 'ICOR' " +
                        "            AND TR.WORK_ORDER = '" + reference + "' " +
                        "            AND TR.WO_TASK_NO = '" + referenceTask + "' " +
                        "            AND TR.RESOURCE_TYPE = '" + resourceCode + "' " +
                        "     GROUP BY " +
                        "         TR.DSTRCT_CODE, " +
                        "         TR.WORK_ORDER, " +
                        "         TR.WO_TASK_NO, " +
                        "         TR.RESOURCE_TYPE " +
                        " ),RES_EST AS ( " +
                        "     SELECT " +
                        "         TASKS.DSTRCT_CODE, " +
                        "         TASKS.WORK_ORDER, " +
                        "         TASKS.WO_TASK_NO, " +
                        "         RS.RESOURCE_TYPE   RES_CODE, " +
                        "         TT.TABLE_DESC      RES_DESC, " +
                        "         TO_NUMBER(RS.CREW_SIZE) QTY_REQ, " +
                        "         RS.EST_RESRCE_HRS " +
                        "     FROM " +
                        "         TASKS " +
                        "         INNER JOIN ELLIPSE.MSF735 RS " +
                        "         ON RS.KEY_735_ID LIKE 'ICOR" + reference + referenceTask + "%' " +
                        "         INNER JOIN ELLIPSE.MSF010 TT " +
                        "         ON TT.TABLE_CODE = RS.RESOURCE_TYPE " +
                        "            AND TT.TABLE_TYPE = 'TT' " +
                        "     WHERE " +
                        "         TASKS.DSTRCT_CODE = 'ICOR' " +
                        "         AND RS.RESOURCE_TYPE = '" + resourceCode + "' " +
                        "         AND RS.REC_735_TYPE IN ('WT', 'ST') " +
                        " ),TABLA_REC AS ( " +
                        "     SELECT " +
                        "         RES_EST.DSTRCT_CODE, " +
                        "         DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, " +
                        "         DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, " +
                        "         DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, " +
                        "         RES_EST.QTY_REQ, " +
                        "         DECODE(RES_EST.EST_RESRCE_HRS,NULL,0,RES_EST.EST_RESRCE_HRS) EST_RESRCE_HRS, " +
                        "         DECODE(RES_REAL.ACT_RESRCE_HRS,NULL,0,RES_REAL.ACT_RESRCE_HRS) ACT_RESRCE_HRS " +
                        "     FROM " +
                        "         RES_REAL " +
                        "         FULL JOIN RES_EST " +
                        "         ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE " +
                        "            AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER " +
                        "            AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO " +
                        "            AND RES_REAL.RES_CODE = RES_EST.RES_CODE " +
                        " )SELECT " +
                        "     SHIFT_TASKS.WORK_GROUP, " +
                        "     SHIFT_TASKS.WORK_ORDER, " +
                        "     SHIFT_TASKS.WO_TASK_NO, " +
                        "     SHIFT_TASKS.WO_TASK_DESC, " +
                        "     SHIFT_TASKS.SHIFT, " +
                        "     SHIFT_TASKS.PLAN_STR_DATE, " +
                        "     SHIFT_TASKS.PLAN_FIN_DATE, " +
                        "     SHIFT_TASKS.TSK_DUR_HOURS, " +
                        "     ROUND(24 * ( SHIFT_TASKS.PLAN_FIN_DATE - SHIFT_TASKS.PLAN_STR_DATE ),2) SHIFT_TSK_DUR_HOURS, " +
                        "     TABLA_REC.RES_CODE, " +
                        "     TABLA_REC.QTY_REQ, " +
                        "     TABLA_REC.EST_RESRCE_HRS, " +
                        "     TABLA_REC.ACT_RESRCE_HRS, " +
                        "     DECODE(SHIFT_TASKS.TSK_DUR_HOURS, 0, 0, ROUND(TABLA_REC.EST_RESRCE_HRS * ( 24 * ( SHIFT_TASKS.PLAN_FIN_DATE - SHIFT_TASKS.PLAN_STR_DATE ) / SHIFT_TASKS.TSK_DUR_HOURS ),2)) SHIFT_LAB_HOURS " +
                        " FROM " +
                        "     SHIFT_TASKS " +
                        "     INNER JOIN TABLA_REC " +
                        "     ON SHIFT_TASKS.WORK_ORDER = TABLA_REC.WORK_ORDER " +
                        "        AND SHIFT_TASKS.WO_TASK_NO = TABLA_REC.WO_TASK_NO " +
                        "        AND SHIFT_TASKS.DSTRCT_CODE = TABLA_REC.DSTRCT_CODE ";

            return query;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using System.Xml.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseStandardJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;

namespace EllipseJobsClassLibrary
{
    public class JobActions
    {
        public static List<Jobs> FetchJobsPost(EllipseFunctions ef, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate)
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

            var jobs = xElement.Descendants("dto").Select(dto => new Jobs
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
                OriginalPlannedStartDate = DateTime.ParseExact((string)dto.Element("originalPlannedStartDate"), "yyyyMMdd", CultureInfo.InvariantCulture),
                PlanPriority = (string)dto.Element("planPriority"),
                PlanStrDate = DateTime.ParseExact((string)dto.Element("planStrDate"), "yyyyMMdd", CultureInfo.InvariantCulture),
                RaisedDate = (string)dto.Element("raisedDate"),
                Reference = (string)dto.Element("reference"),
                StdJobNo = (string)dto.Element("stdJobNo"),
                StdJobTask = (string)dto.Element("woTaskNo"),
                WoDesc = (string)dto.Element("woDesc"),
                WoStatusM = (string)dto.Element("woStatusM"),
                WoStatusU = (string)dto.Element("woStatusU"),
                WoType = (string)dto.Element("woType"),
                WorkGroup = (string)dto.Element("workGroup"),
                WorkOrder = (string)dto.Element("workOrder"),
                WoTaskNo = (string)dto.Element("woTaskNo")
            }).ToList();

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
                                                    EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDecimal(req.HrsReq) : 0,
                                                    RealLabourHours = !string.IsNullOrEmpty(req.HrsReal) ? Convert.ToDecimal(req.HrsReal) : 0
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

                    foreach (var requirement in from req in reqList
                                                let requirement = new LabourResources
                                                {
                                                    WorkGroup = req.WorkGroup,
                                                    ResourceCode = req.ReqCode,
                                                    EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDecimal(req.HrsReq) : 0,
                                                    RealLabourHours = 0
                                                }
                                                where req.ReqType == "LAB"
                                                select requirement)
                    {
                        job.LabourResourcesList.Add(requirement);
                    }
                }
            }

            return jobs;
        }

        public static List<PSoftLabourDetails> GetPsoftResources(string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var ef = new EllipseFunctions();
            ef.SetDBSettings(Environments.SigcorProductivo);
            var sqlQuery = Queries.GetPsoftResourcesQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, startDate, endDate);
            var drResources = ef.GetQueryResult(sqlQuery);
            var list = new List<PSoftLabourDetails>();

            if (drResources == null || drResources.IsClosed || !drResources.HasRows) return list;
            while (drResources.Read())
            {
                var res = new PSoftLabourDetails
                {
                    WorkGroup = drResources["GRUPO"].ToString().Trim(),
                    EmployeeId = drResources["EMPLOYEE_ID"].ToString().Trim(),
                    Name = drResources["NAME"].ToString().Trim(),
                    Code = drResources["RECURSO"].ToString().Trim(),
                    Date = drResources["FECHA"].ToString().Trim(),
                    Hours = Convert.ToDecimal(drResources["HORAS"].ToString().Trim())
                };
                list.Add(res);
            }

            return list;
        }

    }

    public class Queries
    {
        public static string GetPsoftResourcesQuery(string dbReference, string dbLink, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var groupList = new List<string>();

            if (primakeryKey == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList.Add(primaryValue);
            else if (primakeryKey == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == primaryValue).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == primaryValue).Select(g => g.Name).ToList();

            var query = "";
            query = query + "SELECT ";
            query = query + "    WE.WORK_GROUP GRUPO, ";
            query = query + "    EMP.RESOURCE_TYPE RECURSO, ";
            query = query + "    EMP.EMPLOYEE_ID, ";
            query = query + "    TRIM(EMP.FIRST_NAME) || ' ' || TRIM(EMP.SURNAME) NAME, ";
            query = query + "    TO_CHAR(TURNOS.FEC_JORND, 'YYYYMMDD') FECHA, ";
            query = query + "    SUM( (TURNOS.HORAS) * (1 - ( (WG.BDOWN_ALLOW_PC + WG.ASSIGN_OTH_PC) / 100) ) ) HORAS ";
            query = query + "FROM ";
            query = query + "    " + dbReference + ".MSF810" + dbLink + " EMP ";
            query = query + "     INNER JOIN " + dbReference + ".MSF723" + dbLink + " WE ";
            query = query + "      ON EMP.EMPLOYEE_ID = WE.EMPLOYEE_ID ";
            query = query + "    INNER JOIN SIGMDC.MDC_EXPLOTACION TURNOS ";
            query = query + "    ON LPAD(EMP.EMPLOYEE_ID, 11, '0') = LPAD(TURNOS.EMPLID, 11, '0') ";
            query = query + "    INNER JOIN " + dbReference + ".MSF720" + dbLink + " WG ";
            query = query + "     ON WE.WORK_GROUP = WG.WORK_GROUP ";
            query = query + "WHERE ";
            query = query + "    WE.WORK_GROUP IN (" + groupList.Aggregate("", (current, g) => current + "'" + g + "'") + ") ";
            query = query + "    AND TO_CHAR(TURNOS.FEC_JORND, 'YYYYMMDD') BETWEEN '" + startDate + "' AND '" + endDate + "' ";
            query = query + "    AND WE.STOP_DT_REVSD = '00000000' ";
            query = query + "    AND TRIM(EMP.RESOURCE_TYPE)IS NOT NULL ";
            query = query + "    AND TRIM(EMP.RESOURCE_TYPE) <> 'SMPT' ";
            query = query + "GROUP BY WE.WORK_GROUP, EMP.RESOURCE_TYPE, EMP.EMPLOYEE_ID, TRIM(EMP.FIRST_NAME) || ' ' || TRIM(EMP.SURNAME), TO_CHAR(TURNOS.FEC_JORND,'YYYYMMDD') ";
            query = query + "ORDER BY 1,2,3,4 ";

            return query;
        }
    }
}

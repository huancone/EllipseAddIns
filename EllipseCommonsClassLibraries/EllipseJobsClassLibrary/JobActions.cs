using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using System.Xml.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseJobsClassLibrary.JobsMWPService;
using EllipseWorkOrdersClassLibrary;
using EllipseStandardJobsClassLibrary;

namespace EllipseJobsClassLibrary
{
    public class JobActions
    {
        public static List<Jobs> FetchJobs(EllipseFunctions ef, string urlService, OperationContext opSheet, string district, string dateInclude, int searchCriteriaKey1, string searchCriteriaValue1, string startDate, string endDate)
        {
            var proxyJobs = new JobsMWPService.JobsMWPService();//ejecuta las acciones del servicio
            var requestJobs = new JobsMWPSearchParam();
            var groupList = new List<string>();

            proxyJobs.Url = urlService + "/JobsMWPService";

            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList.Add(searchCriteriaValue1);
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList();

            requestJobs.workGroups = groupList.ToArray();
            requestJobs.planStrDate = Convert.ToDateTime(startDate.Substring(4, 2) + "/" + startDate.Substring(6, 2) + "/" + startDate.Substring(0, 4));
            requestJobs.planFinDate = Convert.ToDateTime(endDate.Substring(4, 2) + "/" + endDate.Substring(6, 2) + "/" + endDate.Substring(0, 4));
            requestJobs.planStrDateSpecified = true;
            requestJobs.planFinDateSpecified = true;

            switch (dateInclude)
            {
                case "Backlog":
                    requestJobs.dateIncludes = "BI";
                    break;
                case "Unscheduled":
                    requestJobs.dateIncludes = "UI";
                    break;
                case "Backlog and Unscheduled":
                    requestJobs.dateIncludes = "BU";
                    break;
                case "Backlog Only":
                    requestJobs.dateIncludes = "BO";
                    break;
                case "Unscheduled Only":
                    requestJobs.dateIncludes = "UO";
                    break;
                case "Backlog and Unscheduled Only":
                    requestJobs.dateIncludes = "UB";
                    break;
            }

            requestJobs.dstrctCode = district;
            requestJobs.overlappingDateSearch = true;
            requestJobs.datePreset = "N";
            requestJobs.dateIncrement = "1";
            requestJobs.dateIncrementUnit = "D";
            requestJobs.allDistricts = false;
            requestJobs.searchEntity = "A";
            requestJobs.matchOnChildren = true;
            requestJobs.includeProjectHierarchy = false;
            requestJobs.displaySuppressed = false;
            requestJobs.includeEquipmentHierarchy = false;
            requestJobs.includeSubLists = false;
            requestJobs.woStatusMSearch = "U";
            requestJobs.excludeWorkOrderType = false;
            requestJobs.excludeMaintenanceType = false;
            requestJobs.attachedToOutage = false;
            requestJobs.includePreferedEGI = false;
            requestJobs.resourceTotalsOnly = false;
            requestJobs.resourceWorkGroupTotalsOnly = false;
            requestJobs.resourceCrewTotalsOnly = false;
            requestJobs.resourceDisableAvailabilityCache = false;
            requestJobs.enableSuppressedWithResourceBalancing = false;
            requestJobs.retrieveResourceRequirements = false;

            var result = proxyJobs.jobsSearch(opSheet, requestJobs, new JobsMWPDTO()).ToList();

            //obtengo el pronostico de ordenes y tareas
            var jobListResult = result.Select(r => new Jobs
            {
                DstrctCode = r.jobsMWPDTO.dstrctCode,
                WorkGroup = r.jobsMWPDTO.workGroup,
                EquipNo = r.jobsMWPDTO.equipNo,
                ItemName1 = r.jobsMWPDTO.itemName1,
                MaintSchTask = r.jobsMWPDTO.maintSchTask,
                StdJobNo = r.jobsMWPDTO.stdJobNo,
                WorkOrder = r.jobsMWPDTO.workOrder,
                WoDesc = r.jobsMWPDTO.woDesc,
                MaintType = r.jobsMWPDTO.maintType,
                WoType = r.jobsMWPDTO.woType,
                OrigPriority = r.jobsMWPDTO.origPriority,
                OriginalPlannedStartDate = r.jobsMWPDTO.originalPlannedStartDate,
                PlanStrDate = r.jobsMWPDTO.planStrDate,
                JobTaskList = new List<WorkOrderTask>()
            }).ToList();

            foreach (var job in jobListResult)
            {
                //si es una orden de trabajo.
                if (job.WorkOrder != null)
                {
                    //obtengo la lista de tareas de la orden de trabajo

                    var woTaskList = WorkOrderActions.FetchWorkOrderTask(ef, job.DstrctCode, job.WorkOrder);

                    //recorro la lista de ordenes de trabajo.
                    foreach (var woTask in woTaskList)
                    {
                        if (woTask.ClosedStatus == "C") continue;
                        var task = new WorkOrderTask
                        {
                            DistrictCode = woTask.DistrictCode,
                            WorkGroup = woTask.WorkGroup,
                            WorkOrder = woTask.WorkOrder,
                            WoTaskNo = woTask.WoTaskNo,
                            WoTaskDesc = woTask.WoTaskDesc,
                            EstimatedMachHrs = woTask.EstimatedMachHrs,
                            EstimatedDurationsHrs = woTask.EstimatedDurationsHrs,
                            LabourResourcesList = new List<LabourResources>()
                        };

                        //Obtengo la lista de requerimientos de la tarea
                        var reqList = WorkOrderActions.FetchTaskRequirements(ef, task.DistrictCode, task.WorkGroup, task.WorkOrder, task.WoTaskNo);

                        //recorro la lista de Requerimientos
                        foreach (var requirement in from req in reqList
                                                    let requirement = new LabourResources
                                                    {
                                                        ResourceCode = req.ReqCode,
                                                        EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDecimal(req.HrsReq) : 0,
                                                        RealLabourHours = !string.IsNullOrEmpty(req.HrsReal) ? Convert.ToDecimal(req.HrsReal) : 0
                                                    }
                                                    where req.ReqType == "LAB"
                                                    select requirement)
                        {
                            task.LabourResourcesList.Add(requirement);
                        }

                        //Agrego cada tarea a la lista de tareas de la orden
                        job.JobTaskList.Add(task);
                    }
                }
                else if (job.StdJobNo != null)
                {
                    //obtengo la lista de tareas de la orden de trabajo
                    var stdTaskList = StandardJobActions.FetchStandardJobTask(ef, job.DstrctCode, job.WorkGroup, job.StdJobNo);

                    //recorro la lista de ordenes de trabajo.
                    foreach (var std in stdTaskList)
                    {
                        var task = new WorkOrderTask
                        {
                            DistrictCode = std.DistrictCode,
                            WorkGroup = std.WorkGroup,
                            WorkOrder = std.StandardJob,
                            WoTaskNo = std.SjTaskNo,
                            WoTaskDesc = std.SjTaskDesc,
                            EstimatedMachHrs = std.EstimatedMachHrs,
                            EstimatedDurationsHrs = std.EstimatedDurationsHrs,
                            LabourResourcesList = new List<LabourResources>()
                        };

                        //Obtengo la lista de requerimientos de la tarea
                        var reqList = WorkOrderActions.FetchTaskRequirements(ef, task.DistrictCode, task.WorkGroup, task.WorkOrder, task.WoTaskNo);

                        //recorro la lista de Requerimientos
                        foreach (var requirement in from req in reqList
                                                    let requirement = new LabourResources
                                                    {
                                                        ResourceCode = req.ReqCode,
                                                        EstimatedLabourHours = !string.IsNullOrEmpty(req.HrsReq) ? Convert.ToDecimal(req.HrsReq) : 0,
                                                        RealLabourHours = !string.IsNullOrEmpty(req.HrsReal) ? Convert.ToDecimal(req.HrsReq) : 0
                                                    }
                                                    where req.ReqType == "LAB"
                                                    select requirement)
                        {
                            task.LabourResourcesList.Add(requirement);
                        }
                        //Agrego cada tarea a la lista de tareas de la orden
                        job.JobTaskList.Add(task);
                    }
                }
            }

            return jobListResult;
        }

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
            requestXml = requestXml + "    <actions>";
            requestXml = requestXml + "        <action>";
            requestXml = requestXml + "            <name>service</name>";
            requestXml = requestXml + "            <data>";
            requestXml = requestXml + "                <name>com.mincom.ellipse.service.m8mwp.jobsmwp.JobsMWPService</name>";
            requestXml = requestXml + "                <operation>jobsSearch</operation>";
            requestXml = requestXml + "                <returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "                <dto   uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "                    <overlappingDateSearch>Y</overlappingDateSearch>";
            requestXml = requestXml + "                    <datePreset>N</datePreset>";
            requestXml = requestXml + "                    <dateIncrement>1</dateIncrement>";
            requestXml = requestXml + "                    <dateIncrementUnit>D</dateIncrementUnit>";
            requestXml = requestXml + "                    <planStrDate>" + startDate.Substring(4, 2) + "/" + startDate.Substring(6, 2) + "/" + startDate.Substring(0, 4) + "</planStrDate>";
            requestXml = requestXml + "                    <planFinDate>" + endDate.Substring(4, 2) + "/" + endDate.Substring(6, 2) + "/" + endDate.Substring(0, 4) + "</planFinDate>";
            requestXml = requestXml + "                    <dateIncludes>" + dateInclude + "</dateIncludes>";
            requestXml = requestXml + "                    <allDistricts>" + district + "</allDistricts>";
            requestXml = requestXml + "                    <searchEntity>A</searchEntity>";
            requestXml = requestXml + "                    <matchOnChildren>Y</matchOnChildren>";
            requestXml = requestXml + "                    <workGroups>";
            requestXml = groupList.Aggregate(requestXml, (current, @group) => current + "                        <item>" + @group + "</item>");
            requestXml = requestXml + "                    </workGroups>";
            requestXml = requestXml + "                    <includeProjectHierarchy>N</includeProjectHierarchy>";
            requestXml = requestXml + "                    <displaySuppressed>N</displaySuppressed>";
            requestXml = requestXml + "                    <includeEquipmentHierarchy>N</includeEquipmentHierarchy>";
            requestXml = requestXml + "                    <includeSubLists>N</includeSubLists>";
            requestXml = requestXml + "                    <woStatusMSearch>U</woStatusMSearch>";
            requestXml = requestXml + "                    <excludeWorkOrderType>N</excludeWorkOrderType>";
            requestXml = requestXml + "                    <excludeMaintenanceType>N</excludeMaintenanceType>";
            requestXml = requestXml + "                    <attachedToOutage>N</attachedToOutage>";
            requestXml = requestXml + "                    <includePreferedEGI>N</includePreferedEGI>";
            requestXml = requestXml + "                    <resourceTotalsOnly>N</resourceTotalsOnly>";
            requestXml = requestXml + "                    <resourceWorkGroupTotalsOnly>N</resourceWorkGroupTotalsOnly>";
            requestXml = requestXml + "                    <resourceCrewTotalsOnly>N</resourceCrewTotalsOnly>";
            requestXml = requestXml + "                    <resourceDisableAvailabilityCache>N</resourceDisableAvailabilityCache>";
            requestXml = requestXml + "                    <enableSuppressedWithResourceBalancing>N</enableSuppressedWithResourceBalancing>";
            requestXml = requestXml + "                    <retrieveResourceRequirements>N</retrieveResourceRequirements>";
            requestXml = requestXml + "               </dto>";
            requestXml = requestXml + "                <maxInstances>1000</maxInstances>";
            requestXml = requestXml + "           </data>";
            requestXml = requestXml + "            <id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "       </action>";
            requestXml = requestXml + "   </actions>";
            requestXml = requestXml + "    <chains />";
            requestXml = requestXml + "    <connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "    <application>msewjo</application>";
            requestXml = requestXml + "    <applicationPage>results</applicationPage>";
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

            var xmlDoc = XDocument.Parse(responseDto.ResponseXML.ToString());
            
            var customers = 
                from dto in xmlDoc.Descendants("interaction").Descendants("actions").Descendants("action").Descendants("data").Descendants("results")
                select new Jobs()
                {
                   
                       AssignPerson = (string)dto.Attribute("assignPerson"),
                       DstrctAcctCode = (string)dto.Attribute("dstrctAcctCode"),
                       DstrctCode = (string)dto.Attribute("dstrctCode"),
                       EquipNo = (string)dto.Attribute("equipNo"),
                       ItemName1 = (string)dto.Attribute("itemName1"),
                       ItemName2 = (string)dto.Attribute("itemName2"),
                       JobId = (string)dto.Attribute("jobId"),
                       JobParentId = (string)dto.Attribute("jobParentId"),
                       JobType = (string)dto.Attribute("jobType"),
                       MaintSchTask = (string)dto.Attribute("maintSchTask"),
                       MaintType = (string)dto.Attribute("maintType"),
                       MstReference = (string)dto.Attribute("mstReference"),
                       OrigPriority = (string)dto.Attribute("origPriority"),
                       OriginalPlannedStartDate = Convert.ToDateTime((string)dto.Attribute("originalPlannedStartDate")),
                       PlanPriority = (string)dto.Attribute("planPriority"),
                       PlanStrDate = Convert.ToDateTime((string)dto.Attribute("planStrDate")),
                       RaisedDate = (string)dto.Attribute("raisedDate"),
                       Reference = (string)dto.Attribute("reference"),
                       StdJobNo = (string)dto.Attribute("stdJobNo"),
                       WoDesc = (string)dto.Attribute("woDesc"),
                       WoStatusM = (string)dto.Attribute("woStatusM"),
                       WoStatusU = (string)dto.Attribute("woStatusU"),
                       WoType = (string)dto.Attribute("woType"),
                       WorkGroup = (string)dto.Attribute("workGroup"),
                       WorkOrder = (string)dto.Attribute("workOrder")
             };

            return customers.ToList();
        }
    }
}

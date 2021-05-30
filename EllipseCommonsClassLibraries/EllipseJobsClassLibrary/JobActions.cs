using System;
using System.Collections.Generic;
using System.Linq;
using EllipseJobsClassLibrary.WorkOrderTaskMWPService;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;
using EllipseStandardJobsClassLibrary;
using EllipseWorkOrdersClassLibrary;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary.Ellipse.Connections;
using Debugger = SharedClassLibrary.Utilities.Debugger;

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

            if (!string.IsNullOrWhiteSpace(searchParam.DateIncludes))
                searchParam.DateIncludes = searchParam.DateIncludes.ToUpper();
            switch (searchParam.DateIncludes)
            {
                case "BACKLOG":
                    searchParam.DateIncludes = "BI";
                    break;
                case "UNSCHEDULED":
                    searchParam.DateIncludes = "UI";
                    break;
                case "BACKLOG AND UNSCHEDULED":
                    searchParam.DateIncludes = "BU";
                    break;
                case "BACKLOG ONLY":
                    searchParam.DateIncludes = "BO";
                    break;
                case "UNSCHEDULED ONLY":
                    searchParam.DateIncludes = "UO";
                    break;
                case "BACKLOG AND UNSCHEDULED ONLY":
                    searchParam.DateIncludes = "UB";
                    break;
            }

            if (!string.IsNullOrWhiteSpace(searchParam.SearchEntity))
                searchParam.SearchEntity = searchParam.SearchEntity.ToUpper();
            switch (searchParam.SearchEntity)
            {
                case "WORK ORDERS ONLY":
                    searchParam.SearchEntity = "W";
                    break;
                case "MST FORECAST ONLY":
                    searchParam.SearchEntity = "M";
                    break;
                case "WORK ORDERS AND MST FORECAST":
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
        public static List<JobTask> FetchJobsTasks(EllipseFunctions ef, string urlService, WorkOrderTaskMWPService.OperationContext opContext, TaskSearchParam searchParam)
        {

            if (!string.IsNullOrWhiteSpace(searchParam.DateInclude))
                searchParam.DateInclude = searchParam.DateInclude.ToUpper();
            switch (searchParam.DateInclude)
            {
                case "BACKLOG":
                    searchParam.DateInclude = "BI";
                    break;
                case "UNSCHEDULED":
                    searchParam.DateInclude = "UI";
                    break;
                case "BACKLOG AND UNSCHEDULED":
                    searchParam.DateInclude = "BU";
                    break;
                case "BACKLOG ONLY":
                    searchParam.DateInclude = "BO";
                    break;
                case "UNSCHEDULED ONLY":
                    searchParam.DateInclude = "UO";
                    break;
                case "BACKLOG AND UNSCHEDULED ONLY":
                    searchParam.DateInclude = "UB";
                    break;
            }

            if (!string.IsNullOrWhiteSpace(searchParam.SearchEntity))
                searchParam.SearchEntity = searchParam.SearchEntity.ToUpper();
            switch (searchParam.SearchEntity)
            {
                case "WORK ORDERS ONLY":
                    searchParam.SearchEntity = "W";
                    break;
                case "MST FORECAST ONLY":
                    searchParam.SearchEntity = "M";
                    break;
                case "WORK ORDERS AND MST FORECAST":
                    searchParam.SearchEntity = "A";
                    break;
            }

            using (var taskService = new WorkOrderTaskMWPService.WorkOrderTaskMWPService())
            {
                taskService.Url = urlService + "/WorkOrderTaskMWPService";

                
                var taskSearchParams = new TasksMWPSearchParam();

                taskSearchParams.taskSearchType = "T";
                taskSearchParams.isTaskSearch = true;
                taskSearchParams.isTaskSearchSpecified = true; ;
                taskSearchParams.workOrderSearchMethod = "EM";
                taskSearchParams.taskDatePreset = "N";
                taskSearchParams.taskDateIncrement = "1";
                taskSearchParams.taskDateIncrementUnit = "D";
                if (!string.IsNullOrWhiteSpace(searchParam.StartDate))
                    taskSearchParams.startDate = MyUtilities.ToDate(searchParam.StartDate);
                taskSearchParams.startDateSpecified = !string.IsNullOrWhiteSpace(searchParam.StartDate);
                if (!string.IsNullOrWhiteSpace(searchParam.FinishDate))
                    taskSearchParams.finishDate = MyUtilities.ToDate(searchParam.FinishDate);
                taskSearchParams.finishDateSpecified = !string.IsNullOrWhiteSpace(searchParam.FinishDate);
                taskSearchParams.allDistrictsForTasks = false;
                taskSearchParams.allDistrictsForTasksSpecified = true;
                taskSearchParams.dstrctCode = searchParam.District;
                if (searchParam.WorkGroups != null && searchParam.WorkGroups.Count > 0)
                    taskSearchParams.workGroupsForTasks = searchParam.WorkGroups.ToArray();
                taskSearchParams.status = "N";
                taskSearchParams.unassigned = false;
                taskSearchParams.unassignedSpecified = true;
                taskSearchParams.overlappingDateSearch = searchParam.OverlappingDates;
                taskSearchParams.overlappingDateSearchSpecified = true;
                taskSearchParams.status = "N";
                taskSearchParams.datePreset = "T";
                taskSearchParams.dateIncrement = "1";
                taskSearchParams.dateIncrementUnit = "D";
                taskSearchParams.dateIncludes = searchParam.DateInclude;
                taskSearchParams.matchOnChildren = false;
                taskSearchParams.matchOnChildrenSpecified = true;
                taskSearchParams.includeProjectHierarchy = false;
                taskSearchParams.includeProjectHierarchySpecified = true;
                taskSearchParams.includeMSTis = searchParam.IncludeMst;
                taskSearchParams.includeMSTisSpecified = true;
                taskSearchParams.displayMSTiTaskDetails = false;
                taskSearchParams.displayMSTiTaskDetailsSpecified = true;
                taskSearchParams.includeEquipmentHierarchy = false;
                taskSearchParams.includeEquipmentHierarchySpecified = true;
                taskSearchParams.includeSubLists = false;
                taskSearchParams.includeSubListsSpecified = true;
                taskSearchParams.woStatusMSearch = "U";
                taskSearchParams.excludeWorkOrderType = false;
                taskSearchParams.excludeWorkOrderTypeSpecified = true;
                taskSearchParams.excludeMaintenanceType = false;
                taskSearchParams.excludeMaintenanceTypeSpecified = true;
                taskSearchParams.attachedToOutage = false;
                taskSearchParams.attachedToOutageSpecified = true;
                taskSearchParams.includePreferedEGI = false;
                taskSearchParams.includePreferedEGISpecified = true;
                taskSearchParams.crewTotalsOnly = false;
                taskSearchParams.crewTotalsOnlySpecified = true;
                taskSearchParams.searchEntity = searchParam.SearchEntity;

                //taskSearchParams.equipmentNumber = searchParam.EquipmentNumber;
                //taskSearchParams.taskEquipmentNumber = searchParam.TaskEquipmentNumber;
                //taskSearchParams.data1732 = searchParam.EquipmentNumber;

                var restartTask = new TasksMWPDTO();
                var reply = taskService.tasksSearch(opContext, taskSearchParams, restartTask);


                if (reply == null)
                    throw new Exception("TaskSearch Error. Couldn't receive reply from service.");
                var errorMessages = "";

                var jobTasks = new List<JobTask>();
                foreach (var item in reply)
                {
                    if (item.errors != null)
                        errorMessages = item.errors.Aggregate(errorMessages, (current, err) => current + (err.messageId + ": " + err.messageText + "\n"));

                    if (!string.IsNullOrWhiteSpace(errorMessages))
                        throw new Exception(errorMessages);

                    var task = new JobTask(item.tasksMWPDTO);


                    jobTasks.Add(task);
                }
                if (!string.IsNullOrWhiteSpace(errorMessages))
                    throw new Exception(errorMessages);

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
                                EstimatedLabourHours = MyUtilities.ToDouble(req.UnitsQty, IxConversionConstant.DefaultNullAndEmpty),
                                RealLabourHours = MyUtilities.ToDouble(req.RealQty, IxConversionConstant.DefaultNullAndEmpty)
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

                if (!searchParam.AdditionalInformation)
                    return jobTasks;


                foreach (var task in jobTasks)
                {
                    try
                    {
                        var taskAdd = GetJobTaskAdditional(ef, task);
                        task.Additional = taskAdd;
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("JobActions.cs:GetJobTaskAdditional()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                        //ignored;
                    }
                }

                return jobTasks;
            }
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


            var taskAdd = new JobTaskAdditional(drConn);

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
                var res = new DailyJobs(drResources);
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

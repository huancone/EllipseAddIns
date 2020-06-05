using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseJobsClassLibrary;
using JobsMWPService = EllipseJobsClassLibrary.JobsMWPService;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public static class PlannerActions
    {
        public static List<PlannerItem> FetchSigmanPhotoItems(EllipseFunctions ef, string district, string monitoringPeriod, string workGroup)
        {
            var sqlQuery = Queries.GetFetchSigmanPhotoQuery(ef.dbReference, ef.dbLink, district, monitoringPeriod, workGroup);
            var drItems = ef.GetQueryResult(sqlQuery);
            var list = new List<PlannerItem>();

            if (drItems == null || drItems.IsClosed || !drItems.HasRows)
                return list;
            while (drItems.Read())
            {
                var item = new PlannerItem
                {
                    MonitoringPeriod = drItems["PERIODO_MONITOREO"].ToString().Trim(),
                    WorkGroup = drItems["WORK_GROUP"].ToString().Trim(),
                    EquipNo = drItems["EQUIP_NO"].ToString().Trim(),
                    CompCode = drItems["COMPONENT_CODE"].ToString().Trim(),
                    CompModCode = drItems["MODIFIED_CODE"].ToString().Trim(),
                    //workOrder = drItems["DSTRCT_CODE"].ToString().Trim(),
                    MaintSchedTask = drItems["MAINT_SCH_TASK"].ToString().Trim(),

                    //creationDate = drItems["DSTRCT_CODE"].ToString().Trim(),
                    //planDate = drItems["DSTRCT_CODE"].ToString().Trim(),
                    NextSchedDate = drItems["NEXT_SCH_DATE"].ToString().Trim(),
                    LastPerfDate = drItems["LAST_PERF_DATE"].ToString().Trim(),

                    //DurationHours = drItems["DSTRCT_CODE"].ToString().Trim(),
                    //LaboutHors = drItems["DSTRCT_CODE"].ToString().Trim(),

                    //originatorUser = drItems["DSTRCT_CODE"].ToString().Trim(),
                    //originatorPosition = drItems["DSTRCT_CODE"].ToString().Trim(),
                    OriginatorItemDate = drItems["FECHA_FOTO"].ToString().Trim(),
                    //lastModUser = drItems["DSTRCT_CODE"].ToString().Trim(),
                    //lastModPosition = drItems["DSTRCT_CODE"].ToString().Trim(),
                    //lastModItemDate = drItems["DSTRCT_CODE"].ToString().Trim(),

                    //itemStatus = drItems["DSTRCT_CODE"].ToString().Trim(),

                };

                list.Add(item);
            }
            return list;
        }

        public static List<PlannerItem> FetchEllipsePlannerItems(string urlService, string district, string position, string startDate, string finishDate, int workGroupCriteriaKey, string workGroupCriteriaValue, string searchEntities, string additionalJobs)
        {
            var plannerList = new List<PlannerItem>();

            var opContext = new JobsMWPService.OperationContext
            {
                district = district,
                position = position,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            List<string> groupList = null;
            if (workGroupCriteriaKey == SearchFieldCriteriaType.Area.Key)
            {
                groupList = Groups.GetWorkGroupList(workGroupCriteriaValue).Select(g => g.Name).ToList(); ;
            }
            else
            {
                groupList = new List<string>();
                groupList.Add(workGroupCriteriaValue);
            }
            

            var searchParam = new JobSearchParam();
            searchParam.PlanStrDate = startDate;
            searchParam.PlanFinDate = finishDate;
            searchParam.WorkGroups = groupList.ToArray();
            searchParam.DateIncludes = additionalJobs;
            searchParam.SearchEntity = searchEntities;


            var jobList = JobActions.FetchJobs(urlService, opContext, searchParam);

            foreach (var job in jobList)
            {
                var item = new PlannerItem();
                item.WorkGroup = job.WorkGroup;
                item.WorkOrder = job.WorkOrder;
            }
            return plannerList;
        }
    }
}

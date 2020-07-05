using System;
using System.Collections.Generic;
using System.Windows.Forms.VisualStyles;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;
using EllipseEquipmentClassLibrary;
using EllipseJobsClassLibrary;
using EllipseMaintSchedTaskClassLibrary;
using JobsMWPService = EllipseJobsClassLibrary.JobsMWPService;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public static class PlannerActions
    {
        public static int DeleteSigmanTask(EllipseFunctions ef, JobSearchParam searchParam)
        {

            var sqlQuery = Queries.GetDeleteTaskQuery(ef.DbReference, ef.DbLink, searchParam.DateTypeSearch, searchParam.PlanStrDate, searchParam.PlanFinDate, searchParam.WorkGroups);

            return ef.ExecuteQuery(sqlQuery);
        }
        public static int DisableSigmanTask(EllipseFunctions ef, JobSearchParam searchParam)
        {

            var sqlQuery = Queries.GetDisableTaskStatusQuery(ef.DbReference, ef.DbLink, searchParam.DateTypeSearch, searchParam.PlanStrDate, searchParam.PlanFinDate, searchParam.WorkGroups);

            return ef.ExecuteQuery(sqlQuery);
        }
        public static int InsertItemIntoSigman(EllipseFunctions ef, PlannerItem item)
        {
            if (item != null && !string.IsNullOrWhiteSpace(item.WorkOrder))
            {
                long number1;
                if (long.TryParse(item.WorkOrder, out number1))
                    item.WorkOrder = item.WorkOrder.PadLeft(8, '0');
            }
            if (item != null && !string.IsNullOrWhiteSpace(item.EquipNo))
            {
                var eqList = EquipmentActions.GetEquipmentList(ef, item.DistrictCode, item.EquipNo);
                if (eqList != null && eqList.Count > 0)
                    item.EquipNo = eqList[0];
                else
                    throw new ArgumentException("No se ha podido encontrar en Ellipse el equipo ingresado");
            }
            var sqlQuery = Queries.InsertSigmanItemQuery(ef.DbReference, ef.DbLink, item);
            return ef.ExecuteQuery(sqlQuery);
        }
        public static List<PlannerItem> FetchSigmanPhotoItems(EllipseFunctions ef, string district, JobSearchParam searchParam)
        {
            var sqlQuery = Queries.GetFetchSigmanPhotoQuery(ef.DbReference, ef.DbLink, searchParam.SearchEntity, searchParam.DateTypeSearch, searchParam.PlanStrDate, searchParam.PlanFinDate, searchParam.WorkGroups);
            var drItems = ef.GetQueryResult(sqlQuery);
            var list = new List<PlannerItem>();

            if (drItems == null || drItems.IsClosed || !drItems.HasRows)
                return list;
            while (drItems.Read())
            {
                var item = new PlannerItem
                {
                    Period = "" + drItems["PERIOD"].ToString().Trim(),
                    WorkGroup = "" + drItems["WORK_GROUP"].ToString().Trim(),
                    EquipNo = "" + drItems["EQUIP_NO"].ToString().Trim(),
                    CompCode = "" + drItems["COMP_CODE"].ToString().Trim(),
                    CompModCode = "" + drItems["COMP_MOD_CODE"].ToString().Trim(),
                    WorkOrder = "" + drItems["WORK_ORDER"].ToString().Trim(),
                    MaintSchedTask = "" + drItems["MAINT_SCH_TASK"].ToString().Trim(),

                    RaisedDate = "" + drItems["RAISED_DATE"].ToString().Trim(),
                    PlanDate = "" + drItems["PLAN_STR_DATE"].ToString().Trim(),
                    NextSchedDate = "" + drItems["NEXT_SCH_DATE"].ToString().Trim(),
                    LastPerfDate = "" + drItems["LAST_PERF_DATE"].ToString().Trim(),

                    DurationHours = "" + drItems["DURATION_HOURS"].ToString().Trim(),
                    LabourHours = "" + drItems["LABOUR_HOURS"].ToString().Trim(),

                    LastModUser = "" + drItems["LAST_MOD_USER"].ToString().Trim(),
                    LastModItemDate = "" + drItems["LAST_MOD_DATE"].ToString().Trim(),

                    RecordStatus = "" + drItems["RECORD_STATUS"].ToString().Trim(),
                };

                list.Add(item);
            }
            return list;
        }

        public static List<PlannerItem> FetchEllipsePlannerItems(EllipseFunctions ef, string urlService, string district, string position, JobSearchParam searchParam, bool ignoreNextTask)
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




            var jobList = JobActions.FetchJobs(urlService, opContext, searchParam);

            foreach (var job in jobList)
            {
                var item = new PlannerItem();
                item.WorkGroup = job.WorkGroup;
                item.EquipNo  = "" + job.EquipNo;
                item.CompCode  = "" + job.CompCode;
                item.CompModCode  = "" + job.CompModCode;
                item.WorkOrder  = "" + job.WorkOrder;
                item.MaintSchedTask  = "" + job.MaintSchTask;
                item.Period  = "" + (!string.IsNullOrWhiteSpace(job.PlanStrDate) ? job.PlanStrDate.Substring(0, 6) : job.PlanStrDate);
                item.RaisedDate  = "" + job.RaisedDate;
                item.PlanDate  = "" + job.PlanStrDate;
                item.NextSchedDate = "";
                item.LastPerfDate  = "" + job.LastPerformedDate;
                item.DurationHours  = "" + job.EstDurHrs;
                item.LabourHours  = "" + job.EstLabHrs;
                
                var forecastSearch = new MstForecast();
                forecastSearch.CompCode = item.CompCode;
                forecastSearch.EquipNo = item.EquipNo;
                forecastSearch.MaintSchTask = item.MaintSchedTask;
                forecastSearch.CompModCode = item.CompModCode;
                forecastSearch.HideSuppressed = "Y";
                forecastSearch.Ninstances = "10";
                forecastSearch.Rec700Type = "ES";
                forecastSearch.ShowRelated = "N";

                if (!ignoreNextTask)
                {
                    try
                    {
                        var fcList = MstActions.ForecastMaintenanceScheduleTaskPost(ef, forecastSearch);
                        foreach (var mst in fcList)
                        {
                            if (MyUtilities.ToInteger(mst.PlanStrDate) > MyUtilities.ToInteger(item.PlanDate))
                            {
                                item.NextSchedDate = "" + mst.PlanStrDate;
                                item.LastPerfDate = "" + (string.IsNullOrWhiteSpace(item.LastPerfDate) ? mst.LastPerformedDate : item.LastPerfDate);
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        item.NextSchedDate = ex.Message;
                    }
                }
                else
                    item.NextSchedDate = "IGNORED DATE";

                plannerList.Add(item);
            }
            return plannerList;
        }

    }
}

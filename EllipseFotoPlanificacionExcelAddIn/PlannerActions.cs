using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseJobsClassLibrary;
using EllipseMaintSchedTaskClassLibrary;
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

            var fcOpContext = MstActions.GetMstServiceOperationContext(district, position);

            foreach (var job in jobList)
            {
                var item = new PlannerItem();
                item.WorkGroup = job.WorkGroup;
                item.EquipNo  = "" + job.EquipNo;
                item.CompCode  = "" + job.CompCode;
                item.CompModCode  = "" + job.CompModCode;
                item.WorkOrder  = "" + job.WorkOrder;
                item.MaintSchedTask  = "" + job.MaintSchTask;
                item.MonitoringPeriod  = !string.IsNullOrWhiteSpace(job.PlanStrDate) ? "" + job.PlanStrDate.Substring(0, 6) : job.PlanStrDate;
                item.CreationDate  = "" + job.RaisedDate;
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
                                item.NextSchedDate = mst.PlanStrDate;
                                item.LastPerfDate = string.IsNullOrWhiteSpace(item.LastPerfDate) ? mst.LastPerformedDate : item.LastPerfDate;
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

        public static List<PlannerItem> Test(EllipseFunctions ef, string urlService, string district, string position, string startDate, string finishDate, int workGroupCriteriaKey, string workGroupCriteriaValue, string searchEntities, string additionalJobs)
        {
            var plannerList = new List<PlannerItem>();

            var fcOpContext = MstActions.GetMstServiceOperationContext(district, position);

            var item = new PlannerItem();
            item.WorkGroup = "CTC";
            item.EquipNo = "1400000";
            item.CompCode = "";
            item.CompModCode = "";
            item.WorkOrder = "";
            item.MaintSchedTask = "CV3";
            item.MonitoringPeriod = "20200630";
            item.CreationDate = "";
            item.PlanDate = "20200630";
            item.NextSchedDate = "";
            item.LastPerfDate = "20200530";
            item.DurationHours = "31.5";
            item.LabourHours = "";

            var forecastSearch = new MstForecast();
            forecastSearch.CompCode = item.CompCode;
            forecastSearch.EquipNo = item.EquipNo;
            forecastSearch.MaintSchTask = item.MaintSchedTask;
            forecastSearch.CompModCode = item.CompModCode;
            forecastSearch.HideSuppressed = "Y";
            forecastSearch.Ninstances = "3";
            forecastSearch.Rec700Type = "ES";
            forecastSearch.ShowRelated = "N";

            try
            {
                /*var fcList = MstActions.ForecastMaintenanceScheduleTask(urlService, fcOpContext, forecastSearch);
                
                if (fcList != null && fcList.Count >= 1)
                {
                    item.NextSchedDate = fcList[0].PlanStrDate;
                    item.LastPerfDate = string.IsNullOrWhiteSpace(item.LastPerfDate) ? fcList[0].LastPerformedDate : item.LastPerfDate;
                }
                */
                
               MstActions.ForecastMaintenanceScheduleTaskPost(ef, forecastSearch);
            }
            catch (Exception ex)
            {
                item.NextSchedDate = ex.Message;
            }

            plannerList.Add(item);
            
            return plannerList;
        }
    }
}

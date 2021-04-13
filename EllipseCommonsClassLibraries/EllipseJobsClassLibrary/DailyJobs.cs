using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class DailyJobs
    {
        public string WorkGroup { get; set; }
        public string WorkOrder { get; set; }
        public string WoTaskNo { get; set; }
        public string WoTaskDesc { get; set; }
        public string Shift { get; set; }
        public string PlanStrDate { get; set; }
        public string PlanStrTime { get; set; }
        public string PlanFinDate { get; set; }
        public string PlanFinTime { get; set; }
        public string EstimatedDurationsHrs { get; set; }
        public string EstimatedShiftDurationsHrs { get; set; }
        public string ResourceCode { get; set; }
        public string QuantityRequired { get; set; }
        public string EstimatedResourceHours { get; set; }
        public string ActualResourceHours { get; set; }
        public string ShiftLabourHours { get; set; }

        public DailyJobs()
        {

        }
        public DailyJobs(IDataRecord dr)
        {
            WorkGroup = dr["WORK_GROUP"].ToString().Trim();
            WorkOrder = dr["WORK_ORDER"].ToString().Trim();
            WoTaskNo = dr["WO_TASK_NO"].ToString().Trim();
            WoTaskDesc = dr["WO_TASK_DESC"].ToString().Trim();
            Shift = dr["SHIFT"].ToString().Trim();
            PlanStrDate = dr["PLAN_STR_DATE"].ToString().Trim();
            PlanFinDate = dr["PLAN_FIN_DATE"].ToString().Trim();
            EstimatedDurationsHrs = dr["TSK_DUR_HOURS"].ToString().Trim();
            EstimatedShiftDurationsHrs = dr["SHIFT_TSK_DUR_HOURS"].ToString().Trim();
            ResourceCode = dr["RES_CODE"].ToString().Trim();
            ShiftLabourHours = dr["SHIFT_LAB_HOURS"].ToString().Trim();
        }
    }
}

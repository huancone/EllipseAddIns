using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class JobTaskAdditional
    {
        public string DistrictCode;
        public string WorkOrder;
        public string TaskNo;
        public string EquipNo;
        public string CompCode;
        public string CompModCode;
        public string MaintScheduleTask;
        public string StandardJobNo;
        public string PlanStartDate;
        public string OriginalSchedDate;
        public string RequiredStartDate;
        public string RequiredByDate;
        public string CompletedCode;
        public string WorkOrderAssignPerson;
        public string AssignPerson;
        public string MaintenanceType;
        public string WorkOrderType;
        public string JobDescCode;
        public string EquipPrimaryStatType;
        public string ScheduleStatValue;
        public string ActualStatValue;
        public string MinSchedDate;
        public string MaxSchedDate;
        public string MinSchedStat;
        public string MaxSchedStat;

        public JobTaskAdditional()
        {

        }

        public JobTaskAdditional(IDataRecord dr)
        {
            DistrictCode = dr["DSTRCT_CODE"].ToString().Trim();
            WorkOrder = dr["WORK_GROUP"].ToString().Trim();
            TaskNo = dr["WO_TASK_NO"].ToString().Trim();
            EquipNo = dr["EQUIP_NO"].ToString().Trim();
            CompCode = dr["COMP_CODE"].ToString().Trim();
            CompModCode = dr["COMP_MOD_CODE"].ToString().Trim();
            MaintScheduleTask = dr["MAINT_SCH_TASK"].ToString().Trim();
            StandardJobNo = dr["STD_JOB_NO"].ToString().Trim();
            PlanStartDate = dr["PLAN_STR_DATE"].ToString().Trim();
            OriginalSchedDate = dr["ORIG_SCHED_DATE"].ToString().Trim();
            RequiredStartDate = dr["REQ_START_DATE"].ToString().Trim();
            RequiredByDate = dr["REQ_BY_DATE"].ToString().Trim();
            CompletedCode = dr["COMPLETED_CODE"].ToString().Trim();
            WorkOrderAssignPerson = dr["WO_ASSIGN_PERSON"].ToString().Trim();
            AssignPerson = dr["ASSIGN_PERSON"].ToString().Trim();
            MaintenanceType = dr["MAINT_TYPE"].ToString().Trim();
            JobDescCode = dr["JOB_DESC_CODE"].ToString().Trim();
            WorkOrderType = dr["WO_TYPE"].ToString().Trim();
            EquipPrimaryStatType = dr["EQ_STAT_TYPE_PR"].ToString().Trim();
            ScheduleStatValue = dr["SCHED_STAT_VALUE"].ToString().Trim();
            ActualStatValue = dr["ACTUAL_STAT_VALUE"].ToString().Trim();
            MinSchedDate = dr["MIN_SCHED_DT"].ToString().Trim();
            MaxSchedDate = dr["MAX_SCHED_DT"].ToString().Trim();
            MinSchedStat = dr["MIN_SCH_STAT"].ToString().Trim();
            MaxSchedStat = dr["MAX_SCH_STAT"].ToString().Trim();
        }
    }
}

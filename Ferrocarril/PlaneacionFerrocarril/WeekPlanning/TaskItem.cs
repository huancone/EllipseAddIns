using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace PlaneacionFerrocarril
{
    public class TaskItem
    {
        public string District;
        public string StdWo;
        public string TaskNo;
        public string TaskDescription;
        public string WorkGroup;
        public string EquipNo;
        public string EquipDesc;
        public string Mst;
        public string NextSchedule;
        public string TaskStatus;
        public string ResType;
        public string ActResHours;
        public string EstResHours;

        public TaskItem()
        {

        }

        public TaskItem(IDataRecord dr)
        {
            District= dr["DSTRCT_CODE"].ToString();
            StdWo= dr["STD_WO"].ToString();
            TaskNo= dr["TASK"].ToString();
            TaskDescription= dr["DESCR"].ToString();
            WorkGroup= dr["WORK_GROUP"].ToString();
            EquipNo= dr["EQUIP_NO"].ToString();
            EquipDesc= dr["ITEM_NAME_1"].ToString();
            Mst= dr["MAINT_SCH_TASK"].ToString();
            NextSchedule= dr["NEXT_SCH"].ToString();
            TaskStatus= dr["STATUS"].ToString();
            ResType= ("" + dr["RESOURCE_TYPE"].ToString()).Trim();
            ActResHours= dr["ACT_RESRCE_HRS"].ToString();
            EstResHours= dr["EST_RESRCE_HRS"].ToString();
        }
    }
}

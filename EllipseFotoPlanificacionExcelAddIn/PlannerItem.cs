using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public class PlannerItem
    {
        public string MonitoringPeriod;
        public string WorkGroup;
        public string EquipNo;
        public string CompCode;
        public string CompModCode;
        public string WorkOrder;
        public string MaintSchedTask;

        public string CreationDate;
        public string PlanDate;
        public string NextSchedDate;
        public string LastPerfDate;

        public string DurationHours;
        public string LabourHours;

        public string OriginatorUser;
        public string OriginatorPosition;
        public string OriginatorItemDate;
        public string LastModUser;
        public string LastModPosition;
        public string LastModItemDate;

        public string ItemStatus;
    }
}

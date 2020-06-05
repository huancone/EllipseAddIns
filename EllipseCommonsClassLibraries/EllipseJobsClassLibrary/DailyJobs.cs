using System;
using System.Collections.Generic;
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
    }
}

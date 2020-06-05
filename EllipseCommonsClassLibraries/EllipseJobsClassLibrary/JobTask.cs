using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class JobTask
    {
        public string AssignPerson { get; set; }
        public string DstrctAcctCode { get; set; }
        public string DstrctCode { get; set; }
        public string EquipNo { get; set; }
        public string ItemName1 { get; set; }
        public string ItemName2 { get; set; }
        public string JobId { get; set; }
        public string JobParentId { get; set; }
        public string JobType { get; set; }
        public string MaintSchTask { get; set; }
        public string MaintType { get; set; }
        public string MstReference { get; set; }
        public string OrigPriority { get; set; }
        public string OriginalPlannedStartDate { get; set; }
        public string PlanPriority { get; set; }
        public string PlanStrDate { get; set; }
        public string PlanStrTime { get; set; }
        public string PlanFinDate { get; set; }
        public string PlanFinTime { get; set; }
        public string EstimatedDurationsHrs { get; set; }
        public string RaisedDate { get; set; }
        public string Reference { get; set; }
        public string StdJobNo { get; set; }
        public string StdJobTask { get; set; }
        public string WoDesc { get; set; }
        public string WoStatusM { get; set; }
        public string WoStatusU { get; set; }
        public string WoType { get; set; }
        public string WorkGroup { get; set; }
        public string WorkOrder { get; set; }
        public string WoTaskNo { get; set; }
        public string WoTaskDesc { get; set; }
        public string EstimatedMachHrs { get; set; }
        public string Shift { get; set; }
        public List<LabourResources> LabourResourcesList { get; set; }
    }
}

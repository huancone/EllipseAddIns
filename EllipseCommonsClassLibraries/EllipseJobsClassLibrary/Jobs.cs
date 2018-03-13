using System;
using System.Collections.Generic;

namespace EllipseJobsClassLibrary
{
    public class Jobs
    {
        public string AssignPerson { get; set; }
        public string DstrctAcctCode { get; set; }
        public string DstrctCode { get; set; }
        public string EquipNo{ get; set; }
        public string ItemName1{ get; set; }
        public string ItemName2{ get; set; }
        public string JobId{ get; set; }
        public string JobParentId{ get; set; }
        public string JobType{ get; set; }
        public string MaintSchTask{ get; set; }
        public string MaintType{ get; set; }
        public string MstReference{ get; set; }
        public string OrigPriority{ get; set; }
        public DateTime OriginalPlannedStartDate{ get; set; }
        public string PlanPriority{ get; set; }
        public DateTime PlanStrDate{ get; set; }
        public string RaisedDate{ get; set; }
        public string Reference{ get; set; }
        public string StdJobNo{ get; set; }
        public string StdJobTask { get; set; }
        public string WoDesc{ get; set; }
        public string WoStatusM{ get; set; }
        public string WoStatusU{ get; set; }
        public string WoType{ get; set; }
        public string WorkGroup{ get; set; }
        public string WorkOrder{ get; set; }
        public string WoTaskNo { get; set; }
        public string WoTaskDesc { get; set; }
        public string EstimatedMachHrs { get; set; }
        public string EstimatedDurationsHrs { get; set; }
        public List<LabourResources> LabourResourcesList { get; set; }
    }

    public class LabourResources
    {

        public string WorkGroup { get; set; }
        public string ResourceCode { get; set; }
        public decimal EstimatedLabourHours { get; set; }
        public decimal RealLabourHours { get; set; }
        public decimal AvailableLabourHours { get; set; }
    }

    public class PSoftLabourDetails
    {
        public string WorkGroup { get; set; }
        public string Code { get; set; }
        public string Date { get; set; }
        public decimal Hours { get; set; }
        public string EmployeeId { get; set; }
        public string Name { get; set; }
    }

    public class EllipseLabourDetails
    {
        public string WorkGroup { get; set; }
        public string Code { get; set; }
        public decimal Quantity { get; set; }
        public decimal Hours { get; set; }
    }
}

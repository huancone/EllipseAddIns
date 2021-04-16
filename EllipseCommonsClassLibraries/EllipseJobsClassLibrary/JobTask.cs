using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseJobsClassLibrary.WorkOrderTaskMWPService;
using System.Globalization;
using SharedClassLibrary.Utilities;

namespace EllipseJobsClassLibrary
{
    public class JobTask
    {
        public JobTaskAdditional Additional;
        public string AssignPerson { get; set; }
        public string DstrctAcctCode { get; set; }
        public string DstrctCode { get; set; }
        public string EquipNo { get; set; }
        public string CompCode { get; set; }
        public string CompModCode { get; set; }
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
        public string WoStatusMDescription { get; set; }
        public string WoStatusU { get; set; }
        public string WoStatusUDescription { get; set; }
        public string WoType { get; set; }
        public string WorkGroup { get; set; }
        public string WorkOrder { get; set; }
        public string WoTaskNo { get; set; }
        public string WoTaskDesc { get; set; }
        public string EstimatedMachHrs { get; set; }
        public string Shift { get; set; }
        public List<LabourResources> LabourResourcesList { get; set; }

        public JobTask()
        {

        }

        public JobTask(TasksMWPDTO taskMwpDto)
        {
            AssignPerson = taskMwpDto.assignPerson;
            DstrctAcctCode = taskMwpDto.dstrctAcctCode;
            DstrctCode = taskMwpDto.dstrctCode;
            EquipNo = taskMwpDto.equipNo;
            CompCode = taskMwpDto.compCode;
            CompModCode = taskMwpDto.compModCode;
            ItemName1 = taskMwpDto.itemName1;
            ItemName2 = taskMwpDto.itemName2;
            JobId = taskMwpDto.jobId;
            JobParentId = taskMwpDto.jobParentId;
            JobType = taskMwpDto.jobType;
            MaintSchTask = taskMwpDto.maintSchTask;
            MaintType = taskMwpDto.maintType;
            MstReference = taskMwpDto.mstReference;
            OrigPriority = taskMwpDto.origPriority;
            OriginalPlannedStartDate = MyUtilities.ToString(taskMwpDto.originalPlannedStartDate);
            PlanPriority = taskMwpDto.planPriority;
            PlanStrDate = MyUtilities.ToString(taskMwpDto.planStrDate);
            PlanStrTime = taskMwpDto.planStrTime;
            PlanFinDate = MyUtilities.ToString(taskMwpDto.planFinDate);
            PlanFinTime = taskMwpDto.planFinTime;
            EstimatedDurationsHrs = taskMwpDto.estDurHrs.ToString(CultureInfo.InvariantCulture);
            RaisedDate = MyUtilities.ToString(taskMwpDto.raisedDate);
            Reference = taskMwpDto.reference;
            StdJobNo = taskMwpDto.stdJobNo;
            StdJobTask = taskMwpDto.WOTaskNo;
            WoStatusM = taskMwpDto.woStatusM;
            WoStatusMDescription = taskMwpDto.woStatusMDescription;
            WoStatusU = taskMwpDto.woStatusU;
            WoStatusUDescription = taskMwpDto.woStatusUDescription;
            WoType = taskMwpDto.woType;
            WorkGroup = taskMwpDto.workGroup;
            WorkOrder = taskMwpDto.workOrder;
            WoDesc = taskMwpDto.woDesc;
            WoTaskNo = taskMwpDto.WOTaskNo;
            WoTaskDesc = taskMwpDto.taskDescription;
        }
    }
}

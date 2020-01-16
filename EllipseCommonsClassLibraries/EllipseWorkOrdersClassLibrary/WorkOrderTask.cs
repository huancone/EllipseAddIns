using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseWorkOrdersClassLibrary
{
    public class WorkOrderTask
    {
        public string DistrictCode;
        public string WorkGroup;
        public string WorkOrder;
        public string WoTaskNo;
        public string WoTaskDesc;
        public string JobDescCode;
        public string SafetyInstr;
        public string CompleteInstr;
        public string ComplTextCode;
        public string AssignPerson;
        public string EstMachHrs;
        public string PlanStartDate;
        public string PlanStartTime;
        public string PlanFinishDate;
        public string PlanFinishTime;
        public string EquipGrpId;
        public string AplType;
        public string CompCode;
        public string CompModCode;
        public string AplSeqNo;
        public string WorkOrderDescription;
        public string EstimatedMachHrs;
        public string EstimatedDurationsHrs;
        public string NoLabor;
        public string NoMaterial;
        public string AplEquipmentGrpId;
        public string AplCompCode;
        public string AplCompModCode;
        public string EstimatedMachHrsSpecified;
        public string EstimatedDurationsHrsSpecified;
        public string ExtTaskText;
        public string CompleteTaskText;
        public string TaskStatusM;
        public string ClosedStatus;
        public string CompletedCode;
        public string CompletedBy;
        public string ClosedDate;

        public WorkOrderTaskService.WorkOrderDTO WorkOrderDto { get; private set; }

        public WorkOrderTaskService.WorkOrderDTO GetWorkOrderDto()
        {
            return WorkOrderDto ?? (WorkOrderDto = new WorkOrderTaskService.WorkOrderDTO());
        }

        public WorkOrderTaskService.WorkOrderDTO SetWorkOrderDto(string prefix, string no)
        {
            WorkOrderDto = WorkOrderActions.GetNewWorkOrderTaskDto(prefix, no);
            return WorkOrderDto;
        }

        public WorkOrderTaskService.WorkOrderDTO SetWorkOrderDto(string no)
        {
            WorkOrderDto = WorkOrderActions.GetNewWorkOrderTaskDto(no);
            return WorkOrderDto;
        }

        public WorkOrderTaskService.WorkOrderDTO SetWorkOrderTaskDto(WorkOrderTaskService.WorkOrderDTO wo)
        {
            WorkOrderDto = wo;
            return WorkOrderDto;
        }
    }

    public class TaskRequirement
    {
        public string WorkOrder;
        public string DistrictCode;
        public string WorkGroup;
        public string WoTaskDesc;
        public string WoTaskNo;
        public string ReqType;
        public string SeqNo;
        public string ReqCode;
        public string ReqDesc;
        public string QtyReq;
        public string QtyIss;
        public string HrsReq;
        public string HrsReal;
        public string UoM;
    }
}

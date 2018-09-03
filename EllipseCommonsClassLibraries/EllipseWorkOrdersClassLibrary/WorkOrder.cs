using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using WorkOrderTaskService = EllipseWorkOrdersClassLibrary.WorkOrderTaskService;

namespace EllipseWorkOrdersClassLibrary
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WorkOrder
    {
        public string districtCode;
        public string workGroup;
        private WorkOrderService.WorkOrderDTO workOrderDTO;
        public string workOrderDesc;
        public string workOrderStatusM;
        public string equipmentNo;
        public string equipmentRef;
        public string compCode;
        public string compModCode;
        public string workOrderType;
        public string maintenanceType;

        public string workOrderStatusU;
        public string raisedDate;
        public string raisedTime;
        public string originatorId;
        public string origPriority;
        public string origDocType;
        public string origDocNo;
        public string requestId;

        public string stdJobNo;
        public string maintSchTask;
        public string assignPerson;
        public string planPriority;
        public string requisitionStartDate;
        public string requisitionStartTime;
        public string requiredByDate;
        public string requiredByTime;
        public string planStrDate;
        public string planStrTime;
        public string planFinDate;
        public string planFinTime;
        public string unitOfWork;
        public string unitsRequired;
        public string pcComplete;
        public string unitsComplete;
        private WorkOrderService.WorkOrderDTO relatedWoDTO;
        public string accountCode;
        public string projectNo;
        public string parentWo;
        public string autoRequisitionInd;
        public string failurePart;
        public string jobCode1;
        public string jobCode2;
        public string jobCode3;
        public string jobCode4;
        public string jobCode5;
        public string jobCode6;
        public string jobCode7;
        public string jobCode8;
        public string jobCode9;
        public string jobCode10;
        public string jobCodeFlag; //informativo propio de la clase para indicar si tiene o no tiene al menos un jobCode
        public string completedCode;
        public string completedBy;
        public string completeTextFlag;
        public string closeCommitDate;
        //Location
        public string location;
        public string locationFr;
        public string noticeLocn;
        //Valores Calculados, Estimados y Actual
        //Se entiende calculado CALC como un valor estimado calculado. Se entiende estimado EST como un valor estimado manual. Se entiende actual ACT como el valor actual real
        //El valor CALC y EST para Horas de Duración es el mismo campo EST y es independiente del flag
        public string calculatedDurationsFlag;
        public string estimatedDurationsHrs;
        public string actualDurationsHrs;

        public string calculatedEquipmentFlag;
        public string calculatedEquipmentCost;
        public string estimatedEquipmentCost;
        public string actualEquipmentCost;

        public string calculatedLabFlag;
        public string calculatedLabHrs;
        public string estimatedLabHrs;
        public string actualLabHrs;
        public string calculatedLabCost;
        public string estimatedLabCost;
        public string actualLabCost;

        public string calculatedMatFlag;
        public string calculatedMatCost;
        public string estimatedMatCost;
        public string actualMatCost;

        public string calculatedOtherFlag;
        public string calculatedOtherCost;
        public string estimatedOtherCost;
        public string actualOtherCost;

        public string finalCosts;

        private ExtendedDescription _extendedDescription;

        /// <summary>
        /// Obtiene los campos de WorkOrderDTO para las acciones requeridas por el servicio
        /// </summary>
        /// <returns>WorkOrderService.WorkOrderDTO: arreglo(no, prefix)</returns>
        public WorkOrderService.WorkOrderDTO GetWorkOrderDto()
        {
            return workOrderDTO ?? (workOrderDTO = new WorkOrderService.WorkOrderDTO());
        }

        public WorkOrderService.WorkOrderDTO SetWorkOrderDto(string prefix, string no)
        {
            workOrderDTO = WorkOrderActions.GetNewWorkOrderDto(prefix, no);
            return workOrderDTO;
        }

        public WorkOrderService.WorkOrderDTO SetWorkOrderDto(string no)
        {
            workOrderDTO = WorkOrderActions.GetNewWorkOrderDto(no);
            return workOrderDTO;
        }

        public WorkOrderService.WorkOrderDTO SetWorkOrderDto(WorkOrderService.WorkOrderDTO wo)
        {
            workOrderDTO = wo;
            return workOrderDTO;
        }

        public WorkOrderService.WorkOrderDTO GetRelatedWoDto()
        {
            return relatedWoDTO ?? (relatedWoDTO = new WorkOrderService.WorkOrderDTO());
        }

        public WorkOrderService.WorkOrderDTO SetRelatedWoDto(string no)
        {
            relatedWoDTO = WorkOrderActions.GetNewWorkOrderDto(no);
            return relatedWoDTO;
        }

        public WorkOrderService.WorkOrderDTO SetRelatedWoDto(string prefix, string no)
        {
            relatedWoDTO = WorkOrderActions.GetNewWorkOrderDto(prefix, no);
            return relatedWoDTO;
        }

        public void SetStatus(string statusName)
        {
            if (!string.IsNullOrEmpty(WoStatusList.GetStatusCode(statusName)))
                workOrderStatusM = WoStatusList.GetStatusCode(statusName);
        }

        public ExtendedDescription GetExtendedDescription(string urlService, WorkOrderService.OperationContext opContext)
        {
            if (_extendedDescription != null) return _extendedDescription;

            _extendedDescription = WorkOrderActions.GetWorkOrderExtendedDescription(urlService, opContext, districtCode,
                GetWorkOrderDto().prefix + GetWorkOrderDto().no);

            return _extendedDescription;
        }

        public void SetExtendedDescription(string header, string body)
        {
            if (_extendedDescription == null)
                _extendedDescription = new ExtendedDescription();
            _extendedDescription.Header = header;
            _extendedDescription.Body = body;
        }
    }

    public class ExtendedDescription
    {
        public string Header;
        public string Body;
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WorkOrderCompleteAtributes
    {
        public string districtCode;
        public WorkOrderService.WorkOrderDTO workOrder;
        public string completedBy;
        public string completedCode;
        public string closedDate;
        public string closedTime;
        public string outServDate;
        public string outServTime;
        public bool crteInsitu;
        public bool crteInsituSpecified;
        public string earnCode;
        public string failurePart;
        public decimal hoursCompleted;
        public bool hoursCompletedSpecified;
        public string completeCommentToAppend;
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WorkOrderDuration
    {
        public string jobDurationsCode;
        public string jobDurationsDate;
        public string jobDurationsStart;
        public string jobDurationsFinish;
        public decimal jobDurationsSeqNo;
        public bool jobDurationsSeqNoSpecified;
        public decimal jobDurationsHours;
        public bool jobDurationsHoursSpecified;

        public WorkOrderService.DurationsDTO GetDurationDto()
        {
            var duration = new WorkOrderService.DurationsDTO
            {
                jobDurationsCode = jobDurationsCode,
                jobDurationsDate = jobDurationsDate,
                jobDurationsStart = jobDurationsStart,
                jobDurationsFinish = jobDurationsFinish,
                jobDurationsSeqNo = jobDurationsSeqNo,
                jobDurationsSeqNoSpecified = jobDurationsSeqNoSpecified,
                jobDurationsHours = jobDurationsHours,
                jobDurationsHoursSpecified = jobDurationsHoursSpecified
            };

            return duration;
        }

        public void SetDurationFromDto(WorkOrderService.DurationsDTO duration)
        {
            jobDurationsCode = duration.jobDurationsCode;
            jobDurationsDate = duration.jobDurationsDate;
            jobDurationsStart = duration.jobDurationsStart;
            jobDurationsFinish = duration.jobDurationsFinish;
            jobDurationsSeqNo = duration.jobDurationsSeqNo;
            jobDurationsSeqNoSpecified = duration.jobDurationsSeqNoSpecified;
            jobDurationsHours = duration.jobDurationsHours;
            jobDurationsHoursSpecified = duration.jobDurationsHoursSpecified;
        }
    }

    public static class WoStatusList
    {
        public static string Open = "OPEN";
        public static string OpenCode = "O";
        public static string Authorized = "AUTHORIZED";
        public static string AuthorizedCode = "A";
        public static string Closed = "CLOSED";
        public static string ClosedCode = "C";
        public static string Cancelled = "CANCELLED";
        public static string CancelledCode = "L";
        public static string InWork = "IN_WORK";
        public static string InWorkCode = "W";
        public static string Estimated = "ESTIMATED";
        public static string EstimatedCode = "E";

        public static string Uncompleted = "UNCOMPLETED";

        public static string GetStatusCode(string statusName)
        {
            if (statusName == Open)
                return OpenCode;
            if (statusName == Authorized)
                return AuthorizedCode;
            if (statusName == Closed)
                return ClosedCode;
            if (statusName == Cancelled)
                return CancelledCode;
            if (statusName == InWork)
                return InWorkCode;
            if (statusName == Estimated)
                return EstimatedCode;
            return null;
        }

        public static string GetStatusName(string statusCode)
        {
            if (statusCode == OpenCode)
                return Open;
            if (statusCode == AuthorizedCode)
                return Authorized;
            if (statusCode == ClosedCode)
                return Closed;
            if (statusCode == CancelledCode)
                return Cancelled;
            if (statusCode == InWorkCode)
                return InWork;
            if (statusCode == EstimatedCode)
                return Estimated;
            return null;
        }

        public static List<string> GetStatusNames(bool uncompletedCustom = false)
        {
            if (uncompletedCustom)
                return new List<string> { Open, Authorized, Closed, Cancelled, InWork, Estimated, Uncompleted };
            return new List<string> { Open, Authorized, Closed, Cancelled, InWork, Estimated };
        }

        public static List<string> GetStatusCodes()
        {
            var list = new List<string> { OpenCode, AuthorizedCode, ClosedCode, CancelledCode, InWorkCode, EstimatedCode };
            return list;
        }

        public static List<string> GetUncompletedStatusNames()
        {
            var list = new List<string> { Open, Authorized, InWork, Estimated };
            return list;
        }

        public static List<string> GetUncompletedStatusCodes()
        {
            var list = new List<string> { OpenCode, AuthorizedCode, InWorkCode, EstimatedCode };
            return list;
        }


    }

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
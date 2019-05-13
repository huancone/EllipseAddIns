using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary
{
    public class RequisitionHeader
    {
        public string DistrictCode;
        public string IndSerie;
        public string IreqNo;
        public string IreqType;
        public string RequestedBy;
        public string RequiredByPos;
        public string IssTranType;
        public string OrigWhouseId;
        public string PriorityCode;

        public string CostDistrictA;
        public string WorkOrderA;
        public string WorkProjectIndA;//Solo para el MSO140
        public string EquipmentA;
        public string ProjectA;
        public string CostCentreA;

        public string DelivInstrA;
        public string DelivInstrB;

        public string AllocPcA;
        public string RequiredByDate;
        public string AnswerB;
        public string AnswerD;

        public bool PartIssue;
        public bool PartIssueSpecified;
        public bool ProtectedInd;
        public bool PickTaskReq;
        public bool ProtectedIndSpecified;

        /// <summary>
        /// Compara el objeto encabezado RequisitionHeader con otro encabezado. Devuelve true si el encabezado es igual a objectHeader
        /// </summary>
        /// <param name="objectHeader"></param>
        /// <returns>bool: true si objectHeader es igual</returns>
        public bool Equals(RequisitionHeader objectHeader)
        {
            return DistrictCode == objectHeader.DistrictCode &&
                   IndSerie == objectHeader.IndSerie &&
                   //IreqNo == objectHeader.IreqNo && //este no se debe comparar
                   IreqType == objectHeader.IreqType &&
                   RequestedBy == objectHeader.RequestedBy &&
                   RequiredByPos == objectHeader.RequiredByPos &&
                   IssTranType == objectHeader.IssTranType &&
                   OrigWhouseId == objectHeader.OrigWhouseId &&
                   PriorityCode == objectHeader.PriorityCode &&
                   CostDistrictA == objectHeader.CostDistrictA &&
                   WorkOrderA == objectHeader.WorkOrderA &&
                   EquipmentA == objectHeader.EquipmentA &&
                   ProjectA == objectHeader.ProjectA &&
                   CostCentreA == objectHeader.CostCentreA &&
                   DelivInstrA == objectHeader.DelivInstrA &&
                   DelivInstrB == objectHeader.DelivInstrB &&
                   AllocPcA == objectHeader.AllocPcA &&
                   RequiredByDate == objectHeader.RequiredByDate &&
                   AnswerB == objectHeader.AnswerB &&
                   AnswerD == objectHeader.AnswerD &&
                   PartIssue == objectHeader.PartIssue &&
                   ProtectedInd == objectHeader.ProtectedInd &&
                   PickTaskReq == objectHeader.PickTaskReq &&
                   ProtectedIndSpecified == objectHeader.ProtectedIndSpecified;
        }

        public RequisitionService.RequisitionServiceCreateHeaderRequestDTO GetCreateHeaderRequest()
        {
            var request = new RequisitionService.RequisitionServiceCreateHeaderRequestDTO
            {
                districtCode = DistrictCode,
                ireqNo = IreqNo,
                ireqType = IreqType,
                requestedBy = RequestedBy,
                requiredByPos = RequiredByPos,
                issTranType = IssTranType,
                origWhouseId = OrigWhouseId,
                priorityCode = PriorityCode,
                costDistrictA = CostDistrictA,
                workOrderA = GetNewWorkOrderDto(WorkOrderA),
                equipmentA = EquipmentA,
                projectA = ProjectA,
                costCentreA = CostCentreA,
                delivInstrA = DelivInstrA,
                delivInstrB = DelivInstrB,
                allocPcA = AllocPcA,
                requiredByDate = RequiredByDate,
                answerB = AnswerB,
                answerD = AnswerD,
                partIssue = PartIssue,
                partIssueSpecified = PartIssueSpecified,
                protectedInd = ProtectedInd,
                pickTaskReq = PickTaskReq,
                protectedIndSpecified = ProtectedIndSpecified
            };

            return request;
        }
        public RequisitionService.RequisitionServiceCreateHeaderReplyDTO GetCreateHeaderReply()
        {
            var request = new RequisitionService.RequisitionServiceCreateHeaderReplyDTO
            {
                districtCode = DistrictCode,
                ireqNo = IreqNo,
                ireqType = IreqType,
                requestedBy = RequestedBy,
                requiredByPos = RequiredByPos,
                issTranType = IssTranType,
                origWhouseId = OrigWhouseId,
                priorityCode = PriorityCode,
                costDistrictA = CostDistrictA,
                workOrderA = GetNewWorkOrderDto(WorkOrderA),
                equipmentA = EquipmentA,
                projectA = ProjectA,
                costCentreA = CostCentreA,
                delivInstrA = DelivInstrA,
                delivInstrB = DelivInstrB,
                allocPcA = AllocPcA,
                requiredByDate = RequiredByDate,
                answerB = AnswerB,
                answerD = AnswerD,
                partIssue = PartIssue,
                partIssueSpecified = PartIssueSpecified,
                protectedInd = ProtectedInd,
                pickTaskReq = PickTaskReq,
                protectedIndSpecified = ProtectedIndSpecified
            };

            return request;
        }

        /// <summary>
        /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
        /// </summary>
        /// <param name="no">string: Número de la orden de trabajo</param>
        /// <returns>WorkOrderDTO</returns>
        public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string no)
        {
            var workOrderDto = new RequisitionService.WorkOrderDTO();
            if (string.IsNullOrWhiteSpace(no)) return workOrderDto;

            no = no.Trim();
            if (no.Length < 3)
                throw new Exception(@"El número de orden no corresponde a una orden válida");
            workOrderDto.prefix = no.Substring(0, 2);
            workOrderDto.no = no.Substring(2, no.Length - 2);
            return workOrderDto;
        }
        /// <summary>
        /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
        /// </summary>
        /// <param name="prefix">string: prefijo de la orden de trabajo</param>
        /// <param name="no">string: Número de la orden de trabajo</param>
        /// <returns>WorkOrderDTO</returns>
        public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string prefix, string no)
        {
            var workOrderDto = new RequisitionService.WorkOrderDTO
            {
                prefix = prefix,
                no = no
            };

            return workOrderDto;
        }

    }

}

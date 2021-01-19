using System;
using EllipseLabourCostingExcelAddIn.EquipHireTranService;


namespace EllipseLabourCostingExcelAddIn.EquipmentHireLibrary
{
    public static class EquipmentHireActions
    {
        public static EquipHireTranServiceCreateReplyDTO CreateEquipmentHire(string urlService, OperationContext opContext, EquipmentHire equipmentHire, bool replaceExisting = true)
        {
            var service = new EquipHireTranService.EquipHireTranService { Url = urlService + "/EquipHireTran" };

            var request = new EquipHireTranServiceCreateRequestDTO
            {
                //transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                tranDate = equipmentHire.TransactionDate,
                employee = equipmentHire.EmployeeId,
                equipmentRef = equipmentHire.EquipmentReference,
                statValue = !string.IsNullOrWhiteSpace(equipmentHire.Value) ? Convert.ToDecimal(equipmentHire.Value) : default,
                statValueSpecified = !string.IsNullOrWhiteSpace(equipmentHire.Value),
                prodDataType = equipmentHire.StatisticType,
                projectNo = equipmentHire.ProjectNo,
                workOrder = GetNewWorkOrderDto(equipmentHire.WorkOrder),
                WOTaskNo = equipmentHire.Task,
                accountCode = equipmentHire.AccountCode
            };

            var result = service.create(opContext, request);
            return result;
        }
        public static EquipHireTranServiceDeleteReplyDTO DeleteEquipmentHire(string urlService, OperationContext opContext, EquipmentHire equipmentHire, bool replaceExisting = true)
        {
            var service = new EquipHireTranService.EquipHireTranService { Url = urlService + "/EquipHireTran" };

            var request = new EquipHireTranServiceDeleteRequestDTO
            {
                //transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                tranDate = equipmentHire.TransactionDate,
                employee = equipmentHire.EmployeeId,
                tranSeqNo = !string.IsNullOrWhiteSpace(equipmentHire.TranSequence) ? Convert.ToDecimal(equipmentHire.TranSequence) : default,
                tranSeqNoSpecified = !string.IsNullOrWhiteSpace(equipmentHire.TranSequence),
            };

            var result = service.delete(opContext, request);
            return result;
        }
        //Se toma la misma función de la librería WorkOrderClassLibrary.WorkOrderActions.GetNewWorkOrderDto
        public static WorkOrderDTO GetNewWorkOrderDto(string no)
        {
            var workOrderDto = new WorkOrderDTO();
            if (string.IsNullOrWhiteSpace(no)) return workOrderDto;

            no = no.Trim();
            if (no.Length < 3)
                throw new Exception(@"El número de orden no corresponde a una orden válida");
            workOrderDto.prefix = no.Substring(0, 2);
            workOrderDto.no = no.Substring(2, no.Length - 2);
            return workOrderDto;
        }
    }
}

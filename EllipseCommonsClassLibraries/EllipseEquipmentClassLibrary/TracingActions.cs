using System;
using System.Linq;
using EllipseEquipmentClassLibrary.EquipTraceService;

namespace EllipseEquipmentClassLibrary
{
    public static class TracingActions
    {
        public static bool Fitment(EquipTraceService.OperationContext operationContext, string urlService, string instEquipmentRef, string compCode, string modCode, string fitEquipmentRef, string fitDate, string refType, string refNumber)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceFitRequestDTO()
            {
                installEquipmentRef = instEquipmentRef,
                compCode = compCode,
                modCode = modCode,
                fitEquipmentRef = fitEquipmentRef,
                ETDate = fitDate,
                refType = refType,
                refNum = refNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.fit(operationContext, request);
            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return true;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }

        public static bool Defitment(EquipTraceService.OperationContext operationContext, string urlService, string instEquipmentRef, string compCode, string modCode, string fitEquipmentRef, string defitDate, string refType, string refNumber)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceDefitRequestDTO()
            {
                installEquipmentRef = instEquipmentRef,
                compCode = compCode,
                modCode = modCode,
                fitEquipmentRef = fitEquipmentRef,
                ETDate = defitDate,
                refType = refType,
                refNum = refNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.defit(operationContext, request);
            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return true;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }
    }

}

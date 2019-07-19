using System;
using System.Linq;
using EllipseEquipmentClassLibrary.EquipTraceService;

namespace EllipseEquipmentClassLibrary
{
    public static class TracingActions
    {
        /*
            A ORIGINAL PURCHASE
            B FITMENT
            C DEFITMENT
            D REBUILD OFFSITE
            E REBUILT ONSITE
            F CREDIT REQUISITION
            G W/HOUSE REQUISITION
            H EXCHANGED
            I SCRAPPED
            J DISSASEMBLED
            K SOLD
            L REBUILT IN SITU
            M RETURN FROM SUPPLIER
            N REPAIR IN SITU
            O REPAIR UNFITTED
            P INSPECT IN SITU
            Q REPAIR INFFITED OFF
        */

        public static TracingItem Fitment(EquipTraceService.OperationContext operationContext, string urlService, TracingItem traceRegister)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceFitRequestDTO()
            {
                installEquipmentRef = traceRegister.InstEquipmentNo,
                compCode = traceRegister.ComponentCode,
                modCode = traceRegister.ModifierCode,
                fitEquipmentRef = traceRegister.FitEquipmentNo,
                ETDate = traceRegister.Date,
                refType = traceRegister.ReferenceType,
                refNum = traceRegister.ReferenceNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.fit(operationContext, request);

            var replyRegister = new TracingItem();
            replyRegister.InstEquipmentNo = reply.installEquipmentRef;
            replyRegister.ComponentCode = reply.compCode;
            replyRegister.ModifierCode = reply.modCode;
            replyRegister.TracingAction = reply.actionDesc;
            replyRegister.FitEquipmentNo = reply.fitEquipmentRef;
            replyRegister.Date = reply.ETDate;
            //replyRegister.SequenceNumber = reply;
            replyRegister.ReferenceType = reply.refType;
            replyRegister.ReferenceNumber = reply.refNum;


            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return replyRegister;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }

        public static TracingItem Defitment(EquipTraceService.OperationContext operationContext, string urlService, TracingItem traceRegister)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceDefitRequestDTO()
            {
                installEquipmentRef = traceRegister.InstEquipmentNo,
                compCode = traceRegister.ComponentCode,
                modCode = traceRegister.ModifierCode,
                fitEquipmentRef = traceRegister.FitEquipmentNo,
                ETDate = traceRegister.Date,
                refType = traceRegister.ReferenceType,
                refNum = traceRegister.ReferenceNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.defit(operationContext, request);

            var replyRegister = new TracingItem();
            replyRegister.InstEquipmentNo = reply.installEquipmentRef;
            replyRegister.ComponentCode = reply.compCode;
            replyRegister.ModifierCode = reply.modCode;
            replyRegister.TracingAction = reply.actionDesc;
            replyRegister.FitEquipmentNo = reply.fitEquipmentRef;
            //replyRegister.Date = reply.ETDate;
            //replyRegister.SequenceNumber = reply;
            replyRegister.ReferenceType = reply.refType;
            replyRegister.ReferenceNumber = reply.refNum;


            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return replyRegister;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }

        public static TracingItem Delete(EquipTraceService.OperationContext operationContext, string urlService, TracingItem traceRegister)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            
            var attributeList = new EquipTraceService.Attribute[5];
            attributeList[0] = new EquipTraceService.Attribute
            {
                name = "equipmentNo",
                value = traceRegister.InstEquipmentNo
            };
            attributeList[1] = new EquipTraceService.Attribute
            {
                name = "profileCompCode",
                value = traceRegister.ComponentCode
            };
            attributeList[2] = new EquipTraceService.Attribute
            {
                name = "profileModCode",
                value = traceRegister.ModifierCode
            };
            attributeList[3] = new EquipTraceService.Attribute
            {
                name = "refNum",
                value = traceRegister.ReferenceNumber
            };
            attributeList[4] = new EquipTraceService.Attribute
            {
                name = "refType",
                value = traceRegister.ReferenceType
            };
            var request = new EquipTraceServiceDeleteTracingActionRequestDTO()
            {
                fitEquipmentNo = traceRegister.FitEquipmentNo,
                ETDate = traceRegister.Date,
                seqNum = !string.IsNullOrWhiteSpace(traceRegister.SequenceNumber) ? Convert.ToDecimal(traceRegister.SequenceNumber) : default(decimal),
                seqNumSpecified = string.IsNullOrWhiteSpace(traceRegister.SequenceNumber) ? false : true,
                traceActn = traceRegister.TracingAction,
                customAttributes = attributeList
            };

            proxyEquip.Url = urlService + "/EquipTrace";
            
            var reply = proxyEquip.deleteTracingAction(operationContext, request);
            var replyRegister = new TracingItem();
            replyRegister.InstEquipmentNo = reply.installEquipment;
            replyRegister.ComponentCode = reply.compCode;
            replyRegister.ModifierCode = reply.modCode;
            replyRegister.TracingAction = reply.actionDesc;
            replyRegister.FitEquipmentNo = reply.fitEquipmentNo;
            //replyRegister.Date = reply.ETDate;
            //replyRegister.SequenceNumber = reply.;
            replyRegister.ReferenceType = reply.refType;
            replyRegister.ReferenceNumber = reply.refNum;

            var information = reply.warningsAndInformation;
            if (information == null || information.Length <= 0)
                return replyRegister;
            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }
    }

}

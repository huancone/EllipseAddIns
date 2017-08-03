﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using System.Diagnostics.CodeAnalysis;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using EllipseReferenceCodesClassLibrary;
using EllipseStdTextClassLibrary;

namespace EllipseWorkOrdersClassLibrary
{
    [SuppressMessage("ReSharper", "ForCanBeConvertedToForeach")]
    public static class WorkOrderActions
    {
        public static List<WorkOrder> FetchWorkOrder(EllipseFunctions ef, string district, int primakeryKey, string primaryValue, int secondarykey, string secondaryValue, int dateKey, string startDate, string endDate, string woStatus)
        {
            var sqlQuery = Queries.GetFetchWoQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, secondarykey, secondaryValue, dateKey, startDate, endDate, woStatus);
            var drWorkOrder = ef.GetQueryResult(sqlQuery);
            var list = new List<WorkOrder>();

            if (drWorkOrder == null || drWorkOrder.IsClosed || !drWorkOrder.HasRows) return list;
            while (drWorkOrder.Read())
            {
                var order = new WorkOrder
                {
                    districtCode = drWorkOrder["DSTRCT_CODE"].ToString().Trim(),
                    workGroup = drWorkOrder["WORK_GROUP"].ToString().Trim(),
                    workOrderStatusM = drWorkOrder["WO_STATUS_M"].ToString().Trim(),
                    workOrderDesc = drWorkOrder["WO_DESC"].ToString().Trim(),
                    equipmentNo = drWorkOrder["EQUIP_NO"].ToString().Trim(),
                    compCode = drWorkOrder["COMP_CODE"].ToString().Trim(),
                    compModCode = drWorkOrder["COMP_MOD_CODE"].ToString().Trim(),
                    workOrderType = drWorkOrder["WO_TYPE"].ToString().Trim(),
                    maintenanceType = drWorkOrder["MAINT_TYPE"].ToString().Trim(),
                    workOrderStatusU = drWorkOrder["WO_STATUS_U"].ToString().Trim(),
                    raisedDate = drWorkOrder["RAISED_DATE"].ToString().Trim(),
                    raisedTime = drWorkOrder["RAISED_TIME"].ToString().Trim(),
                    originatorId = drWorkOrder["ORIGINATOR_ID"].ToString().Trim(),
                    origPriority = drWorkOrder["ORIG_PRIORITY"].ToString().Trim(),
                    origDocType = drWorkOrder["ORIG_DOC_TYPE"].ToString().Trim(),
                    origDocNo = drWorkOrder["ORIG_DOC_NO"].ToString().Trim(),
                    requestId = drWorkOrder["REQUEST_ID"].ToString().Trim(),
                    stdJobNo = drWorkOrder["STD_JOB_NO"].ToString().Trim(),
                    maintSchTask = drWorkOrder["MAINT_SCH_TASK"].ToString().Trim(),
                    autoRequisitionInd = drWorkOrder["AUTO_REQ_IND"].ToString().Trim(),
                    assignPerson = drWorkOrder["ASSIGN_PERSON"].ToString().Trim(),
                    planPriority = drWorkOrder["PLAN_PRIORITY"].ToString().Trim(),
                    requisitionStartDate = drWorkOrder["REQ_START_DATE"].ToString().Trim(),
                    requisitionStartTime = drWorkOrder["REQ_START_TIME"].ToString().Trim(),
                    requiredByDate = drWorkOrder["REQ_BY_DATE"].ToString().Trim(),
                    requiredByTime = drWorkOrder["REQ_BY_TIME"].ToString().Trim(),
                    planStrDate = drWorkOrder["PLAN_STR_DATE"].ToString().Trim(),
                    planStrTime = drWorkOrder["PLAN_STR_TIME"].ToString().Trim(),
                    planFinDate = drWorkOrder["PLAN_FIN_DATE"].ToString().Trim(),
                    planFinTime = drWorkOrder["PLAN_FIN_TIME"].ToString().Trim(),
                    unitOfWork = drWorkOrder["UNIT_OF_WORK"].ToString().Trim(),
                    unitsRequired = drWorkOrder["UNITS_REQUIRED"].ToString().Trim(),
                    pcComplete = drWorkOrder["PC_COMPLETE"].ToString().Trim(),
                    unitsComplete = drWorkOrder["UNITS_COMPLETE"].ToString().Trim(),
                    accountCode = drWorkOrder["DSTRCT_ACCT_CODE"].ToString().Trim(),
                    projectNo = drWorkOrder["PROJECT_NO"].ToString().Trim(),
                    parentWo = drWorkOrder["PARENT_WO"].ToString().Trim(),
                    failurePart = drWorkOrder["FAILURE_PART"].ToString().Trim(),
                    jobCode1 = drWorkOrder["WO_JOB_CODEX1"].ToString().Trim(),
                    jobCode2 = drWorkOrder["WO_JOB_CODEX2"].ToString().Trim(),
                    jobCode3 = drWorkOrder["WO_JOB_CODEX3"].ToString().Trim(),
                    jobCode4 = drWorkOrder["WO_JOB_CODEX4"].ToString().Trim(),
                    jobCode5 = drWorkOrder["WO_JOB_CODEX5"].ToString().Trim(),
                    jobCode6 = drWorkOrder["WO_JOB_CODEX6"].ToString().Trim(),
                    jobCode7 = drWorkOrder["WO_JOB_CODEX7"].ToString().Trim(),
                    jobCode8 = drWorkOrder["WO_JOB_CODEX8"].ToString().Trim(),
                    jobCode9 = drWorkOrder["WO_JOB_CODEX9"].ToString().Trim(),
                    jobCode10 = drWorkOrder["WO_JOB_CODEX10"].ToString().Trim(),
                    jobCodeFlag = drWorkOrder["JOB_CODES"].ToString().Trim(),
                    completedCode = drWorkOrder["COMPLETED_CODE"].ToString().Trim(),
                    completedBy = drWorkOrder["COMPLETED_BY"].ToString().Trim(),
                    completeTextFlag = drWorkOrder["COMPLETE_TEXT_FLAG"].ToString().Trim(),
                    closeCommitDate = drWorkOrder["CLOSED_DT"].ToString().Trim(),
                    locationFr = drWorkOrder["LOCATION_FR"].ToString().Trim(),
                    location = drWorkOrder["LOCATION"].ToString().Trim(),
                    noticeLocn = drWorkOrder["NOTICE_LOCN"].ToString().Trim(),
                    calculatedDurationsFlag = drWorkOrder["CALC_DUR_HRS_SW"].ToString().Trim(),
                    estimatedDurationsHrs = drWorkOrder["EST_DUR_HRS"].ToString().Trim(),
                    actualDurationsHrs = drWorkOrder["ACT_DUR_HRS"].ToString().Trim(),
                    calculatedLabFlag = drWorkOrder["RES_UPDATE_FLAG"].ToString().Trim(),
                    estimatedLabHrs = drWorkOrder["EST_LAB_HRS"].ToString().Trim(),
                    calculatedLabHrs = drWorkOrder["CALC_LAB_HRS"].ToString().Trim(),
                    actualLabHrs = drWorkOrder["ACT_LAB_HRS"].ToString().Trim(),
                    estimatedLabCost = drWorkOrder["EST_LAB_COST"].ToString().Trim(),
                    calculatedLabCost = drWorkOrder["CALC_LAB_COST"].ToString().Trim(),
                    actualLabCost = drWorkOrder["ACT_LAB_COST"].ToString().Trim(),
                    calculatedMatFlag = drWorkOrder["MAT_UPDATE_FLAG"].ToString().Trim(),
                    estimatedMatCost = drWorkOrder["EST_MAT_COST"].ToString().Trim(),
                    calculatedMatCost = drWorkOrder["CALC_MAT_COST"].ToString().Trim(),
                    actualMatCost = drWorkOrder["ACT_MAT_COST"].ToString().Trim(),
                    calculatedEquipmentFlag = drWorkOrder["EQUIP_UPDATE_FLAG"].ToString().Trim(),
                    estimatedEquipmentCost = drWorkOrder["EST_EQUIP_COST"].ToString().Trim(),
                    calculatedEquipmentCost = drWorkOrder["CALC_EQUIP_COST"].ToString().Trim(),
                    actualEquipmentCost = drWorkOrder["ACT_EQUIP_COST"].ToString().Trim(),
                    estimatedOtherCost = drWorkOrder["EST_OTHER_COST"].ToString().Trim(),
                    actualOtherCost = drWorkOrder["ACT_OTHER_COST"].ToString().Trim(),
                    finalCosts = drWorkOrder["FINAL_COSTS"].ToString().Trim()
                };
                order.SetWorkOrderDto(drWorkOrder["WORK_ORDER"].ToString().Trim());
                order.SetRelatedWoDto(drWorkOrder["RELATED_WO"].ToString().Trim());
                list.Add(order);
            }

            return list;
        }

        public static WorkOrder FetchWorkOrder(EllipseFunctions ef, string district, WorkOrderDTO workOrder)
        {
            return FetchWorkOrder(ef, district, workOrder.prefix + workOrder.no);
        }
        public static WorkOrder FetchWorkOrder(EllipseFunctions ef, string district, string workOrder)
        {
            long number1;
            if (long.TryParse(workOrder, out number1))
                workOrder = workOrder.PadLeft(8, '0');

            var sqlQuery = Queries.GetFetchWoQuery(ef.dbReference, ef.dbLink, district, workOrder);
            var drWorkOrder = ef.GetQueryResult(sqlQuery);

            if (drWorkOrder == null || drWorkOrder.IsClosed || !drWorkOrder.HasRows || !drWorkOrder.Read()) return null;

            var order = new WorkOrder
            {
                districtCode = drWorkOrder["DSTRCT_CODE"].ToString().Trim(),
                workGroup = drWorkOrder["WORK_GROUP"].ToString().Trim(),
                workOrderStatusM = drWorkOrder["WO_STATUS_M"].ToString().Trim(),
                workOrderDesc = drWorkOrder["WO_DESC"].ToString().Trim(),
                equipmentNo = drWorkOrder["EQUIP_NO"].ToString().Trim(),
                compCode = drWorkOrder["COMP_CODE"].ToString().Trim(),
                compModCode = drWorkOrder["COMP_MOD_CODE"].ToString().Trim(),
                workOrderType = drWorkOrder["WO_TYPE"].ToString().Trim(),
                maintenanceType = drWorkOrder["MAINT_TYPE"].ToString().Trim(),
                workOrderStatusU = drWorkOrder["WO_STATUS_U"].ToString().Trim(),
                raisedDate = drWorkOrder["RAISED_DATE"].ToString().Trim(),
                raisedTime = drWorkOrder["RAISED_TIME"].ToString().Trim(),
                originatorId = drWorkOrder["ORIGINATOR_ID"].ToString().Trim(),
                origPriority = drWorkOrder["ORIG_PRIORITY"].ToString().Trim(),
                origDocType = drWorkOrder["ORIG_DOC_TYPE"].ToString().Trim(),
                origDocNo = drWorkOrder["ORIG_DOC_NO"].ToString().Trim(),
                requestId = drWorkOrder["REQUEST_ID"].ToString().Trim(),
                stdJobNo = drWorkOrder["STD_JOB_NO"].ToString().Trim(),
                maintSchTask = drWorkOrder["MAINT_SCH_TASK"].ToString().Trim(),
                autoRequisitionInd = drWorkOrder["AUTO_REQ_IND"].ToString().Trim(),
                assignPerson = drWorkOrder["ASSIGN_PERSON"].ToString().Trim(),
                planPriority = drWorkOrder["PLAN_PRIORITY"].ToString().Trim(),
                requisitionStartDate = drWorkOrder["REQ_START_DATE"].ToString().Trim(),
                requisitionStartTime = drWorkOrder["REQ_START_TIME"].ToString().Trim(),
                requiredByDate = drWorkOrder["REQ_BY_DATE"].ToString().Trim(),
                requiredByTime = drWorkOrder["REQ_BY_TIME"].ToString().Trim(),
                planStrDate = drWorkOrder["PLAN_STR_DATE"].ToString().Trim(),
                planStrTime = drWorkOrder["PLAN_STR_TIME"].ToString().Trim(),
                planFinDate = drWorkOrder["PLAN_FIN_DATE"].ToString().Trim(),
                planFinTime = drWorkOrder["PLAN_FIN_TIME"].ToString().Trim(),
                unitOfWork = drWorkOrder["UNIT_OF_WORK"].ToString().Trim(),
                unitsRequired = drWorkOrder["UNITS_REQUIRED"].ToString().Trim(),
                pcComplete = drWorkOrder["PC_COMPLETE"].ToString().Trim(),
                unitsComplete = drWorkOrder["UNITS_COMPLETE"].ToString().Trim(),
                accountCode = drWorkOrder["DSTRCT_ACCT_CODE"].ToString().Trim(),
                projectNo = drWorkOrder["PROJECT_NO"].ToString().Trim(),
                parentWo = drWorkOrder["PARENT_WO"].ToString().Trim(),
                failurePart = drWorkOrder["FAILURE_PART"].ToString().Trim(),
                jobCode1 = drWorkOrder["WO_JOB_CODEX1"].ToString().Trim(),
                jobCode2 = drWorkOrder["WO_JOB_CODEX2"].ToString().Trim(),
                jobCode3 = drWorkOrder["WO_JOB_CODEX3"].ToString().Trim(),
                jobCode4 = drWorkOrder["WO_JOB_CODEX4"].ToString().Trim(),
                jobCode5 = drWorkOrder["WO_JOB_CODEX5"].ToString().Trim(),
                jobCode6 = drWorkOrder["WO_JOB_CODEX6"].ToString().Trim(),
                jobCode7 = drWorkOrder["WO_JOB_CODEX7"].ToString().Trim(),
                jobCode8 = drWorkOrder["WO_JOB_CODEX8"].ToString().Trim(),
                jobCode9 = drWorkOrder["WO_JOB_CODEX9"].ToString().Trim(),
                jobCode10 = drWorkOrder["WO_JOB_CODEX10"].ToString().Trim(),
                jobCodeFlag = drWorkOrder["JOB_CODES"].ToString().Trim(),
                completedCode = drWorkOrder["COMPLETED_CODE"].ToString().Trim(),
                completedBy = drWorkOrder["COMPLETED_BY"].ToString().Trim(),
                completeTextFlag = drWorkOrder["COMPLETE_TEXT_FLAG"].ToString().Trim(),
                closeCommitDate = drWorkOrder["CLOSED_DT"].ToString().Trim(),
                locationFr = drWorkOrder["LOCATION_FR"].ToString().Trim(),
                location = drWorkOrder["LOCATION"].ToString().Trim(),
                noticeLocn = drWorkOrder["NOTICE_LOCN"].ToString().Trim(),
                calculatedDurationsFlag = drWorkOrder["CALC_DUR_HRS_SW"].ToString().Trim(),
                estimatedDurationsHrs = drWorkOrder["EST_DUR_HRS"].ToString().Trim(),
                actualDurationsHrs = drWorkOrder["ACT_DUR_HRS"].ToString().Trim(),
                calculatedLabFlag = drWorkOrder["RES_UPDATE_FLAG"].ToString().Trim(),
                estimatedLabHrs = drWorkOrder["EST_LAB_HRS"].ToString().Trim(),
                calculatedLabHrs = drWorkOrder["CALC_LAB_HRS"].ToString().Trim(),
                actualLabHrs = drWorkOrder["ACT_LAB_HRS"].ToString().Trim(),
                estimatedLabCost = drWorkOrder["EST_LAB_COST"].ToString().Trim(),
                calculatedLabCost = drWorkOrder["CALC_LAB_COST"].ToString().Trim(),
                actualLabCost = drWorkOrder["ACT_LAB_COST"].ToString().Trim(),
                calculatedMatFlag = drWorkOrder["MAT_UPDATE_FLAG"].ToString().Trim(),
                estimatedMatCost = drWorkOrder["EST_MAT_COST"].ToString().Trim(),
                calculatedMatCost = drWorkOrder["CALC_MAT_COST"].ToString().Trim(),
                actualMatCost = drWorkOrder["ACT_MAT_COST"].ToString().Trim(),
                calculatedEquipmentFlag = drWorkOrder["EQUIP_UPDATE_FLAG"].ToString().Trim(),
                estimatedEquipmentCost = drWorkOrder["EST_EQUIP_COST"].ToString().Trim(),
                calculatedEquipmentCost = drWorkOrder["CALC_EQUIP_COST"].ToString().Trim(),
                actualEquipmentCost = drWorkOrder["ACT_EQUIP_COST"].ToString().Trim(),
                estimatedOtherCost = drWorkOrder["EST_OTHER_COST"].ToString().Trim(),
                actualOtherCost = drWorkOrder["ACT_OTHER_COST"].ToString().Trim(),
                finalCosts = drWorkOrder["FINAL_COSTS"].ToString().Trim()
            };
            order.SetWorkOrderDto(drWorkOrder["WORK_ORDER"].ToString().Trim());
            order.SetRelatedWoDto(drWorkOrder["RELATED_WO"].ToString().Trim());
            return order;
        }
        public static ExtendedDescription GetWOrkOrderExtendedDescription(string urlService, OperationContext opContext, string district, string workOrder)
        {
            var description = new ExtendedDescription();
            var stdTextOpContext = StdText.GetStdTextOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);

            var stdTextId = "WO" + district + workOrder;

            description.Header = StdText.GetHeader(urlService, stdTextOpContext, stdTextId);
            description.Body = StdText.GetText(urlService, stdTextOpContext, stdTextId);
            return description;
        }

        public static ReplyMessage UpdateWorkOrderExtendedDescription(string urlService, OperationContext opContext, string district, string workOrder, ExtendedDescription description)
        {
            return UpdateWorkOrderExtendedDescription(urlService, opContext, district, workOrder, description.Header, description.Body);
        }

        public static ReplyMessage UpdateWorkOrderExtendedDescription(string urlService, OperationContext opContext, string district, string workOrder, string headerText, string bodyText)
        {
            var reply = new ReplyMessage();
            try
            {
                var stdTextOpContext = StdText.GetCustomOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
                var stdTextId = "WO" + district + workOrder;
                bool headerReply = true, bodyReply = true;
                if (!string.IsNullOrEmpty(headerText))
                    headerReply = StdText.SetHeader(urlService, stdTextOpContext, stdTextId, headerText);
                if (!string.IsNullOrEmpty(bodyText))
                    bodyReply = StdText.SetText(urlService, stdTextOpContext, stdTextId, bodyText);

                if (headerReply && bodyReply)
                    return reply;
                var errorList = new List<string>();
                if (!headerReply)
                    errorList.Add("No se pudo actualizar el encabezado de texto del StdText WO" + workOrder);
                if (!bodyReply)
                    errorList.Add("No se pudo actualizar el cuerpo de texto del StdText WO" + workOrder);
                reply.Errors = errorList.ToArray();
                reply.Message = "Error al actualizar el texto extendido de orden " + workOrder;


            }
            catch (Exception ex)
            {
                Debugger.LogError("WorkOrder.UpdateWorkOrderExtendedDescription()", ex.Message);
                reply.Message = "Error al actualizar el texto extendido de orden " + workOrder;
                var errorList = new List<string> { "No se pudo actualizar el texto del StdText WO" + workOrder };
                reply.Errors = errorList.ToArray();
            }
            return reply;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkOrder")</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de Operación de la WorkOrder</param>
        /// <param name="wo">WorkOrder: WorkOrder</param>
        public static WorkOrderServiceCreateReplyDTO CreateWorkOrder(string urlService, OperationContext opContext, WorkOrder wo)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio
            var requestWo = new WorkOrderServiceCreateRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = wo.districtCode ?? requestWo.districtCode;
            requestWo.workGroup = wo.workGroup ?? requestWo.workGroup;
            if (string.IsNullOrWhiteSpace(wo.GetWorkOrderDto().no) && !string.IsNullOrWhiteSpace(wo.GetWorkOrderDto().prefix))
                requestWo.workOrderPrefix = wo.GetWorkOrderDto().prefix;
            else
                requestWo.workOrder = wo.GetWorkOrderDto();


            requestWo.workOrderDesc = wo.workOrderDesc ?? requestWo.workOrderDesc;
            //workOrderStatusM //n/a para crear
            requestWo.equipmentNo = wo.equipmentNo ?? requestWo.equipmentNo;
            requestWo.equipmentRef = wo.equipmentRef ?? requestWo.equipmentRef;
            requestWo.compCode = wo.compCode ?? requestWo.compCode;
            requestWo.compModCode = wo.compModCode ?? requestWo.compModCode;
            requestWo.workOrderType = wo.workOrderType ?? requestWo.workOrderType;
            requestWo.maintenanceType = wo.maintenanceType ?? requestWo.maintenanceType;
            requestWo.workOrderStatusU = wo.workOrderStatusU ?? requestWo.workOrderStatusU;

            requestWo.raisedDate = wo.raisedDate ?? requestWo.raisedDate;
            requestWo.raisedTime = wo.raisedTime ?? requestWo.raisedTime;
            requestWo.originatorId = wo.originatorId ?? requestWo.originatorId;
            requestWo.origPriority = wo.origPriority ?? requestWo.origPriority;
            requestWo.origDocType = wo.origDocType ?? requestWo.origDocType;
            requestWo.origDocNo = wo.origDocNo ?? requestWo.origDocNo;
            requestWo.relatedWo = wo.GetRelatedWoDto();
            requestWo.requestId = wo.requestId ?? requestWo.requestId;

            requestWo.stdJobNo = wo.stdJobNo ?? requestWo.stdJobNo;
            requestWo.maintenanceSchedTask = wo.maintSchTask ?? requestWo.maintenanceSchedTask;
            requestWo.autoRequisitionInd = !string.IsNullOrWhiteSpace(wo.autoRequisitionInd) && wo.autoRequisitionInd.Equals("Y");
            requestWo.assignPerson = wo.assignPerson ?? requestWo.assignPerson;
            requestWo.planPriority = wo.planPriority ?? requestWo.planPriority;
            requestWo.requisitionStartDate = wo.requisitionStartDate ?? requestWo.requisitionStartDate;
            requestWo.requisitionStartTime = wo.requisitionStartTime ?? requestWo.requisitionStartTime;
            requestWo.requiredByDate = wo.requiredByDate ?? requestWo.requiredByDate;
            requestWo.requiredByTime = wo.requiredByTime ?? requestWo.requiredByTime;
            requestWo.planStrDate = wo.planStrDate ?? requestWo.planStrDate;
            requestWo.planStrTime = wo.planStrTime ?? requestWo.planStrTime;
            requestWo.planFinDate = wo.planFinDate ?? requestWo.planFinDate;
            requestWo.planFinTime = wo.planFinTime ?? requestWo.planFinTime;
            requestWo.unitOfWork = wo.unitOfWork ?? requestWo.unitOfWork;
            requestWo.unitsRequired = !string.IsNullOrWhiteSpace(wo.unitsRequired) ? Convert.ToDecimal(wo.unitsRequired) : default(decimal);
            requestWo.unitsRequiredSpecified = !string.IsNullOrEmpty(wo.unitsRequired);
            requestWo.accountCode = wo.accountCode ?? requestWo.accountCode;
            requestWo.projectNo = wo.projectNo ?? requestWo.projectNo;
            requestWo.parentWo = wo.parentWo ?? requestWo.parentWo;

            requestWo.failurePart = wo.failurePart ?? requestWo.failurePart;
            requestWo.jobCode1 = wo.jobCode1 ?? requestWo.jobCode1;
            requestWo.jobCode2 = wo.jobCode2 ?? requestWo.jobCode2;
            requestWo.jobCode3 = wo.jobCode3 ?? requestWo.jobCode3;
            requestWo.jobCode4 = wo.jobCode4 ?? requestWo.jobCode4;
            requestWo.jobCode5 = wo.jobCode5 ?? requestWo.jobCode5;
            requestWo.jobCode6 = wo.jobCode6 ?? requestWo.jobCode6;
            requestWo.jobCode7 = wo.jobCode7 ?? requestWo.jobCode7;
            requestWo.jobCode8 = wo.jobCode8 ?? requestWo.jobCode8;
            requestWo.jobCode9 = wo.jobCode9 ?? requestWo.jobCode9;
            requestWo.jobCode10 = wo.jobCode10 ?? requestWo.jobCode10;
            requestWo.locationFr = wo.locationFr ?? requestWo.locationFr;
            requestWo.location = wo.location ?? requestWo.location;
            requestWo.noticeLocn = wo.noticeLocn ?? requestWo.noticeLocn;

            requestWo.calculatedDurationsFlag = Utils.IsTrue(wo.calculatedDurationsFlag, true);
            requestWo.calculatedDurationsFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedDurationsFlag);
            requestWo.calculatedLabFlag = Utils.IsTrue(wo.calculatedLabFlag, true);
            requestWo.calculatedLabFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedLabFlag);
            requestWo.calculatedMatFlag = Utils.IsTrue(wo.calculatedMatFlag, true);
            requestWo.calculatedMatFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedMatFlag);
            requestWo.calculatedEquipmentFlag = Utils.IsTrue(wo.calculatedEquipmentFlag, true);
            requestWo.calculatedEquipmentFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedEquipmentFlag);
            requestWo.calculatedOtherFlag = Utils.IsTrue(wo.calculatedOtherFlag, true);
            requestWo.calculatedOtherFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedOtherFlag);
            //se envía la acción
            return proxyWo.create(opContext, requestWo);
        }

        /// <summary>
        /// Actualiza/Mofica una Orden de Trabajo existente
        /// </summary>
        /// <param name="urlService">String: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkOrder")</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de Operación de la WorkOrder</param>
        /// <param name="wo"></param>
        public static WorkOrderServiceModifyReplyDTO ModifyWorkOrder(string urlService, OperationContext opContext, WorkOrder wo)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceModifyRequestDTO();

            proxyWo.Url = urlService;
            //se cargan los parámetros de la orden
            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = wo.districtCode ?? requestWo.districtCode;

            requestWo.workGroup = wo.workGroup ?? requestWo.workGroup;
            requestWo.workOrder = wo.GetWorkOrderDto();
            requestWo.workOrderDesc = wo.workOrderDesc ?? requestWo.workOrderDesc;
            //workOrderStatusM //n/a para modificar
            requestWo.equipmentNo = wo.equipmentNo ?? requestWo.equipmentNo;
            requestWo.equipmentRef = wo.equipmentRef ?? requestWo.equipmentRef;
            requestWo.compCode = wo.compCode ?? requestWo.compCode;
            requestWo.compModCode = wo.compModCode ?? requestWo.compModCode;
            requestWo.workOrderType = wo.workOrderType ?? requestWo.workOrderType;
            requestWo.maintenanceType = wo.maintenanceType ?? requestWo.maintenanceType;
            requestWo.workOrderStatusU = wo.workOrderStatusU ?? requestWo.workOrderStatusU;

            requestWo.raisedDate = wo.raisedDate ?? requestWo.raisedDate;
            requestWo.raisedTime = wo.raisedTime ?? requestWo.raisedTime;
            requestWo.originatorId = wo.originatorId ?? requestWo.originatorId;
            requestWo.origPriority = wo.origPriority ?? requestWo.origPriority;
            requestWo.origDocType = wo.origDocType ?? requestWo.origDocType;
            requestWo.origDocNo = wo.origDocNo ?? requestWo.origDocNo;
            requestWo.relatedWo = wo.GetRelatedWoDto();
            requestWo.requestId = wo.requestId ?? requestWo.requestId;

            requestWo.stdJobNo = wo.stdJobNo ?? requestWo.stdJobNo;
            requestWo.maintenanceSchedTask = wo.maintSchTask ?? requestWo.maintenanceSchedTask;
            requestWo.autoRequisitionInd = !string.IsNullOrWhiteSpace(wo.autoRequisitionInd) && wo.autoRequisitionInd.Equals("Y");
            requestWo.assignPerson = wo.assignPerson ?? requestWo.assignPerson;
            requestWo.planPriority = wo.planPriority ?? requestWo.planPriority;

            requestWo.requisitionStartDate = wo.requisitionStartDate ?? requestWo.requisitionStartDate;
            requestWo.requisitionStartTime = wo.requisitionStartTime ?? requestWo.requisitionStartTime;
            requestWo.requiredByDate = wo.requiredByDate ?? requestWo.requiredByDate;
            requestWo.requiredByTime = wo.requiredByTime ?? requestWo.requiredByTime;
            requestWo.planStrDate = wo.planStrDate ?? requestWo.planStrDate;
            requestWo.planStrTime = wo.planStrTime ?? requestWo.planStrTime;
            requestWo.planFinDate = wo.planFinDate ?? requestWo.planFinDate;
            requestWo.planFinTime = wo.planFinTime ?? requestWo.planFinTime;

            requestWo.unitOfWork = wo.unitOfWork ?? requestWo.unitOfWork;
            requestWo.unitsRequired = !string.IsNullOrWhiteSpace(wo.unitsRequired) ? Convert.ToDecimal(wo.unitsRequired) : default(decimal);
            requestWo.unitsRequiredSpecified = !string.IsNullOrEmpty(wo.unitsRequired);

            requestWo.accountCode = wo.accountCode ?? requestWo.accountCode;
            requestWo.projectNo = wo.projectNo ?? requestWo.projectNo;
            requestWo.parentWo = wo.parentWo ?? requestWo.parentWo;

            requestWo.failurePart = wo.failurePart ?? requestWo.failurePart;
            requestWo.jobCode1 = wo.jobCode1 ?? requestWo.jobCode1;
            requestWo.jobCode2 = wo.jobCode2 ?? requestWo.jobCode2;
            requestWo.jobCode3 = wo.jobCode3 ?? requestWo.jobCode3;
            requestWo.jobCode4 = wo.jobCode4 ?? requestWo.jobCode4;
            requestWo.jobCode5 = wo.jobCode5 ?? requestWo.jobCode5;
            requestWo.jobCode6 = wo.jobCode6 ?? requestWo.jobCode6;
            requestWo.jobCode7 = wo.jobCode7 ?? requestWo.jobCode7;
            requestWo.jobCode8 = wo.jobCode8 ?? requestWo.jobCode8;
            requestWo.jobCode9 = wo.jobCode9 ?? requestWo.jobCode9;
            requestWo.jobCode10 = wo.jobCode10 ?? requestWo.jobCode10;
            requestWo.locationFr = wo.locationFr ?? requestWo.locationFr;
            requestWo.location = wo.location ?? requestWo.location;
            requestWo.noticeLocn = wo.noticeLocn ?? requestWo.noticeLocn;

            requestWo.calculatedDurationsFlag = Convert.ToBoolean(wo.calculatedDurationsFlag);
            requestWo.calculatedDurationsFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedDurationsFlag);
            //
            if (wo.calculatedLabFlag == null && wo.calculatedMatFlag == null && wo.calculatedEquipmentFlag == null && wo.calculatedOtherFlag == null)
                return proxyWo.modify(opContext, requestWo);

            var requestEstimates = new WorkOrderServiceUpdateEstimatesRequestDTO
            {
                districtCode = wo.districtCode,
                workOrder = wo.GetWorkOrderDto(),
                calculatedLabFlag = Convert.ToBoolean(wo.calculatedLabFlag),
                calculatedLabFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedLabFlag),
                calculatedMatFlag = Convert.ToBoolean(wo.calculatedMatFlag),
                calculatedMatFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedMatFlag),
                calculatedEquipmentFlag = Convert.ToBoolean(wo.calculatedEquipmentFlag),
                calculatedEquipmentFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedEquipmentFlag),
                calculatedOtherFlag = Convert.ToBoolean(wo.calculatedOtherFlag),
                calculatedOtherFlagSpecified = !string.IsNullOrWhiteSpace(wo.calculatedOtherFlag),
            };

            proxyWo.updateEstimates(opContext, requestEstimates);
            //se envía la acción final
            return proxyWo.modify(opContext, requestWo);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="urlService"></param>
        /// <param name="opContext"></param>
        /// <param name="wo">WorkOrderCompleteAtributes:</param>
        /// <param name="appendCloseComment">bool: true para adicionar el texto de completeCommentToAppend a los comentarios de cierre</param>
        /// <returns></returns>
        public static WorkOrderServiceCompleteReplyDTO CompleteWorkOrder(string urlService, OperationContext opContext, WorkOrderCompleteAtributes wo, bool appendCloseComment = true)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceCompleteRequestDTO();
            proxyWo.Url = urlService + "/WorkOrder";
            long number1;
            if (long.TryParse("" + wo.workOrder.prefix + wo.workOrder.no, out number1))
                wo.workOrder = GetNewWorkOrderDto(("" + wo.workOrder.prefix + wo.workOrder.no).PadLeft(8, '0'));

            //se cargan los parámetros de la orden
            requestWo.workOrder = wo.workOrder;
            requestWo.districtCode = wo.districtCode;
            requestWo.completedBy = wo.completedBy;
            requestWo.completedCode = wo.completedCode;
            requestWo.closedDate = wo.closedDate;
            requestWo.closedTime = wo.closedTime;
            requestWo.outServDate = wo.outServDate;
            requestWo.outServTime = wo.outServTime;
            requestWo.earnCode = wo.earnCode;
            requestWo.failurePart = wo.failurePart;
            requestWo.hoursCompleted = wo.hoursCompleted;
            requestWo.hoursCompletedSpecified = wo.hoursCompletedSpecified;
            requestWo.crteInsitu = wo.crteInsitu;
            requestWo.crteInsituSpecified = wo.crteInsituSpecified;

            //se envía la acción
            var replyWo = proxyWo.complete(opContext, requestWo);

            //comentario
            if (!appendCloseComment || string.IsNullOrWhiteSpace(wo.completeCommentToAppend)) return replyWo;
            AppendTextToCloseComment(urlService, opContext, replyWo.districtCode, replyWo.workOrder.prefix + replyWo.workOrder.no, wo.completeCommentToAppend);
            //
            return replyWo;
        }

        public static void AppendTextToCloseComment(string urlService, OperationContext opContext, string district, string workOrder, string textToAppend)
        {
            var stdTextId = "CW" + district + workOrder;

            var stdTextCopc = StdText.GetCustomOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            var woCompleteComment = StdText.GetText(urlService, stdTextCopc, stdTextId);

            StdText.SetText(urlService, stdTextCopc, stdTextId, woCompleteComment + "\n" + textToAppend);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="wo">WorkOrder: Elemento WorkOrder com los campos de districtCode, workOrderDTO</param>
        /// <returns></returns>
        public static WorkOrderServiceReopenReplyDTO ReOpenWorkOrder(string urlService, OperationContext opContext, WorkOrder wo)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceReopenRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = wo.districtCode;
            requestWo.workOrder = wo.GetWorkOrderDto();
            //se envía la acción
            var replyWo = proxyWo.reopen(opContext, requestWo);
            //
            return replyWo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="wo">WorkOrder: Elemento WorkOrder com los campos de districtCode, workOrderDTO</param>
        /// <returns></returns>
        public static WorkOrderServiceFinaliseReplyDTO FinalizeWorkOrder(string urlService, OperationContext opContext, WorkOrder wo)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceFinaliseRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = wo.districtCode;
            requestWo.workOrder = wo.GetWorkOrderDto();
            requestWo.finalCosts = true;
            requestWo.finalCostsSpecified = true;
            //se envía la acción
            var replyWo = proxyWo.finalise(opContext, requestWo);
            //
            return replyWo;
        }

        public static string GetWorkOrderCloseText(string urlService, string districtCode, string position, bool returnWarnings, WorkOrderDTO wo)
        {
            //comentario
            var stdTextId = "CW" + districtCode + wo.prefix + wo.no;
            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);
            return StdText.GetText(urlService, stdTextCopc, stdTextId);
        }

        public static void SetWorkOrderCloseText(string urlService, string districtCode, string position, bool returnWarnings, WorkOrderDTO wo, string woCloseText)
        {
            //comentario
            var stdTextId = "CW" + districtCode + wo.prefix + wo.no;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, woCloseText);
        }
        /// <summary>
        /// Crea un nuevo registro de duración para una orden de trabajo especificada
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="districtCode">string: Código del distrito</param>
        /// <param name="workOrder">WorkOrderService.WorkOrderDTO: Orden a la que se le realizará la acción</param>
        /// <param name="duration">WorkOrderDuration: duracion que se le creará a la orden</param>
        /// <returns>/// <returns>WorkOrderService.WorkOrderServiceCreateWorkOrderDurationReplyDTO: respuesta del proceso</returns></returns>
        public static WorkOrderServiceCreateWorkOrderDurationReplyDTO CreateWorkOrderDuration(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder, WorkOrderDuration duration)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceCreateWorkOrderDurationRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            if (string.IsNullOrWhiteSpace(districtCode))
                throw new Exception("DISTRICT REQUIRED");
            requestWo.districtCode = districtCode;
            requestWo.workOrder = workOrder;
            requestWo.durations = new DurationsDTO[1];
            requestWo.durations[0] = duration.GetDurationDto();

            //se envía la acción
            var replyWo = proxyWo.createWorkOrderDuration(opContext, requestWo);
            //
            return replyWo;
        }

        /// <summary>
        /// Modifica los registros de duraciones para una orden de trabajo especificada
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="districtCode">string: Código del distrito</param>
        /// <param name="workOrder">WorkOrderService.WorkOrderDTO: Orden a la que se le realizará la acción</param>
        /// <param name="durations">WorkOrderDuration[]: Arreglo de duraciones que se le asignarán a la orden</param>
        /// <returns>WorkOrderService.WorkOrderServiceModifyWorkOrderDurationReplyDTO: respuesta del proceso</returns>
        public static WorkOrderServiceModifyWorkOrderDurationReplyDTO ModifyWorkOrderDuration(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder, WorkOrderDuration[] durations)
        {
            //Observación: Los campos de una duración son Fecha, Código, Hora Inicial, Hora final. El campo de secuencia no incide en el funcionamiento del proceso
            //por esta razón, no se puede determinar una modificación singular, sino que debe modificarse el arreglo con los campos especificados.
            //No se puede establecer una llave lógica que determine o no si al modificar se está hablando de la misma duración.
            //Para este caso, utilice el método sobrecargado que tiene Duración a Modificar y Nueva Duración que hace uso de eliminar y crear porque el campo de fecha no se deja modificar
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceModifyWorkOrderDurationRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = districtCode;
            requestWo.workOrder = workOrder;

            requestWo.durations = new DurationsDTO[durations.Length];
            for (var i = 0; i < durations.Length; i++)
                requestWo.durations[i] = durations[i].GetDurationDto();

            //se envía la acción
            var replyWo = proxyWo.modifyWorkOrderDuration(opContext, requestWo);
            //
            return replyWo;
        }

        /// <summary>
        /// Modifica un registro de duración para una orden de trabajo especificada. Este método hace uso de las funciones de eliminar y crear
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="districtCode">string: Código del distrito</param>
        /// <param name="workOrder">WorkOrderService.WorkOrderDTO: Orden a la que se le realizará la acción</param>
        /// <param name="oldDuration">WorkOrderDuration: Duración a modificar</param>
        /// <param name="newDuration">WorkOrderDuration: Nuevos parámetros de duración</param>
        /// <returns>WorkOrderService.WorkOrderServiceCreateWorkOrderDurationReplyDTO: respuesta del proceso</returns>
        public static WorkOrderServiceCreateWorkOrderDurationReplyDTO ModifyWorkOrderDuration(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder, WorkOrderDuration oldDuration, WorkOrderDuration newDuration)
        {
            //Observación: Los campos de una duración son Fecha, Código, Hora Inicial, Hora final. El campo de secuencia no incide en el funcionamiento del proceso
            //por esta razón, no se puede determinar una modificación singular, sino que debe modificarse completamente el arreglo con los campos especificados.
            //No se puede establecer una llave lógica que determine o no si al modificar se está hablando de la misma duración.
            //Para este caso, utilice el método sobrecargado que tiene Duración a Modificar y Nueva Duración que hace uso de eliminar y crear porque el campo de fecha no se deja modificar
            var proxyWo = new WorkOrderService.WorkOrderService { Url = urlService + "/WorkOrder" };
            //ejecuta las acciones del 


            //se consultan las duraciones existentes
            var requestRwo = new WorkOrderServiceRetrieveWorkOrderDurationRequestDTO
            {
                districtCode = districtCode,
                workOrder = workOrder
            };
            var replyRwo = proxyWo.retrieveWorkOrderDuration(opContext, requestRwo);

            //se ubica el elemento que se quiere modificar
            for (var i = 0; i < replyRwo.durations.Length; i++)
            {
                if (replyRwo.durations[i].jobDurationsDate != oldDuration.jobDurationsDate ||
                    replyRwo.durations[i].jobDurationsCode != oldDuration.jobDurationsCode ||
                    replyRwo.durations[i].jobDurationsStart != oldDuration.jobDurationsStart ||
                    replyRwo.durations[i].jobDurationsFinish != oldDuration.jobDurationsFinish) continue;
                var delDuration = new WorkOrderDuration();
                delDuration.SetDurationFromDto(replyRwo.durations[i]);
                DeleteWorkOrderDuration(urlService, opContext, districtCode, workOrder, delDuration);

                return CreateWorkOrderDuration(urlService, opContext, districtCode, workOrder, newDuration);
            }
            return null;
        }

        /// <summary>
        /// Elimina un registro de duración para una orden de trabajo especificada
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="districtCode">string: Código del distrito</param>
        /// <param name="workOrder">WorkOrderService.WorkOrderDTO: Orden a la que se le realizará la acción</param>
        /// <param name="duration">WorkOrderDuration: duracion que se le eliminará a la orden</param>
        /// <returns>WorkOrderService.WorkOrderServiceDeleteWorkOrderDurationReplyDTO: Respuesta del proceso</returns>
        public static WorkOrderServiceDeleteWorkOrderDurationReplyDTO DeleteWorkOrderDuration(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder, WorkOrderDuration duration)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceDeleteWorkOrderDurationRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se consultan las duraciones existentes
            var requestRwo = new WorkOrderServiceRetrieveWorkOrderDurationRequestDTO
            {
                districtCode = districtCode,
                workOrder = workOrder
            };
            var replyRwo = proxyWo.retrieveWorkOrderDuration(opContext, requestRwo);

            //se ubica el elemento que se quiere eliminar
            var delDuration = new DurationsDTO();

            for (var i = 0; i < replyRwo.durations.Length; i++)
            {
                if (replyRwo.durations[i].jobDurationsDate == duration.jobDurationsDate
                    && replyRwo.durations[i].jobDurationsCode == duration.jobDurationsCode
                    && replyRwo.durations[i].jobDurationsStart == duration.jobDurationsStart
                    && replyRwo.durations[i].jobDurationsFinish == duration.jobDurationsFinish)
                {
                    delDuration = replyRwo.durations[i];
                    break;
                }
            }

            //se cargan los parámetros de la orden
            requestWo.districtCode = districtCode;
            requestWo.workOrder = workOrder;

            requestWo.durations = new DurationsDTO[1];
            requestWo.durations[0] = delDuration;

            //se envía la acción
            var replyWo = proxyWo.deleteWorkOrderDuration(opContext, requestWo);
            //
            return replyWo;
        }

        /// <summary>
        /// Obtiene el listado de duraciones de una orden de trabajo
        /// </summary>
        /// <param name="urlService">string: URL a los web services de Ellipse</param>
        /// <param name="opContext">WorkOrderService.OperationContext: Contexto de operación</param>
        /// <param name="districtCode">string: Código del distrito</param>
        /// <param name="workOrder">WorkOrderService.WorkOrderDTO: Orden a la que se le realizará la acción</param>
        /// <returns>WorkOrderDuration []: arreglo con los elementos de duración</returns>
        public static WorkOrderDuration[] GetWorkOrderDurations(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio
            var requestWo = new WorkOrderServiceRetrieveWorkOrderDurationRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se consultan las duraciones existentes
            requestWo.districtCode = districtCode;
            requestWo.workOrder = workOrder;
            var replyWo = proxyWo.retrieveWorkOrderDuration(opContext, requestWo);

            var durationList = new List<WorkOrderDuration>();

            // ReSharper disable once LoopCanBeConvertedToQuery
            for (var i = 0; i < replyWo.durations.Length; i++)
            {
                if (replyWo.durations[i].jobDurationsDate == "")
                    break;
                var dur = new WorkOrderDuration
                {
                    jobDurationsCode = replyWo.durations[i].jobDurationsCode,
                    jobDurationsDate = replyWo.durations[i].jobDurationsDate,
                    jobDurationsStart = replyWo.durations[i].jobDurationsStart,
                    jobDurationsFinish = replyWo.durations[i].jobDurationsFinish,
                    jobDurationsSeqNo = replyWo.durations[i].jobDurationsSeqNo,
                    jobDurationsSeqNoSpecified = replyWo.durations[i].jobDurationsSeqNoSpecified,
                    jobDurationsHours = replyWo.durations[i].jobDurationsHours,
                    jobDurationsHoursSpecified = replyWo.durations[i].jobDurationsHoursSpecified
                };
                durationList.Add(dur);
            }


            return durationList.ToArray();
        }

        public static WorkOrderReferenceCodes GetWorkOrderReferenceCodes(EllipseFunctions ef, string urlService, OperationContext opContext, string district, string workOrder)
        {

            var woRefCodes = new WorkOrderReferenceCodes();

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            const string entityType = "WKO";
            var entityValue = "1" + district + workOrder;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            var newef = new EllipseFunctions(ef);
            newef.SetConnectionPoolingType(false);
            //
            var item001 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "001");
            var item002 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "002", "001");
            var item003 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "003", "001");
            var item005 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "005", "001");
            var item006 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "006", "001");
            var item007 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "007", "001");
            var item008 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "008", "001");
            var item009 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "009", "001");
            var item010 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "010", "001");
            var item011 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "011", "001");
            var item012 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "012", "001");
            var item013 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "013", "001");
            var item014 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "014", "001");
            var item015 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "015", "001");
            var item016 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "016", "001");
            var item017 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "017", "001");
            var item018 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "018", "001");
            var item019 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "019", "001");
            var item020 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "020", "001");
            var item021 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "021", "001");
            var item022 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "022", "001");
            var item024 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "024", "001");
            var item025 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "025", "001");
            var item026 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "026", "001");
            var item029 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "029", "001");
            var item030 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "030", "001");
            var item031 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "031", "001");



            woRefCodes.WorkRequest = item001.RefCode; //001_9001
            woRefCodes.ComentariosDuraciones = item002.RefCode; //002_9001
            woRefCodes.ComentariosDuracionesText = item002.StdText; //002_9001
            woRefCodes.EmpleadoId = item003.RefCode; //003_001
            woRefCodes.NroComponente = item005.RefCode; //005_9001
            woRefCodes.P1EqLivMed = item006.RefCode; //006_001
            woRefCodes.P2EqMovilMinero = item007.RefCode; //007_9001
            woRefCodes.P3ManejoSustPeligrosa = item008.RefCode; //008_9001
            woRefCodes.P4GuardasEquipo = item009.RefCode; //009_9001
            woRefCodes.P5Aislamiento = item010.RefCode; //010_9001
            woRefCodes.P6TrabajosAltura = item011.RefCode; //011_9001
            woRefCodes.P7ManejoCargas = item012.RefCode; //012_9001
            woRefCodes.ProyectoIcn = item013.RefCode; //013_9001
            woRefCodes.Reembolsable = item014.RefCode; //014_9001
            woRefCodes.FechaNoConforme = item015.RefCode; //015_9001
            woRefCodes.FechaNoConformeText = item015.StdText; //015_9001
            woRefCodes.NoConforme = item016.RefCode; //016_001
            woRefCodes.FechaEjecucion = item017.RefCode; //017_001
            woRefCodes.HoraIngreso = item018.RefCode; //018_9001
            woRefCodes.HoraSalida = item019.RefCode; //019_9001
            woRefCodes.NombreBuque = item020.RefCode; //020_9001
            woRefCodes.CalificacionEncuesta = item021.RefCode; //021_001
            woRefCodes.TareaCritica = item022.RefCode; //022_001
            woRefCodes.Garantia = item024.RefCode; //024_9001
            woRefCodes.GarantiaText = item024.StdText; //024_9001
            woRefCodes.CodigoCertificacion = item025.RefCode; //025_001
            woRefCodes.FechaEntrega = item026.RefCode; //026_001
            woRefCodes.RelacionarEv = item029.RefCode; //029_001
            woRefCodes.Departamento = item030.RefCode; //030_9001
            woRefCodes.Localizacion = item031.RefCode; //031_001

            newef.CloseConnection();
            return woRefCodes;
        }

        public static ReplyMessage UpdateWorkOrderReferenceCodes(EllipseFunctions eFunctions, string urlService, OperationContext opContext, string district, string workOrder, WorkOrderReferenceCodes woRefCodes)
        {
            var reply = new ReplyMessage();
            var error = new List<string>();

            const string entityType = "WKO";
            var entityValue = "1" + district + workOrder;
            var itemList = new List<ReferenceCodeItem>();

            var item001 = new ReferenceCodeItem(entityType, entityValue, "001", "001", woRefCodes.WorkRequest) { ShortName = "WorkRequest" };
            var item002 = new ReferenceCodeItem(entityType, entityValue, "002", "001", woRefCodes.ComentariosDuraciones, null, woRefCodes.ComentariosDuracionesText) { ShortName = "ComentariosDur" };
            var item003 = new ReferenceCodeItem(entityType, entityValue, "003", "001", woRefCodes.EmpleadoId) { ShortName = "EmpleadoId" };
            var item005 = new ReferenceCodeItem(entityType, entityValue, "005", "001", woRefCodes.NroComponente) { ShortName = "NroComponente" };
            var item006 = new ReferenceCodeItem(entityType, entityValue, "006", "001", woRefCodes.P1EqLivMed) { ShortName = "P1EqLivMed" };
            var item007 = new ReferenceCodeItem(entityType, entityValue, "007", "001", woRefCodes.P2EqMovilMinero) { ShortName = "P2EqMovilMinero" };
            var item008 = new ReferenceCodeItem(entityType, entityValue, "008", "001", woRefCodes.P3ManejoSustPeligrosa) { ShortName = "P3ManejoSustPelig" };
            var item009 = new ReferenceCodeItem(entityType, entityValue, "009", "001", woRefCodes.P4GuardasEquipo) { ShortName = "P4GuardasEquipo" };
            var item010 = new ReferenceCodeItem(entityType, entityValue, "010", "001", woRefCodes.P5Aislamiento) { ShortName = "P5Aislamiento" };
            var item011 = new ReferenceCodeItem(entityType, entityValue, "011", "001", woRefCodes.P6TrabajosAltura) { ShortName = "P6TrabajosAltura" };
            var item012 = new ReferenceCodeItem(entityType, entityValue, "012", "001", woRefCodes.P7ManejoCargas) { ShortName = "P7ManejoCargas" };
            var item013 = new ReferenceCodeItem(entityType, entityValue, "013", "001", woRefCodes.ProyectoIcn) { ShortName = "ProyectoIcn" };
            var item014 = new ReferenceCodeItem(entityType, entityValue, "014", "001", woRefCodes.Reembolsable) { ShortName = "Reembolsable" };
            var item015 = new ReferenceCodeItem(entityType, entityValue, "015", "001", woRefCodes.FechaNoConforme, null, woRefCodes.FechaNoConformeText) { ShortName = "FechaNoConforme" };
            var item016 = new ReferenceCodeItem(entityType, entityValue, "016", "001", woRefCodes.NoConforme) { ShortName = "NoConforme" };
            var item017 = new ReferenceCodeItem(entityType, entityValue, "017", "001", woRefCodes.FechaEjecucion) { ShortName = "FechaEjecucion" };
            var item018 = new ReferenceCodeItem(entityType, entityValue, "018", "001", woRefCodes.HoraIngreso) { ShortName = "HoraIngreso" };
            var item019 = new ReferenceCodeItem(entityType, entityValue, "019", "001", woRefCodes.HoraSalida) { ShortName = "HoraSalida" };
            var item020 = new ReferenceCodeItem(entityType, entityValue, "020", "001", woRefCodes.NombreBuque) { ShortName = "NombreBuque" };
            var item021 = new ReferenceCodeItem(entityType, entityValue, "021", "001", woRefCodes.CalificacionEncuesta) { ShortName = "CalifEncuesta" };
            var item022 = new ReferenceCodeItem(entityType, entityValue, "022", "001", woRefCodes.TareaCritica) { ShortName = "TareaCritica" };
            var item024 = new ReferenceCodeItem(entityType, entityValue, "024", "001", woRefCodes.Garantia, null, woRefCodes.GarantiaText) { ShortName = "Garantia" };
            var item025 = new ReferenceCodeItem(entityType, entityValue, "025", "001", woRefCodes.CodigoCertificacion) { ShortName = "CodCertificacion" };
            var item026 = new ReferenceCodeItem(entityType, entityValue, "026", "001", woRefCodes.FechaEntrega) { ShortName = "FechaEntrega" };
            var item029 = new ReferenceCodeItem(entityType, entityValue, "029", "001", woRefCodes.RelacionarEv) { ShortName = "RelacionarEv" };
            var item030 = new ReferenceCodeItem(entityType, entityValue, "030", "001", woRefCodes.Departamento) { ShortName = "Departamento" };
            var item031 = new ReferenceCodeItem(entityType, entityValue, "031", "001", woRefCodes.Localizacion) { ShortName = "Localizacion" };

            itemList.Add(item001);
            itemList.Add(item002);
            itemList.Add(item003);
            itemList.Add(item005);
            itemList.Add(item006);
            itemList.Add(item007);
            itemList.Add(item008);
            itemList.Add(item009);
            itemList.Add(item010);
            itemList.Add(item011);
            itemList.Add(item012);
            itemList.Add(item013);
            itemList.Add(item014);
            itemList.Add(item015);
            itemList.Add(item016);
            itemList.Add(item017);
            itemList.Add(item018);
            itemList.Add(item019);
            itemList.Add(item020);
            itemList.Add(item021);
            itemList.Add(item022);
            itemList.Add(item024);
            itemList.Add(item025);
            itemList.Add(item026);
            itemList.Add(item029);
            itemList.Add(item030);
            itemList.Add(item031);

            var refCodeOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);

            foreach (var item in itemList)
            {
                try
                {
                    if (item.RefCode == null)
                        continue;
                    var replyRefCode = ReferenceCodeActions.ModifyRefCode(eFunctions, urlService, refCodeOpContext, item);
                    var stdTextId = replyRefCode.stdTxtKey;
                    if (string.IsNullOrWhiteSpace(stdTextId))
                        throw new Exception("No se recibió respuesta");
                }
                catch (Exception ex)
                {
                    error.Add("Error al actualizar " + item.ShortName + ": " + ex.Message);
                }
            }

            reply.Errors = error.ToArray();
            return reply;
        }

        public static WorkOrderServiceRecordWorkProgressReplyDTO RecordWorkProgress(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder, string percentComplete, string unitsComplete, string unitsRequired = null)
        {
            var proxyWo = new WorkOrderService.WorkOrderService();//ejecuta las acciones del servicio

            var requestWo = new WorkOrderServiceRecordWorkProgressRequestDTO();

            proxyWo.Url = urlService + "/WorkOrder";

            //se cargan los parámetros de la orden
            requestWo.districtCode = districtCode;
            requestWo.workOrder = workOrder;
            requestWo.pcComplete = !string.IsNullOrWhiteSpace(percentComplete) ? Convert.ToDecimal(percentComplete) : default(decimal);
            requestWo.pcCompleteSpecified = !string.IsNullOrEmpty(percentComplete);
            requestWo.unitsComplete = !string.IsNullOrWhiteSpace(unitsComplete) ? Convert.ToDecimal(unitsComplete) : default(decimal);
            requestWo.unitsCompleteSpecified = !string.IsNullOrEmpty(unitsComplete);
            requestWo.unitsRequired = !string.IsNullOrWhiteSpace(unitsRequired) ? Convert.ToDecimal(unitsRequired) : default(decimal);
            requestWo.unitsRequiredSpecified = !string.IsNullOrEmpty(unitsRequired);
            //se envía la acción
            var replyWo = proxyWo.recordWorkProgress(opContext, requestWo);
            //
            return replyWo;
        }
        public static WorkOrderServiceRecordWorkProgressReplyDTO CompleteWorkProgress(string urlService, OperationContext opContext, string districtCode, WorkOrderDTO workOrder)
        {
            return RecordWorkProgress(urlService, opContext, districtCode, workOrder, "100", null);
        }

        /// <summary>
        /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
        /// </summary>
        /// <param name="no">string: Número de la orden de trabajo</param>
        /// <returns>WorkOrderService.WorkOrderDTO</returns>
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
        /// <summary>
        /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
        /// </summary>
        /// <param name="prefix">string: prefijo de la orden de trabajo</param>
        /// <param name="no">string: Número de la orden de trabajo</param>
        /// <returns>WorkOrderService.WorkOrderDTO</returns>
        public static WorkOrderDTO GetNewWorkOrderDto(string prefix, string no)
        {
            var workOrderDto = new WorkOrderDTO
            {
                prefix = prefix,
                no = no
            };

            return workOrderDto;
        }

        /// <summary>
        /// Valida si el uso del estado de usuario de orden está bien relacionado para tiempo de apertura de una orden 
        /// </summary>
        /// <param name="raisedDate">string: fecha en formato yyyyMMdd de apertura de una orden</param>
        /// <param name="daysAllowed">int: número de días válidos en el que una orden puede estar abierta sin necesidad de justificar su estado</param>
        /// <returns></returns>
        public static bool ValidateUserStatus(string raisedDate, int daysAllowed)
        {
            var date = DateTime.ParseExact(raisedDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            return DateTime.Today.Subtract(date).TotalDays <= daysAllowed;
        }

        /// <summary>
        /// Obtiene el listado {código, descripción} de los User Status Codes (MSF010WS)
        /// </summary>
        /// <param name="ef"></param>
        /// <returns>List[EllipseCodeItem]: Diccionario{codigo, descripción} del listado de códigos</returns>
        public static List<EllipseCodeItem> GetUserStatusCodeList(EllipseFunctions ef)
        {
            return ef.GetItemCodes("WS");
        }
        public class Queries
        {
            /// <summary>
            /// Obtiene el Query para la consulta de una o más órdenes de trabajo
            /// </summary>
            /// <param name="dbReference">string: Referencia de base de datos (Ej: MIMSPROD, ELLIPSE) </param>
            /// <param name="dbLink">string: link de conexión de base de datos (Ej: @MLDBMIMS)</param>
            /// <param name="districtCode">string: distrito de consulta. Si es nulo se consulta para todos los distritos</param>
            /// <param name="searchCriteriaKey1">int: Indica el tipo de búsqueda según la clase SearchFieldCriteriaType.Type.Key. Valor por defecto (0 - None). (Ej: SearchFieldCriteriaType.WorkGroup.Key) </param>
            /// <param name="searchCriteriaValue1"></param>
            /// <param name="searchCriteriaKey2"></param>
            /// <param name="searchCriteriaValue2"></param>
            /// <param name="dateCriteriaKey"></param>
            /// <param name="startDate">string: fecha en format yyyyMMdd para parámetro de fecha inicial. Predeterminado inicio del año</param>
            /// <param name="endDate">string: fecha en format yyyyMMdd para parámetro de fecha final. Predeterminado fecha de hoy</param>
            /// <param name="woStatus">string: especifica qué estado de la orden se va a consultar WoStatusList.StatusName. Si es nulo se consulta cualquier estado></param>
            /// <returns>string: Query de consulta para ejecución</returns>
            public static string GetFetchWoQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, int dateCriteriaKey, string startDate, string endDate, string woStatus)
            {
                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(districtCode))
                    districtCode = " IN (" + Utils.GetListInSeparator(DistrictConstants.GetDistrictList(), ",", "'") + ")";
                else
                    districtCode = " = '" + districtCode + "'";

                var queryCriteria1 = "";
                //establecemos los parámetros del criterio 1
                if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.WORK_GROUP = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.EQUIP_NO = '" + searchCriteriaValue1 + "'"; //Falta buscar el equip ref //to do
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.ORIGINATOR_ID = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.COMPLETED_BY = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND TRIM(SUBSTR(WO.DSTRCT_ACCT_CODE, 5)) = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkRequest.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.REQUEST_ID = '" + searchCriteriaValue1.PadLeft(12, '0') + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.PARENT_WO = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                    else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "')";
                }
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                    else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                }
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND EQ.EQUIP_CLASS = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
                else
                    queryCriteria1 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
                //

                var queryCriteria2 = "";
                //establecemos los parámetros del criterio 2
                if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.WORK_GROUP = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.EQUIP_NO = '" + searchCriteriaValue2 + "'"; //Falta buscar el equip ref //to do
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.ORIGINATOR_ID = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND WO.COMPLETED_BY = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND TRIM(SUBSTR(WO.DSTRCT_ACCT_CODE, 5)) = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkRequest.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.REQUEST_ID = '" + searchCriteriaValue2.PadLeft(12, '0') + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.PARENT_WO = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "'";
                    else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "'";
                }
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "'";
                    else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "'";
                }
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND EQ.EQUIP_CLASS = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
                //

                //establecemos los parámetros de estado de orden
                string statusRequirement;
                if (string.IsNullOrEmpty(woStatus))
                    statusRequirement = "";
                else if (woStatus == WoStatusList.Uncompleted)
                    statusRequirement = " AND WO.WO_STATUS_M IN (" + Utils.GetListInSeparator(WoStatusList.GetUncompletedStatusCodes(), ",", "'") + ")";
                else if (WoStatusList.GetStatusNames().Contains(woStatus))
                    statusRequirement = " AND WO.WO_STATUS_M = '" + WoStatusList.GetStatusCode(woStatus) + "'";
                else
                    statusRequirement = "";

                //establecemos los parámetros para el rango de fechas
                string dateParameters;
                if (string.IsNullOrEmpty(startDate))
                    startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                if (string.IsNullOrEmpty(endDate))
                    endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                if (dateCriteriaKey == SearchDateCriteriaType.Raised.Key)
                    dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.Closed.Key)
                    dateParameters = " AND WO.CLOSED_DT BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.RequiredBy.Key)
                    dateParameters = " AND WO.REQ_BY_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.RequiredStart.Key)
                    dateParameters = " AND WO.REQ_START_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.PlannedStart.Key)
                    dateParameters = " AND WO.PLAN_STR_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.PlannedFinnish.Key)
                    dateParameters = " AND WO.PLAN_FIN_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.LastModified.Key)
                    dateParameters = " AND WO.LAST_MOD_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.NotFinalized.Key)
                    dateParameters = " AND WO.CLOSED_DT BETWEEN '" + startDate + "' AND '" + endDate + "' AND WO.FINAL_COSTS <> 'Y'";
                else
                    dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                //escribimos el query
                var query = "" +
                            " SELECT" +
                            " WO.DSTRCT_CODE, WO.WORK_GROUP, WO.WORK_ORDER, WO.WO_STATUS_M, WO.WO_DESC, " +
                            " WO.EQUIP_NO, WO.COMP_CODE, WO.COMP_MOD_CODE, WO.LOCATION, WO.RAISED_DATE, WO.RAISED_TIME," +
                            " WO.ORIGINATOR_ID, WO.ORIG_PRIORITY, WO.ORIG_DOC_TYPE, WO.ORIG_DOC_NO, WO.REQUEST_ID, WO.MSSS_STATUS_IND," +
                            " WO.WO_TYPE, WO.MAINT_TYPE, WO.WO_STATUS_U, WO.STD_JOB_NO, WO.MAINT_SCH_TASK, WO.AUTO_REQ_IND, WO.ASSIGN_PERSON, WO.PLAN_PRIORITY, WO.CLOSED_COMMIT_DT, WO.UNIT_OF_WORK, WO.UNITS_REQUIRED, FAILURE_PART, PC_COMPLETE, UNITS_COMPLETE, WO.RELATED_WO," +
                            " WO.REQ_START_DATE, WO.REQ_START_TIME, WO.REQ_BY_DATE, WO.REQ_BY_TIME, WO.PLAN_STR_DATE, WO.PLAN_STR_TIME, WO.PLAN_FIN_DATE, WO.PLAN_FIN_TIME," +
                            " SUBSTR(WO.DSTRCT_ACCT_CODE, 5) DSTRCT_ACCT_CODE, WO.PROJECT_NO, WO.PARENT_WO," +
                            " WO.WO_JOB_CODEX1, WO.WO_JOB_CODEX2, WO.WO_JOB_CODEX3, WO.WO_JOB_CODEX4, WO.WO_JOB_CODEX5, WO.WO_JOB_CODEX6, WO.WO_JOB_CODEX7, WO.WO_JOB_CODEX8, WO.WO_JOB_CODEX9, WO.WO_JOB_CODEX10," +
                            " CASE WHEN TRIM(WO.WO_JOB_CODEX1||WO.WO_JOB_CODEX2||WO.WO_JOB_CODEX3||WO.WO_JOB_CODEX4||WO.WO_JOB_CODEX5||WO.WO_JOB_CODEX6||WO.WO_JOB_CODEX7||WO.WO_JOB_CODEX8||WO.WO_JOB_CODEX9||WO.WO_JOB_CODEX10) IS NULL THEN 'N' ELSE 'Y' END JOB_CODES," +
                            " WO.COMPLETED_CODE, WO.COMPLETED_BY," +
                            " CASE WHEN WO.DSTRCT_CODE || WO.WORK_ORDER NOT IN (SELECT STV.STD_KEY FROM " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " STV WHERE STV.STD_TEXT_CODE=('CW')) THEN 'N' ELSE 'Y' END COMPLETE_TEXT_FLAG," +
                            " WO.CLOSED_DT," +
                            " WOEST.CALC_DUR_HRS_SW, WOEST.EST_DUR_HRS, WOEST.ACT_DUR_HRS," +
                            " WOEST.RES_UPDATE_FLAG, WOEST.EST_LAB_HRS, WOEST.CALC_LAB_HRS, WOEST.ACT_LAB_HRS, WOEST.EST_LAB_COST, WOEST.CALC_LAB_COST, WOEST.ACT_LAB_COST," +
                            " WOEST.MAT_UPDATE_FLAG, WOEST.EST_MAT_COST, WOEST.CALC_MAT_COST, WOEST.ACT_MAT_COST," +
                            " WOEST.EQUIP_UPDATE_FLAG, WOEST.EST_EQUIP_COST, WOEST.CALC_EQUIP_COST, WOEST.ACT_EQUIP_COST," +
                            " WOEST.EST_OTHER_COST, WOEST.ACT_OTHER_COST," +
                            " WO.LOCATION_FR, WO.LOCATION, WO.NOTICE_LOCN," +
                            " WO.LAST_MOD_DATE, WO.FINAL_COSTS," +
                            " EQ.EQUIP_CLASS, EQ.EQUIP_GRP_ID, EQ.PARENT_EQUIP" +
                            " FROM" +
                            " " + dbReference + ".MSF620" + dbLink + " WO LEFT JOIN " + dbReference + ".MSF621" + dbLink + " WOEST ON (WO.WORK_ORDER    = WOEST.WORK_ORDER AND WO.DSTRCT_CODE = WOEST.DSTRCT_CODE)" +
                            " " + "LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON WO.EQUIP_NO = EQ.EQUIP_NO" +
                            " WHERE" +
                            " " + queryCriteria1 +
                            " " + queryCriteria2 +
                            " " + statusRequirement +
                            " AND WO.DSTRCT_CODE " + districtCode +
                            dateParameters;

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            /// <summary>
            /// Obtiene el Query para la consulta de una o más órdenes de trabajo
            /// </summary>
            /// <param name="dbReference">string: Referencia de base de datos (Ej: MIMSPROD, ELLIPSE) </param>
            /// <param name="dbLink">string: link de conexión de base de datos (Ej: @MLDBMIMS)</param>
            /// <param name="districtCode">string: distrito de consulta. Si es nulo se consulta para todos los distritos</param>
            /// <param name="workOrder">string: número de la orden de trabajo</param>
            /// <returns>string: Query de consulta para ejecución</returns>
            public static string GetFetchWoQuery(string dbReference, string dbLink, string districtCode, string workOrder)
            {
                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(districtCode))
                    districtCode = " IN (" + Utils.GetListInSeparator(DistrictConstants.GetDistrictList(), ",", "'") + ")";
                else
                    districtCode = " = '" + districtCode + "'";

                //escribimos el query
                var query = "" +
                            " SELECT" +
                            " WO.DSTRCT_CODE, WO.WORK_GROUP, WO.WORK_ORDER, WO.WO_STATUS_M, WO.WO_DESC, " +
                            " WO.EQUIP_NO, WO.COMP_CODE, WO.COMP_MOD_CODE, WO.LOCATION, WO.RAISED_DATE, WO.RAISED_TIME," +
                            " WO.ORIGINATOR_ID, WO.ORIG_PRIORITY, WO.ORIG_DOC_TYPE, WO.ORIG_DOC_NO, WO.REQUEST_ID, WO.MSSS_STATUS_IND," +
                            " WO.WO_TYPE, WO.MAINT_TYPE, WO.WO_STATUS_U, WO.STD_JOB_NO, WO.MAINT_SCH_TASK, WO.AUTO_REQ_IND, WO.ASSIGN_PERSON, WO.PLAN_PRIORITY, WO.CLOSED_COMMIT_DT, WO.UNIT_OF_WORK, WO.UNITS_REQUIRED, FAILURE_PART, PC_COMPLETE, UNITS_COMPLETE, WO.RELATED_WO," +
                            " WO.REQ_START_DATE, WO.REQ_START_TIME, WO.REQ_BY_DATE, WO.REQ_BY_TIME, WO.PLAN_STR_DATE, WO.PLAN_STR_TIME, WO.PLAN_FIN_DATE, WO.PLAN_FIN_TIME," +
                            " SUBSTR(WO.DSTRCT_ACCT_CODE, 5) DSTRCT_ACCT_CODE, WO.PROJECT_NO, WO.PARENT_WO," +
                            " WO.WO_JOB_CODEX1, WO.WO_JOB_CODEX2, WO.WO_JOB_CODEX3, WO.WO_JOB_CODEX4, WO.WO_JOB_CODEX5, WO.WO_JOB_CODEX6, WO.WO_JOB_CODEX7, WO.WO_JOB_CODEX8, WO.WO_JOB_CODEX9, WO.WO_JOB_CODEX10," +
                            " CASE WHEN TRIM(WO.WO_JOB_CODEX1||WO.WO_JOB_CODEX2||WO.WO_JOB_CODEX3||WO.WO_JOB_CODEX4||WO.WO_JOB_CODEX5||WO.WO_JOB_CODEX6||WO.WO_JOB_CODEX7||WO.WO_JOB_CODEX8||WO.WO_JOB_CODEX9||WO.WO_JOB_CODEX10) IS NULL THEN 'N' ELSE 'Y' END JOB_CODES," +
                            " WO.COMPLETED_CODE, WO.COMPLETED_BY," +
                            " CASE WHEN WO.DSTRCT_CODE || WO.WORK_ORDER NOT IN (SELECT STV.STD_KEY FROM " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " STV WHERE STV.STD_TEXT_CODE=('CW')) THEN 'N' ELSE 'Y' END COMPLETE_TEXT_FLAG," +
                            " WO.CLOSED_DT," +
                            " WOEST.CALC_DUR_HRS_SW, WOEST.EST_DUR_HRS, WOEST.ACT_DUR_HRS," +
                            " WOEST.RES_UPDATE_FLAG, WOEST.EST_LAB_HRS, WOEST.CALC_LAB_HRS, WOEST.ACT_LAB_HRS, WOEST.EST_LAB_COST, WOEST.CALC_LAB_COST, WOEST.ACT_LAB_COST," +
                            " WOEST.MAT_UPDATE_FLAG, WOEST.EST_MAT_COST, WOEST.CALC_MAT_COST, WOEST.ACT_MAT_COST," +
                            " WOEST.EQUIP_UPDATE_FLAG, WOEST.EST_EQUIP_COST, WOEST.CALC_EQUIP_COST, WOEST.ACT_EQUIP_COST," +
                            " WOEST.EST_OTHER_COST, WOEST.ACT_OTHER_COST," +
                            " WO.LOCATION_FR, WO.LOCATION, WO.NOTICE_LOCN," +
                            " WO.LAST_MOD_DATE, WO.FINAL_COSTS" +
                            " FROM" +
                            " " + dbReference + ".MSF620" + dbLink + " WO LEFT JOIN " + dbReference + ".MSF621" + dbLink + " WOEST ON (WO.WORK_ORDER    = WOEST.WORK_ORDER AND WO.DSTRCT_CODE = WOEST.DSTRCT_CODE)" +
                            " WHERE" +
                            " WO.WORK_ORDER = '" + workOrder + "'" +
                            " AND WO.DSTRCT_CODE " + districtCode;

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }


        }
    }
}
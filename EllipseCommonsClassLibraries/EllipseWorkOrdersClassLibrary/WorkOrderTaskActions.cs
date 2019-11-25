using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Utilities;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary.EquipmentReqmntsService;
using EllipseWorkOrdersClassLibrary.MaterialReqmntsService;
using EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using EllipseWorkOrdersClassLibrary.WorkOrderTaskService;

namespace EllipseWorkOrdersClassLibrary
{
    public static class WorkOrderTaskActions
    {
        public static string Close = "Close";
        public static string ReOpen = "ReOpen";
        public static string Create = "Create";
        public static string Modify = "Modify";
        public static string Delete = "Delete";

        public static List<string> GetActionsList()
        {
            var list = new List<string>();
            list.Add(Create);
            list.Add(Modify);
            list.Add(Delete);
            list.Add(Close);
            list.Add(ReOpen);

            return list;
        }

        public static List<TaskRequirement> FetchTaskRequirements(EllipseFunctions ef, string districtCode, string workGroup, string stdJob, string reqType = "ALL", string taskNo = null)
        {
            var sqlQuery = Queries.GetFetchWoTaskRealRequirementsQuery(ef.dbReference, ef.dbLink, districtCode, stdJob, reqType, string.IsNullOrWhiteSpace(taskNo) ? null : taskNo.PadLeft(3, '0'));
            var woTaskDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<TaskRequirement>();

            if (woTaskDataReader == null || woTaskDataReader.IsClosed || !woTaskDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
            while (woTaskDataReader.Read())
            {
                var taskReq = new TaskRequirement
                {
                    ReqType = "" + woTaskDataReader["REQ_TYPE"].ToString().Trim(),                  //REQ_TYPE
                    DistrictCode = "" + woTaskDataReader["DSTRCT_CODE"].ToString().Trim(),          //DSTRCT_CODE
                    WorkGroup = "" + woTaskDataReader["WORK_GROUP"].ToString().Trim(),              //WORK_GROUP
                    WorkOrder = "" + woTaskDataReader["WORK_ORDER"].ToString().Trim(),              //WORK_ORDER
                    WoTaskNo = "" + woTaskDataReader["WO_TASK_NO"].ToString().Trim(),               //WO_TASK_NO
                    WoTaskDesc = "" + woTaskDataReader["WO_TASK_DESC"].ToString().Trim(),           //WO_TASK_DESC
                    SeqNo = "" + woTaskDataReader["SEQ_NO"].ToString().Trim(),                      //SEQ_NO
                    ReqCode = "" + woTaskDataReader["RES_CODE"].ToString().Trim(),                  //RES_CODE
                    ReqDesc = "" + woTaskDataReader["RES_DESC"].ToString().Trim(),                  //RES_DESC
                    UoM = "" + woTaskDataReader["UNITS"].ToString().Trim(),                         //UNITS
                    QtyReq = "" + woTaskDataReader["QTY_REQ"].ToString().Trim(),                    //QTY_REQ
                    QtyIss = "" + woTaskDataReader["QTY_ISS"].ToString().Trim(),                    //QTY_ISS
                    HrsReq = "" + woTaskDataReader["EST_RESRCE_HRS"].ToString().Trim(), //EST_RESRCE_HRS
                    HrsReal = "" + woTaskDataReader["ACT_RESRCE_HRS"].ToString().Trim()  //ACT_RESRCE_HRS
                };
                list.Add(taskReq);
            }
            ef.CloseConnection();
           
            return list;
        }

        public static void ModifyWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            var proxyWoTask = new WorkOrderTaskService.WorkOrderTaskService();//ejecuta las acciones del servicio
            var requestWoTask = new WorkOrderTaskServiceModifyRequestDTO();

            //se cargan los parámetros de la orden
            proxyWoTask.Url = urlService + "/WorkOrderTaskService";

            //se cargan los parámetros de la orden
            requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
            requestWoTask.workGroup = woTask.WorkGroup ?? requestWoTask.workGroup;
            requestWoTask.workOrder = woTask.WorkOrderDto ?? requestWoTask.workOrder;
            requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');
            requestWoTask.WOTaskDesc = woTask.WoTaskDesc ?? requestWoTask.WOTaskDesc;
            requestWoTask.jobDescCode = woTask.JobDescCode ?? requestWoTask.jobDescCode;
            requestWoTask.safetyInstr = woTask.SafetyInstr ?? requestWoTask.safetyInstr;
            requestWoTask.completeInstr = woTask.CompleteInstr ?? requestWoTask.completeInstr;
            requestWoTask.complTextCode = woTask.ComplTextCode ?? requestWoTask.complTextCode;
            requestWoTask.assignPerson = woTask.AssignPerson ?? requestWoTask.assignPerson;
            requestWoTask.estimatedMachHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedMachHrs) ? Convert.ToDecimal(woTask.EstimatedMachHrs) : default(decimal);
            requestWoTask.estimatedMachHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedMachHrs);
            requestWoTask.planStrDate = woTask.PlanStartDate ?? requestWoTask.planStrDate;
            requestWoTask.tskDurationsHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs) ? Convert.ToDecimal(woTask.EstimatedDurationsHrs) : default(decimal);
            requestWoTask.tskDurationsHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs);
            requestWoTask.APLEquipmentGrpId = woTask.AplEquipmentGrpId ?? requestWoTask.APLEquipmentGrpId;
            requestWoTask.APLType = woTask.AplType ?? requestWoTask.APLType;
            requestWoTask.APLCompCode = woTask.AplCompCode ?? requestWoTask.APLCompCode;
            requestWoTask.APLCompModCode = woTask.AplCompModCode ?? requestWoTask.APLCompModCode;
            requestWoTask.APLSeqNo = woTask.AplSeqNo ?? requestWoTask.APLSeqNo;

            proxyWoTask.modify(opContext, requestWoTask);
        }

        public static void CreateWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            var proxywoTask = new WorkOrderTaskService.WorkOrderTaskService();//ejecuta las acciones del servicio
            var requestWoTask = new WorkOrderTaskServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            proxywoTask.Url = urlService + "/WorkOrderTaskService";

            //se cargan los parámetros de la orden
            requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
            requestWoTask.workGroup = woTask.WorkGroup ?? requestWoTask.workGroup;
            requestWoTask.workOrder = woTask.WorkOrderDto ?? requestWoTask.workOrder;

            requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');
            requestWoTask.WOTaskDesc = woTask.WoTaskDesc ?? requestWoTask.WOTaskDesc;
            requestWoTask.jobDescCode = woTask.JobDescCode ?? requestWoTask.jobDescCode;
            requestWoTask.safetyInstr = woTask.SafetyInstr ?? requestWoTask.safetyInstr;
            requestWoTask.completeInstr = woTask.CompleteInstr ?? requestWoTask.completeInstr;
            requestWoTask.complTextCode = woTask.ComplTextCode ?? requestWoTask.complTextCode;
            requestWoTask.assignPerson = woTask.AssignPerson ?? requestWoTask.assignPerson;
            requestWoTask.estimatedMachHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedMachHrs) ? Convert.ToDecimal(woTask.EstimatedMachHrs) : default(decimal);
            requestWoTask.estimatedMachHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedMachHrs);
            requestWoTask.tskDurationsHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs) ? Convert.ToDecimal(woTask.EstimatedDurationsHrs) : default(decimal);
            requestWoTask.tskDurationsHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs);
            requestWoTask.APLEquipmentGrpId = woTask.AplEquipmentGrpId ?? requestWoTask.APLEquipmentGrpId;
            requestWoTask.APLType = woTask.AplType ?? requestWoTask.APLType;
            requestWoTask.APLCompCode = woTask.AplCompCode ?? requestWoTask.APLCompCode;
            requestWoTask.APLCompModCode = woTask.AplCompModCode ?? requestWoTask.APLCompModCode;
            requestWoTask.APLSeqNo = woTask.AplSeqNo ?? requestWoTask.APLSeqNo;

            proxywoTask.create(opContext, requestWoTask);
        }


        public static void DeleteWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            var proxywoTask = new WorkOrderTaskService.WorkOrderTaskService();//ejecuta las acciones del servicio
            var requestWoTask = new WorkOrderTaskServiceDeleteRequestDTO();
            var requestWoTaskList = new List<WorkOrderTaskServiceDeleteRequestDTO>();

            //se cargan los parámetros de la orden
            proxywoTask.Url = urlService + "/WorkOrderTaskService";

            //se cargan los parámetros de la orden
            requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
            requestWoTask.workOrder = woTask.WorkOrderDto ?? requestWoTask.workOrder;
            requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');

            requestWoTaskList.Add(requestWoTask);

            proxywoTask.multipleDelete(opContext, requestWoTaskList.ToArray());
        }

        public static ReplyMessage CompleteWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            var proxywoTask = new WorkOrderTaskService.WorkOrderTaskService();//ejecuta las acciones del servicio
            var requestWoTask = new WorkOrderTaskServiceCompleteRequestDTO();
            var requestWoTaskList = new List<WorkOrderTaskServiceCompleteRequestDTO>();

            //se cargan los parámetros de la orden
            proxywoTask.Url = urlService + "/WorkOrderTaskService";

            //se cargan los parámetros de la orden
            requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
            requestWoTask.workOrder = woTask.SetWorkOrderDto(woTask.WorkOrder);
            requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');
            requestWoTask.completedCode = string.IsNullOrWhiteSpace(woTask.CompletedCode) ? "06" : woTask.CompletedCode;
            requestWoTask.completedBy = woTask.CompletedBy;
            requestWoTask.closedDt = woTask.ClosedDate;

            var serviceReply = proxywoTask.complete(opContext, requestWoTask);

            var reply = new ReplyMessage();
            reply.Message = "Completed " + serviceReply.workOrder.prefix + serviceReply.workOrder.no + " " + serviceReply.WOTaskNo + " Completed Code " + serviceReply.completedCode + " - " + serviceReply.completedCodeDescription;
            return reply;
        }
        public static ReplyMessage ReOpenWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            var proxywoTask = new WorkOrderTaskService.WorkOrderTaskService();//ejecuta las acciones del servicio
            var requestWoTask = new WorkOrderTaskServiceReopenRequestDTO();

            //se cargan los parámetros de la orden
            proxywoTask.Url = urlService + "/WorkOrderTaskService";

            //se cargan los parámetros de la orden
            requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
            requestWoTask.workOrder = woTask.SetWorkOrderDto(woTask.WorkOrder);
            requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');

            var serviceReply = proxywoTask.reopen(opContext, requestWoTask);

            var reply = new ReplyMessage();
            reply.Message = "ReOpen " + serviceReply.workOrder.prefix + serviceReply.workOrder.no + " " + serviceReply.WOTaskNo;
            return reply;
        }

        public static void SetWorkOrderTaskText(string urlService, string districtCode, string position, bool returnWarnings, WorkOrderTask woTask)
        {
            if (!string.IsNullOrWhiteSpace(woTask.WoTaskNo))
                woTask.WoTaskNo = woTask.WoTaskNo.PadLeft(3, '0');//comentario
            var stdTextId = "WI" + districtCode + woTask.WorkOrder + woTask.WoTaskNo;
            var stdTextCompleteId = "WA" + districtCode + woTask.WorkOrder + woTask.WoTaskNo;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, woTask.ExtTaskText);

            StdText.SetText(urlService, stdTextCopc, stdTextCompleteId, woTask.CompleteTaskText);
        }


        /// <summary>
        /// 
        /// </summary>

        public static void CreateTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new ResourceReqmntsService.ResourceReqmntsService
            {
                Url = urlService + "/ResourceReqmntsService"
            }; //ejecuta las acciones del servicio

            //se cargan los parámetros de la orden

            //se cargan los parámetros de la orden

            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new ResourceReqmntsServiceCreateRequestDTO
            {
                workOrder = new ResourceReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                districtCode = taskReq.DistrictCode,
                workOrderTask = taskReq.WoTaskNo,
                resourceClass = taskReq.ReqCode.Substring(0, 1),
                resourceCode = taskReq.ReqCode.Substring(1),
                quantityRequired = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                quantityRequiredSpecified = !string.IsNullOrWhiteSpace(taskReq.QtyReq),
                hrsReqd = !string.IsNullOrWhiteSpace(taskReq.HrsReq) ? Convert.ToDecimal(taskReq.HrsReq) : default(decimal),
                hrsReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.HrsReq),
                classType = "WT",
                enteredInd = "S"
            };
            proxyTaskReq.create(opContext, requestTaskReq);
        }

        public static void CreateTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new MaterialReqmntsService.MaterialReqmntsService
            {
                Url = urlService + "/MaterialReqmntsService"
            }; //ejecuta las acciones del servicio

            //se cargan los parámetros de la orden

            //se cargan los parámetros de la orden

            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new MaterialReqmntsServiceCreateRequestDTO
            {
                workOrder = new MaterialReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                districtCode = taskReq.DistrictCode,
                workOrderTask = taskReq.WoTaskNo,
                seqNo = taskReq.SeqNo,
                stockCode = taskReq.ReqCode.Substring(1).PadLeft(9, '0'),
                unitQuantityReqdSpecified = true,
                catalogueFlag = true,
                catalogueFlagSpecified = true,
                contestibleFlag = false,
                contestibleFlagSpecified = true,
                classType = "WT",
                enteredInd = "S",
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false,
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
            };
            proxyTaskReq.create(opContext, requestTaskReq);
        }

        public static void CreateTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceCreateRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WorkOrder))
                taskReq.WorkOrder = taskReq.WorkOrder.PadLeft(3, '0');

            var requestTaskReq = new EquipmentReqmntsServiceCreateRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new EquipmentReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = taskReq.WoTaskNo,
                seqNo = taskReq.SeqNo,
                eqptType = taskReq.ReqCode.Substring(1),
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                unitQuantityReqdSpecified = true,
                UOM = taskReq.UoM,
                contestibleFlg = false,
                contestibleFlgSpecified = true,
                classType = "WT",
                enteredInd = "S",
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleCreate(opContext, requestTaskReqList.ToArray());
        }

        public static void ModifyTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new ResourceReqmntsService.ResourceReqmntsService
            {
                Url = urlService + "/ResourceReqmntsService"
            };

            var requestTaskReqList = new List<ResourceReqmntsServiceModifyRequestDTO>();

            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new ResourceReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new ResourceReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = !string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? taskReq.WoTaskNo : null,
                resourceClass = taskReq.ReqCode.Substring(0, 1),
                resourceCode = taskReq.ReqCode.Substring(1),
                quantityRequired = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                quantityRequiredSpecified = taskReq.QtyReq != null,
                hrsReqd = !string.IsNullOrWhiteSpace(taskReq.HrsReq) ? Convert.ToDecimal(taskReq.HrsReq) : default(decimal),
                hrsReqdSpecified = taskReq.HrsReq != null,
                classType = "WT",
                enteredInd = "S"
            };

            requestTaskReqList.Add(requestTaskReq);

            proxyTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static void ModifyTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new MaterialReqmntsService.MaterialReqmntsService
            {
                Url = urlService + "/MaterialReqmntsService"
            };

            var requestTaskReqList = new List<MaterialReqmntsServiceModifyRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new MaterialReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new MaterialReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = !string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? taskReq.WoTaskNo : null,
                seqNo = taskReq.SeqNo,
                stockCode = taskReq.ReqCode.Substring(1).PadLeft(9, '0'),
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                classType = "WT",
                enteredInd = "S",
                unitQuantityReqdSpecified = true,
                catalogueFlag = true,
                catalogueFlagSpecified = true,
                contestibleFlag = false,
                contestibleFlagSpecified = true,
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static void ModifyTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceModifyRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new EquipmentReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new EquipmentReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = !string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? taskReq.WoTaskNo : null,
                seqNo = taskReq.SeqNo,
                eqptType = taskReq.ReqCode.Substring(1),
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                unitQuantityReqdSpecified = true,
                UOM = taskReq.UoM,
                contestibleFlg = false,
                contestibleFlgSpecified = true,
                classType = "WT",
                enteredInd = "S",
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static void DeleteTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new ResourceReqmntsService.ResourceReqmntsService
            {
                Url = urlService + "/ResourceReqmntsService"
            };

            var requestTaskReqList = new List<ResourceReqmntsServiceDeleteRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new ResourceReqmntsServiceDeleteRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new ResourceReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = Convert.ToString(Convert.ToDecimal(taskReq.WoTaskNo), CultureInfo.InvariantCulture),
                resourceClass = taskReq.ReqCode.Substring(0, 1),
                resourceCode = taskReq.ReqCode.Substring(1),
                classType = "WT",
                enteredInd = "S"
            };
            requestTaskReqList.Add(requestTaskReq);

            proxyTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static void DeleteTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new MaterialReqmntsService.MaterialReqmntsService
            {
                Url = urlService + "/MaterialReqmntsService"
            };

            var requestTaskReqList = new List<MaterialReqmntsServiceDeleteRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new MaterialReqmntsServiceDeleteRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new MaterialReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = Convert.ToString(Convert.ToDecimal(taskReq.WoTaskNo), CultureInfo.InvariantCulture),
                seqNo = taskReq.SeqNo,
                classType = "WT",
                enteredInd = "S",
                CUItemNoSpecified = false,
                JEItemNoSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static void DeleteTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceDeleteRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.WoTaskNo))
                taskReq.WoTaskNo = taskReq.WoTaskNo.PadLeft(3, '0');

            var requestTaskReq = new EquipmentReqmntsServiceDeleteRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                workOrder = new EquipmentReqmntsService.WorkOrderDTO
                {
                    prefix = taskReq.WorkOrder.Substring(0, 2),
                    no = taskReq.WorkOrder.Substring(2, taskReq.WorkOrder.Length - 2),
                },
                workOrderTask = Convert.ToString(Convert.ToDecimal(taskReq.WoTaskNo), CultureInfo.InvariantCulture),
                seqNo = taskReq.SeqNo,
                operationTypeEQP = taskReq.ReqCode,
                classType = "WT",
                enteredInd = "S",
                CUItemNoSpecified = false,
                JEItemNoSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static List<WorkOrderTask> FetchWorkOrderTask(EllipseFunctions ef, string districtCode, string workOrder, string woTaskNo)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchWorkOrderTasksQuery(ef.dbReference, ef.dbLink, districtCode, workOrder, woTaskNo));

            var list = new List<WorkOrderTask>();

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
            while (stdDataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var task = new WorkOrderTask();

                task.DistrictCode = "" + stdDataReader["DSTRCT_CODE"].ToString().Trim();
                task.WorkGroup = "" + stdDataReader["WORK_GROUP"].ToString().Trim();
                task.WorkOrder = "" + stdDataReader["WORK_ORDER"].ToString().Trim();
                task.WorkOrderDescription = "" + stdDataReader["WO_DESC"].ToString().Trim();

                task.WoTaskNo = "" + stdDataReader["WO_TASK_NO"].ToString().Trim();
                task.WoTaskDesc = "" + stdDataReader["WO_TASK_DESC"].ToString().Trim();
                task.JobDescCode = "" + stdDataReader["JOB_DESC_CODE"].ToString().Trim();
                task.SafetyInstr = "" + stdDataReader["SAFETY_INSTR"].ToString().Trim();
                task.CompleteInstr = "" + stdDataReader["COMPLETE_INSTR"].ToString().Trim();
                task.ComplTextCode = "" + stdDataReader["COMPL_TEXT_CDE"].ToString().Trim();

                task.AssignPerson = "" + stdDataReader["ASSIGN_PERSON"].ToString().Trim();
                task.EstimatedMachHrs = "" + stdDataReader["EST_MACH_HRS"].ToString().Trim();
                task.PlanStartDate = "" + stdDataReader["PLAN_STR_DATE"].ToString().Trim();

                task.EstimatedDurationsHrs = "" + stdDataReader["TSK_DUR_HOURS"].ToString().Trim();
                task.NoLabor = "" + stdDataReader["NO_REC_LABOR"].ToString().Trim();
                task.NoMaterial = "" + stdDataReader["NO_REC_MATERIAL"].ToString().Trim();

                task.AplEquipmentGrpId = "" + stdDataReader["EQUIP_GRP_ID"].ToString().Trim();
                task.AplType = "" + stdDataReader["APL_TYPE"].ToString().Trim();
                task.AplCompCode = "" + stdDataReader["COMP_CODE"].ToString().Trim();
                task.AplCompModCode = "" + stdDataReader["COMP_MOD_CODE"].ToString().Trim();
                task.AplSeqNo = "" + stdDataReader["APL_SEQ_NO"].ToString().Trim();
                task.ClosedStatus = "" + stdDataReader["CLOSED_STATUS"].ToString().Trim();

                list.Add(task);
            }
            ef.CloseConnection();
            return list;
        }

        public static class Queries
        {

            public static string GetFetchWorkOrderTasksQuery(string dbReference, string dbLink, string districtCode, string workOrder, string woTaskNo)
            {
                var query = "" +
                            "SELECT " +
                            "	WO.DSTRCT_CODE, " +
                            "	WO.WORK_GROUP, " +
                            "	WO.WORK_ORDER, " +
                            "	WO.WO_DESC, " +
                            "	WT.WO_TASK_NO, " +
                            "	WT.WO_TASK_DESC, " +
                            "	WT.JOB_DESC_CODE, " +
                            "	WT.SAFETY_INSTR, " +
                            "	WT.COMPLETE_INSTR, " +
                            "	WT.COMPL_TEXT_CDE, " +
                            "	WT.ASSIGN_PERSON, " +
                            "	WT.EST_MACH_HRS, " +
                            "	WT.TSK_DUR_HOURS, " +
                            "	WT.PLAN_STR_DATE, " +
                            "	WT.EQUIP_GRP_ID, " +
                            "	WT.APL_TYPE, " +
                            "	WT.COMP_CODE, " +
                            "	WT.COMP_MOD_CODE, " +
                            "	WT.APL_SEQ_NO, " +
                            "	WT.CLOSED_STATUS, " +
                            "	( " +
                            "		SELECT " +
                            "			COUNT(*) LABOR " +
                            "		FROM " +
                            "			ELLIPSE.MSF623 TSK " +
                            "			INNER JOIN ELLIPSE.MSF735 RS " +
                            "			ON RS.KEY_735_ID     = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                            "			   AND RS.REC_735_TYPE   = 'WT' " +
                            "		WHERE " +
                            "			TSK.WORK_ORDER = WO.WORK_ORDER " +
                            "			AND   TSK.WO_TASK_NO = WT.WO_TASK_NO " +
                            "	)NO_REC_LABOR, " +
                            "	( " +
                            "		SELECT " +
                            "			COUNT(*) MATER " +
                            "		FROM " +
                            "			ELLIPSE.MSF623 TSK " +
                            "			INNER JOIN ELLIPSE.MSF734 RS " +
                            "			ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                            "			   AND RS.CLASS_TYPE   = 'WT' " +
                            "		WHERE " +
                            "			TSK.WORK_ORDER = WO.WORK_ORDER " +
                            "			AND   TSK.WO_TASK_NO = WT.WO_TASK_NO " +
                            "	)NO_REC_MATERIAL " +
                            "FROM " +
                            "	" + dbReference + ".MSF620" + dbLink + " WO " +
                            "	INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT " +
                            "	ON WO.WORK_ORDER    = WT.WORK_ORDER " +
                            "	   AND WO.DSTRCT_CODE   = WT.DSTRCT_CODE " +
                            "	   AND WO.WORK_ORDER    = '" + workOrder + "'" +
                            "	   AND WO.DSTRCT_CODE   = '" + districtCode + "'";
                if (woTaskNo != "")
                {
                    query = query + " AND WT.WO_TASK_NO   = " + woTaskNo + " ";
                }

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFetchWoTaskRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
            {
                var query = "" +
                            "SELECT " +
                            "	'LAB' REQ_TYPE, " +
                            "	TSK.DSTRCT_CODE, " +
                            "	TSK.WORK_GROUP, " +
                            "	TSK.WORK_ORDER, " +
                            "	TSK.WO_TASK_NO, " +
                            "	TSK.WO_TASK_DESC, " +
                            "	'N/A' SEQ_NO, " +
                            "	RS.RESOURCE_TYPE RES_CODE, " +
                            "	TO_NUMBER(RS.CREW_SIZE) QTY_REQ, " +
                            "	RS.EST_RESRCE_HRS HRS_QTY, " +
                            "	TT.TABLE_DESC RES_DESC, " +
                            "	'' UNITS " +
                            "FROM " +
                            "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                            "	INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS " +
                            "	ON RS.KEY_735_ID     = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                            "	   AND RS.REC_735_TYPE   = 'WT' " +
                            "	INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT " +
                            "	ON TT.TABLE_CODE   = RS.RESOURCE_TYPE " +
                            "	   AND TT.TABLE_TYPE   = 'TT' " +
                            "WHERE " +
                            "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                            "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                            "	AND   TSK.WO_TASK_NO = '" + taskNo + "' " +
                            "UNION ALL " +
                            "SELECT " +
                            "	'MAT' REQ_TYPE, " +
                            "	TSK.DSTRCT_CODE, " +
                            "	TSK.WORK_GROUP, " +
                            "	TSK.WORK_ORDER, " +
                            "	TSK.WO_TASK_NO, " +
                            "	TSK.WO_TASK_DESC, " +
                            "	RS.SEQNCE_NO SEQ_NO, " +
                            "	RS.STOCK_CODE RES_CODE, " +
                            "	RS.UNIT_QTY_REQD QTY_REQ, " +
                            "	0 HRS_QTY, " +
                            "	SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, " +
                            "	'' UNITS " +
                            "FROM " +
                            "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                            "	INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS " +
                            "	ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                            "	LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT " +
                            "	ON RS.STOCK_CODE   = SCT.STOCK_CODE " +
                            "	    AND RS.CLASS_TYPE   = 'WT' " +
                            "WHERE " +
                            "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                            "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                            "	AND   TSK.WO_TASK_NO = '" + taskNo + "' " +
                            "UNION ALL " +
                            "SELECT " +
                            "	'EQU' REQ_TYPE, " +
                            "	TSK.DSTRCT_CODE, " +
                            "	TSK.WORK_GROUP, " +
                            "	TSK.WORK_ORDER, " +
                            "	TSK.WO_TASK_NO, " +
                            "	TSK.WO_TASK_DESC, " +
                            "	RS.SEQNCE_NO SEQ_NO, " +
                            "	RS.EQPT_TYPE RES_CODE, " +
                            "	RS.QTY_REQ, " +
                            "	RS.UNIT_QTY_REQD HRS_QTY, " +
                            "	ET.TABLE_DESC RES_DESC, " +
                            "	RS.UOM UNITS " +
                            "FROM " +
                            "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                            "	INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS " +
                            "	ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                            "	   AND RS.CLASS_TYPE   = 'WT' " +
                            "	INNER JOIN " + dbReference + ".MSF010" + dbLink + " ET " +
                            "	ON RS.EQPT_TYPE   = ET.TABLE_CODE " +
                            "	   AND TABLE_TYPE     = 'ET' " +
                            "WHERE " +
                            "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                            "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                            "	AND   TSK.WO_TASK_NO = '" + taskNo + "'";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;

            }

            public static string GetFetchWoTaskRealRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string reqType = "ALL", string taskNo = null)
            {
                var query = "";
                switch (reqType)
                {
                    case "LAB":
                        {
                            query = "WITH RES_REAL AS ( ";
                            query += "    SELECT ";
                            query += "        TR.DSTRCT_CODE, ";
                            query += "        WT.WORK_GROUP, ";
                            query += "        TR.WORK_ORDER, ";
                            query += "        TR.WO_TASK_NO, ";
                            query += "        WT.WO_TASK_DESC, ";
                            query += "        TR.RESOURCE_TYPE RES_CODE, ";
                            query += "        TT.TABLE_DESC RES_DESC, ";
                            query += "        SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS ";
                            query += "    FROM ";
                            query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                            query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                            query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                            query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                            query += "        ON TT.TABLE_CODE = TR.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                            query += "        INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT ";
                            query += "        ON WT.DSTRCT_CODE = TR.DSTRCT_CODE AND WT.WORK_ORDER = TR.WORK_ORDER AND WT.WO_TASK_NO = TR.WO_TASK_NO ";
                            query += "    WHERE ";
                            query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' ";
                            if (taskNo != null)
                                query += "AND   TR.WO_TASK_NO = '" + taskNo + "' ";
                            query += "    GROUP BY ";
                            query += "        TR.DSTRCT_CODE, ";
                            query += "        WT.WORK_GROUP, ";
                            query += "        TR.WORK_ORDER, ";
                            query += "        TR.WO_TASK_NO, ";
                            query += "        WT.WO_TASK_DESC, ";
                            query += "        TR.RESOURCE_TYPE, ";
                            query += "        TT.TABLE_DESC ";
                            query += "),RES_EST AS ( ";
                            query += "    SELECT ";
                            query += "        TSK.DSTRCT_CODE, ";
                            query += "        TSK.WORK_GROUP, ";
                            query += "        TSK.WORK_ORDER, ";
                            query += "        TSK.WO_TASK_NO, ";
                            query += "        TSK.WO_TASK_DESC, ";
                            query += "        RS.RESOURCE_TYPE RES_CODE, ";
                            query += "        TT.TABLE_DESC RES_DESC, ";
                            query += "        TO_NUMBER(RS.CREW_SIZE) QTY_REQ, ";
                            query += "        RS.EST_RESRCE_HRS ";
                            query += "    FROM ";
                            query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                            query += "        INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS ";
                            query += "        ON RS.KEY_735_ID = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.REC_735_TYPE = 'WT' ";
                            query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                            query += "        ON TT.TABLE_CODE = RS.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                            query += "    WHERE ";
                            query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                            if (taskNo != null)
                                query += "      AND   TSK.WO_TASK_NO = '" + taskNo + "' ";
                            query += "),TABLA_REC AS ( ";
                            query += "    SELECT ";
                            query += "        DECODE(RES_EST.DSTRCT_CODE,NULL,RES_REAL.DSTRCT_CODE,RES_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                            query += "        DECODE(RES_EST.WORK_GROUP,NULL,RES_REAL.WORK_GROUP,RES_EST.WORK_GROUP) WORK_GROUP, ";
                            query += "        DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, ";
                            query += "        DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, ";
                            query += "        DECODE(RES_EST.WO_TASK_DESC,NULL,RES_REAL.WO_TASK_DESC,RES_EST.WO_TASK_DESC) WO_TASK_DESC, ";
                            query += "        DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, ";
                            query += "        DECODE(RES_EST.RES_DESC,NULL,RES_REAL.RES_DESC,RES_EST.RES_DESC) RES_DESC, ";
                            query += "        RES_EST.QTY_REQ, ";
                            query += "        RES_REAL.ACT_RESRCE_HRS, ";
                            query += "        RES_EST.EST_RESRCE_HRS ";
                            query += "    FROM ";
                            query += "        RES_REAL ";
                            query += "        FULL JOIN RES_EST ";
                            query += "        ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO AND RES_REAL.RES_CODE = RES_EST.RES_CODE ";
                            query += ") SELECT ";
                            query += "    'LAB' REQ_TYPE, ";
                            query += "    TABLA_REC.DSTRCT_CODE, ";
                            query += "    TABLA_REC.WORK_GROUP, ";
                            query += "    TABLA_REC.WORK_ORDER, ";
                            query += "    TABLA_REC.WO_TASK_NO, ";
                            query += "    TABLA_REC.WO_TASK_DESC, ";
                            query += "    '' SEQ_NO, ";
                            query += "    TABLA_REC.RES_CODE, ";
                            query += "    TABLA_REC.RES_DESC, ";
                            query += "    '' UNITS, ";
                            query += "    TABLA_REC.QTY_REQ, ";
                            query += "    NULL QTY_ISS, ";
                            query += "    DECODE(TABLA_REC.EST_RESRCE_HRS, NULL, 0, TABLA_REC.EST_RESRCE_HRS) EST_RESRCE_HRS, ";
                            query += "    DECODE(TABLA_REC.ACT_RESRCE_HRS, NULL, 0, TABLA_REC.ACT_RESRCE_HRS) ACT_RESRCE_HRS ";
                            query += "FROM ";
                            query += "    TABLA_REC ";
                            break;
                        }
                    case "MAT":
                        {
                            query = "WITH MAT_REAL AS ( ";
                            query += "    SELECT ";
                            query += "        TR.DSTRCT_CODE, ";
                            query += "        WO.WORK_GROUP, ";
                            query += "        TR.WORK_ORDER, ";
                            query += "        TR.STOCK_CODE AS RES_CODE, ";
                            query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                            query += "        SCT.UNIT_OF_ISSUE UNITS, ";
                            query += "        SUM(TR.QUANTITY_ISS) QTY_ISS ";
                            query += "    FROM ";
                            query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                            query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                            query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                            query += "        LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                            query += "        ON TR.STOCK_CODE = SCT.STOCK_CODE ";
                            query += "        INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO ";
                            query += "        ON WO.DSTRCT_CODE = TR.DSTRCT_CODE AND WO.WORK_ORDER = TR.WORK_ORDER ";
                            query += "    WHERE ";
                            query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' AND   TX.REC900_TYPE = 'S' ";
                            query += "    GROUP BY ";
                            query += "        TR.DSTRCT_CODE, ";
                            query += "        WO.WORK_GROUP, ";
                            query += "        TR.WORK_ORDER, ";
                            query += "        TR.STOCK_CODE, ";
                            query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME, ";
                            query += "        SCT.UNIT_OF_ISSUE ";
                            query += "),MAT_EST AS ( ";
                            query += "    SELECT ";
                            query += "        TSK.DSTRCT_CODE, ";
                            query += "        TSK.WORK_GROUP, ";
                            query += "        TSK.WORK_ORDER, ";
                            query += "        TSK.WO_TASK_NO, ";
                            query += "        TSK.WO_TASK_DESC, ";
                            query += "        RS.SEQNCE_NO SEQ_NO, ";
                            query += "        RS.STOCK_CODE RES_CODE, ";
                            query += "        RS.UNIT_QTY_REQD QTY_REQ, ";
                            query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                            query += "        SCT.UNIT_OF_ISSUE UNITS ";
                            query += "    FROM ";
                            query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                            query += "        INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS ";
                            query += "        ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO ";
                            query += "        INNER JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                            query += "        ON RS.STOCK_CODE = SCT.STOCK_CODE AND RS.CLASS_TYPE = 'WT' ";
                            query += "    WHERE ";
                            query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                            if (taskNo != null) query += "           AND  TSK.WO_TASK_NO = '" + taskNo + "' ";
                            query += "),TABLA_MAT AS ( ";
                            query += "    SELECT ";
                            query += "        DECODE(MAT_EST.DSTRCT_CODE,NULL,MAT_REAL.DSTRCT_CODE,MAT_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                            query += "        DECODE(MAT_EST.WORK_GROUP,NULL,MAT_REAL.WORK_GROUP,MAT_EST.WORK_GROUP) WORK_GROUP, ";
                            query += "        DECODE(MAT_EST.WORK_ORDER,NULL,MAT_REAL.WORK_ORDER,MAT_EST.WORK_ORDER) WORK_ORDER, ";
                            query += "        MAT_EST.WO_TASK_NO, ";
                            query += "        MAT_EST.WO_TASK_DESC, ";
                            query += "        MAT_EST.SEQ_NO, ";
                            query += "        DECODE(MAT_EST.RES_CODE,NULL,MAT_REAL.RES_CODE,MAT_EST.RES_CODE) RES_CODE, ";
                            query += "        DECODE(MAT_EST.RES_DESC,NULL,MAT_REAL.RES_DESC,MAT_EST.RES_DESC) RES_DESC, ";
                            query += "        DECODE(MAT_EST.UNITS,NULL,MAT_REAL.UNITS,MAT_EST.UNITS) UNITS, ";
                            query += "        MAT_EST.QTY_REQ, ";
                            query += "        MAT_REAL.QTY_ISS ";
                            query += "    FROM ";
                            query += "        MAT_REAL ";
                            query += "        FULL JOIN MAT_EST ";
                            query += "        ON MAT_REAL.DSTRCT_CODE = MAT_EST.DSTRCT_CODE AND MAT_REAL.WORK_ORDER = MAT_EST.WORK_ORDER AND MAT_REAL.RES_CODE = MAT_EST.RES_CODE ";
                            query += ")SELECT ";
                            query += "    'MAT' REQ_TYPE, ";
                            query += "    TABLA_MAT.DSTRCT_CODE, ";
                            query += "    TABLA_MAT.WORK_GROUP, ";
                            query += "    TABLA_MAT.WORK_ORDER, ";
                            query += "    TABLA_MAT.WO_TASK_NO, ";
                            query += "    TABLA_MAT.WO_TASK_DESC, ";
                            query += "    TABLA_MAT.SEQ_NO, ";
                            query += "    TABLA_MAT.RES_CODE, ";
                            query += "    TABLA_MAT.RES_DESC, ";
                            query += "    DECODE(TABLA_MAT.UNITS, NULL, '', TABLA_MAT.UNITS) UNITS, ";
                            query += "    TABLA_MAT.QTY_REQ, ";
                            query += "    DECODE(TABLA_MAT.QTY_ISS, NULL, 0,TABLA_MAT.QTY_ISS) QTY_ISS, ";
                            query += "    0 EST_RESRCE_HRS, ";
                            query += "    0 ACT_RESRCE_HRS ";
                            query += "  FROM ";
                            query += "    TABLA_MAT ";
                            break;
                        }
                    default:
                        query = "WITH MAT_REAL AS ( ";
                        query += "    SELECT ";
                        query += "        TR.DSTRCT_CODE, ";
                        query += "        WO.WORK_GROUP, ";
                        query += "        TR.WORK_ORDER, ";
                        query += "        TR.STOCK_CODE AS RES_CODE, ";
                        query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                        query += "        SCT.UNIT_OF_ISSUE UNITS, ";
                        query += "        SUM(TR.QUANTITY_ISS) QTY_ISS ";
                        query += "    FROM ";
                        query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                        query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                        query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                        query += "        LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                        query += "        ON TR.STOCK_CODE = SCT.STOCK_CODE ";
                        query += "        INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO ";
                        query += "        ON WO.DSTRCT_CODE = TR.DSTRCT_CODE AND WO.WORK_ORDER = TR.WORK_ORDER ";
                        query += "    WHERE ";
                        query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' AND   TX.REC900_TYPE = 'S' ";
                        query += "    GROUP BY ";
                        query += "        TR.DSTRCT_CODE, ";
                        query += "        WO.WORK_GROUP, ";
                        query += "        TR.WORK_ORDER, ";
                        query += "        TR.STOCK_CODE, ";
                        query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME, ";
                        query += "        SCT.UNIT_OF_ISSUE ";
                        query += "),MAT_EST AS ( ";
                        query += "    SELECT ";
                        query += "        TSK.DSTRCT_CODE, ";
                        query += "        TSK.WORK_GROUP, ";
                        query += "        TSK.WORK_ORDER, ";
                        query += "        TSK.WO_TASK_NO, ";
                        query += "        TSK.WO_TASK_DESC, ";
                        query += "        RS.SEQNCE_NO SEQ_NO, ";
                        query += "        RS.STOCK_CODE RES_CODE, ";
                        query += "        RS.UNIT_QTY_REQD QTY_REQ, ";
                        query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                        query += "        SCT.UNIT_OF_ISSUE UNITS ";
                        query += "    FROM ";
                        query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                        query += "        INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS ";
                        query += "        ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO ";
                        query += "        INNER JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                        query += "        ON RS.STOCK_CODE = SCT.STOCK_CODE AND RS.CLASS_TYPE = 'WT' ";
                        query += "    WHERE ";
                        query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                        if (taskNo != null) query += "           AND  TSK.WO_TASK_NO = '" + taskNo + "' ";
                        query += "),TABLA_MAT AS ( ";
                        query += "    SELECT ";
                        query += "        DECODE(MAT_EST.DSTRCT_CODE,NULL,MAT_REAL.DSTRCT_CODE,MAT_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                        query += "        DECODE(MAT_EST.WORK_GROUP,NULL,MAT_REAL.WORK_GROUP,MAT_EST.WORK_GROUP) WORK_GROUP, ";
                        query += "        DECODE(MAT_EST.WORK_ORDER,NULL,MAT_REAL.WORK_ORDER,MAT_EST.WORK_ORDER) WORK_ORDER, ";
                        query += "        MAT_EST.WO_TASK_NO, ";
                        query += "        MAT_EST.WO_TASK_DESC, ";
                        query += "        MAT_EST.SEQ_NO, ";
                        query += "        DECODE(MAT_EST.RES_CODE,NULL,MAT_REAL.RES_CODE,MAT_EST.RES_CODE) RES_CODE, ";
                        query += "        DECODE(MAT_EST.RES_DESC,NULL,MAT_REAL.RES_DESC,MAT_EST.RES_DESC) RES_DESC, ";
                        query += "        DECODE(MAT_EST.UNITS,NULL,MAT_REAL.UNITS,MAT_EST.UNITS) UNITS, ";
                        query += "        MAT_EST.QTY_REQ, ";
                        query += "        MAT_REAL.QTY_ISS ";
                        query += "    FROM ";
                        query += "        MAT_REAL ";
                        query += "        FULL JOIN MAT_EST ";
                        query += "        ON MAT_REAL.DSTRCT_CODE = MAT_EST.DSTRCT_CODE AND MAT_REAL.WORK_ORDER = MAT_EST.WORK_ORDER AND MAT_REAL.RES_CODE = MAT_EST.RES_CODE ";
                        query += "),RES_REAL AS ( ";
                        query += "    SELECT ";
                        query += "        TR.DSTRCT_CODE, ";
                        query += "        WT.WORK_GROUP, ";
                        query += "        TR.WORK_ORDER, ";
                        query += "        TR.WO_TASK_NO, ";
                        query += "        WT.WO_TASK_DESC, ";
                        query += "        TR.RESOURCE_TYPE RES_CODE, ";
                        query += "        TT.TABLE_DESC RES_DESC, ";
                        query += "        SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS ";
                        query += "    FROM ";
                        query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                        query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                        query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                        query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                        query += "        ON TT.TABLE_CODE = TR.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                        query += "        INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT ";
                        query += "        ON WT.DSTRCT_CODE = TR.DSTRCT_CODE AND WT.WORK_ORDER = TR.WORK_ORDER AND WT.WO_TASK_NO = TR.WO_TASK_NO ";
                        query += "    WHERE ";
                        query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' ";
                        if (taskNo != null)
                            query += "AND   TR.WO_TASK_NO = '" + taskNo + "' ";
                        query += "    GROUP BY ";
                        query += "        TR.DSTRCT_CODE, ";
                        query += "        WT.WORK_GROUP, ";
                        query += "        TR.WORK_ORDER, ";
                        query += "        TR.WO_TASK_NO, ";
                        query += "        WT.WO_TASK_DESC, ";
                        query += "        TR.RESOURCE_TYPE, ";
                        query += "        TT.TABLE_DESC ";
                        query += "),RES_EST AS ( ";
                        query += "    SELECT ";
                        query += "        TSK.DSTRCT_CODE, ";
                        query += "        TSK.WORK_GROUP, ";
                        query += "        TSK.WORK_ORDER, ";
                        query += "        TSK.WO_TASK_NO, ";
                        query += "        TSK.WO_TASK_DESC, ";
                        query += "        RS.RESOURCE_TYPE RES_CODE, ";
                        query += "        TT.TABLE_DESC RES_DESC, ";
                        query += "        TO_NUMBER(RS.CREW_SIZE) QTY_REQ, ";
                        query += "        RS.EST_RESRCE_HRS ";
                        query += "    FROM ";
                        query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                        query += "        INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS ";
                        query += "        ON RS.KEY_735_ID = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.REC_735_TYPE = 'WT' ";
                        query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                        query += "        ON TT.TABLE_CODE = RS.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                        query += "    WHERE ";
                        query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                        if (taskNo != null)
                            query += "      AND   TSK.WO_TASK_NO = '" + taskNo + "' ";
                        query += "),TABLA_REC AS ( ";
                        query += "    SELECT ";
                        query += "        DECODE(RES_EST.DSTRCT_CODE,NULL,RES_REAL.DSTRCT_CODE,RES_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                        query += "        DECODE(RES_EST.WORK_GROUP,NULL,RES_REAL.WORK_GROUP,RES_EST.WORK_GROUP) WORK_GROUP, ";
                        query += "        DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, ";
                        query += "        DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, ";
                        query += "        DECODE(RES_EST.WO_TASK_DESC,NULL,RES_REAL.WO_TASK_DESC,RES_EST.WO_TASK_DESC) WO_TASK_DESC, ";
                        query += "        DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, ";
                        query += "        DECODE(RES_EST.RES_DESC,NULL,RES_REAL.RES_DESC,RES_EST.RES_DESC) RES_DESC, ";
                        query += "        RES_EST.QTY_REQ, ";
                        query += "        RES_REAL.ACT_RESRCE_HRS, ";
                        query += "        RES_EST.EST_RESRCE_HRS ";
                        query += "    FROM ";
                        query += "        RES_REAL ";
                        query += "        FULL JOIN RES_EST ";
                        query += "        ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO AND RES_REAL.RES_CODE = RES_EST.RES_CODE ";
                        query += ") SELECT ";
                        query += "    'MAT' REQ_TYPE, ";
                        query += "    TABLA_MAT.DSTRCT_CODE, ";
                        query += "    TABLA_MAT.WORK_GROUP, ";
                        query += "    TABLA_MAT.WORK_ORDER, ";
                        query += "    TABLA_MAT.WO_TASK_NO, ";
                        query += "    TABLA_MAT.WO_TASK_DESC, ";
                        query += "    TABLA_MAT.SEQ_NO, ";
                        query += "    TABLA_MAT.RES_CODE, ";
                        query += "    TABLA_MAT.RES_DESC, ";
                        query += "    DECODE(TABLA_MAT.UNITS, NULL, '', TABLA_MAT.UNITS) UNITS, ";
                        query += "    TABLA_MAT.QTY_REQ, ";
                        query += "    DECODE(TABLA_MAT.QTY_ISS, NULL, 0,TABLA_MAT.QTY_ISS) QTY_ISS, ";
                        query += "    0 EST_RESRCE_HRS, ";
                        query += "    0 ACT_RESRCE_HRS ";
                        query += "  FROM ";
                        query += "    TABLA_MAT ";
                        query += "UNION ALL ";
                        query += "SELECT ";
                        query += "    'LAB' REQ_TYPE, ";
                        query += "    TABLA_REC.DSTRCT_CODE, ";
                        query += "    TABLA_REC.WORK_GROUP, ";
                        query += "    TABLA_REC.WORK_ORDER, ";
                        query += "    TABLA_REC.WO_TASK_NO, ";
                        query += "    TABLA_REC.WO_TASK_DESC, ";
                        query += "    '' SEQ_NO, ";
                        query += "    TABLA_REC.RES_CODE, ";
                        query += "    TABLA_REC.RES_DESC, ";
                        query += "    '' UNITS, ";
                        query += "    TABLA_REC.QTY_REQ, ";
                        query += "    NULL QTY_ISS, ";
                        query += "    DECODE(TABLA_REC.EST_RESRCE_HRS, NULL, 0, TABLA_REC.EST_RESRCE_HRS) EST_RESRCE_HRS, ";
                        query += "    DECODE(TABLA_REC.ACT_RESRCE_HRS, NULL, 0, TABLA_REC.ACT_RESRCE_HRS) ACT_RESRCE_HRS ";
                        query += "FROM ";
                        query += "    TABLA_REC ";
                        query += "UNION ALL ";
                        query += "SELECT ";
                        query += "    'EQU' REQ_TYPE, ";
                        query += "    TSK.DSTRCT_CODE, ";
                        query += "    TSK.WORK_GROUP, ";
                        query += "    TSK.WORK_ORDER, ";
                        query += "    TSK.WO_TASK_NO, ";
                        query += "    TSK.WO_TASK_DESC, ";
                        query += "    RS.SEQNCE_NO SEQ_NO, ";
                        query += "    RS.EQPT_TYPE RES_CODE, ";
                        query += "    ET.TABLE_DESC RES_DESC, ";
                        query += "    RS.UOM UNITS, ";
                        query += "    RS.QTY_REQ, ";
                        query += "    0 QTY_ISS, ";
                        query += "    DECODE(RS.UNIT_QTY_REQD, NULL, 0, RS.UNIT_QTY_REQD) EST_RESRCE_HRS, ";
                        query += "    0 ACT_RESRCE_HRS ";
                        query += "FROM ";
                        query += "    " + dbReference + ".MSF623" + dbLink + " TSK ";
                        query += "    INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS ";
                        query += "    ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.CLASS_TYPE = 'WT' ";
                        query += "    INNER JOIN " + dbReference + ".MSF010" + dbLink + " ET ";
                        query += "    ON RS.EQPT_TYPE = ET.TABLE_CODE AND TABLE_TYPE = 'ET' ";
                        query += "WHERE ";
                        query += "    TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                        if (taskNo != null)
                            query += "AND   TSK.WO_TASK_NO = '" + taskNo + "'";
                        break;
                }

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                return query;

            }
        }
    }
}

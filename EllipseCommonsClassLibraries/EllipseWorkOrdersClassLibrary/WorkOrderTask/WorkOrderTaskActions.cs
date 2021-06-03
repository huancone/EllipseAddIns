using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Classes;
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

        public static List<TaskRequirement> FetchRequirements(EllipseFunctions ef, string districtCode, string workOrder, string reqType, string taskNo)
        {
            string sqlQuery;
            if (string.IsNullOrWhiteSpace(taskNo))
                sqlQuery = Queries.GetFetchWoRequirementsQuery(ef.DbReference, ef.DbLink, districtCode, workOrder, reqType, string.IsNullOrWhiteSpace(taskNo) ? null : taskNo.PadLeft(3, '0'));
            else
                sqlQuery = Queries.GetFetchWoTaskRequirementsQuery(ef.DbReference, ef.DbLink, districtCode, workOrder, reqType, string.IsNullOrWhiteSpace(taskNo) ? null : taskNo.PadLeft(3, '0'));
            var woTaskDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<TaskRequirement>();

            if (woTaskDataReader == null || woTaskDataReader.IsClosed)
                return list;
            
            while (woTaskDataReader.Read())
            {
                var taskReq = new TaskRequirement
                {
                    ReqType = "" + woTaskDataReader["REQ_TYPE"].ToString().Trim(),                  
                    DistrictCode = "" + woTaskDataReader["DSTRCT_CODE"].ToString().Trim(),          
                    WorkGroup = "" + woTaskDataReader["WORK_GROUP"].ToString().Trim(),              
                    WorkOrder = "" + woTaskDataReader["WORK_ORDER"].ToString().Trim(),              
                    WoTaskNo = "" + woTaskDataReader["WO_TASK_NO"].ToString().Trim(),               
                    WoTaskDesc = "" + woTaskDataReader["WO_TASK_DESC"].ToString().Trim(),           
                    SeqNo = "" + woTaskDataReader["SEQ_NO"].ToString().Trim(),                      
                    ReqCode = "" + woTaskDataReader["RES_CODE"].ToString().Trim(),                  
                    ReqDesc = "" + woTaskDataReader["RES_DESC"].ToString().Trim(),                  
                    UoM = "" + woTaskDataReader["UNITS"].ToString().Trim(),                         
                    EstSize = "" + woTaskDataReader["EST_SIZE"].ToString().Trim(),                  
                    UnitsQty = "" + woTaskDataReader["UNITS_QTY"].ToString().Trim(),                    
                    RealQty = "" + woTaskDataReader["REAL_QTY"].ToString().Trim(),
                    SharedTasks = "" + woTaskDataReader["SHARED_TASKS"].ToString().Trim(),
                    
                };
                list.Add(taskReq);
            }

            return list;
        }

        public static WorkOrderTaskServiceModifyReplyDTO ModifyWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            using (var serviceWoTask = new WorkOrderTaskService.WorkOrderTaskService())
            {
                var requestWoTask = new WorkOrderTaskServiceModifyRequestDTO();

                //se cargan los parámetros de la orden
                serviceWoTask.Url = urlService + "/WorkOrderTaskService";

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
                requestWoTask.planFinDate = woTask.PlanFinishDate ?? requestWoTask.planFinDate;
                requestWoTask.planStrTime = woTask.PlanStartTime ?? requestWoTask.planStrTime;
                requestWoTask.planFinTime = woTask.PlanFinishTime ?? requestWoTask.planFinTime;
                requestWoTask.tskDurationsHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs) ? Convert.ToDecimal(woTask.EstimatedDurationsHrs) : default(decimal);
                requestWoTask.tskDurationsHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs);
                requestWoTask.APLEquipmentGrpId = woTask.AplEquipmentGrpId ?? requestWoTask.APLEquipmentGrpId;
                requestWoTask.APLType = woTask.AplType ?? requestWoTask.APLType;
                requestWoTask.APLCompCode = woTask.AplCompCode ?? requestWoTask.APLCompCode;
                requestWoTask.APLCompModCode = woTask.AplCompModCode ?? requestWoTask.APLCompModCode;
                requestWoTask.APLSeqNo = woTask.AplSeqNo ?? requestWoTask.APLSeqNo;

                return serviceWoTask.modify(opContext, requestWoTask);
            }
                
        }

        public static WorkOrderTaskServiceCreateReplyDTO CreateWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            using (var serviceWoTask = new WorkOrderTaskService.WorkOrderTaskService())
            {
                var requestWoTask = new WorkOrderTaskServiceCreateRequestDTO();

                //se cargan los parámetros de la orden
                serviceWoTask.Url = urlService + "/WorkOrderTaskService";

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
                requestWoTask.planStrTime = woTask.PlanStartTime ?? requestWoTask.planStrTime;
                requestWoTask.planFinDate = woTask.PlanFinishDate ?? requestWoTask.planFinDate;
                requestWoTask.planFinTime = woTask.PlanFinishTime ?? requestWoTask.planFinTime;
                requestWoTask.tskDurationsHrs = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs) ? Convert.ToDecimal(woTask.EstimatedDurationsHrs) : default(decimal);
                requestWoTask.tskDurationsHrsSpecified = !string.IsNullOrWhiteSpace(woTask.EstimatedDurationsHrs);
                requestWoTask.APLEquipmentGrpId = woTask.AplEquipmentGrpId ?? requestWoTask.APLEquipmentGrpId;
                requestWoTask.APLType = woTask.AplType ?? requestWoTask.APLType;
                requestWoTask.APLCompCode = woTask.AplCompCode ?? requestWoTask.APLCompCode;
                requestWoTask.APLCompModCode = woTask.AplCompModCode ?? requestWoTask.APLCompModCode;
                requestWoTask.APLSeqNo = woTask.AplSeqNo ?? requestWoTask.APLSeqNo;

                return serviceWoTask.create(opContext, requestWoTask);
            }
        }


        public static WorkOrderTaskServiceDeleteReplyCollectionDTO DeleteWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            using (var serviceWoTask = new WorkOrderTaskService.WorkOrderTaskService())
            {
                var requestWoTask = new WorkOrderTaskServiceDeleteRequestDTO();
                var requestWoTaskList = new List<WorkOrderTaskServiceDeleteRequestDTO>();

                //se cargan los parámetros de la orden
                serviceWoTask.Url = urlService + "/WorkOrderTaskService";

                //se cargan los parámetros de la orden
                requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
                requestWoTask.workOrder = woTask.WorkOrderDto ?? requestWoTask.workOrder;
                requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');

                requestWoTaskList.Add(requestWoTask);

                return serviceWoTask.multipleDelete(opContext, requestWoTaskList.ToArray());
            }
        }

        public static ReplyMessage CompleteWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            using (var serviceWoTask = new WorkOrderTaskService.WorkOrderTaskService())
            {
                var requestWoTask = new WorkOrderTaskServiceCompleteRequestDTO();
                var requestWoTaskList = new List<WorkOrderTaskServiceCompleteRequestDTO>();

                //se cargan los parámetros de la orden
                serviceWoTask.Url = urlService + "/WorkOrderTaskService";

                //se cargan los parámetros de la orden
                requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
                requestWoTask.workOrder = woTask.SetWorkOrderDto(woTask.WorkOrder);
                requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');
                requestWoTask.completedCode = string.IsNullOrWhiteSpace(woTask.CompletedCode) ? "06" : woTask.CompletedCode;
                requestWoTask.completedBy = woTask.CompletedBy;
                requestWoTask.closedDt = woTask.ClosedDate;

                var serviceReply = serviceWoTask.complete(opContext, requestWoTask);

                var reply = new ReplyMessage();
                reply.Message = "Completed " + serviceReply.workOrder.prefix + serviceReply.workOrder.no + " " + serviceReply.WOTaskNo + " Completed Code " + serviceReply.completedCode + " - " + serviceReply.completedCodeDescription;
                return reply;
            }
        }
        public static ReplyMessage ReOpenWorkOrderTask(string urlService, WorkOrderTaskService.OperationContext opContext, WorkOrderTask woTask)
        {
            using (var serviceWoTask = new WorkOrderTaskService.WorkOrderTaskService())
            {
                var requestWoTask = new WorkOrderTaskServiceReopenRequestDTO();

                //se cargan los parámetros de la orden
                serviceWoTask.Url = urlService + "/WorkOrderTaskService";

                //se cargan los parámetros de la orden
                requestWoTask.districtCode = woTask.DistrictCode ?? requestWoTask.districtCode;
                requestWoTask.workOrder = woTask.SetWorkOrderDto(woTask.WorkOrder);
                requestWoTask.WOTaskNo = woTask.WoTaskNo ?? requestWoTask.WOTaskNo.PadLeft(3, '0');

                var serviceReply = serviceWoTask.reopen(opContext, requestWoTask);

                var reply = new ReplyMessage();
                reply.Message = "ReOpen " + serviceReply.workOrder.prefix + serviceReply.workOrder.no + " " + serviceReply.WOTaskNo;
                return reply;
            }
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

        public static ResourceReqmntsServiceCreateReplyDTO CreateTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new ResourceReqmntsService.ResourceReqmntsService
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
                quantityRequired = !string.IsNullOrWhiteSpace(taskReq.EstSize) ? Convert.ToDecimal(taskReq.EstSize) : default(decimal),
                quantityRequiredSpecified = !string.IsNullOrWhiteSpace(taskReq.EstSize),
                hrsReqd = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                hrsReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
                classType = "WT",
                enteredInd = "S"
            };
            return serviceTaskReq.create(opContext, requestTaskReq);
        }

        public static MaterialReqmntsServiceCreateReplyDTO CreateTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new MaterialReqmntsService.MaterialReqmntsService
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
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                unitQuantityReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
            };
            return serviceTaskReq.create(opContext, requestTaskReq);
        }

        public static EquipmentReqmntsServiceCreateReplyCollectionDTO CreateTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
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
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.EstSize) ? Convert.ToDecimal(taskReq.EstSize) : default(decimal),
                unitQuantityReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.EstSize),
                quantityRequired = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                quantityRequiredSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
                UOM = taskReq.UoM,
                contestibleFlg = false,
                contestibleFlgSpecified = true,
                classType = "WT",
                enteredInd = "S",
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false,
            };

            requestTaskReqList.Add(requestTaskReq);
            return serviceTaskReq.multipleCreate(opContext, requestTaskReqList.ToArray());
        }

        public static ResourceReqmntsServiceModifyReplyCollectionDTO ModifyTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new ResourceReqmntsService.ResourceReqmntsService
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
                quantityRequired = !string.IsNullOrWhiteSpace(taskReq.EstSize) ? Convert.ToDecimal(taskReq.EstSize) : default(decimal),
                quantityRequiredSpecified = !string.IsNullOrWhiteSpace(taskReq.EstSize),
                hrsReqd = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                hrsReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
                classType = "WT",
                enteredInd = "S"
            };

            requestTaskReqList.Add(requestTaskReq);

            return serviceTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static MaterialReqmntsServiceModifyReplyCollectionDTO ModifyTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new MaterialReqmntsService.MaterialReqmntsService
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
                classType = "WT",
                enteredInd = "S",
                catalogueFlag = true,
                catalogueFlagSpecified = true,
                contestibleFlag = false,
                contestibleFlagSpecified = true,
                totalOnlyFlg = true,
                CUItemNoSpecified = false,
                JEItemNoSpecified = false,
                fixedAmountSpecified = false,
                rateAmountSpecified = false,
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                unitQuantityReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
            };

            requestTaskReqList.Add(requestTaskReq);
            return serviceTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static EquipmentReqmntsServiceModifyReplyCollectionDTO ModifyTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
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
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.EstSize) ? Convert.ToDecimal(taskReq.EstSize) : default(decimal),
                unitQuantityReqdSpecified = !string.IsNullOrWhiteSpace(taskReq.EstSize),
                quantityRequired = !string.IsNullOrWhiteSpace(taskReq.UnitsQty) ? Convert.ToDecimal(taskReq.UnitsQty) : default(decimal),
                quantityRequiredSpecified = !string.IsNullOrWhiteSpace(taskReq.UnitsQty),
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
            return serviceTaskReq.multipleModify(opContext, requestTaskReqList.ToArray());
        }

        public static ResourceReqmntsServiceDeleteReplyCollectionDTO DeleteTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new ResourceReqmntsService.ResourceReqmntsService
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

            return serviceTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static MaterialReqmntsServiceDeleteReplyCollectionDTO DeleteTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new MaterialReqmntsService.MaterialReqmntsService
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
            return serviceTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static EquipmentReqmntsServiceDeleteReplyCollectionDTO DeleteTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var serviceTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService
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
            return serviceTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static List<WorkOrderTask> FetchWorkOrderTask(EllipseFunctions ef, string districtCode, string workOrder, string woTaskNo)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchWorkOrderTasksQuery(ef.DbReference, ef.DbLink, districtCode, workOrder, woTaskNo));

            var list = new List<WorkOrderTask>();

            if (stdDataReader == null || stdDataReader.IsClosed)
                return list;

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
                task.PlanStartTime = "" + stdDataReader["PLAN_STR_TIME"].ToString().Trim();
                task.PlanFinishDate = "" + stdDataReader["PLAN_FIN_DATE"].ToString().Trim();
                task.PlanFinishTime = "" + stdDataReader["PLAN_FIN_TIME"].ToString().Trim();

                task.EstimatedDurationsHrs = "" + stdDataReader["TSK_DUR_HOURS"].ToString().Trim();
                task.NoLabor = "" + stdDataReader["NO_REC_LABOR"].ToString().Trim();
                task.NoMaterial = "" + stdDataReader["NO_REC_MATERIAL"].ToString().Trim();

                task.AplEquipmentGrpId = "" + stdDataReader["EQUIP_GRP_ID"].ToString().Trim();
                task.AplType = "" + stdDataReader["APL_TYPE"].ToString().Trim();
                task.AplCompCode = "" + stdDataReader["COMP_CODE"].ToString().Trim();
                task.AplCompModCode = "" + stdDataReader["COMP_MOD_CODE"].ToString().Trim();
                task.AplSeqNo = "" + stdDataReader["APL_SEQ_NO"].ToString().Trim();
                task.TaskStatusM = "" + stdDataReader["TASK_STATUS_M"].ToString().Trim();
                task.ClosedStatus = "" + stdDataReader["CLOSED_STATUS"].ToString().Trim();
                task.CompletedCode = "" + stdDataReader["COMPLETED_CODE"].ToString().Trim();
                task.CompletedBy = "" + stdDataReader["COMPLETED_BY"].ToString().Trim();
                task.ClosedDate = "" + stdDataReader["CLOSED_DT"].ToString().Trim();

                list.Add(task);
            }

            return list;
        }

        
    }
}

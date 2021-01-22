using System;
using System.Collections.Generic;
using System.Globalization;
using EllipseStandardJobsClassLibrary.EquipmentReqmntsService;
using EllipseStandardJobsClassLibrary.ResourceReqmntsService;
using EllipseStandardJobsClassLibrary.StandardJobTaskService;
using EllipseStdTextClassLibrary;

namespace EllipseStandardJobsClassLibrary
{
    public static class StandardJobTaskActions
    {
        public static KeyValuePair<string, string> Create = new KeyValuePair<string, string>("C", "Create");
        public static KeyValuePair<string, string> Modify = new KeyValuePair<string, string>("M", "Modify");
        public static KeyValuePair<string, string> Delete = new KeyValuePair<string, string>("D", "Delete");

        public static List<KeyValuePair<string, string>> GetTaskActionCodes()
        {
            var list = new List<KeyValuePair<string, string>> { Create, Modify, Delete };

            return list;
        }

        /// <summary>
        /// Actualiza un standardJob con la información de la clase StandardJob
        /// </summary>
        /// <param name="urlService">string: URL del servicio</param>
        /// <param name="opContext">StandardJobService.OperationContext: Contexto de operación de Ellipse</param>
        /// <param name="stdTask">StandardJobTask: Objeto clase de StandardJobTask</param>
        /// <param name="overrideActive">bool: (Opcional) Indica si se actualiza sin importar si el Standard está activo/inactivo</param>
        public static void ModifyStandardJobTask(string urlService, StandardJobTaskService.OperationContext opContext, StandardJobTask stdTask, bool overrideActive = false)
        {
            var proxyStdTask = new StandardJobTaskService.StandardJobTaskService();//ejecuta las acciones del servicio
            var requestStdTask = new StandardJobTaskServiceModifyRequestDTO();

            //se cargan los parámetros de la orden
            proxyStdTask.Url = urlService + "/StandardJobServiceTask";

            //se cargan los parámetros de la orden
            requestStdTask.districtCode = stdTask.DistrictCode ?? requestStdTask.districtCode;
            requestStdTask.workGroup = stdTask.WorkGroup ?? requestStdTask.workGroup;
            requestStdTask.standardJob = stdTask.StandardJob ?? requestStdTask.standardJob;

            requestStdTask.SJTaskNo = stdTask.SjTaskNo ?? requestStdTask.SJTaskNo.PadLeft(3, '0');
            requestStdTask.SJTaskDesc = stdTask.SjTaskDesc ?? requestStdTask.SJTaskDesc;
            requestStdTask.jobDescCode = stdTask.JobDescCode ?? requestStdTask.jobDescCode;
            requestStdTask.safetyInstr = stdTask.SafetyInstr ?? requestStdTask.safetyInstr;
            requestStdTask.completeInstr = stdTask.CompleteInstr ?? requestStdTask.completeInstr;
            requestStdTask.complTextCode = stdTask.ComplTextCode ?? requestStdTask.complTextCode;
            requestStdTask.assignPerson = stdTask.AssignPerson ?? requestStdTask.assignPerson;
            requestStdTask.estimatedMachHrs = !string.IsNullOrEmpty(stdTask.EstimatedMachHrs) ? Convert.ToDecimal(stdTask.EstimatedMachHrs) : default(decimal);
            requestStdTask.estimatedMachHrsSpecified = !string.IsNullOrEmpty(stdTask.EstimatedMachHrs);
            requestStdTask.unitOfWork = stdTask.UnitOfWork ?? requestStdTask.unitOfWork;
            requestStdTask.unitsRequired = !string.IsNullOrEmpty(stdTask.UnitsRequired) ? Convert.ToDecimal(stdTask.UnitsRequired) : default(decimal);
            requestStdTask.unitsRequiredSpecified = !string.IsNullOrEmpty(stdTask.UnitsRequired);
            requestStdTask.unitsPerDay = !string.IsNullOrEmpty(stdTask.UnitsPerDay) ? Convert.ToDecimal(stdTask.UnitsPerDay) : default(decimal);
            requestStdTask.unitsPerDaySpecified = !string.IsNullOrEmpty(stdTask.UnitsPerDay);
            requestStdTask.estimatedDurationsHrs = !string.IsNullOrEmpty(stdTask.EstimatedDurationsHrs) ? Convert.ToDecimal(stdTask.EstimatedDurationsHrs) : default(decimal);
            requestStdTask.estimatedDurationsHrsSpecified = !string.IsNullOrEmpty(stdTask.EstimatedDurationsHrs);
            requestStdTask.APLEquipmentGrpId = stdTask.AplEquipmentGrpId ?? requestStdTask.APLEquipmentGrpId;
            requestStdTask.APLType = stdTask.AplType ?? requestStdTask.APLType;
            requestStdTask.APLCompCode = stdTask.AplCompCode ?? requestStdTask.APLCompCode;
            requestStdTask.APLCompModCode = stdTask.AplCompModCode ?? requestStdTask.APLCompModCode;
            requestStdTask.APLSeqNo = stdTask.AplSeqNo ?? requestStdTask.APLSeqNo;

            proxyStdTask.modify(opContext, requestStdTask);
        }

        /// <summary>
        /// Actualiza un standardJob con la información de la clase StandardJob
        /// </summary>
        /// <param name="urlService">string: URL del servicio</param>
        /// <param name="opContext">StandardJobService.OperationContext: Contexto de operación de Ellipse</param>
        /// <param name="stdTask">StandardJobTask: Objeto clase de StandardJobTask</param>
        /// <param name="overrideActive">bool: (Opcional) Indica si se actualiza sin importar si el Standard está activo/inactivo</param>
        public static void CreateStandardJobTask(string urlService, StandardJobTaskService.OperationContext opContext, StandardJobTask stdTask, bool overrideActive = false)
        {
            var proxyStdTask = new StandardJobTaskService.StandardJobTaskService();//ejecuta las acciones del servicio
            var requestStdTask = new StandardJobTaskServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            proxyStdTask.Url = urlService + "/StandardJobServiceTask";

            //se cargan los parámetros de la orden
            requestStdTask.districtCode = stdTask.DistrictCode ?? requestStdTask.districtCode;
            requestStdTask.workGroup = stdTask.WorkGroup ?? requestStdTask.workGroup;
            requestStdTask.standardJob = stdTask.StandardJob ?? requestStdTask.standardJob;

            requestStdTask.SJTaskNo = stdTask.SjTaskNo ?? requestStdTask.SJTaskNo.PadLeft(3, '0');
            requestStdTask.SJTaskDesc = stdTask.SjTaskDesc ?? requestStdTask.SJTaskDesc;
            requestStdTask.jobDescCode = stdTask.JobDescCode ?? requestStdTask.jobDescCode;
            requestStdTask.safetyInstr = stdTask.SafetyInstr ?? requestStdTask.safetyInstr;
            requestStdTask.completeInstr = stdTask.CompleteInstr ?? requestStdTask.completeInstr;
            requestStdTask.complTextCode = stdTask.ComplTextCode ?? requestStdTask.complTextCode;
            requestStdTask.assignPerson = stdTask.AssignPerson ?? requestStdTask.assignPerson;
            requestStdTask.estimatedMachHrs = stdTask.EstimatedMachHrs != null ? Convert.ToDecimal(stdTask.EstimatedMachHrs) : default(decimal);
            requestStdTask.estimatedMachHrsSpecified = stdTask.EstimatedMachHrs != null;
            requestStdTask.unitOfWork = stdTask.UnitOfWork ?? requestStdTask.unitOfWork;
            requestStdTask.unitsRequired = stdTask.UnitsRequired != null ? Convert.ToDecimal(stdTask.UnitsRequired) : default(decimal);
            requestStdTask.unitsRequiredSpecified = stdTask.UnitsRequired != null;
            requestStdTask.unitsPerDay = stdTask.UnitsPerDay != null ? Convert.ToDecimal(stdTask.UnitsPerDay) : default(decimal);
            requestStdTask.unitsPerDaySpecified = stdTask.UnitsPerDay != null;
            requestStdTask.estimatedDurationsHrs = stdTask.EstimatedDurationsHrs != null ? Convert.ToDecimal(stdTask.EstimatedDurationsHrs) : default(decimal);
            requestStdTask.estimatedDurationsHrsSpecified = stdTask.EstimatedDurationsHrs != null;
            requestStdTask.APLEquipmentGrpId = stdTask.AplEquipmentGrpId ?? requestStdTask.APLEquipmentGrpId;
            requestStdTask.APLEGIRef = stdTask.AplEgiRef ?? requestStdTask.APLEGIRef;
            requestStdTask.APLType = stdTask.AplType ?? requestStdTask.APLType;
            requestStdTask.APLCompCode = stdTask.AplCompCode ?? requestStdTask.APLCompCode;
            requestStdTask.APLCompModCode = stdTask.AplCompModCode ?? requestStdTask.APLCompModCode;
            requestStdTask.APLSeqNo = stdTask.AplSeqNo ?? requestStdTask.APLSeqNo;

            proxyStdTask.create(opContext, requestStdTask);
        }

        public static void DeleteStandardJobTask(string urlService, StandardJobTaskService.OperationContext opContext, StandardJobTask stdTask)
        {
            var proxyStdTask = new StandardJobTaskService.StandardJobTaskService();//ejecuta las acciones del servicio
            var requestStdTask = new StandardJobTaskServiceDeleteRequestDTO();

            proxyStdTask.Url = urlService + "/StandardJobServiceTask";
            requestStdTask.districtCode = stdTask.DistrictCode ?? requestStdTask.districtCode;
            requestStdTask.standardJob = stdTask.StandardJob ?? requestStdTask.standardJob;
            requestStdTask.SJTaskNo = stdTask.SjTaskNo ?? requestStdTask.SJTaskNo.PadLeft(3, '0');
            proxyStdTask.delete(opContext, requestStdTask);
        }

        public static void SetStandardJobTaskText(string urlService, string districtCode, string position, bool returnWarnings, StandardJobTask stdTask)
        {
            if (!string.IsNullOrWhiteSpace(stdTask.SjTaskNo))
                stdTask.SjTaskNo = stdTask.SjTaskNo.PadLeft(3, '0');//comentario
            var stdTextId = "JI" + districtCode + stdTask.StandardJob + stdTask.SjTaskNo;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, stdTask.ExtTaskText);
        }


        public static void CreateTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq, bool overrideActive = false)
        {
            var serviceTaskReq = new ResourceReqmntsService.ResourceReqmntsService();//ejecuta las acciones del servicio
            var requestTaskReq = new ResourceReqmntsServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            serviceTaskReq.Url = urlService + "/ResourceReqmntsService";

            //se cargan los parámetros de la orden

            requestTaskReq.districtCode = taskReq.DistrictCode ?? requestTaskReq.districtCode;
            requestTaskReq.stdJobNo = taskReq.StandardJob ?? requestTaskReq.stdJobNo;
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            requestTaskReq.SJTaskNo = taskReq.SJTaskNo ?? requestTaskReq.SJTaskNo;

            requestTaskReq.resourceClass = taskReq.ReqCode.Substring(0, 1);
            requestTaskReq.resourceCode = taskReq.ReqCode.Substring(1);
            requestTaskReq.quantityRequired = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal);
            requestTaskReq.quantityRequiredSpecified = true;
            requestTaskReq.hrsReqd = taskReq.HrsReq != null ? Convert.ToDecimal(taskReq.HrsReq) : default(decimal);
            requestTaskReq.hrsReqdSpecified = true;
            requestTaskReq.classType = "ST";
            requestTaskReq.enteredInd = "S";

            serviceTaskReq.create(opContext, requestTaskReq);
        }

        public static void CreateTaskMaterial(string urlService, MaterialReqmntsService.OperationContext opContext, TaskRequirement taskReq, bool overrideActive = false)
        {
            var proxyTaskReq = new MaterialReqmntsService.MaterialReqmntsService();//ejecuta las acciones del servicio
            var requestTaskReq = new MaterialReqmntsService.MaterialReqmntsServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            proxyTaskReq.Url = urlService + "/MaterialReqmntsService";

            //se cargan los parámetros de la orden

            requestTaskReq.districtCode = taskReq.DistrictCode ?? requestTaskReq.districtCode;
            requestTaskReq.stdJobNo = taskReq.StandardJob ?? requestTaskReq.stdJobNo;
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            requestTaskReq.SJTaskNo = taskReq.SJTaskNo ?? requestTaskReq.SJTaskNo;

            requestTaskReq.seqNo = taskReq.SeqNo ?? requestTaskReq.seqNo;
            requestTaskReq.stockCode = taskReq.ReqCode.Substring(1).PadLeft(9, '0');
            requestTaskReq.unitQuantityReqd = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal);
            requestTaskReq.unitQuantityReqdSpecified = true;
            requestTaskReq.catalogueFlag = true;
            requestTaskReq.catalogueFlagSpecified = true;
            requestTaskReq.classType = "ST";
            requestTaskReq.contestibleFlag = false;
            requestTaskReq.contestibleFlagSpecified = true;
            requestTaskReq.enteredInd = "S";
            requestTaskReq.totalOnlyFlg = true;

            requestTaskReq.CUItemNoSpecified = false;
            requestTaskReq.JEItemNoSpecified = false;
            requestTaskReq.fixedAmountSpecified = false;
            requestTaskReq.rateAmountSpecified = false;


            proxyTaskReq.create(opContext, requestTaskReq);
        }

        public static void ModifyTaskResource(string urlService, ResourceReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new ResourceReqmntsService.ResourceReqmntsService
            {
                Url = urlService + "/ResourceReqmntsService"
            };

            var requestTaskReqList = new List<ResourceReqmntsServiceModifyRequestDTO>();

            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new ResourceReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = !string.IsNullOrWhiteSpace(taskReq.SJTaskNo) ? taskReq.SJTaskNo : null,
                resourceClass = taskReq.ReqCode.Substring(0, 1),
                resourceCode = taskReq.ReqCode.Substring(1),
                quantityRequired = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                quantityRequiredSpecified = true,
                hrsReqd = taskReq.HrsReq != null ? Convert.ToDecimal(taskReq.HrsReq) : default(decimal),
                hrsReqdSpecified = true,
                classType = "ST",
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

            var requestTaskReqList = new List<MaterialReqmntsService.MaterialReqmntsServiceModifyRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new MaterialReqmntsService.MaterialReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = !string.IsNullOrWhiteSpace(taskReq.SJTaskNo) ? taskReq.SJTaskNo : null,
                seqNo = taskReq.SeqNo,
                stockCode = taskReq.ReqCode.Substring(1).PadLeft(9, '0'),
                unitQuantityReqd = !string.IsNullOrWhiteSpace(taskReq.QtyReq) ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                unitQuantityReqdSpecified = true,
                catalogueFlag = true,
                catalogueFlagSpecified = true,
                classType = "ST",
                contestibleFlag = false,
                contestibleFlagSpecified = true,
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
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new ResourceReqmntsServiceDeleteRequestDTO()
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = Convert.ToString(Convert.ToDecimal(taskReq.SJTaskNo), CultureInfo.InvariantCulture),
                resourceClass = taskReq.ReqCode.Substring(0, 1),
                resourceCode = taskReq.ReqCode.Substring(1),
                classType = "ST",
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

            var requestTaskReqList = new List<MaterialReqmntsService.MaterialReqmntsServiceDeleteRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new MaterialReqmntsService.MaterialReqmntsServiceDeleteRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = Convert.ToString(Convert.ToDecimal(taskReq.SJTaskNo), CultureInfo.InvariantCulture),
                seqNo = taskReq.SeqNo,
                classType = "ST",
                enteredInd = "S",
                CUItemNoSpecified = false,
                JEItemNoSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }

        public static void CreateTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService()
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceCreateRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new EquipmentReqmntsServiceCreateRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = !string.IsNullOrWhiteSpace(taskReq.SJTaskNo) ? taskReq.SJTaskNo : null,
                seqNo = taskReq.SeqNo,
                eqptType = taskReq.ReqCode.Substring(1),
                unitQuantityReqd = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                unitQuantityReqdSpecified = true,
                UOM = taskReq.UoM,
                contestibleFlg = false,
                contestibleFlgSpecified = true,
                classType = "ST",
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

        public static void ModifyTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService()
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceModifyRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new EquipmentReqmntsServiceModifyRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = !string.IsNullOrWhiteSpace(taskReq.SJTaskNo) ? taskReq.SJTaskNo : null,
                seqNo = taskReq.SeqNo,
                eqptType = taskReq.ReqCode.Substring(1),
                unitQuantityReqd = taskReq.QtyReq != null ? Convert.ToDecimal(taskReq.QtyReq) : default(decimal),
                unitQuantityReqdSpecified = true,
                UOM = taskReq.UoM,
                contestibleFlg = false,
                contestibleFlgSpecified = true,
                classType = "ST",
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

        public static void DeleteTaskEquipment(string urlService, EquipmentReqmntsService.OperationContext opContext, TaskRequirement taskReq)
        {
            var proxyTaskReq = new EquipmentReqmntsService.EquipmentReqmntsService()
            {
                Url = urlService + "/EquipmentReqmntsService"
            };

            var requestTaskReqList = new List<EquipmentReqmntsServiceDeleteRequestDTO>();
            if (!string.IsNullOrWhiteSpace(taskReq.SJTaskNo))
                taskReq.SJTaskNo = taskReq.SJTaskNo.PadLeft(3, '0');
            var requestTaskReq = new EquipmentReqmntsServiceDeleteRequestDTO
            {
                districtCode = taskReq.DistrictCode,
                stdJobNo = taskReq.StandardJob,
                SJTaskNo = Convert.ToString(Convert.ToDecimal(taskReq.SJTaskNo), CultureInfo.InvariantCulture),
                seqNo = taskReq.SeqNo,
                operationTypeEQP = taskReq.ReqCode,
                classType = "ST",
                enteredInd = "S",
                CUItemNoSpecified = false,
                JEItemNoSpecified = false
            };

            requestTaskReqList.Add(requestTaskReq);
            proxyTaskReq.multipleDelete(opContext, requestTaskReqList.ToArray());
        }
    }
}

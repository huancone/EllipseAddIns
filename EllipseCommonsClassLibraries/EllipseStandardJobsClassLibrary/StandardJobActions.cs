using System;
using System.Collections.Generic;
using EllipseReferenceCodesClassLibrary;
using EllipseStandardJobsClassLibrary.StandardJobService;
using EllipseStdTextClassLibrary;
using SharedClassLibrary.Ellipse;

namespace EllipseStandardJobsClassLibrary
{

    public static class StandardJobActions
    {
        /// <summary>
        /// Actualiza un standardJob con la información de la clase StandardJob
        /// </summary>
        /// <param name="urlService">string: URL del servicio</param>
        /// <param name="opContext">StandardJobService.OperationContext: Contexto de operación de Ellipse</param>
        /// <param name="stdJob">StandardJob: Objeto clase de StandardJob</param>
        /// <param name="overrideActive">bool: (Opcional) Indica si se actualiza sin importar si el Standard está activo/inactivo</param>
        public static void ModifyStandardJob(string urlService, StandardJobService.OperationContext opContext, StandardJob stdJob, bool overrideActive = false)
        {
            var proxyStdJob = new StandardJobService.StandardJobService();//ejecuta las acciones del servicio
            var requestStdJob = new StandardJobServiceModifyRequestDTO();

            //se cargan los parámetros de la orden
            proxyStdJob.Url = urlService + "/StandardJobService";

            //se cargan los parámetros de la orden
            requestStdJob.districtCode = stdJob.DistrictCode ?? requestStdJob.districtCode;
            requestStdJob.workGroup = stdJob.WorkGroup ?? requestStdJob.workGroup;
            requestStdJob.standardJob = stdJob.StandardJobNo ?? requestStdJob.standardJob;
            requestStdJob.status = stdJob.Status ?? requestStdJob.status;
            requestStdJob.status = string.IsNullOrWhiteSpace(requestStdJob.status) ? "I" : requestStdJob.status;
            requestStdJob.standardJobDescription = stdJob.StandardJobDescription ?? requestStdJob.standardJobDescription;

            requestStdJob.originatorId = stdJob.OriginatorId ?? requestStdJob.originatorId;
            requestStdJob.assignPerson = stdJob.AssignPerson ?? requestStdJob.assignPerson;
            requestStdJob.origPriority = stdJob.OrigPriority ?? requestStdJob.origPriority;
            requestStdJob.workOrderType = stdJob.WorkOrderType ?? requestStdJob.workOrderType;
            requestStdJob.maintenanceType = stdJob.MaintenanceType ?? requestStdJob.maintenanceType;
            requestStdJob.compCode = stdJob.CompCode ?? requestStdJob.compCode;
            requestStdJob.compModCode = stdJob.CompModCode ?? requestStdJob.compModCode;
            requestStdJob.unitOfWork = stdJob.UnitOfWork ?? requestStdJob.unitOfWork;
            requestStdJob.unitsRequired = stdJob.UnitsRequired != null ? Convert.ToDecimal(stdJob.UnitsRequired) : default(decimal);
            requestStdJob.unitsRequiredSpecified = stdJob.UnitsRequired != null;

            requestStdJob.accountCode = stdJob.AccountCode ?? requestStdJob.accountCode;
            requestStdJob.reallocAccCode = stdJob.ReallocAccCode ?? requestStdJob.reallocAccCode;
            requestStdJob.projectNo = stdJob.ProjectNo ?? requestStdJob.projectNo;
            requestStdJob.estimatedOtherCost = stdJob.EstimatedOtherCost != null ? Convert.ToDecimal(stdJob.EstimatedOtherCost) : default(decimal);
            requestStdJob.estimatedOtherCostSpecified = stdJob.EstimatedOtherCost != null;
            requestStdJob.estimatedDurationsHrs = stdJob.EstimatedDurationsHrs != null ? Convert.ToDecimal(stdJob.EstimatedDurationsHrs) : default(decimal);
            requestStdJob.estimatedDurationsHrsSpecified = stdJob.EstimatedDurationsHrs != null;
            requestStdJob.calculatedDurationsHrsFlg = Convert.ToBoolean(stdJob.CalculatedDurationsHrsFlg);
            requestStdJob.calculatedDurationsHrsFlgSpecified = stdJob.CalculatedDurationsHrsFlg != null;
            requestStdJob.resUpdateFlag = Convert.ToBoolean(stdJob.ResUpdateFlag);
            requestStdJob.resUpdateFlagSpecified = stdJob.ResUpdateFlag != null;
            requestStdJob.matUpdateFlag = Convert.ToBoolean(stdJob.MatUpdateFlag);
            requestStdJob.matUpdateFlagSpecified = stdJob.MatUpdateFlag != null;
            requestStdJob.equipmentUpdateFlag = Convert.ToBoolean(stdJob.EquipmentUpdateFlag);
            requestStdJob.equipmentUpdateFlagSpecified = stdJob.EquipmentUpdateFlag != null;
            requestStdJob.otherUpdateFlag = Convert.ToBoolean(stdJob.OtherUpdateFlag);
            requestStdJob.otherUpdateFlagSpecified = stdJob.OtherUpdateFlag != null;
            //requestStdJob.totalUpdateFlag = true;//no se encuentran en la base de datos. Al parecer son calculados según la selección de los anteriores
            //requestStdJob.totalUpdateFlagSpecified = true;

            requestStdJob.jobCode1 = stdJob.JobCode1 ?? requestStdJob.jobCode1;
            requestStdJob.jobCode2 = stdJob.JobCode2 ?? requestStdJob.jobCode2;
            requestStdJob.jobCode3 = stdJob.JobCode3 ?? requestStdJob.jobCode3;
            requestStdJob.jobCode4 = stdJob.JobCode4 ?? requestStdJob.jobCode4;
            requestStdJob.jobCode5 = stdJob.JobCode5 ?? requestStdJob.jobCode5;
            requestStdJob.jobCode6 = stdJob.JobCode6 ?? requestStdJob.jobCode6;
            requestStdJob.jobCode7 = stdJob.JobCode7 ?? requestStdJob.jobCode7;
            requestStdJob.jobCode8 = stdJob.JobCode8 ?? requestStdJob.jobCode8;
            requestStdJob.jobCode9 = stdJob.JobCode9 ?? requestStdJob.jobCode9;
            requestStdJob.jobCode10 = stdJob.JobCode10 ?? requestStdJob.jobCode10;
            //

            requestStdJob.accountCode = stdJob.AccountCode ?? requestStdJob.accountCode;

            //para poder modificar el standard debe estar inactivo
            var stdStatus = requestStdJob.status;
            if (overrideActive)
                if (!requestStdJob.status.Equals("I"))
                    requestStdJob.status = "I";

            //se envía la acción
            proxyStdJob.modify(opContext, requestStdJob);

            var requestEstimates = new StandardJobServiceUpdateEstimatesRequestDTO
            {
                districtCode = requestStdJob.districtCode,
                standardJob = requestStdJob.standardJob,
                equipmentUpdateFlag = requestStdJob.equipmentUpdateFlag,
                equipmentUpdateFlagSpecified = requestStdJob.equipmentUpdateFlagSpecified,
                estimatedDurationsHrs = requestStdJob.estimatedDurationsHrs,
                estimatedDurationsHrsSpecified = requestStdJob.estimatedDurationsHrsSpecified,
                resUpdateFlag = requestStdJob.resUpdateFlag,
                resUpdateFlagSpecified = requestStdJob.resUpdateFlagSpecified,
                matUpdateFlag = requestStdJob.matUpdateFlag,
                matUpdateFlagSpecified = requestStdJob.matUpdateFlagSpecified,
                //totalUpdateFlag = requestStdJob.totalUpdateFlag,//no se encuentran en la base de datos. Al parecer son calculados según la selección de los anteriores
                //totalUpdateFlagSpecified = requestStdJob.totalUpdateFlagSpecified,
                otherUpdateFlag = requestStdJob.otherUpdateFlag,
                otherUpdateFlagSpecified = requestStdJob.otherUpdateFlagSpecified
            };

            proxyStdJob.updateEstimates(opContext, requestEstimates);



            //restaura al estado original
            if (overrideActive && !stdStatus.Equals("I"))
                UpdateStandardJobStatus(urlService, opContext, stdJob, stdStatus);
        }

        /// <summary>
        /// Actualiza un standardJob con la información de la clase StandardJob
        /// </summary>
        /// <param name="urlService">string: URL del servicio</param>
        /// <param name="opContext">StandardJobService.OperationContext: Contexto de operación de Ellipse</param>
        /// <param name="stdJob">StandardJob: Objeto clase de StandardJob</param>
        public static void CreateStandardJob(string urlService, StandardJobService.OperationContext opContext, StandardJob stdJob)
        {
            var proxyStdJob = new StandardJobService.StandardJobService();//ejecuta las acciones del servicio
            var requestStdJob = new StandardJobServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            proxyStdJob.Url = urlService + "/StandardJobService";

            //se cargan los parámetros de la orden
            requestStdJob.districtCode = stdJob.DistrictCode ?? requestStdJob.districtCode;
            requestStdJob.workGroup = stdJob.WorkGroup ?? requestStdJob.workGroup;
            requestStdJob.standardJob = stdJob.StandardJobNo ?? requestStdJob.standardJob;
            requestStdJob.status = stdJob.Status ?? requestStdJob.status;
            requestStdJob.status = string.IsNullOrWhiteSpace(requestStdJob.status) ? "I" : requestStdJob.status;
            requestStdJob.standardJobDescription = stdJob.StandardJobDescription ?? requestStdJob.standardJobDescription;

            requestStdJob.originatorId = stdJob.OriginatorId ?? requestStdJob.originatorId;
            requestStdJob.assignPerson = stdJob.AssignPerson ?? requestStdJob.assignPerson;
            requestStdJob.origPriority = stdJob.OrigPriority ?? requestStdJob.origPriority;
            requestStdJob.workOrderType = stdJob.WorkOrderType ?? requestStdJob.workOrderType;
            requestStdJob.maintenanceType = stdJob.MaintenanceType ?? requestStdJob.maintenanceType;
            requestStdJob.compCode = stdJob.CompCode ?? requestStdJob.compCode;
            requestStdJob.compModCode = stdJob.CompModCode ?? requestStdJob.compModCode;

            requestStdJob.unitOfWork = !string.IsNullOrWhiteSpace(stdJob.UnitOfWork) ? stdJob.UnitOfWork : null;
            requestStdJob.unitsRequired = !string.IsNullOrWhiteSpace(stdJob.UnitOfWork) ? Convert.ToDecimal(stdJob.UnitsRequired) : default(decimal); ;
            requestStdJob.unitsRequiredSpecified = !string.IsNullOrWhiteSpace(stdJob.UnitOfWork);

            requestStdJob.accountCode = stdJob.AccountCode ?? requestStdJob.accountCode;
            requestStdJob.reallocAccCode = stdJob.ReallocAccCode ?? requestStdJob.reallocAccCode;
            requestStdJob.projectNo = stdJob.ProjectNo ?? requestStdJob.projectNo;
            requestStdJob.estimatedOtherCost = !string.IsNullOrWhiteSpace(stdJob.EstimatedOtherCost) ? Convert.ToDecimal(stdJob.EstimatedOtherCost) : default(decimal);
            requestStdJob.estimatedOtherCostSpecified = !string.IsNullOrWhiteSpace(stdJob.EstimatedOtherCost);
            requestStdJob.estimatedDurationsHrs = !string.IsNullOrWhiteSpace(stdJob.EstimatedDurationsHrs) ? Convert.ToDecimal(stdJob.EstimatedDurationsHrs) : default(decimal);
            requestStdJob.estimatedDurationsHrsSpecified = !string.IsNullOrWhiteSpace(stdJob.EstimatedDurationsHrs);
            requestStdJob.calculatedDurationsHrsFlg = Convert.ToBoolean(stdJob.CalculatedDurationsHrsFlg);
            requestStdJob.calculatedDurationsHrsFlgSpecified = !string.IsNullOrWhiteSpace(stdJob.CalculatedDurationsHrsFlg);
            requestStdJob.resUpdateFlag = Convert.ToBoolean(stdJob.ResUpdateFlag);
            requestStdJob.resUpdateFlagSpecified = !string.IsNullOrWhiteSpace(stdJob.ResUpdateFlag);
            requestStdJob.matUpdateFlag = Convert.ToBoolean(stdJob.MatUpdateFlag);
            requestStdJob.matUpdateFlagSpecified = !string.IsNullOrWhiteSpace(stdJob.MatUpdateFlag);
            requestStdJob.equipmentUpdateFlag = Convert.ToBoolean(stdJob.EquipmentUpdateFlag);
            requestStdJob.equipmentUpdateFlagSpecified = !string.IsNullOrWhiteSpace(stdJob.EquipmentUpdateFlag);
            requestStdJob.otherUpdateFlag = Convert.ToBoolean(stdJob.OtherUpdateFlag);
            requestStdJob.otherUpdateFlagSpecified = !string.IsNullOrWhiteSpace(stdJob.OtherUpdateFlag);
            //requestStdJob.totalUpdateFlag = true;//no se encuentran en la base de datos. Al parecer son calculados según la selección de los anteriores
            //requestStdJob.totalUpdateFlagSpecified = true;

            requestStdJob.jobCode1 = stdJob.JobCode1 ?? requestStdJob.jobCode1;
            requestStdJob.jobCode2 = stdJob.JobCode2 ?? requestStdJob.jobCode2;
            requestStdJob.jobCode3 = stdJob.JobCode3 ?? requestStdJob.jobCode3;
            requestStdJob.jobCode4 = stdJob.JobCode4 ?? requestStdJob.jobCode4;
            requestStdJob.jobCode5 = stdJob.JobCode5 ?? requestStdJob.jobCode5;
            requestStdJob.jobCode6 = stdJob.JobCode6 ?? requestStdJob.jobCode6;
            requestStdJob.jobCode7 = stdJob.JobCode7 ?? requestStdJob.jobCode7;
            requestStdJob.jobCode8 = stdJob.JobCode8 ?? requestStdJob.jobCode8;
            requestStdJob.jobCode9 = stdJob.JobCode9 ?? requestStdJob.jobCode9;
            requestStdJob.jobCode10 = stdJob.JobCode10 ?? requestStdJob.jobCode10;
            //

            //se envía la acción
            proxyStdJob.create(opContext, requestStdJob);

            var requestEstimates = new StandardJobServiceUpdateEstimatesRequestDTO
            {
                districtCode = requestStdJob.districtCode,
                standardJob = requestStdJob.standardJob,
                equipmentUpdateFlag = requestStdJob.equipmentUpdateFlag,
                equipmentUpdateFlagSpecified = requestStdJob.equipmentUpdateFlagSpecified,
                estimatedDurationsHrs = requestStdJob.estimatedDurationsHrs,
                estimatedDurationsHrsSpecified = requestStdJob.estimatedDurationsHrsSpecified,
                resUpdateFlag = requestStdJob.resUpdateFlag,
                resUpdateFlagSpecified = requestStdJob.resUpdateFlagSpecified,
                matUpdateFlag = requestStdJob.matUpdateFlag,
                matUpdateFlagSpecified = requestStdJob.matUpdateFlagSpecified,
                //totalUpdateFlag = requestStdJob.totalUpdateFlag,//no se encuentran en la base de datos. Al parecer son calculados según la selección de los anteriores
                //totalUpdateFlagSpecified = requestStdJob.totalUpdateFlagSpecified,
                otherUpdateFlag = requestStdJob.otherUpdateFlag,
                otherUpdateFlagSpecified = requestStdJob.otherUpdateFlagSpecified
            };

            proxyStdJob.updateEstimates(opContext, requestEstimates);

            if (stdJob.Status != null && !stdJob.Status.Equals("I"))
                UpdateStandardJobStatus(urlService, opContext, stdJob, stdJob.Status);
        }

        public static string UpdateStandardJobStatus(string urlService, StandardJobService.OperationContext opContext, StandardJob stdJob, string status)
        {
            //se envía la acción
            var requestStatus = new StandardJobServiceUpdateStatusRequestDTO
            {
                standardJob = stdJob.StandardJobNo,
                status = status
            };

            var proxyStdJob = new StandardJobService.StandardJobService
            {
                Url = urlService + "/StandardJobService"
            };

            var reply = proxyStdJob.updateStatus(opContext, requestStatus);
            return reply.status;
        }

        public static List<StandardJob> FetchStandardJob(EllipseFunctions ef, string districtCode, string workGroup, bool quickReview = false)
        {

            var sqlQuery = quickReview
                ? Queries.GetFetchQuickStandardQuery(ef.DbReference, ef.DbLink, districtCode, workGroup)
                : Queries.GetFetchStandardQuery(ef.DbReference, ef.DbLink, districtCode, workGroup);

            var stdDataReader =
                ef.GetQueryResult(sqlQuery);

            var list = new List<StandardJob>();

            if (stdDataReader == null || stdDataReader.IsClosed)
                return list;

            while (stdDataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var job = new StandardJob();
                job.DistrictCode = stdDataReader["DSTRCT_CODE"].ToString().Trim();
                job.WorkGroup = stdDataReader["WORK_GROUP"].ToString().Trim();
                job.StandardJobNo = stdDataReader["STD_JOB_NO"].ToString().Trim();
                job.Status = stdDataReader["SJ_ACTIVE_STATUS"].ToString().Trim(); //TO DO
                job.StandardJobDescription = stdDataReader["STD_JOB_DESC"].ToString().Trim();

                job.NoWos = stdDataReader["USO_OTS"].ToString().Trim();
                job.NoMsts = stdDataReader["USO_MSTS"].ToString().Trim();
                job.LastUse = stdDataReader["ULTIMO_USO"].ToString().Trim();
                job.NoTasks = stdDataReader["NO_OF_TASKS"].ToString().Trim();

                job.OriginatorId = stdDataReader["ORIGINATOR_ID"].ToString().Trim();
                job.AssignPerson = stdDataReader["ASSIGN_PERSON"].ToString().Trim();
                job.OrigPriority = stdDataReader["ORIG_PRIORITY"].ToString().Trim();
                job.WorkOrderType = stdDataReader["WO_TYPE"].ToString().Trim();
                job.MaintenanceType = stdDataReader["MAINT_TYPE"].ToString().Trim();
                job.CompCode = stdDataReader["COMP_CODE"].ToString().Trim();
                job.CompModCode = stdDataReader["COMP_MOD_CODE"].ToString().Trim();
                job.UnitOfWork = stdDataReader["UNIT_OF_WORK"].ToString().Trim();
                //El valor en base de datos de Units Required es cero cuando el UoW es nulo, pero si se envía un cero en actualización pide el UoW
                job.UnitsRequired = string.IsNullOrWhiteSpace(stdDataReader["UNIT_OF_WORK"].ToString()) && stdDataReader["UNITS_REQUIRED"].ToString().Trim() == "0" ? null : stdDataReader["UNITS_REQUIRED"].ToString().Trim();
                job.CalculatedDurationsHrsFlg = stdDataReader["CALC_DUR_HRS_SW"].ToString().Trim();
                job.EstimatedDurationsHrs = stdDataReader["EST_DUR_HRS"].ToString().Trim();

                job.AccountCode = stdDataReader["ACCOUNT_CODE"].ToString().Trim();
                job.ReallocAccCode = stdDataReader["REALL_ACCT_CDE"].ToString().Trim();
                job.ProjectNo = stdDataReader["PROJECT_NO"].ToString().Trim();

                job.EstimatedOtherCost = stdDataReader["EST_OTHER_COST"].ToString().Trim();
                job.CalculatedLabHrs = stdDataReader["CALC_LAB_HRS"].ToString().Trim();
                job.CalculatedLabCost = stdDataReader["CALC_LAB_COST"].ToString().Trim();
                job.CalculatedMatCost = stdDataReader["CALC_MAT_COST"].ToString().Trim();
                job.CalculatedEquipmentCost = stdDataReader["CALC_EQUIP_COST"].ToString().Trim();

                job.ResUpdateFlag = stdDataReader["RES_UPDATE_FLAG"].ToString().Trim();
                job.MatUpdateFlag = stdDataReader["MAT_UPDATE_FLAG"].ToString().Trim();
                job.EquipmentUpdateFlag = stdDataReader["EQUIP_UPDATE_FLAG"].ToString().Trim();

                job.JobCode1 = stdDataReader["WO_JOB_CODEX1"].ToString().Trim();
                job.JobCode2 = stdDataReader["WO_JOB_CODEX2"].ToString().Trim();
                job.JobCode3 = stdDataReader["WO_JOB_CODEX3"].ToString().Trim();
                job.JobCode4 = stdDataReader["WO_JOB_CODEX4"].ToString().Trim();
                job.JobCode5 = stdDataReader["WO_JOB_CODEX5"].ToString().Trim();
                job.JobCode6 = stdDataReader["WO_JOB_CODEX6"].ToString().Trim();
                job.JobCode7 = stdDataReader["WO_JOB_CODEX7"].ToString().Trim();
                job.JobCode8 = stdDataReader["WO_JOB_CODEX8"].ToString().Trim();
                job.JobCode9 = stdDataReader["WO_JOB_CODEX9"].ToString().Trim();
                job.JobCode10 = stdDataReader["WO_JOB_CODEX10"].ToString().Trim();

                list.Add(job);
            }

            return list;
        }

        public static StandardJob FetchStandardJob(EllipseFunctions ef, string districtCode, string workGroup, string stdJob)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchStandardQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, stdJob));

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.Read())
                return null;

            var job = new StandardJob
            {
                DistrictCode = stdDataReader["DSTRCT_CODE"].ToString().Trim(),
                WorkGroup = stdDataReader["WORK_GROUP"].ToString().Trim(),
                StandardJobNo = stdDataReader["STD_JOB_NO"].ToString().Trim(),
                Status = stdDataReader["SJ_ACTIVE_STATUS"].ToString().Trim(),
                StandardJobDescription = stdDataReader["STD_JOB_DESC"].ToString().Trim(),
                NoWos = stdDataReader["USO_OTS"].ToString().Trim(),
                NoMsts = stdDataReader["USO_MSTS"].ToString().Trim(),
                LastUse = stdDataReader["ULTIMO_USO"].ToString().Trim(),
                NoTasks = stdDataReader["NO_OF_TASKS"].ToString().Trim(),
                OriginatorId = stdDataReader["ORIGINATOR_ID"].ToString().Trim(),
                AssignPerson = stdDataReader["ASSIGN_PERSON"].ToString().Trim(),
                OrigPriority = stdDataReader["ORIG_PRIORITY"].ToString().Trim(),
                WorkOrderType = stdDataReader["WO_TYPE"].ToString().Trim(),
                MaintenanceType = stdDataReader["MAINT_TYPE"].ToString().Trim(),
                CompCode = stdDataReader["COMP_CODE"].ToString().Trim(),
                CompModCode = stdDataReader["COMP_MOD_CODE"].ToString().Trim(),
                UnitOfWork = stdDataReader["UNIT_OF_WORK"].ToString().Trim(),
                //El valor en base de datos de Units Required es cero cuando el UoW es nulo, pero si se envía un cero en actualización pide el UoW
                UnitsRequired = string.IsNullOrWhiteSpace(stdDataReader["UNIT_OF_WORK"].ToString()) && stdDataReader["UNITS_REQUIRED"].ToString().Trim() == "0" ? null : stdDataReader["UNITS_REQUIRED"].ToString().Trim(),
                CalculatedDurationsHrsFlg = stdDataReader["CALC_DUR_HRS_SW"].ToString().Trim(),
                EstimatedDurationsHrs = stdDataReader["EST_DUR_HRS"].ToString().Trim(),
                AccountCode = stdDataReader["ACCOUNT_CODE"].ToString().Trim(),
                ReallocAccCode = stdDataReader["REALL_ACCT_CDE"].ToString().Trim(),
                ProjectNo = stdDataReader["PROJECT_NO"].ToString().Trim(),
                EstimatedOtherCost = stdDataReader["EST_OTHER_COST"].ToString().Trim(),
                ResUpdateFlag = stdDataReader["RES_UPDATE_FLAG"].ToString().Trim(),
                MatUpdateFlag = stdDataReader["MAT_UPDATE_FLAG"].ToString().Trim(),
                EquipmentUpdateFlag = stdDataReader["EQUIP_UPDATE_FLAG"].ToString().Trim(),
                CalculatedLabHrs = stdDataReader["CALC_LAB_HRS"].ToString().Trim(),
                CalculatedLabCost = stdDataReader["CALC_LAB_COST"].ToString().Trim(),
                CalculatedMatCost = stdDataReader["CALC_MAT_COST"].ToString().Trim(),
                CalculatedEquipmentCost = stdDataReader["CALC_EQUIP_COST"].ToString().Trim(),
                JobCode1 = stdDataReader["WO_JOB_CODEX1"].ToString().Trim(),
                JobCode2 = stdDataReader["WO_JOB_CODEX2"].ToString().Trim(),
                JobCode3 = stdDataReader["WO_JOB_CODEX3"].ToString().Trim(),
                JobCode4 = stdDataReader["WO_JOB_CODEX4"].ToString().Trim(),
                JobCode5 = stdDataReader["WO_JOB_CODEX5"].ToString().Trim(),
                JobCode6 = stdDataReader["WO_JOB_CODEX6"].ToString().Trim(),
                JobCode7 = stdDataReader["WO_JOB_CODEX7"].ToString().Trim(),
                JobCode8 = stdDataReader["WO_JOB_CODEX8"].ToString().Trim(),
                JobCode9 = stdDataReader["WO_JOB_CODEX9"].ToString().Trim(),
                JobCode10 = stdDataReader["WO_JOB_CODEX10"].ToString().Trim()
            };


            //job.totalUpdateFlag =; //no se encuentran en la base de datos. Al parecer son calculados según la selección de los anteriores
            return job;
        }

        public static List<StandardJobTask> FetchStandardJobTask(EllipseFunctions ef, string districtCode, string workGroup, string stdJob)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchStandardJobTasksQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, stdJob));

            var list = new List<StandardJobTask>();

            if (stdDataReader == null || stdDataReader.IsClosed)
                return list;

            while (stdDataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var task = new StandardJobTask();

                task.DistrictCode = "" + stdDataReader["DSTRCT_CODE"].ToString().Trim();
                task.WorkGroup = "" + stdDataReader["WORK_GROUP"].ToString().Trim();
                task.StandardJob = "" + stdDataReader["STD_JOB_NO"].ToString().Trim();
                task.StandardJobDescription = "" + stdDataReader["STD_JOB_DESC"].ToString().Trim();

                task.SjTaskNo = "" + stdDataReader["STD_JOB_TASK"].ToString().Trim();
                task.SjTaskDesc = "" + stdDataReader["SJ_TASK_DESC"].ToString().Trim();
                task.JobDescCode = "" + stdDataReader["JOB_DESC_CODE"].ToString().Trim();
                task.SafetyInstr = "" + stdDataReader["SAFETY_INSTR"].ToString().Trim();
                task.CompleteInstr = "" + stdDataReader["COMPLETE_INSTR"].ToString().Trim();
                task.ComplTextCode = "" + stdDataReader["COMPL_TEXT_CDE"].ToString().Trim();

                task.AssignPerson = "" + stdDataReader["ASSIGN_PERSON"].ToString().Trim();
                task.EstimatedMachHrs = "" + stdDataReader["EST_MACH_HRS"].ToString().Trim();
                task.UnitOfWork = "" + stdDataReader["UNIT_OF_WORK"].ToString().Trim();
                //El valor en base de datos de Units Required es cero cuando el UoW es nulo, pero si se envía un cero en actualización pide el UoW
                task.UnitsRequired = "" + (string.IsNullOrWhiteSpace(stdDataReader["UNIT_OF_WORK"].ToString()) && stdDataReader["UNITS_REQUIRED"].ToString().Trim() == "0" ? null : stdDataReader["UNITS_REQUIRED"].ToString().Trim());
                task.UnitsPerDay = "" + stdDataReader["UNITS_PER_DAY"].ToString().Trim();

                task.EstimatedDurationsHrs = "" + stdDataReader["EST_DUR_HRS"].ToString().Trim();
                task.NoLabor = "" + stdDataReader["NO_REC_LABOR"].ToString().Trim();
                task.NoMaterial = "" + stdDataReader["NO_REC_MATERIAL"].ToString().Trim();

                task.AplEquipmentGrpId = "" + stdDataReader["EQUIP_GRP_ID"].ToString().Trim();
                task.AplType = "" + stdDataReader["APL_TYPE"].ToString().Trim();
                task.AplCompCode = "" + stdDataReader["COMP_CODE"].ToString().Trim();
                task.AplCompModCode = "" + stdDataReader["COMP_MOD_CODE"].ToString().Trim();
                task.AplSeqNo = "" + stdDataReader["APL_SEQ_NO"].ToString().Trim();

                list.Add(task);
            }

            return list;
        }

        public static List<TaskRequirement> FetchTaskRequirements(EllipseFunctions ef, string districtCode, string workGroup, string stdJob, string taskNo = null)
        {
            var sqlQuery = (taskNo == null) ? Queries.GetFetchStdJobTaskRequirementsQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, stdJob)
                                          : Queries.GetFetchStdJobTaskRequirementsQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, stdJob, taskNo.PadLeft(3, '0'));

            var stdDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<TaskRequirement>();

            if (stdDataReader == null || stdDataReader.IsClosed)
                return list;

            while (stdDataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var taskReq = new TaskRequirement();

                taskReq.DistrictCode = "" + stdDataReader["DSTRCT_CODE"].ToString().Trim();
                taskReq.WorkGroup = "" + stdDataReader["WORK_GROUP"].ToString().Trim();
                taskReq.StandardJob = "" + stdDataReader["STD_JOB_NO"].ToString().Trim();
                taskReq.SJTaskNo = "" + stdDataReader["STD_JOB_TASK"].ToString().Trim();
                taskReq.SJTaskDesc = "" + stdDataReader["SJ_TASK_DESC"].ToString().Trim();
                taskReq.ReqType = "" + stdDataReader["REQ_TYPE"].ToString().Trim();
                taskReq.SeqNo = "" + stdDataReader["SEQ_NO"].ToString().Trim();
                taskReq.ReqCode = "" + stdDataReader["RES_CODE"].ToString().Trim();
                taskReq.ReqDesc = "" + stdDataReader["RES_DESC"].ToString().Trim();
                taskReq.QtyReq = "" + stdDataReader["QTY_REQ"].ToString().Trim();
                taskReq.HrsReq = "" + stdDataReader["HRS_QTY"].ToString().Trim();

                list.Add(taskReq);
            }

            return list;
        }


        public static void SetStandardJobText(string urlService, string districtCode, string position, bool returnWarnings, StandardJob stdJob)
        {
            //comentario
            var stdTextId = "SJ" + districtCode + stdJob.StandardJobNo;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, stdJob.ExtText);
        }


        public static StandardJobReferenceCodes GetStandardJobReferenceCodes(EllipseFunctions eFunctions, string urlService, StandardJobService.OperationContext opSheet, string district, string standardJobNo)
        {
            var stdRefCodes = new StandardJobReferenceCodes();

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(opSheet.district, opSheet.position, opSheet.maxInstances, opSheet.returnWarnings);
            const string entityType = "WKO";
            var entityValue = "2" + district + standardJobNo;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            var newef = new EllipseFunctions(eFunctions);
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



            stdRefCodes.WorkRequest = item001.RefCode; //001_9001
            stdRefCodes.ComentariosDuraciones = item002.RefCode; //002_9001
            stdRefCodes.ComentariosDuracionesText = item002.StdText; //002_9001
            stdRefCodes.EmpleadoId = item003.RefCode; //003_001
            stdRefCodes.NroComponente = item005.RefCode; //005_9001
            stdRefCodes.P1EqLivMed = item006.RefCode; //006_001
            stdRefCodes.P2EqMovilMinero = item007.RefCode; //007_9001
            stdRefCodes.P3ManejoSustPeligrosa = item008.RefCode; //008_9001
            stdRefCodes.P4GuardasEquipo = item009.RefCode; //009_9001
            stdRefCodes.P5Aislamiento = item010.RefCode; //010_9001
            stdRefCodes.P6TrabajosAltura = item011.RefCode; //011_9001
            stdRefCodes.P7ManejoCargas = item012.RefCode; //012_9001
            stdRefCodes.ProyectoIcn = item013.RefCode; //013_9001
            stdRefCodes.Reembolsable = item014.RefCode; //014_9001
            stdRefCodes.FechaNoConforme = item015.RefCode; //015_9001
            stdRefCodes.FechaNoConformeText = item015.StdText; //015_9001
            stdRefCodes.NoConforme = item016.RefCode; //016_001
            stdRefCodes.FechaEjecucion = item017.RefCode; //017_001
            stdRefCodes.HoraIngreso = item018.RefCode; //018_9001
            stdRefCodes.HoraSalida = item019.RefCode; //019_9001
            stdRefCodes.NombreBuque = item020.RefCode; //020_9001
            stdRefCodes.CalificacionEncuesta = item021.RefCode; //021_001
            stdRefCodes.TareaCritica = item022.RefCode; //022_001
            stdRefCodes.Garantia = item024.RefCode; //024_9001
            stdRefCodes.GarantiaText = item024.StdText; //024_9001
            stdRefCodes.CodigoCertificacion = item025.RefCode; //025_001
            stdRefCodes.FechaEntrega = item026.RefCode; //026_001
            stdRefCodes.RelacionarEv = item029.RefCode; //029_001
            stdRefCodes.Departamento = item030.RefCode; //030_9001
            stdRefCodes.Localizacion = item031.RefCode; //031_9001

            return stdRefCodes;
        }

        public static bool UpdateWorkOrderReferenceCodes(EllipseFunctions eFunctions, string urlService, StandardJobService.OperationContext opContext, string district, string standardJobNo, StandardJobReferenceCodes stdJobReferenceCodes)
        {
            const string entityType = "WKO";
            var entityValue = "2" + district + standardJobNo;
            var itemList = new List<ReferenceCodeItem>();

            var item001 = new ReferenceCodeItem(entityType, entityValue, "001", "001", stdJobReferenceCodes.WorkRequest);
            var item002 = new ReferenceCodeItem(entityType, entityValue, "002", "001", stdJobReferenceCodes.ComentariosDuraciones, null, stdJobReferenceCodes.ComentariosDuracionesText);
            var item003 = new ReferenceCodeItem(entityType, entityValue, "003", "001", stdJobReferenceCodes.EmpleadoId);
            var item005 = new ReferenceCodeItem(entityType, entityValue, "005", "001", stdJobReferenceCodes.NroComponente);
            var item006 = new ReferenceCodeItem(entityType, entityValue, "006", "001", stdJobReferenceCodes.P1EqLivMed);
            var item007 = new ReferenceCodeItem(entityType, entityValue, "007", "001", stdJobReferenceCodes.P2EqMovilMinero);
            var item008 = new ReferenceCodeItem(entityType, entityValue, "008", "001", stdJobReferenceCodes.P3ManejoSustPeligrosa);
            var item009 = new ReferenceCodeItem(entityType, entityValue, "009", "001", stdJobReferenceCodes.P4GuardasEquipo);
            var item010 = new ReferenceCodeItem(entityType, entityValue, "010", "001", stdJobReferenceCodes.P5Aislamiento);
            var item011 = new ReferenceCodeItem(entityType, entityValue, "011", "001", stdJobReferenceCodes.P6TrabajosAltura);
            var item012 = new ReferenceCodeItem(entityType, entityValue, "012", "001", stdJobReferenceCodes.P7ManejoCargas);
            var item013 = new ReferenceCodeItem(entityType, entityValue, "013", "001", stdJobReferenceCodes.ProyectoIcn);
            var item014 = new ReferenceCodeItem(entityType, entityValue, "014", "001", stdJobReferenceCodes.Reembolsable);
            var item015 = new ReferenceCodeItem(entityType, entityValue, "015", "001", stdJobReferenceCodes.FechaNoConforme, null, stdJobReferenceCodes.FechaNoConformeText);
            var item016 = new ReferenceCodeItem(entityType, entityValue, "016", "001", stdJobReferenceCodes.NoConforme);
            var item017 = new ReferenceCodeItem(entityType, entityValue, "017", "001", stdJobReferenceCodes.FechaEjecucion);
            var item018 = new ReferenceCodeItem(entityType, entityValue, "018", "001", stdJobReferenceCodes.HoraIngreso);
            var item019 = new ReferenceCodeItem(entityType, entityValue, "019", "001", stdJobReferenceCodes.HoraSalida);
            var item020 = new ReferenceCodeItem(entityType, entityValue, "020", "001", stdJobReferenceCodes.NombreBuque);
            var item021 = new ReferenceCodeItem(entityType, entityValue, "021", "001", stdJobReferenceCodes.CalificacionEncuesta);
            var item022 = new ReferenceCodeItem(entityType, entityValue, "022", "001", stdJobReferenceCodes.TareaCritica);
            var item024 = new ReferenceCodeItem(entityType, entityValue, "024", "001", stdJobReferenceCodes.Garantia, null, stdJobReferenceCodes.GarantiaText);
            var item025 = new ReferenceCodeItem(entityType, entityValue, "025", "001", stdJobReferenceCodes.CodigoCertificacion);
            var item026 = new ReferenceCodeItem(entityType, entityValue, "026", "001", stdJobReferenceCodes.FechaEntrega);
            var item029 = new ReferenceCodeItem(entityType, entityValue, "029", "001", stdJobReferenceCodes.RelacionarEv);
            var item030 = new ReferenceCodeItem(entityType, entityValue, "030", "001", stdJobReferenceCodes.Departamento);
            var item031 = new ReferenceCodeItem(entityType, entityValue, "031", "001", stdJobReferenceCodes.Localizacion);

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

            var refOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);

            foreach (var item in itemList)
                ReferenceCodeActions.ModifyRefCode(eFunctions, urlService, refOpContext, item);
            return true;
        }

        public static List<StandardJobEquipments> RetrieveStandardJobEquipments(EllipseFunctions eFunctions, string urlService, StandardJobService.OperationContext opSheet, string district, string standardJobNo)
        {
            var proxyRetrieveEquipment = new StandardJobService.StandardJobService() { Url = urlService + "/StandardJobService" };


            var standardJobEquipments = new List<StandardJobEquipments>();
            var RetrieveEquipmentRequest = new StandardJobServiceRetrieveEquipmentRequestDTO()
            {
                districtCode = district,
                standardJob = standardJobNo
            };

            var reply = proxyRetrieveEquipment.retrieveEquipment(opSheet, RetrieveEquipmentRequest);

            for (var i = 0; i < reply.equipmentRelation.Length; i++)
            {
                var eq = new StandardJobEquipments();
                eq.DistrictCode = district;
                eq.StandardJob = standardJobNo;
                eq.EquipmentGrpId = reply.equipmentRelation[i].equipmentGrpId;
                eq.EquipmentNo = reply.equipmentRelation[i].equipmentNo;
                eq.EquipmentDescription = reply.equipmentRelation[i].equipmentDescription;
                eq.CompCode = reply.equipmentRelation[i].compCode;
                eq.CompCodeDescription = reply.equipmentRelation[i].compCodeDescription;
                eq.ModCode = reply.equipmentRelation[i].modCode;
                eq.ModCodeDescription = reply.equipmentRelation[i].modCodeDescription;


                if (!string.IsNullOrWhiteSpace(reply.equipmentRelation[i].equipmentNo) || !string.IsNullOrWhiteSpace(reply.equipmentRelation[i].equipmentGrpId))
                {
                    standardJobEquipments.Add(eq);
                }
            }

            return standardJobEquipments;
        }

        public static void CreateStandardJobEquipment(EllipseFunctions eFunctions, string urlService, StandardJobService.OperationContext opSheet, StandardJobEquipments stdEquipment)
        {
            var proxyRetrieveEquipment = new StandardJobService.StandardJobService() { Url = urlService + "/StandardJobService" };

            var requestParametersList = new List<StandardJobServiceAddEquipmentRequestDTO>();
            var equipmentRelationList = new List<EquipmentRelationDTO>();

            equipmentRelationList.Add(new EquipmentRelationDTO
            {
                equipmentGrpId = stdEquipment.EquipmentGrpId,
                equipmentNo = stdEquipment.EquipmentNo,
                equipmentRef = stdEquipment.EquipmentNo,
                compCode = stdEquipment.CompCode,
                modCode = stdEquipment.ModCode
            });

            requestParametersList.Add(new StandardJobServiceAddEquipmentRequestDTO
            {

                districtCode = stdEquipment.DistrictCode,
                standardJob = stdEquipment.StandardJob,
                equipmentRelation = equipmentRelationList.ToArray()
            });

            proxyRetrieveEquipment.multipleAddEquipment(opSheet, requestParametersList.ToArray());
        }

        public static void DeleteStandardJobEquipment(EllipseFunctions eFunctions, string urlService, StandardJobService.OperationContext opSheet, StandardJobEquipments stdEquipment)
        {
            var proxyRetrieveEquipment = new StandardJobService.StandardJobService() { Url = urlService + "/StandardJobService" };

            var standardJobEquipments = new List<StandardJobEquipments>();
            var requestParametersList = new List<StandardJobServiceDeleteEquipmentRequestDTO>();
            var equipmentRelationList = new List<EquipmentRelationDTO>();

            equipmentRelationList.Add(new EquipmentRelationDTO
            {
                equipmentGrpId = stdEquipment.EquipmentGrpId,
                equipmentNo = stdEquipment.EquipmentNo,
                equipmentRef = stdEquipment.EquipmentNo,
                compCode = stdEquipment.CompCode,
                modCode = stdEquipment.ModCode
            });

            requestParametersList.Add(new StandardJobServiceDeleteEquipmentRequestDTO
            {

                districtCode = stdEquipment.DistrictCode,
                standardJob = stdEquipment.StandardJob,
                equipmentRelation = equipmentRelationList.ToArray()
            });

            proxyRetrieveEquipment.multipleDeleteEquipment(opSheet, requestParametersList.ToArray());
        }
    }

}

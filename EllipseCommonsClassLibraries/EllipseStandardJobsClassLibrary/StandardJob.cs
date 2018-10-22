using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseReferenceCodesClassLibrary;
using EllipseStdTextClassLibrary;
using EllipseStandardJobsClassLibrary.StandardJobService;
using EllipseStandardJobsClassLibrary.StandardJobTaskService;
using EllipseStandardJobsClassLibrary.ResourceReqmntsService;
using EllipseStandardJobsClassLibrary.EquipmentReqmntsService;

using OperationContext = EllipseStandardJobsClassLibrary.ResourceReqmntsService.OperationContext;

namespace EllipseStandardJobsClassLibrary
{
    public class StandardJob
    {
        public WOJobCodesDTO[] WOjCode;
        public StandardJobServiceRetrieveReplyDTO StandardJobDto { get; set; }
        public string WOjCodeActualCount;
        public string WOjCodeActualCountSpecified;
        public string WOjCodeRestart;
        public string AccountCode;
        public string AssignPerson;
        public string CalculatedDurationsHrsFlg;
        public string CalculatedDurationsHrsFlgSpecified;
        public string CompCode;
        public string CompModCode;
        public string DistrictCode;
        public string EquipmentUpdateFlag;
        public string EquipmentUpdateFlagSpecified;
        public string EstimatedDurationsHrs;
        public string EstimatedDurationsHrsSpecified;
        public string EstimatedEquipmentCost;
        public string EstimatedEquipmentCostSpecified;
        public string EstimatedLabCost;
        public string EstimatedLabCostSpecified;
        public string EstimatedLabHrs;
        public string EstimatedLabHrsSpecified;
        public string EstimatedMatCost;
        public string EstimatedMatCostSpecified;
        public string EstimatedOtherCost;
        public string EstimatedOtherCostSpecified;
        public string EstimatedTotalCost;
        public string EstimatedTotalCostSpecified;
        public string CalculatedDurationsHrs;
        public string CalculatedEquipmentCost;
        public string CalculatedLabCost;
        public string CalculatedLabHrs;
        public string CalculatedMatCost;
        public string GlobalInd;
        public string GlobalIndSpecified;
        public string JobCode1;
        public string JobCode2;
        public string JobCode3;
        public string JobCode4;
        public string JobCode5;
        public string JobCode6;
        public string JobCode7;
        public string JobCode8;
        public string JobCode9;
        public string JobCode10;
        public string LocationFrom;
        public string LocationTo;
        public string MaintenanceType;
        public string MatUpdateFlag;
        public string MatUpdateFlagSpecified;
        public string OrigPriority;
        public string OriginatorId;
        public string OtherUpdateFlag;
        public string OtherUpdateFlagSpecified;
        public string PaperHist;
        public string PaperHistSpecified;
        public string Position;
        public string ProjectNo;
        public string ReallocAccCode;
        public string RecallTime;
        public string RecallTimeSpecified;
        public string ResUpdateFlag;
        public string ResUpdateFlagSpecified;
        public string RiskCode1;
        public string RiskCode2;
        public string RiskCode3;
        public string RiskCode4;
        public string RiskCode5;
        public string RiskCode6;
        public string RiskCode7;
        public string RiskCode8;
        public string RiskCode9;
        public string RiskCode10;
        public string ShutdownEquipmentNo;
        public string ShutdownEquipmentRef;
        public string ShutdownType;
        public string StandardJobNo;
        public string StandardJobDescription;
        public string Status;
        public string TotalUpdateFlag;
        public string TotalUpdateFlagSpecified;
        public string UnitOfWork;
        public string UnitsRequired;
        public string UnitsRequiredSpecified;
        public string WorkGroup;
        public string WorkOrderType;
        //
        public string LastUse;
        public string NoTasks;
        public string NoWos;
        public string NoMsts;

        public string ExtText { get; set; }


    }

    public class StandardJobTask
    {
        public string AplCompCode;
        public string AplCompModCode;
        public string AplEgiRef;
        public string AplEquipmentGrpId;
        public string AplSeqNo;
        public string AplType;
        public string SjTaskDesc;
        public string SjTaskNo;
        public string AssignPerson;
        public string ComplTextCode;
        public string CompleteInstr;
        public string CrewType;
        public string DistrictCode;
        public string EstimatedDurationsHrs;
        public string EstimatedDurationsHrsSpecified;
        public string EstimatedMachHrs;
        public string EstimatedMachHrsSpecified;
        public string GlobalInd;
        public string GlobalIndSpecified;
        public string JobDescCode;
        public string LinkedTask;
        public string LinkedTaskSpecified;
        public string LocationFrom;
        public string LocationTo;
        public string PlanOffset;
        public string PlanOffsetSpecified;
        public string SafetyInstr;
        public string StandardJob;
        public string StandardJobDescription;
        public string UnitOfWork;
        public string UnitsPerDay;
        public string UnitsPerDaySpecified;
        public string UnitsRequired;
        public string UnitsRequiredSpecified;
        public string WorkActualClassif;
        public string WorkCentre;
        public string WorkGroup;
        //
        public string NoLabor;
        public string NoMaterial;
        public string NoEquipment;
        public string ExtTaskText { get; set; }
    }

    public class TaskRequirement
    {
        public string StandardJob;
        public string DistrictCode;
        public string WorkGroup;
        public string SJTaskDesc;
        public string SJTaskNo;
        public string ReqType;
        public string SeqNo;
        public string ReqCode;
        public string ReqDesc;
        public string QtyReq;
        public string HrsReq;
        public string UoM;
    }

    public class StandardJobReferenceCodes
    {
        //public string ExtendedDescriptionHeader;
        //public string ExtendedDescriptionBody;
        public string WorkRequest;//001_9001
        //public string WorkRequestText;//001_9001
        public string ComentariosDuraciones;//002_9001
        public string ComentariosDuracionesText;//002_9001
        public string EmpleadoId;//003_9001 //Antiguamente Comentario
        //public string ComentariosText;//003_9001 //Deshabilitado
        //public string ComentarioApertura;//004_9001
        //public string ComentarioAperturaText;//004_9001
        public string NroComponente;//005_9001
        //public string NroComponenteText;//005_9001
        public string P1EqLivMed;//006_001
        public string P2EqMovilMinero;//007_9001
        //public string P2EqMovilMineroText;//007_9001
        public string P3ManejoSustPeligrosa;//008_9001
        //public string P3ManejoSustPeligrosaText;//008_9001
        public string P4GuardasEquipo;//009_9001
        //public string P4GuardasEquipoText;//009_9001
        public string P5Aislamiento;//010_9001
        //public string P5AislamientoText;//010_9001
        public string P6TrabajosAltura;//011_9001
        //public string P6TrabajosAlturaText;//011_9001
        public string P7ManejoCargas;//012_9001
        //public string P7ManejoCargasText;//012_9001
        public string ProyectoIcn;//013_9001
        //public string ProyectoIcnText;//013_9001
        public string Reembolsable;//014_9001
        //public string ReembolsableText;//014_9001
        public string FechaNoConforme;//015_9001
        public string FechaNoConformeText;//015_9001
        public string NoConforme;//016_001
        public string FechaEjecucion;//017_001
        public string HoraIngreso;//018_9001
        //public string HoraIngresoText;//018_9001
        public string HoraSalida;//019_9001
        //public string HoraSalidaText;//019_9001
        public string NombreBuque;//020_9001
        //public string NombreBuqueText;//020_9001
        public string CalificacionEncuesta;//021_001
        public string TareaCritica;//022_001
        //public string DuracionMtg;//023_001
        public string Garantia;//024_9001
        public string GarantiaText;//024_9001
        public string CodigoCertificacion;//025_001
        public string FechaEntrega;//026_001
        //public string EnInterventoria;//027_001
        //public string FechaDevolucion;//028_001
        public string RelacionarEv;//029_001
        public string Departamento;//030_9001
        //public string DepartamentoText;//030_9001
        public string Localizacion;//031_001
    }
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
            requestStdJob.estimatedOtherCostSpecified = stdJob.EstimatedOtherCostSpecified != null;
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
                ? Queries.GetFetchQuickStandardQuery(ef.dbReference, ef.dbLink, districtCode, workGroup)
                : Queries.GetFetchStandardQuery(ef.dbReference, ef.dbLink, districtCode, workGroup);

            var stdDataReader =
                ef.GetQueryResult(sqlQuery);

            var list = new List<StandardJob>();

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
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
                job.UnitsRequired = stdDataReader["UNITS_REQUIRED"].ToString().Trim();
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
            ef.CloseConnection();
            return list;
        }

        public static StandardJob FetchStandardJob(EllipseFunctions ef, string districtCode, string workGroup, string stdJob)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchStandardQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, stdJob));

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.HasRows || !stdDataReader.Read())
            {
                ef.CloseConnection();
                return null;
            }
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
                UnitsRequired = stdDataReader["UNITS_REQUIRED"].ToString().Trim(),
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

            ef.CloseConnection();
            return job;
        }

        public static List<StandardJobTask> FetchStandardJobTask(EllipseFunctions ef, string districtCode, string workGroup, string stdJob)
        {
            var stdDataReader =
                ef.GetQueryResult(Queries.GetFetchStandardJobTasksQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, stdJob));

            var list = new List<StandardJobTask>();

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
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
                task.UnitsRequired = "" + stdDataReader["UNITS_REQUIRED"].ToString().Trim();
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
            ef.CloseConnection();
            return list;
        }

        public static List<TaskRequirement> FetchTaskRequirements(EllipseFunctions ef, string districtCode, string workGroup, string stdJob, string taskNo = null)
        {
            var sqlQuery = (taskNo == null) ? Queries.GetFetchStdJobTaskRequirementsQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, stdJob)
                                          : Queries.GetFetchStdJobTaskRequirementsQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, stdJob, taskNo.PadLeft(3, '0'));
            
            var stdDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<TaskRequirement>();

            if (stdDataReader == null || stdDataReader.IsClosed || !stdDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
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
            ef.CloseConnection();
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
            requestStdTask.estimatedDurationsHrs =!string.IsNullOrEmpty(stdTask.EstimatedDurationsHrs) ? Convert.ToDecimal(stdTask.EstimatedDurationsHrs) : default(decimal);
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
            requestStdTask.APLType = stdTask.AplType ?? requestStdTask.APLType;
            requestStdTask.APLCompCode = stdTask.AplCompCode ?? requestStdTask.APLCompCode;
            requestStdTask.APLCompModCode = stdTask.AplCompModCode ?? requestStdTask.APLCompModCode;
            requestStdTask.APLSeqNo = stdTask.AplSeqNo ?? requestStdTask.APLSeqNo;

            proxyStdTask.create(opContext, requestStdTask);
        }

        public static void ModifyStandardJobTaskPost(EllipseFunctions ef, StandardJobTask stdTask)
        {
            ef.InitiatePostConnection();

            if(!string.IsNullOrWhiteSpace(stdTask.SjTaskNo))
                stdTask.SjTaskNo = stdTask.SjTaskNo.PadLeft(3, '0');

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.enterpriseservice.ellipse.standardjobtask.StandardJobTaskService</name>";
            requestXml = requestXml + "				<operation>modify</operation>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto   uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "                 <standardJob>" + stdTask.StandardJob + "</standardJob>";
            requestXml = requestXml + "                 <districtCode>" + stdTask.DistrictCode + "</districtCode>";
            requestXml = requestXml + "                 <safetyInstr>" + stdTask.SafetyInstr + "</safetyInstr>";
            if (MyUtilities.IsTrue(stdTask.UnitsRequiredSpecified))
                requestXml = requestXml + "                 <unitsRequired>" + stdTask.UnitsRequired + "</unitsRequired>";
            requestXml = requestXml + "                 <unitOfWork>" + stdTask.UnitOfWork + "</unitOfWork>";
            if (MyUtilities.IsTrue(stdTask.EstimatedMachHrsSpecified))
                requestXml = requestXml + "                 <estimatedMachHrs>" + stdTask.EstimatedMachHrs + "</estimatedMachHrs>";
            requestXml = requestXml + "                 <sJTaskNo>" + stdTask.SjTaskNo + "</sJTaskNo>";
            requestXml = requestXml + "                 <sJTaskDesc>" + stdTask.SjTaskDesc + "</sJTaskDesc>";
            requestXml = requestXml + "                 <complTextCode>" + stdTask.ComplTextCode + "</complTextCode>";
            requestXml = requestXml + "                 <jobDescCode>" + stdTask.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "                 <workGroup>" + stdTask.WorkGroup + "</workGroup>";
            requestXml = requestXml + "                 <assignPerson>" + stdTask.AssignPerson + "</assignPerson>";
            requestXml = requestXml + "                 <completeInstr>" + stdTask.CompleteInstr + "</completeInstr>";
            if (MyUtilities.IsTrue(stdTask.EstimatedDurationsHrsSpecified))
                requestXml = requestXml + "                 <estimatedDurationsHrs>" + stdTask.EstimatedDurationsHrs + "</estimatedDurationsHrs>";
            if(MyUtilities.IsTrue(stdTask.UnitsPerDaySpecified))
                requestXml = requestXml + "                 <unitsPerDay>" + stdTask.UnitsPerDay + "</unitsPerDay>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>mse693</application>";
            requestXml = requestXml + "	<applicationPage>read</applicationPage>";
            requestXml = requestXml + "	<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";
            requestXml = requestXml.Replace("&", "&amp;");
            var responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
        }

        public static void CreateStandardJobTaskPost(EllipseFunctions ef, StandardJobTask stdTask)
        {

            ef.InitiatePostConnection();
            if (!string.IsNullOrWhiteSpace(stdTask.SjTaskNo))
                stdTask.SjTaskNo = stdTask.SjTaskNo.PadLeft(3, '0');

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.enterpriseservice.ellipse.standardjobtask.StandardJobTaskService</name>";
            requestXml = requestXml + "				<operation>create</operation>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto  uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "                 <standardJob>" + stdTask.StandardJob + "</standardJob>";
            requestXml = requestXml + "                 <districtCode>" + stdTask.DistrictCode + "</districtCode>";
            requestXml = requestXml + "                 <safetyInstr>" + stdTask.SafetyInstr + "</safetyInstr>";
            requestXml = requestXml + "                 <unitOfWork>" + stdTask.UnitOfWork + "</unitOfWork>";
            if (MyUtilities.IsTrue(stdTask.EstimatedMachHrsSpecified))
                requestXml = requestXml + "                 <estimatedMachHrs>" + stdTask.EstimatedMachHrs + "</estimatedMachHrs>";
            requestXml = requestXml + "                 <sJTaskNo>" + stdTask.SjTaskNo + "</sJTaskNo>";
            requestXml = requestXml + "                 <sJTaskDesc>" + stdTask.SjTaskDesc + "</sJTaskDesc>";
            requestXml = requestXml + "                 <complTextCode>" + stdTask.ComplTextCode + "</complTextCode>";
            requestXml = requestXml + "                 <jobDescCode>" + stdTask.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "                 <workGroup>" + stdTask.WorkGroup + "</workGroup>";
            requestXml = requestXml + "                 <assignPerson>" + stdTask.AssignPerson + "</assignPerson>";
            requestXml = requestXml + "                 <completeInstr>" + stdTask.CompleteInstr + "</completeInstr>";
            if (MyUtilities.IsTrue(stdTask.EstimatedDurationsHrsSpecified))
                requestXml = requestXml + "                 <estimatedDurationsHrs>" + stdTask.EstimatedDurationsHrs + "</estimatedDurationsHrs>";
            if (MyUtilities.IsTrue(stdTask.UnitsRequiredSpecified))
                requestXml = requestXml + "                 <unitsRequired>" + stdTask.UnitsRequired + "</unitsRequired>";
            if (MyUtilities.IsTrue(stdTask.UnitsPerDaySpecified))
                requestXml = requestXml + "                 <unitsPerDay>" + stdTask.UnitsPerDay + "</unitsPerDay>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>mse693</application>";
            requestXml = requestXml + "	<applicationPage>read</applicationPage>";
            requestXml = requestXml + "	<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");


            var responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
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

        public static void DeleteStandardJobTaskPost(EllipseFunctions ef, StandardJobTask stdTask)
        {
            ef.InitiatePostConnection();
            if (!string.IsNullOrWhiteSpace(stdTask.SjTaskNo))
                stdTask.SjTaskNo = stdTask.SjTaskNo.PadLeft(3, '0');

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "    <actions>";
            requestXml = requestXml + "        <action>";
            requestXml = requestXml + "            <name>service</name>";
            requestXml = requestXml + "            <data>";
            requestXml = requestXml + "                <name>com.mincom.enterpriseservice.ellipse.standardjobtask.StandardJobTaskService</name>";
            requestXml = requestXml + "                <operation>delete</operation>";
            requestXml = requestXml + "                <className>mfui.actions.tree.node::TreeNodeDeleteAction</className>";
            requestXml = requestXml + "                <returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "                <dto uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "                    <standardJob>" + stdTask.StandardJob + "</standardJob>";
            requestXml = requestXml + "                    <sJTaskNo>" + stdTask.SjTaskNo + "</sJTaskNo>";
            requestXml = requestXml + "                </dto>";
            requestXml = requestXml + "            </data>";
            requestXml = requestXml + "            <id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "        </action>";
            requestXml = requestXml + "    </actions>";
            requestXml = requestXml + "    <chains/>";
            requestXml = requestXml + "    <connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "    <application>mse690</application>";
            requestXml = requestXml + "    <applicationPage>read</applicationPage>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");


            var responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);

        }


        public static void SetStandardJobTaskText(string urlService, string districtCode, string position, bool returnWarnings, StandardJobTask stdTask)
        {
            if (!string.IsNullOrWhiteSpace(stdTask.SjTaskNo))
                stdTask.SjTaskNo = stdTask.SjTaskNo.PadLeft(3, '0');//comentario
            var stdTextId = "JI" + districtCode + stdTask.StandardJob + stdTask.SjTaskNo;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, stdTask.ExtTaskText);
        }

        public static void SetStandardJobText(string urlService, string districtCode, string position, bool returnWarnings, StandardJob stdJob)
        {
            //comentario
            var stdTextId = "SJ" + districtCode + stdJob.StandardJobNo;

            var stdTextCopc = StdText.GetCustomOpContext(districtCode, position, 100, returnWarnings);

            StdText.SetText(urlService, stdTextCopc, stdTextId, stdJob.ExtText);
        }


        public static void CreateTaskResource(string urlService, OperationContext opContext, TaskRequirement taskReq, bool overrideActive = false)
        {
            var proxyTaskReq = new ResourceReqmntsService.ResourceReqmntsService();//ejecuta las acciones del servicio
            var requestTaskReq = new ResourceReqmntsServiceCreateRequestDTO();

            //se cargan los parámetros de la orden
            proxyTaskReq.Url = urlService + "/ResourceReqmntsService";

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

            proxyTaskReq.create(opContext, requestTaskReq);
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

        public static void ModifyTaskResource(string urlService, OperationContext opContext, TaskRequirement taskReq)
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
                SJTaskNo = !string.IsNullOrWhiteSpace(taskReq.SJTaskNo) ? taskReq.SJTaskNo : null ,
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


        public static void DeleteTaskResource(string urlService, OperationContext opContext, TaskRequirement taskReq)
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

        public static class Queries
        {
            public static string GetFetchQuickStandardQuery(string dbReference, string dbLink, string districtCode, string workGroup)
            {
                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(districtCode))
                    districtCode = " AND STD.DSTRCT_CODE IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
                else
                    districtCode = " AND STD.DSTRCT_CODE = '" + districtCode + "'";


                //establecemos los parámetrode de grupo
                if (string.IsNullOrEmpty(workGroup))
                    workGroup = " AND STD.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
                else
                    workGroup = " AND STD.WORK_GROUP = '" + workGroup + "'";



                var query = "" +
                    " SELECT" +
                    "   STD.DSTRCT_CODE, STD.WORK_GROUP, STD.STD_JOB_NO, STD.STD_JOB_DESC, STD.SJ_ACTIVE_STATUS, STD.ORIGINATOR_ID, STD.ORIG_PRIORITY," +
                    "   STD.WO_TYPE, STD.MAINT_TYPE, STD.ASSIGN_PERSON, STD.COMP_CODE, STD.COMP_MOD_CODE, STD.UNIT_OF_WORK, STD.UNITS_REQUIRED," +
                    "   STD.ACCOUNT_CODE, STD.REALL_ACCT_CDE, STD.PROJECT_NO," +
                    "   STD.CALC_DUR_HRS_SW, STD.EST_DUR_HRS, STD.RES_UPDATE_FLAG, STD.EST_LAB_HRS, STD.EST_LAB_COST, STD.MAT_UPDATE_FLAG, STD.EST_MAT_COST, STD.EQUIP_UPDATE_FLAG, STD.EST_EQUIP_COST, STD.EST_OTHER_COST," +
                    "   STD.CALC_LAB_HRS, STD.CALC_LAB_COST, STD.CALC_MAT_COST, STD.CALC_EQUIP_COST," +
                    "   STD.NO_OF_TASKS, 'CONS.RAP.' USO_OTS, 'CONS.RAP.' USO_MSTS, 'CONS.RAP.' ULTIMO_USO," +
                    "   STD.WO_JOB_CODEX1, STD.WO_JOB_CODEX2, STD.WO_JOB_CODEX3, STD.WO_JOB_CODEX4, STD.WO_JOB_CODEX5," +
                    "   STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10" +
                    " FROM" +
                    "   " + dbReference + ".msf690" + dbLink + " STD " +
                    " WHERE" +
                    "" + workGroup +
                    "" + districtCode +
                    " ORDER BY STD.WORK_GROUP, STD.STD_JOB_NO";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFetchStandardQuery(string dbReference, string dbLink, string districtCode, string workGroup)
            {
                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(districtCode))
                    districtCode = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
                else
                    districtCode = " IN ('" + districtCode + "')";


                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(workGroup))
                    workGroup = " IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
                else
                    workGroup = " IN ('" + workGroup + "')";



                var query = "" +
                               " SELECT * FROM (WITH SOT AS (SELECT STD_JOB_NO, MAX(USO_OTS) USO_OTS, MAX(ULTIMO_USO) ULTIMO_USO, MAX(USO_MSTS) USO_MSTS FROM" +
                               "    (SELECT STD_JOB_NO, COUNT(*) USO_OTS, MAX(CREATION_DATE) ULTIMO_USO, 0 AS USO_MSTS FROM " + dbReference + ".MSF620" + dbLink + " WHERE WORK_GROUP " + workGroup + " GROUP BY STD_JOB_NO" +
                               "    UNION ALL SELECT STD_JOB_NO, 0 AS USO_OTS, MAX(LAST_SCH_DATE) AS ULTIMO_USO, COUNT(*) USO_MSTS FROM " + dbReference + ".MSF700" + dbLink + " WHERE WORK_GROUP " + workGroup + " GROUP BY STD_JOB_NO)" +
                               "    GROUP BY STD_JOB_NO)" +
                               "    SELECT" +
                               "    STD.DSTRCT_CODE, STD.WORK_GROUP, STD.STD_JOB_NO, STD.STD_JOB_DESC, STD.SJ_ACTIVE_STATUS, STD.ORIGINATOR_ID, STD.ORIG_PRIORITY," +
                               "    STD.WO_TYPE, STD.MAINT_TYPE, STD.ASSIGN_PERSON, STD.COMP_CODE, STD.COMP_MOD_CODE, STD.UNIT_OF_WORK, STD.UNITS_REQUIRED," +
                               "    STD.ACCOUNT_CODE, STD.REALL_ACCT_CDE, STD.PROJECT_NO," +
                               "    STD.CALC_DUR_HRS_SW, STD.EST_DUR_HRS, STD.RES_UPDATE_FLAG, STD.EST_LAB_HRS, STD.EST_LAB_COST, STD.MAT_UPDATE_FLAG, STD.EST_MAT_COST, STD.EQUIP_UPDATE_FLAG, STD.EST_EQUIP_COST, STD.EST_OTHER_COST," +
                               "    STD.CALC_LAB_HRS, STD.CALC_LAB_COST, STD.CALC_MAT_COST, STD.CALC_EQUIP_COST," +
                               "    STD.NO_OF_TASKS, SOT.USO_OTS, SOT.USO_MSTS, SOT.ULTIMO_USO," +
                               "    STD.WO_JOB_CODEX1, STD.WO_JOB_CODEX2, STD.WO_JOB_CODEX3, STD.WO_JOB_CODEX4, STD.WO_JOB_CODEX5," +
                               "    STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10" +
                               " FROM" +
                               "    " + dbReference + ".msf690" + dbLink + " STD LEFT JOIN SOT ON STD.STD_JOB_NO = SOT.STD_JOB_NO" +
                               " WHERE" +
                               " STD.WORK_GROUP " + workGroup +
                               " AND  STD.DSTRCT_CODE " + districtCode +
                               " ORDER BY STD.WORK_GROUP, STD.STD_JOB_NO)";
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFetchStandardQuery(string dbReference, string dbLink, string districtCode, string workGroup, string standardJob)
            {
                var query = "" +
                               " SELECT * FROM (WITH SOT AS (SELECT STD_JOB_NO, MAX(USO_OTS) USO_OTS, MAX(ULTIMO_USO) ULTIMO_USO, MAX(USO_MSTS) USO_MSTS FROM" +
                               "    (SELECT STD_JOB_NO, COUNT(*) USO_OTS, MAX(CREATION_DATE) ULTIMO_USO, 0 AS USO_MSTS FROM " + dbReference + ".MSF620" + dbLink + " WHERE WORK_GROUP = '" + workGroup + "' GROUP BY STD_JOB_NO" +
                               "    UNION ALL SELECT STD_JOB_NO, 0 AS USO_OTS, MAX(LAST_SCH_DATE) AS ULTIMO_USO, COUNT(*) USO_MSTS FROM " + dbReference + ".MSF700" + dbLink + " WHERE WORK_GROUP = '" + workGroup + "' GROUP BY STD_JOB_NO)" +
                               "    GROUP BY STD_JOB_NO)" +
                               "    SELECT" +
                               "    STD.DSTRCT_CODE, STD.WORK_GROUP, STD.STD_JOB_NO, STD.STD_JOB_DESC, STD.SJ_ACTIVE_STATUS, STD.ORIGINATOR_ID, STD.ORIG_PRIORITY," +
                               "    STD.WO_TYPE, STD.MAINT_TYPE, STD.ASSIGN_PERSON, STD.COMP_CODE, STD.COMP_MOD_CODE, STD.UNIT_OF_WORK, STD.UNITS_REQUIRED," +
                               "    STD.ACCOUNT_CODE, STD.REALL_ACCT_CDE, STD.PROJECT_NO," +
                               "    STD.CALC_DUR_HRS_SW, STD.EST_DUR_HRS, STD.RES_UPDATE_FLAG, STD.EST_LAB_HRS, STD.EST_LAB_COST, STD.MAT_UPDATE_FLAG, STD.EST_MAT_COST, STD.EQUIP_UPDATE_FLAG, STD.EST_EQUIP_COST, STD.EST_OTHER_COST," +
                               "    STD.CALC_LAB_HRS, STD.CALC_LAB_COST, STD.CALC_MAT_COST, STD.CALC_EQUIP_COST," +
                               "    STD.NO_OF_TASKS, SOT.USO_OTS, SOT.USO_MSTS, SOT.ULTIMO_USO," +
                               "    STD.WO_JOB_CODEX1, STD.WO_JOB_CODEX2, STD.WO_JOB_CODEX3, STD.WO_JOB_CODEX4, STD.WO_JOB_CODEX5," +
                               "    STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10" +
                               " FROM" +
                               "    " + dbReference + ".msf690" + dbLink + " STD LEFT JOIN SOT ON STD.STD_JOB_NO = SOT.STD_JOB_NO" +
                               " WHERE" +
                               " STD.WORK_GROUP = '" + workGroup + "'" +
                               " AND STD.DSTRCT_CODE = '" + districtCode + "'" +
                               " AND STD.STD_JOB_NO = '" + standardJob + "'" +
                               " ORDER BY STD.WORK_GROUP, STD.STD_JOB_NO)";
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFetchStandardJobTasksQuery(string dbReference, string dbLink, string districtCode, string workGroup, string standardJob)
            {
                var query = "" +
                    "SELECT" +
                    "    A.DSTRCT_CODE, A.WORK_GROUP, A.STD_JOB_NO, A.STD_JOB_DESC," +
                    "    B.STD_JOB_TASK, B.SJ_TASK_DESC, B.JOB_DESC_CODE, B.SAFETY_INSTR, B.COMPLETE_INSTR, B.COMPL_TEXT_CDE, B.ASSIGN_PERSON, B.EST_MACH_HRS, B.UNIT_OF_WORK, B.UNITS_REQUIRED, B.UNITS_PER_DAY," +
                    "    B.EST_DUR_HRS , B.EQUIP_GRP_ID, B.APL_TYPE, B.COMP_CODE, B.COMP_MOD_CODE, B.APL_SEQ_NO," +
                    "    (SELECT COUNT(*) LABOR FROM " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS ON RS.KEY_735_ID = '" + districtCode + "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK AND RS.REC_735_TYPE = 'ST' INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ON TT.TABLE_CODE = rs.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' WHERE TSK.WORK_GROUP = '" + workGroup + " ' AND TSK.STD_JOB_NO = '" + standardJob + " ' AND TSK.STD_JOB_TASK = B.STD_JOB_TASK) NO_REC_LABOR," +
                    "    (SELECT COUNT(*) MATER FROM " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS ON RS.CLASS_KEY  = '" + districtCode + "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK AND RS.CLASS_TYPE   = 'ST' WHERE TSK.WORK_GROUP = '" + workGroup + " ' AND TSK.STD_JOB_NO = '" + standardJob + " ' AND TSK.STD_JOB_TASK = B.STD_JOB_TASK) NO_REC_MATERIAL" +
                    " FROM" +
                    "    " + dbReference + ".MSF690" + dbLink + " A JOIN " + dbReference + ".MSF693" + dbLink + " B ON A.STD_JOB_NO = B.STD_JOB_NO" +
                    " WHERE A.WORK_GROUP = '" + workGroup + " ' AND A.STD_JOB_NO = '" + standardJob + "'";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }

            public static string GetFetchStdJobTaskRequirementsQuery(string dbReference, string dbLink, string districtCode, string workGroup, string standardJob, string taskNo)
            {
                var query = "" +
                               " SELECT" +
                               " 'LAB' REQ_TYPE, TSK.DSTRCT_CODE, TSK.WORK_GROUP, TSK.STD_JOB_NO, TSK.STD_JOB_TASK, TSK.SJ_TASK_DESC, 'N/A' SEQ_NO, RS.RESOURCE_TYPE RES_CODE, TO_NUMBER(RS.CREW_SIZE) QTY_REQ, RS.EST_RESRCE_HRS HRS_QTY, TT.TABLE_DESC RES_DESC, '' UNITS" +
                               " FROM" +
                               " " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF735" +
                               dbLink + " RS ON RS.KEY_735_ID = '" + districtCode +
                               "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK" +
                               " AND RS.REC_735_TYPE = 'ST' INNER JOIN " + dbReference + ".MSF010" + dbLink +
                               " TT ON TT.TABLE_CODE = RS.RESOURCE_TYPE" +
                               " AND TT.TABLE_TYPE = 'TT'" +
                               " WHERE" +
                               " TRIM(TSK.WORK_GROUP) = '" + workGroup + "' AND TSK.STD_JOB_NO = '" + standardJob +
                               "' AND TSK.STD_JOB_TASK = '" + taskNo + "'" +
                               " UNION ALL" +
                               " SELECT" +
                               " 'MAT' REQ_TYPE, TSK.DSTRCT_CODE, TSK.WORK_GROUP, TSK.STD_JOB_NO, TSK.STD_JOB_TASK, TSK.SJ_TASK_DESC, RS.SEQNCE_NO SEQ_NO, RS.STOCK_CODE RES_CODE, RS.UNIT_QTY_REQD QTY_REQ, 0 HRS_QTY, SCT.DESC_LINEX1||SCT.ITEM_NAME RES_DESC,'' UNITS" +
                               " FROM" +
                               " " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF734" +
                               dbLink + " RS ON RS.CLASS_KEY = '" + districtCode +
                               "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK" +
                               " AND RS.CLASS_TYPE = 'ST' LEFT JOIN " + dbReference + ".MSF100" + dbLink +
                               " SCT ON RS.STOCK_CODE = SCT.STOCK_CODE" +
                               " WHERE" +
                               " TRIM(TSK.WORK_GROUP) = '" + workGroup + "' AND TSK.STD_JOB_NO = '" + standardJob +
                               "' AND TSK.STD_JOB_TASK = '" + taskNo + "'" +
                               " UNION ALL" +
                               " SELECT " +
                               "   'EQU' REQ_TYPE, " +
                               "   TSK.DSTRCT_CODE, " +
                               "   TSK.WORK_GROUP, " +
                               "   TSK.STD_JOB_NO, " +
                               "   TSK.STD_JOB_TASK, " +
                               "   TSK.SJ_TASK_DESC, " +
                               "   RS.SEQNCE_NO SEQ_NO, " +
                               "   RS.EQPT_TYPE RES_CODE, " +
                               "   RS.QTY_REQ, " +
                               "   RS.UNIT_QTY_REQD HRS_QTY, " +
                               "   ET.TABLE_DESC RES_DESC," +
                               "   RS.UOM UNITS " +
                               " FROM " +
                               "   " + dbReference + ".MSF693" + dbLink + " TSK " +
                               " INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS " +
                               " ON " +
                               "   RS.CLASS_KEY = '" + districtCode + "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK " +
                               " AND RS.CLASS_TYPE = 'ST' " +
                               " INNER JOIN ELLIPSE.MSF010 ET " +
                               " ON " +
                               "   RS.EQPT_TYPE = ET.TABLE_CODE " +
                               " WHERE " +
                               "   TSK.WORK_GROUP = '" + workGroup + "' " +
                               " AND TSK.STD_JOB_NO = '" + standardJob + "' " +
                               " AND TSK.STD_JOB_TASK = '" + taskNo + "'" +
                               " AND TABLE_TYPE = 'ET' ";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;

            }

            public static string GetFetchStdJobTaskRequirementsQuery(string dbReference, string dbLink, string districtCode, string workGroup, string standardJob)
            {
                var query = "" +
                               " SELECT" +
                               " 'LAB' REQ_TYPE, TSK.DSTRCT_CODE, TSK.WORK_GROUP, TSK.STD_JOB_NO, TSK.STD_JOB_TASK, TSK.SJ_TASK_DESC, 'N/A' SEQ_NO, RS.RESOURCE_TYPE RES_CODE, TO_NUMBER(RS.CREW_SIZE) QTY_REQ, RS.EST_RESRCE_HRS HRS_QTY, TT.TABLE_DESC RES_DESC, '' UNITS" +
                               " FROM" +
                               " " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF735" +
                               dbLink + " RS ON RS.KEY_735_ID = '" + districtCode +
                               "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK" +
                               " AND RS.REC_735_TYPE = 'ST' INNER JOIN " + dbReference + ".MSF010" + dbLink +
                               " TT ON TT.TABLE_CODE = RS.RESOURCE_TYPE" +
                               " AND TT.TABLE_TYPE = 'TT'" +
                               " WHERE" +
                               " TRIM(TSK.WORK_GROUP) = '" + workGroup + "' AND TSK.STD_JOB_NO = '" + standardJob + "'" +
                               " UNION ALL" +
                               " SELECT" +
                               " 'MAT' REQ_TYPE, TSK.DSTRCT_CODE, TSK.WORK_GROUP, TSK.STD_JOB_NO, TSK.STD_JOB_TASK, TSK.SJ_TASK_DESC, RS.SEQNCE_NO SEQ_NO, RS.STOCK_CODE RES_CODE, RS.UNIT_QTY_REQD QTY_REQ, 0 HRS_QTY, SCT.DESC_LINEX1||SCT.ITEM_NAME RES_DESC,'' UNITS" +
                               " FROM" +
                               " " + dbReference + ".MSF693" + dbLink + " TSK INNER JOIN " + dbReference + ".MSF734" +
                               dbLink + " RS ON RS.CLASS_KEY = '" + districtCode +
                               "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK" +
                               " AND RS.CLASS_TYPE = 'ST' LEFT JOIN " + dbReference + ".MSF100" + dbLink +
                               " SCT ON RS.STOCK_CODE = SCT.STOCK_CODE" +
                               " WHERE" +
                               " TRIM(TSK.WORK_GROUP) = '" + workGroup + "' AND TSK.STD_JOB_NO = '" + standardJob + "'" +
                               " UNION ALL" +
                               " SELECT " +
                               "   'EQU' REQ_TYPE, " +
                               "   TSK.DSTRCT_CODE, " +
                               "   TSK.WORK_GROUP, " +
                               "   TSK.STD_JOB_NO, " +
                               "   TSK.STD_JOB_TASK, " +
                               "   TSK.SJ_TASK_DESC, " +
                               "   RS.SEQNCE_NO SEQ_NO, " +
                               "   RS.EQPT_TYPE RES_CODE, " +
                               "   RS.QTY_REQ, " +
                               "   RS.UNIT_QTY_REQD HRS_QTY, " +
                               "   ET.TABLE_DESC RES_DESC," +
                               "   RS.UOM UNITS " +
                               " FROM " +
                               "   " + dbReference + ".MSF693" + dbLink + " TSK " +
                               " INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS " +
                               " ON " +
                               "   RS.CLASS_KEY = '" + districtCode + "' || TSK.STD_JOB_NO || TSK.STD_JOB_TASK " +
                               " AND RS.CLASS_TYPE = 'ST' " +
                               " INNER JOIN ELLIPSE.MSF010 ET " +
                               " ON " +
                               "   RS.EQPT_TYPE = ET.TABLE_CODE " +
                               " WHERE " +
                               "   TSK.WORK_GROUP = '" + workGroup + "' " +
                               " AND TSK.STD_JOB_NO = '" + standardJob + "' " +
                               " AND TABLE_TYPE = 'ET' ";

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;

            }
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

            newef.CloseConnection();
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
    }
}


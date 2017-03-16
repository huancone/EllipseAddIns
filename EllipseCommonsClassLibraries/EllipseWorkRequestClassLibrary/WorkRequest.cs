using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using EllipseWorkRequestClassLibrary.WorkRequestService;
using EllipseCommonsClassLibrary;

namespace EllipseWorkRequestClassLibrary
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WorkRequest
    {
        public string activityClass;
        public string assignPerson;
        public string classification;
        public string classificationDescription;
        public string closedBy;
        public string closedDate;
        public string closedTime;
        public string contactId;
        public string copyRequestId;
        public string custPOItemNoRef;
        public string custPONoRef;
        public string customerNo;
        public string employee;
        public string equipmentNo;
        public string equipmentRef;
        public string estimateNo;
        public string location;
        public string ownerId;
        public string priorityCode;
        public string priorityCodeDescription;
        public string programCode;
        public string raisedDate;
        public string raisedTime;
        public string raisedUser;
        public string region;
        public string regionDescription;
        public string requestId;
        public string requestIdDescription1;
        public string requestIdDescription2;
        public string requestType;
        public string requestTypeDescription;
        public string requestorId;
        public string requiredByDate;
        public string requiredByTime;
        public string riskCode1;
        public string riskCode10;
        public string riskCode2;
        public string riskCode3;
        public string riskCode4;
        public string riskCode5;
        public string riskCode6;
        public string riskCode7;
        public string riskCode8;
        public string riskCode9;
        public string source;
        public string sourceDescription;
        public string sourceReference;
        public string standardJob;
        public string standardJobDistrict;
        public string userStatus;
        public string userStatusDescription;
        public string workGroup;
        public string status;
        public decimal priorityValue;
        public bool priorityValueFieldSpecified;
        public ServiceLevelAgreement ServiceLevelAgreement;
        public WorkRequestReferenceCodes ReferenceCodes;
        public WorkRequest()
        {
            ServiceLevelAgreement = new ServiceLevelAgreement();
            ReferenceCodes = new WorkRequestReferenceCodes();
        }
    }

    public class ServiceLevelAgreement
    {
        public string ServiceLevel;
        public string FailureCode;
        public string StartDate;
        public string StartTime;
        public string DueDate;
        public string DueDays;
        public string DueTime;
        public string DueHours;
        public string WarnDays;
        public string WarnDate;
        public string WarnTime;
        public string WarnHours;
    }

    public class WorkRequestReferenceCodes
    {
        public string ExtendedDescriptionHeader;
        public string ExtendedDescriptionBody;

        public string StockCode1;//001_9001
        public string StockCode2;//001_9002
        public string StockCode3;//001_9003
        public string StockCode4;//001_9004
        public string StockCode5;//001_9005

        public string StockQuantity1;//001_001
        public string StockQuantity2;//001_002
        public string StockQuantity3;//001_003
        public string StockQuantity4;//001_004
        public string StockQuantity5;//001_005

        public string HorasHombre;//006_9001
        public string HorasQty;//006_001
        public string DuracionTarea;//007_001
        public string EquipoDetenido;//008_001
        public string WorkOrderOrigen;//009_001
        public string RaisedReprogramada;//010_001
        public string CambioHora;//011_001
    }

    public static class WorkRequestActions
    {
        public static List<WorkRequest> FetchWorkRequest(EllipseFunctions ef, string workGroup, string startDate, string endDate, string wrStatus)
        {
            var sqlQuery = Queries.GetFetchWorkRequest(ef.dbReference, ef.dbLink, workGroup, startDate, endDate, wrStatus);
            var drWorkRequest = ef.GetQueryResult(sqlQuery);
            var list = new List<WorkRequest>();

            if (drWorkRequest == null || drWorkRequest.IsClosed || !drWorkRequest.HasRows) return list;
            while (drWorkRequest.Read())
            {
                var request = new WorkRequest
                {
                    workGroup = drWorkRequest["WORK_GROUP"].ToString().Trim(),
                    requestId = drWorkRequest["REQUEST_ID"].ToString().Trim(),
                    status = drWorkRequest["REQUEST_STAT"].ToString().Trim(),
                    requestIdDescription1 = drWorkRequest["SHORT_DESC_1"].ToString().Trim(),
                    requestIdDescription2 = drWorkRequest["SHORT_DESC_2"].ToString().Trim(),
                    equipmentNo = drWorkRequest["EQUIP_NO"].ToString().Trim(),
                    employee = drWorkRequest["EMPLOYEE_ID"].ToString().Trim(),
                    classification = drWorkRequest["WORK_REQ_CLASSIF"].ToString().Trim(),
                    classificationDescription = drWorkRequest["WORK_REQ_CLASSIF_DESC"].ToString().Trim(),
                    requestType = drWorkRequest["WORK_REQ_TYPE"].ToString().Trim(),
                    requestTypeDescription = drWorkRequest["WORK_REQ_TYPE_DESC"].ToString().Trim(),
                    userStatus = drWorkRequest["REQUEST_USTAT"].ToString().Trim(),
                    userStatusDescription = drWorkRequest["REQUEST_USTAT_DESC"].ToString().Trim(),
                    priorityCode = drWorkRequest["PRIORITY_CDE_541"].ToString().Trim(),
                    priorityCodeDescription = drWorkRequest["PRIORITY_CDE_541_DESC"].ToString().Trim(),
                    region = drWorkRequest["REGION"].ToString().Trim(),
                    regionDescription = drWorkRequest["REGION_DESC"].ToString().Trim(),
                    contactId = drWorkRequest["CONTACT_ID"].ToString().Trim(),
                    source = drWorkRequest["WORK_REQ_SOURCE"].ToString().Trim(),
                    sourceDescription = drWorkRequest["WORK_REQ_SOURCE_DESC"].ToString().Trim(),
                    sourceReference = drWorkRequest["SOURCE_REF"].ToString().Trim(),
                    requiredByDate = drWorkRequest["REQUIRED_DATE"].ToString().Trim(),
                    requiredByTime = drWorkRequest["REQUIRED_TIME"].ToString().Trim(),
                    raisedUser = drWorkRequest["CREATION_USER"].ToString().Trim(),
                    raisedDate = drWorkRequest["RAISED_DATE"].ToString().Trim(),
                    raisedTime = drWorkRequest["RAISED_TIME"].ToString().Trim(),
                    closedBy = drWorkRequest["COMPLETED_BY"].ToString().Trim(),
                    closedDate = drWorkRequest["CLOSED_DATE"].ToString().Trim(),
                    closedTime = drWorkRequest["CLOSED_TIME"].ToString().Trim(),
                    assignPerson = drWorkRequest["ASSIGN_PERSON"].ToString().Trim(),
                    ownerId = drWorkRequest["OWNER_ID"].ToString().Trim(),
                    estimateNo = drWorkRequest["ESTIMATE_NO"].ToString().Trim(),
                    standardJob = drWorkRequest["STD_JOB_NO"].ToString().Trim(),
                    standardJobDistrict = drWorkRequest["STD_JOB_DSTRCT"].ToString().Trim(),
                    ServiceLevelAgreement =
                    {
                        ServiceLevel = drWorkRequest["SL_AGREEMENT"].ToString().Trim(),
                        FailureCode = drWorkRequest["SLA_FAILURE_CODE"].ToString().Trim(),
                        StartDate = drWorkRequest["SLA_START_DATE"].ToString().Trim(),
                        StartTime = drWorkRequest["SLA_START_TIME"].ToString().Trim(),
                        DueDate = drWorkRequest["SLA_DUE_DATE"].ToString().Trim(),
                        DueDays = drWorkRequest["SLA_DUE_DAYS"].ToString().Trim(),
                        DueTime = drWorkRequest["SLA_DUE_TIME"].ToString().Trim(),
                        DueHours = drWorkRequest["SLA_DUE_HOURS"].ToString().Trim(),
                        WarnDate = drWorkRequest["SLA_WARN_DATE"].ToString().Trim(),
                        WarnDays = drWorkRequest["SLA_WARN_DAYS"].ToString().Trim(),
                        WarnTime = drWorkRequest["SLA_WARN_TIME"].ToString().Trim(),
                        WarnHours = drWorkRequest["SLA_WARN_HOURS"].ToString().Trim()
                    },
                    ReferenceCodes =
                    {
                        StockCode1 = drWorkRequest["STOCK_CODE1"].ToString().Trim(),
                        StockCode2 = drWorkRequest["STOCK_CODE2"].ToString().Trim(),
                        StockCode3 = drWorkRequest["STOCK_CODE3"].ToString().Trim(),
                        StockCode4 = drWorkRequest["STOCK_CODE4"].ToString().Trim(),
                        StockCode5 = drWorkRequest["STOCK_CODE5"].ToString().Trim(),
                        StockQuantity1 = drWorkRequest["STOCKQTY1"].ToString().Trim(),
                        StockQuantity2 = drWorkRequest["STOCKQTY2"].ToString().Trim(),
                        StockQuantity3 = drWorkRequest["STOCKQTY3"].ToString().Trim(),
                        StockQuantity4 = drWorkRequest["STOCKQTY4"].ToString().Trim(),
                        StockQuantity5 = drWorkRequest["STOCKQTY5"].ToString().Trim(),
                        HorasHombre = drWorkRequest["HORASHOMBRE"].ToString().Trim(),
                        HorasQty = drWorkRequest["HORASHQTY"].ToString().Trim(),
                        DuracionTarea = drWorkRequest["DURACIONTAREA"].ToString().Trim(),
                        EquipoDetenido = drWorkRequest["EQUIPODETENIDO"].ToString().Trim(),
                        WorkOrderOrigen = drWorkRequest["WORKORDERORIGEN"].ToString().Trim(),
                        RaisedReprogramada = drWorkRequest["RAISEDREPROGRAMADA"].ToString().Trim(),
                        CambioHora = drWorkRequest["CAMBIOHORA"].ToString().Trim(),
                        ExtendedDescriptionHeader = drWorkRequest["EXTDESCHEADER"].ToString().Trim(),
                        ExtendedDescriptionBody = drWorkRequest["EXTDESCBODY"].ToString().Trim()
                    }
                };

                list.Add(request);
            }

            return list;
        }

        public static WorkRequest FetchWorkRequest(EllipseFunctions ef, string requestId, bool padNumber = true)
        {
            if (requestId != null && requestId.All(char.IsDigit) && padNumber)
                requestId = requestId.PadLeft(12, '0');

            var sqlQuery = Queries.GetFetchWorkRequest(ef.dbReference, ef.dbLink, requestId);
            var drWorkRequest = ef.GetQueryResult(sqlQuery);
            var request = new WorkRequest();
            if (drWorkRequest == null || drWorkRequest.IsClosed || !drWorkRequest.HasRows) return request;
            drWorkRequest.Read();
            request = new WorkRequest
            {
                workGroup = drWorkRequest["WORK_GROUP"].ToString().Trim(),
                requestId = drWorkRequest["REQUEST_ID"].ToString().Trim(),
                status = drWorkRequest["REQUEST_STAT"].ToString().Trim(),
                requestIdDescription1 = drWorkRequest["SHORT_DESC_1"].ToString().Trim(),
                requestIdDescription2 = drWorkRequest["SHORT_DESC_2"].ToString().Trim(),
                equipmentNo = drWorkRequest["EQUIP_NO"].ToString().Trim(),
                employee = drWorkRequest["EMPLOYEE_ID"].ToString().Trim(),
                classification = drWorkRequest["WORK_REQ_CLASSIF"].ToString().Trim(),
                classificationDescription = drWorkRequest["WORK_REQ_CLASSIF_DESC"].ToString().Trim(),
                requestType = drWorkRequest["WORK_REQ_TYPE"].ToString().Trim(),
                requestTypeDescription = drWorkRequest["WORK_REQ_TYPE_DESC"].ToString().Trim(),
                userStatus = drWorkRequest["REQUEST_USTAT"].ToString().Trim(),
                userStatusDescription = drWorkRequest["REQUEST_USTAT_DESC"].ToString().Trim(),
                priorityCode = drWorkRequest["PRIORITY_CDE_541"].ToString().Trim(),
                priorityCodeDescription = drWorkRequest["PRIORITY_CDE_541_DESC"].ToString().Trim(),
                region = drWorkRequest["REGION"].ToString().Trim(),
                regionDescription = drWorkRequest["REGION_DESC"].ToString().Trim(),
                contactId = drWorkRequest["CONTACT_ID"].ToString().Trim(),
                source = drWorkRequest["WORK_REQ_SOURCE"].ToString().Trim(),
                sourceDescription = drWorkRequest["WORK_REQ_SOURCE_DESC"].ToString().Trim(),
                sourceReference = drWorkRequest["SOURCE_REF"].ToString().Trim(),
                requiredByDate = drWorkRequest["REQUIRED_DATE"].ToString().Trim(),
                requiredByTime = drWorkRequest["REQUIRED_TIME"].ToString().Trim(),
                raisedUser = drWorkRequest["CREATION_USER"].ToString().Trim(),
                raisedDate = drWorkRequest["RAISED_DATE"].ToString().Trim(),
                raisedTime = drWorkRequest["RAISED_TIME"].ToString().Trim(),
                closedBy = drWorkRequest["COMPLETED_BY"].ToString().Trim(),
                closedDate = drWorkRequest["CLOSED_DATE"].ToString().Trim(),
                closedTime = drWorkRequest["CLOSED_TIME"].ToString().Trim(),
                assignPerson = drWorkRequest["ASSIGN_PERSON"].ToString().Trim(),
                ownerId = drWorkRequest["OWNER_ID"].ToString().Trim(),
                estimateNo = drWorkRequest["ESTIMATE_NO"].ToString().Trim(),
                standardJob = drWorkRequest["STD_JOB_NO"].ToString().Trim(),
                standardJobDistrict = drWorkRequest["STD_JOB_DSTRCT"].ToString().Trim(),
                ServiceLevelAgreement =
                {
                    ServiceLevel = drWorkRequest["SL_AGREEMENT"].ToString().Trim(),
                    FailureCode = drWorkRequest["SLA_FAILURE_CODE"].ToString().Trim(),
                    StartDate = drWorkRequest["SLA_START_DATE"].ToString().Trim(),
                    StartTime = drWorkRequest["SLA_START_TIME"].ToString().Trim(),
                    DueDate = drWorkRequest["SLA_DUE_DATE"].ToString().Trim(),
                    DueDays = drWorkRequest["SLA_DUE_DAYS"].ToString().Trim(),
                    DueTime = drWorkRequest["SLA_DUE_TIME"].ToString().Trim(),
                    DueHours = drWorkRequest["SLA_DUE_HOURS"].ToString().Trim(),
                    WarnDate = drWorkRequest["SLA_WARN_DATE"].ToString().Trim(),
                    WarnDays = drWorkRequest["SLA_WARN_DAYS"].ToString().Trim(),
                    WarnTime = drWorkRequest["SLA_WARN_TIME"].ToString().Trim(),
                    WarnHours = drWorkRequest["SLA_WARN_HOURS"].ToString().Trim()
                },
                ReferenceCodes =
                {
                    StockCode1 = drWorkRequest["STOCK_CODE1"].ToString().Trim(),
                    StockCode2 = drWorkRequest["STOCK_CODE2"].ToString().Trim(),
                    StockCode3 = drWorkRequest["STOCK_CODE3"].ToString().Trim(),
                    StockCode4 = drWorkRequest["STOCK_CODE4"].ToString().Trim(),
                    StockCode5 = drWorkRequest["STOCK_CODE5"].ToString().Trim(),
                    StockQuantity1 = drWorkRequest["STOCKQTY1"].ToString().Trim(),
                    StockQuantity2 = drWorkRequest["STOCKQTY2"].ToString().Trim(),
                    StockQuantity3 = drWorkRequest["STOCKQTY3"].ToString().Trim(),
                    StockQuantity4 = drWorkRequest["STOCKQTY4"].ToString().Trim(),
                    StockQuantity5 = drWorkRequest["STOCKQTY5"].ToString().Trim(),
                    HorasHombre = drWorkRequest["HORASHOMBRE"].ToString().Trim(),
                    HorasQty = drWorkRequest["HORASHQTY"].ToString().Trim(),
                    DuracionTarea = drWorkRequest["DURACIONTAREA"].ToString().Trim(),
                    EquipoDetenido = drWorkRequest["EQUIPODETENIDO"].ToString().Trim(),
                    WorkOrderOrigen = drWorkRequest["WORKORDERORIGEN"].ToString().Trim(),
                    RaisedReprogramada = drWorkRequest["RAISEDREPROGRAMADA"].ToString().Trim(),
                    CambioHora = drWorkRequest["CAMBIOHORA"].ToString().Trim(),
                    ExtendedDescriptionHeader = drWorkRequest["EXTDESCHEADER"].ToString().Trim(),
                    ExtendedDescriptionBody = drWorkRequest["EXTDESCBODY"].ToString().Trim()
                }
            };

            return request;
        }

        /// <summary>
        /// Crea un nuevo WorkRequest a partir de los datos ingresados en la clase WorkRequest wr
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="wr">WorkRequest: WorkRequest</param>
        public static WorkRequestServiceCreateReplyDTO CreateWorkRequest(string urlService, OperationContext opContext, WorkRequest wr)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceCreateRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.activityClass = wr.activityClass ?? requestWr.activityClass;
            requestWr.assignPerson= wr.assignPerson ?? requestWr.assignPerson;
            requestWr.classification= wr.classification ?? requestWr.classification;
            requestWr.contactId= wr.contactId ?? requestWr.contactId;
            requestWr.copyRequestId= wr.copyRequestId ?? requestWr.copyRequestId;
            requestWr.custPOItemNoRef= wr.custPOItemNoRef ?? requestWr.custPOItemNoRef;
            requestWr.custPONoRef= wr.custPONoRef ?? requestWr.custPONoRef;
            requestWr.customerNo= wr.customerNo ?? requestWr.customerNo;
            requestWr.employee= wr.employee ?? requestWr.employee;
            requestWr.equipmentNo= wr.equipmentNo ?? requestWr.equipmentNo;
            requestWr.equipmentRef= wr.equipmentRef ?? requestWr.equipmentRef;
            requestWr.estimateNo= wr.estimateNo ?? requestWr.estimateNo;
            requestWr.location= wr.location ?? requestWr.location;
            requestWr.ownerId= wr.ownerId ?? requestWr.ownerId;
            requestWr.priorityCode= wr.priorityCode ?? requestWr.priorityCode;
            requestWr.programCode= wr.programCode ?? requestWr.programCode;
            requestWr.raisedDate= wr.raisedDate ?? requestWr.raisedDate;
            requestWr.raisedTime= wr.raisedTime ?? requestWr.raisedTime;
            requestWr.raisedUser= wr.raisedUser ?? requestWr.raisedUser;
            requestWr.region= wr.region ?? requestWr.region;
            requestWr.requestId= wr.requestId ?? requestWr.requestId;
            requestWr.requestIdDescription1= wr.requestIdDescription1 ?? requestWr.requestIdDescription1;
            requestWr.requestIdDescription2= wr.requestIdDescription2 ?? requestWr.requestIdDescription2;
            requestWr.requestType= wr.requestType ?? requestWr.requestType;
            requestWr.requestorId= wr.requestorId ?? requestWr.requestorId;
            requestWr.requiredByDate= wr.requiredByDate ?? requestWr.requiredByDate;
            requestWr.requiredByTime= wr.requiredByTime ?? requestWr.requiredByTime;
            requestWr.riskCode1= wr.riskCode1 ?? requestWr.riskCode1;
            requestWr.riskCode10= wr.riskCode10 ?? requestWr.riskCode10;
            requestWr.riskCode2= wr.riskCode2 ?? requestWr.riskCode2;
            requestWr.riskCode3= wr.riskCode3 ?? requestWr.riskCode3;
            requestWr.riskCode4= wr.riskCode4 ?? requestWr.riskCode4;
            requestWr.riskCode5= wr.riskCode5 ?? requestWr.riskCode5;
            requestWr.riskCode6= wr.riskCode6 ?? requestWr.riskCode6;
            requestWr.riskCode7= wr.riskCode7 ?? requestWr.riskCode7;
            requestWr.riskCode8= wr.riskCode8 ?? requestWr.riskCode8;
            requestWr.riskCode9= wr.riskCode9 ?? requestWr.riskCode9;
            requestWr.source= wr.source ?? requestWr.source;
            requestWr.sourceReference= wr.sourceReference ?? requestWr.sourceReference;
            requestWr.standardJob= wr.standardJob ?? requestWr.standardJob;
            requestWr.standardJobDistrict= wr.standardJobDistrict ?? requestWr.standardJobDistrict;
            requestWr.userStatus= wr.userStatus ?? requestWr.userStatus;
            requestWr.workGroup= wr.workGroup ?? requestWr.workGroup;
            //se envía la acción
            return proxyWr.create(opContext, requestWr);
        }

        /// <summary>
        /// Modifica un WorkRequest a partir de los datos ingresados en la clase WorkRequest wr
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="wr">WorkRequest: WorkRequest</param>
        public static WorkRequestServiceModifyReplyDTO ModifyWorkRequest(string urlService, OperationContext opContext, WorkRequest wr)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceModifyRequestDTO();
            
            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.activityClass = wr.activityClass ?? requestWr.activityClass;
            requestWr.assignPerson = wr.assignPerson ?? requestWr.assignPerson;
            requestWr.classification = wr.classification ?? requestWr.classification;
            requestWr.contactId = wr.contactId ?? requestWr.contactId;
            requestWr.custPOItemNoRef = wr.custPOItemNoRef ?? requestWr.custPOItemNoRef;
            requestWr.custPONoRef = wr.custPONoRef ?? requestWr.custPONoRef;
            requestWr.customerNo = wr.customerNo ?? requestWr.customerNo;
            requestWr.employee = wr.employee ?? requestWr.employee;
            requestWr.equipmentNo = wr.equipmentNo ?? requestWr.equipmentNo;
            requestWr.equipmentRef = wr.equipmentRef ?? requestWr.equipmentRef;
            requestWr.estimateNo = wr.estimateNo ?? requestWr.estimateNo;
            requestWr.location = wr.location ?? requestWr.location;
            requestWr.ownerId = wr.ownerId ?? requestWr.ownerId;
            requestWr.priorityCode = wr.priorityCode ?? requestWr.priorityCode;
            requestWr.programCode = wr.programCode ?? requestWr.programCode;
            requestWr.region = wr.region ?? requestWr.region;
            requestWr.requestId = wr.requestId ?? requestWr.requestId;
            requestWr.requestIdDescription1 = wr.requestIdDescription1 ?? requestWr.requestIdDescription1;
            requestWr.requestIdDescription2 = wr.requestIdDescription2 ?? requestWr.requestIdDescription2;
            requestWr.requestType = wr.requestType ?? requestWr.requestType;
            requestWr.requestorId = wr.requestorId ?? requestWr.requestorId;
            requestWr.requiredByDate = wr.requiredByDate ?? requestWr.requiredByDate;
            requestWr.requiredByTime = wr.requiredByTime ?? requestWr.requiredByTime;
            requestWr.riskCode1 = wr.riskCode1 ?? requestWr.riskCode1;
            requestWr.riskCode10 = wr.riskCode10 ?? requestWr.riskCode10;
            requestWr.riskCode2 = wr.riskCode2 ?? requestWr.riskCode2;
            requestWr.riskCode3 = wr.riskCode3 ?? requestWr.riskCode3;
            requestWr.riskCode4 = wr.riskCode4 ?? requestWr.riskCode4;
            requestWr.riskCode5 = wr.riskCode5 ?? requestWr.riskCode5;
            requestWr.riskCode6 = wr.riskCode6 ?? requestWr.riskCode6;
            requestWr.riskCode7 = wr.riskCode7 ?? requestWr.riskCode7;
            requestWr.riskCode8 = wr.riskCode8 ?? requestWr.riskCode8;
            requestWr.riskCode9 = wr.riskCode9 ?? requestWr.riskCode9;
            requestWr.source = wr.source ?? requestWr.source;
            requestWr.sourceReference = wr.sourceReference ?? requestWr.sourceReference;
            requestWr.standardJob = wr.standardJob ?? requestWr.standardJob;
            requestWr.standardJobDistrict = wr.standardJobDistrict ?? requestWr.standardJobDistrict;
            requestWr.userStatus = wr.userStatus ?? requestWr.userStatus;
            requestWr.workGroup = wr.workGroup ?? requestWr.workGroup;
            //se envía la acción
            return proxyWr.modify(opContext, requestWr);
        }

        /// <summary>
        /// Elimina un WorkRequest a partir de un id dado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="workRequestId">string: workRequestId a eliminar</param>
        public static WorkRequestServiceDeleteReplyDTO DeleteWorkRequest(string urlService, OperationContext opContext, string workRequestId)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceDeleteRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = workRequestId;
            //se envía la acción
            return proxyWr.delete(opContext, requestWr);
        }

        /// <summary>
        /// Cierra un WorkRequest a partir de un id dado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="workRequestId">string: workRequestId a cerrar</param>
        /// <param name="closedBy">string: nombre de usuario que cierra el Work Request</param>
        /// <param name="closedDate">string: fecha en formato yyyymmdd de cierre del Work Request</param>
        /// <param name="closedTime">string: hora en format hhmmss de cierre del Work Request</param>
        public static WorkRequestServiceCloseReplyDTO CloseWorkRequest(string urlService, OperationContext opContext, string workRequestId, string closedBy, string closedDate, string closedTime = null)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceCloseRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = workRequestId;
            requestWr.closedBy = closedBy;
            requestWr.closedDate = closedDate;
            requestWr.closedTime = closedTime;
            //se envía la acción
            return proxyWr.close(opContext, requestWr);
        }
        /// <summary>
        /// Cierra un WorkRequest a partir de un id dado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="workRequestId">string: workRequestId a cerrar</param>
        public static WorkRequestServiceReopenReplyDTO ReOpenWorkRequest(string urlService, OperationContext opContext, string workRequestId)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceReopenRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = workRequestId;
            //se envía la acción
            return proxyWr.reopen(opContext, requestWr);
        }
        /// <summary>
        /// Establece el Service Level Agreement de un Work Request a partir del SLA ingresado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="workRequestId">string: workRequestId a eliminar</param>
        /// <param name="sla">ServiceLevelAgreement : SLA a establecer</param>
        public static WorkRequestServiceSetSLAReplyDTO SetWorkRequestSla(string urlService, OperationContext opContext, string workRequestId, ServiceLevelAgreement sla)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceSetSLARequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = workRequestId;
            requestWr.SLA = sla.ServiceLevel;
            requestWr.SLAFailureCode = sla.FailureCode;
            requestWr.SLAStartDate = sla.StartDate;
            requestWr.SLAStartTime = !string.IsNullOrWhiteSpace(sla.StartTime) ? sla.StartTime : null;
            requestWr.SLADueDays = !string.IsNullOrWhiteSpace(sla.DueDays) ? Convert.ToDecimal(sla.DueDays) : requestWr.SLADueDays;
            requestWr.SLADueHours= !string.IsNullOrWhiteSpace(sla.DueHours) ? Convert.ToDecimal(sla.DueHours) : requestWr.SLADueHours;
            requestWr.SLAWarnDays = !string.IsNullOrWhiteSpace(sla.WarnDays) ? Convert.ToDecimal(sla.WarnDays) : requestWr.SLAWarnDays;
            requestWr.SLAWarnHours = !string.IsNullOrWhiteSpace(sla.WarnHours) ? Convert.ToDecimal(sla.WarnHours) : requestWr.SLAWarnHours;

            requestWr.SLADueDaysSpecified = !string.IsNullOrWhiteSpace(sla.DueDays);
            requestWr.SLADueHoursSpecified = !string.IsNullOrWhiteSpace(sla.DueHours);
            //se envía la acción
            return proxyWr.setSLA(opContext, requestWr);
        }
        /// <summary>
        /// Establece el Service Level Agreement de un Work Request a partir del SLA ingresado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="workRequestId">string: workRequestId a eliminar</param>
        /// <param name="sla">ServiceLevelAgreement : SLA a establecer</param>
        public static WorkRequestServiceResetSLAReplyDTO ResetWorkRequestSla(string urlService, OperationContext opContext, string workRequestId, ServiceLevelAgreement sla)
        {
            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceResetSLARequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = workRequestId;
            requestWr.SLA = sla.ServiceLevel;
            requestWr.SLAStartDate = sla.StartDate;
            requestWr.SLAStartTime = !string.IsNullOrWhiteSpace(sla.StartTime) ? sla.StartTime : null;
            //se envía la acción
            return proxyWr.resetSLA(opContext, requestWr);
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

        public static class Queries
        {
            public static string GetFetchWorkRequest(string dbReference, string dbLink, string workGroup, string startDate, string endDate, string wrStatus)
            {
                //establecemos los parámetrode de grupo
                if (string.IsNullOrEmpty(workGroup))
                    workGroup = " IN (" + Utils.GetListInSeparator(GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
                else
                    workGroup = " = '" + workGroup + "'";

                //establecemos los parámetros de estado de orden
                string statusRequirement;
                if (string.IsNullOrEmpty(wrStatus))
                    statusRequirement = "";
                else if (wrStatus == WrStatusList.Uncompleted)
                    statusRequirement = " AND WR.REQUEST_STAT IN (" + Utils.GetListInSeparator(WrStatusList.GetUncompletedStatusCodes(), ",", "'") + ")";
                else if (WrStatusList.GetStatusNames().Contains(wrStatus))
                    statusRequirement = " AND WR.REQUEST_STAT = '" + WrStatusList.GetStatusCode(wrStatus) + "'";
                else
                    statusRequirement = "";

                //establecemos los parámetros para el rango de fechas
                if (string.IsNullOrEmpty(startDate))
                    startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                if (string.IsNullOrEmpty(endDate))
                    endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                //escribimos el query
                var query = "" +
                   " SELECT " +
                   "   WR.WORK_GROUP," +
                   "   WR.REQUEST_ID," +
                   "   WR.REQUEST_STAT," +
                   "   WR.SHORT_DESC_1," +
                   "   WR.SHORT_DESC_2," +
                   "   WR.EQUIP_NO,  " +
                   "   WR.EMPLOYEE_ID," +
                   "   WR.WORK_REQ_CLASSIF," +
                   "   RQCL.TABLE_DESC WORK_REQ_CLASSIF_DESC," +
                   "   WR.WORK_REQ_TYPE," +
                   "   RQWO.TABLE_DESC WORK_REQ_TYPE_DESC," +
                   "   WR.REQUEST_USTAT," +
                   "   RQWS.TABLE_DESC REQUEST_USTAT_DESC," +
                   "   WR.PRIORITY_CDE_541," +
                   "   RQPY.TABLE_DESC PRIORITY_CDE_541_DESC," +
                   "   WR.REGION," +
                   "   REGN.TABLE_DESC REGION_DESC," +
                   "   WR.CONTACT_ID," +
                   "   WR.WORK_REQ_SOURCE," +
                   "   RQSC.TABLE_DESC WORK_REQ_SOURCE_DESC," +
                   "   WR.SOURCE_REF," +
                   "   WR.REQUIRED_DATE," +
                   "   WR.REQUIRED_TIME," +
                   "   WR.CREATION_USER," +
                   "   WR.RAISED_DATE," +
                   "   WR.RAISED_TIME," +
                   "   WR.COMPLETED_BY," +
                   "   WR.CLOSED_DATE," +
                   "   WR.CLOSED_TIME," +
                   "   WR.ASSIGN_PERSON," +
                   "   WR.OWNER_ID," +
                   "   WR.ESTIMATE_NO," +
                   "   WR.STD_JOB_NO," +
                   "   WR.STD_JOB_DSTRCT," +
                   "   WR.SL_AGREEMENT," +
                   "   WR.SLA_FAILURE_CODE," +
                   "   WR.SLA_START_DATE," +
                   "   WR.SLA_START_TIME," +
                   "   WR.SLA_DUE_DATE," +
                   "   SUBSTR(WR.SLA_DUE_DAYS, 0, LENGTH(WR.SLA_DUE_DAYS)-1)||SUBSTR(RAWTOHEX(WR.SLA_DUE_DAYS),-1) SLA_DUE_DAYS," +
                   "   WR.SLA_DUE_TIME," +
                   "   WR.SLA_DUE_HOURS," +
                   "   WR.SLA_WARN_DATE," +
                   "   SUBSTR(WR.SLA_WARN_DAYS, 0, LENGTH(WR.SLA_WARN_DAYS)-1)||SUBSTR(RAWTOHEX(WR.SLA_WARN_DAYS),-1) SLA_WARN_DAYS," +
                   "   WR.SLA_WARN_TIME," +
                   "   WR.SLA_WARN_HOURS," +
                   "   (SELECT REPLACE(TRIM(STD_VOLAT_1), '.HEADING ', '') FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ' AND STD_KEY = WR.REQUEST_ID AND STD_LINE_NO = '0000') EXTDESCHEADER, " +
                   "   REPLACE((SELECT LISTAGG(TEXTO,' ') WITHIN GROUP (ORDER BY STD_LINE_NO) FROM (SELECT STD_KEY, STD_LINE_NO,TRIM(STD_VOLAT_1)||' '||TRIM(STD_VOLAT_2)||' '||TRIM(STD_VOLAT_3)||' '||TRIM(STD_VOLAT_4)||' '||TRIM(STD_VOLAT_5) AS TEXTO " +
                   "     FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ') WHERE STD_KEY = WR.REQUEST_ID GROUP BY STD_KEY),                                                                                                                               " +
                   "       '.HEADING '||(SELECT REPLACE(TRIM(STD_VOLAT_1), '.HEADING ', '') FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ' AND STD_KEY   = WR.REQUEST_ID AND STD_LINE_NO     = '0000')||' ','') EXTDESCBODY," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '001') STOCK_CODE1," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '002') STOCK_CODE2," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '003') STOCK_CODE3," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '004') STOCK_CODE4," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '005') STOCK_CODE5," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '001' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY1," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '002' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY2," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '003' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY3," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '004' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY4," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '005' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY5, " +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '006' AND SEQ_NUM = '001') HORASHOMBRE," +
                   "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '006' AND SEQ_NUM = '001' AND ENTITY_VALUE  = WR.REQUEST_ID)) HORASHQTY," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '007' AND SEQ_NUM = '001') DURACIONTAREA," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '008' AND SEQ_NUM = '001') EQUIPODETENIDO," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '009' AND SEQ_NUM = '001') WORKORDERORIGEN," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '010' AND SEQ_NUM = '001') RAISEDREPROGRAMADA," +
                   "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '011' AND SEQ_NUM = '001') CAMBIOHORA" +
                   " FROM" +
                   "   " + dbReference + ".MSF541" + dbLink + " WR" +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " RQCL ON WR.WORK_REQ_CLASSIF = RQCL.TABLE_CODE AND RQCL.TABLE_TYPE = 'RQCL' " +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " RQWO ON WR.WORK_REQ_TYPE = RQWO.TABLE_CODE AND RQWO.TABLE_TYPE = 'WO' " +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " RQWS ON WR.REQUEST_USTAT = RQWS.TABLE_CODE AND RQWS.TABLE_TYPE = 'WS' " +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " RQPY ON WR.PRIORITY_CDE_541 = RQPY.TABLE_CODE AND RQPY.TABLE_TYPE = 'PY' " +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " REGN ON WR.REGION = REGN.TABLE_CODE AND REGN.TABLE_TYPE = 'REGN' " +
                   "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " RQSC ON WR.WORK_REQ_SOURCE = RQSC.TABLE_CODE AND RQSC.TABLE_TYPE = 'RQSC' " +
                   " WHERE" +
                   "   WR.WORK_GROUP " + workGroup +
                   "" + statusRequirement +
                   "   AND WR.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";

                return query;
            }
            public static string GetFetchWorkRequest(string dbReference, string dbLink, string requestId)
            {
                //escribimos el query
                var query = "" +
                            " SELECT " +
                            "   WR.WORK_GROUP," +
                            "   WR.REQUEST_ID," +
                            "   WR.REQUEST_STAT," +
                            "   WR.SHORT_DESC_1," +
                            "   WR.SHORT_DESC_2," +
                            "   WR.EQUIP_NO,  " +
                            "   WR.EMPLOYEE_ID," +
                            "   WR.WORK_REQ_CLASSIF," +
                            "   RQCL.TABLE_DESC WORK_REQ_CLASSIF_DESC," +
                            "   WR.WORK_REQ_TYPE," +
                            "   RQWO.TABLE_DESC WORK_REQ_TYPE_DESC," +
                            "   WR.REQUEST_USTAT," +
                            "   RQWS.TABLE_DESC REQUEST_USTAT_DESC," +
                            "   WR.PRIORITY_CDE_541," +
                            "   RQPY.TABLE_DESC PRIORITY_CDE_541_DESC," +
                            "   WR.REGION," +
                            "   REGN.TABLE_DESC REGION_DESC," +
                            "   WR.CONTACT_ID," +
                            "   WR.WORK_REQ_SOURCE," +
                            "   RQSC.TABLE_DESC WORK_REQ_SOURCE_DESC," +
                            "   WR.SOURCE_REF," +
                            "   WR.REQUIRED_DATE," +
                            "   WR.REQUIRED_TIME," +
                            "   WR.CREATION_USER," +
                            "   WR.RAISED_DATE," +
                            "   WR.RAISED_TIME," +
                            "   WR.COMPLETED_BY," +
                            "   WR.CLOSED_DATE," +
                            "   WR.CLOSED_TIME," +
                            "   WR.ASSIGN_PERSON," +
                            "   WR.OWNER_ID," +
                            "   WR.ESTIMATE_NO," +
                            "   WR.STD_JOB_NO," +
                            "   WR.STD_JOB_DSTRCT," +
                            "   WR.SL_AGREEMENT," +
                            "   WR.SLA_FAILURE_CODE," +
                            "   WR.SLA_START_DATE," +
                            "   WR.SLA_START_TIME," +
                            "   WR.SLA_DUE_DATE," +
                            "   SUBSTR(WR.SLA_DUE_DAYS, 0, LENGTH(WR.SLA_DUE_DAYS)-1)||SUBSTR(RAWTOHEX(WR.SLA_DUE_DAYS),-1) SLA_DUE_DAYS," +
                            "   WR.SLA_DUE_TIME," +
                            "   WR.SLA_DUE_HOURS," +
                            "   WR.SLA_WARN_DATE," +
                            "   SUBSTR(WR.SLA_WARN_DAYS, 0, LENGTH(WR.SLA_WARN_DAYS)-1)||SUBSTR(RAWTOHEX(WR.SLA_WARN_DAYS),-1) SLA_WARN_DAYS," +
                            "   WR.SLA_WARN_TIME," +
                            "   WR.SLA_WARN_HOURS," +
                            "   (SELECT REPLACE(TRIM(STD_VOLAT_1), '.HEADING ', '') FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ' AND STD_KEY = WR.REQUEST_ID AND STD_LINE_NO = '0000') EXTDESCHEADER, " +
                            "   REPLACE((SELECT LISTAGG(TEXTO,' ') WITHIN GROUP (ORDER BY STD_LINE_NO) FROM (SELECT STD_KEY, STD_LINE_NO,TRIM(STD_VOLAT_1)||' '||TRIM(STD_VOLAT_2)||' '||TRIM(STD_VOLAT_3)||' '||TRIM(STD_VOLAT_4)||' '||TRIM(STD_VOLAT_5) AS TEXTO " +
                            "     FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ') WHERE STD_KEY = WR.REQUEST_ID GROUP BY STD_KEY)," +
                            "       '.HEADING '||(SELECT REPLACE(TRIM(STD_VOLAT_1), '.HEADING ', '') FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'WQ' AND STD_KEY   = WR.REQUEST_ID AND STD_LINE_NO     = '0000')||' ','') EXTDESCBODY," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '001') STOCK_CODE1," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '002') STOCK_CODE2," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '003') STOCK_CODE3," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '004') STOCK_CODE4," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '001' AND SEQ_NUM = '005') STOCK_CODE5," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '001' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY1," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '002' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY2," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '003' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY3," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '004' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY4," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '001' AND SEQ_NUM = '005' AND ENTITY_VALUE  = WR.REQUEST_ID)) STOCKQTY5," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '006' AND SEQ_NUM = '001') HORASHOMBRE," +
                            "   (SELECT TRIM(STD_VOLAT_1||STD_VOLAT_2||STD_VOLAT_3||STD_VOLAT_4||STD_VOLAT_5) FROM ELLIPSE.MSF096_STD_VOLAT WHERE STD_TEXT_CODE = 'RC' AND STD_KEY = (SELECT STD_TXT_KEY FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND REF_NO = '006' AND SEQ_NUM = '001' AND ENTITY_VALUE  = WR.REQUEST_ID)) HORASHQTY," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '007' AND SEQ_NUM = '001') DURACIONTAREA," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '008' AND SEQ_NUM = '001') EQUIPODETENIDO," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '009' AND SEQ_NUM = '001') WORKORDERORIGEN," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '010' AND SEQ_NUM = '001') RAISEDREPROGRAMADA," +
                            "   (SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE ENTITY_TYPE = 'WRQ' AND ENTITY_VALUE = WR.REQUEST_ID AND REF_NO = '011' AND SEQ_NUM = '001') CAMBIOHORA" +
                            " FROM" +
                            "   " + dbReference + ".MSF541" + dbLink + " WR" +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " RQCL ON WR.WORK_REQ_CLASSIF = RQCL.TABLE_CODE AND RQCL.TABLE_TYPE = 'RQCL' " +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " RQWO ON WR.WORK_REQ_TYPE = RQWO.TABLE_CODE AND RQWO.TABLE_TYPE = 'WO' " +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " RQWS ON WR.REQUEST_USTAT = RQWS.TABLE_CODE AND RQWS.TABLE_TYPE = 'WS' " +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " RQPY ON WR.PRIORITY_CDE_541 = RQPY.TABLE_CODE AND RQPY.TABLE_TYPE = 'PY' " +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " REGN ON WR.REGION = REGN.TABLE_CODE AND REGN.TABLE_TYPE = 'REGN' " +
                            "     LEFT JOIN " + dbReference + ".MSF010" + dbLink +
                            " RQSC ON WR.WORK_REQ_SOURCE = RQSC.TABLE_CODE AND RQSC.TABLE_TYPE = 'RQSC' " +
                            " WHERE" +
                            "   WR.REQUEST_ID = '" + requestId + "'";
                return query;
            }
        }
    }

    public static class WrStatusList
    {
        public static string Open = "OPEN";
        public static string OpenCode = "O";
        public static string Closed = "CLOSED";
        public static string ClosedCode = "C";
        public static string Cancelled = "CANCELLED";
        public static string CancelledCode = "L";
        public static string InWork = "IN_WORK";
        public static string InWorkCode = "W";
        public static string Estimated = "ESTIMATED";
        public static string EstimatedCode = "E";

        public static string Uncompleted = "UNCOMPLETED";

        public static string GetStatusCode(string statusName)
        {
            if (statusName == Open)
                return OpenCode;
            if (statusName == Closed)
                return ClosedCode;
            if (statusName == Cancelled)
                return CancelledCode;
            if (statusName == InWork)
                return InWorkCode;
            if (statusName == Estimated)
                return EstimatedCode;
            return null;
        }

        public static string GetStatusName(string statusCode)
        {
            if (statusCode == OpenCode)
                return Open;
            if (statusCode == ClosedCode)
                return Closed;
            if (statusCode == CancelledCode)
                return Cancelled;
            if (statusCode == InWorkCode)
                return InWork;
            if (statusCode == EstimatedCode)
                return Estimated;
            return null;
        }

        public static List<string> GetStatusNames()
        {
            var list = new List<string> { Open, Closed, Cancelled, InWork, Estimated };
            return list;
        }
        public static List<string> GetStatusCodes()
        {
            var list = new List<string> { OpenCode, ClosedCode, CancelledCode, InWorkCode, EstimatedCode };
            return list;
        }
        public static List<string> GetUncompletedStatusNames()
        {
            var list = new List<string> { Open, InWork, Estimated };
            return list;
        }
        public static List<string> GetUncompletedStatusCodes()
        {
            var list = new List<string> { OpenCode, InWorkCode, EstimatedCode };
            return list;
        }


    }
}

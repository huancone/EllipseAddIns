using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using EllipseWorkRequestClassLibrary.WorkRequestService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;
using EllipseReferenceCodesClassLibrary;
using EllipseStdTextClassLibrary;

namespace EllipseWorkRequestClassLibrary
{

    public static class WorkRequestActions
    {
        public static List<WorkRequest> FetchWorkRequest(EllipseFunctions ef, int searchCriteria1Key, string searchCriteria1Value, int searchCriteria2Key, string searchCriteria2Value, int dateCriteriaKey, string startDate, string endDate, string wrStatus)
        {
            var sqlQuery = Queries.GetFetchWorkRequest(ef.dbReference, ef.dbLink, searchCriteria1Key,
                searchCriteria1Value, searchCriteria2Key, searchCriteria2Value, dateCriteriaKey, startDate, endDate, wrStatus);
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
                };
                
                list.Add(request);
            }

            return list;
        }

        public static WorkRequest FetchWorkRequest(EllipseFunctions ef, string requestId, bool padNumber = true)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
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
            };

            return request;
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public static WorkRequestReferenceCodes GetWorkRequestReferenceCodes(EllipseFunctions eFunctions, string urlService, OperationContext opContext, string requestId)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var wrRefCodes = new WorkRequestReferenceCodes();

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            const string entityType = "WRQ";
            var entityValue = requestId;

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
            var item001_01 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "001");
            var item001_02 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "002");
            var item001_03 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "003");
            var item001_04 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "004");
            var item001_05 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "005");
            var item006 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "006", "001");
            var item007 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "007", "001");
            var item008 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "008", "001");
            var item009 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "009", "001");
            var item010 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "010", "001");
            var item011 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "011", "001");


            wrRefCodes.StockCode1 = item001_01.RefCode; //001_9001
            wrRefCodes.StockCode1Qty = item001_01.StdText; //001_9001
            wrRefCodes.StockCode2 = item001_02.RefCode; //001_9002
            wrRefCodes.StockCode2Qty = item001_02.StdText; //001_9002
            wrRefCodes.StockCode3 = item001_03.RefCode; //001_9003
            wrRefCodes.StockCode3Qty = item001_03.StdText; //001_9003
            wrRefCodes.StockCode4 = item001_04.RefCode; //001_9004
            wrRefCodes.StockCode4Qty = item001_04.StdText; //001_9004
            wrRefCodes.StockCode5 = item001_05.RefCode; //001_9005
            wrRefCodes.StockCode5Qty = item001_05.StdText; //001_9005
            wrRefCodes.HorasHombre = item006.RefCode; //006_9001
            wrRefCodes.HorasQty = item006.StdText; //006_9001
            wrRefCodes.DuracionTarea = item007.RefCode; //007_001
            wrRefCodes.EquipoDetenido = item008.RefCode; //008_001
            wrRefCodes.RaisedReprogramada = item009.RefCode; //009_001
            wrRefCodes.WorkOrderOrigen = item010.RefCode; //010_001
            wrRefCodes.CambioHora = item011.RefCode; //011_001

            newef.CloseConnection();
            return wrRefCodes;
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
            requestWr.assignPerson = wr.assignPerson ?? requestWr.assignPerson;
            requestWr.classification = wr.classification ?? requestWr.classification;
            requestWr.contactId = wr.contactId ?? requestWr.contactId;
            requestWr.copyRequestId = wr.copyRequestId ?? requestWr.copyRequestId;
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
            requestWr.raisedDate = wr.raisedDate ?? requestWr.raisedDate;
            requestWr.raisedTime = wr.raisedTime ?? requestWr.raisedTime;
            requestWr.raisedUser = wr.raisedUser ?? requestWr.raisedUser;
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
        /// <param name="requestId">string: requestId a eliminar</param>
        public static WorkRequestServiceDeleteReplyDTO DeleteWorkRequest(string urlService, OperationContext opContext, string requestId)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');


            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceDeleteRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = requestId;
            //se envía la acción
            return proxyWr.delete(opContext, requestWr);
        }

        /// <summary>
        /// Cierra un WorkRequest a partir de un id dado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="requestId">string: workRequestId a cerrar</param>
        /// <param name="closedBy">string: nombre de usuario que cierra el Work Request</param>
        /// <param name="closedDate">string: fecha en formato yyyymmdd de cierre del Work Request</param>
        /// <param name="closedTime">string: hora en format hhmmss de cierre del Work Request</param>
        public static WorkRequestServiceCloseReplyDTO CloseWorkRequest(string urlService, OperationContext opContext, string requestId, string closedBy, string closedDate, string closedTime = null)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceCloseRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = requestId;
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
        /// <param name="requestId">string: workRequestId a cerrar</param>
        public static WorkRequestServiceReopenReplyDTO ReOpenWorkRequest(string urlService, OperationContext opContext, string requestId)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceReopenRequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = requestId;
            //se envía la acción
            return proxyWr.reopen(opContext, requestWr);
        }
        /// <summary>
        /// Establece el Service Level Agreement de un Work Request a partir del SLA ingresado
        /// </summary>
        /// <param name="urlService">string: URL del servicio web (ej. "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/WorkRequest")</param>
        /// <param name="opContext">WorkRequestService.OperationContext: Contexto de Operación del WorkRequest</param>
        /// <param name="requestId">string: workRequestId a eliminar</param>
        /// <param name="sla">ServiceLevelAgreement : SLA a establecer</param>
        public static WorkRequestServiceSetSLAReplyDTO SetWorkRequestSla(string urlService, OperationContext opContext, string requestId, ServiceLevelAgreement sla)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');


            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceSetSLARequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = requestId;
            requestWr.SLA = sla.ServiceLevel;
            requestWr.SLAFailureCode = sla.FailureCode;
            requestWr.SLAStartDate = sla.StartDate;
            requestWr.SLAStartTime = !string.IsNullOrWhiteSpace(sla.StartTime) ? sla.StartTime : null;
            requestWr.SLADueDays = !string.IsNullOrWhiteSpace(sla.DueDays) ? Convert.ToDecimal(sla.DueDays) : requestWr.SLADueDays;
            requestWr.SLADueHours = !string.IsNullOrWhiteSpace(sla.DueHours) ? Convert.ToDecimal(sla.DueHours) : requestWr.SLADueHours;
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
        /// <param name="requestId">string: workRequestId a eliminar</param>
        /// <param name="sla">ServiceLevelAgreement : SLA a establecer</param>
        public static WorkRequestServiceResetSLAReplyDTO ResetWorkRequestSla(string urlService, OperationContext opContext, string requestId, ServiceLevelAgreement sla)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var proxyWr = new WorkRequestService.WorkRequestService();//ejecuta las acciones del servicio
            var requestWr = new WorkRequestServiceResetSLARequestDTO();

            proxyWr.Url = urlService + "/WorkRequest";

            //se cargan los parámetros de la orden
            requestWr.requestId = requestId;
            requestWr.SLA = sla.ServiceLevel;
            requestWr.SLAStartDate = sla.StartDate;
            requestWr.SLAStartTime = !string.IsNullOrWhiteSpace(sla.StartTime) ? sla.StartTime : null;
            //se envía la acción
            return proxyWr.resetSLA(opContext, requestWr);
        }
        public static ExtendedDescription GetWorkRequestExtendedDescription(string urlService, OperationContext opContext, string requestId)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');
            var description = new ExtendedDescription();
            var stdTextOpContext = StdText.GetStdTextOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);

            var stdTextId = "WQ" + requestId;

            description.Header = StdText.GetHeader(urlService, stdTextOpContext, stdTextId);
            description.Body = StdText.GetText(urlService, stdTextOpContext, stdTextId);
            return description;
        }

        public static ReplyMessage UpdateWorkRequestExtendedDescription(string urlService, OperationContext opContext, string requestId, ExtendedDescription description)
        {
            return UpdateWorkRequestExtendedDescription(urlService, opContext, requestId, description.Header, description.Body);
        }

        public static ReplyMessage UpdateWorkRequestExtendedDescription(string urlService, OperationContext opContext, string requestId, string headerText, string bodyText)
        {
            var reply = new ReplyMessage();
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');
            try
            {
                var stdTextOpContext = StdText.GetCustomOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
                var stdTextId = "WQ" + requestId;
                bool headerReply = true, bodyReply = true;
                if (!string.IsNullOrEmpty(headerText))
                    headerReply = StdText.SetHeader(urlService, stdTextOpContext, stdTextId, headerText);
                if (!string.IsNullOrEmpty(bodyText))
                    bodyReply = StdText.SetText(urlService, stdTextOpContext, stdTextId, bodyText);

                if (headerReply && bodyReply)
                    return reply;
                var errorList = new List<string>();
                if (!headerReply)
                    errorList.Add("No se pudo actualizar el encabezado de texto del StdText " + stdTextId);
                if (!bodyReply)
                    errorList.Add("No se pudo actualizar el cuerpo de texto del StdText " + stdTextId);
                reply.Errors = errorList.ToArray();
                reply.Message = "Error al actualizar el texto extendido de WorkRequest " + requestId;


            }
            catch (Exception ex)
            {
                Debugger.LogError("WorkRequest.UpdateWorkRequestExtendedDescription()", ex.Message);
                reply.Message = "Error al actualizar el texto extendido de WorkRequest " + requestId;
                var errorList = new List<string> { "No se pudo actualizar el texto del StdText WQ" + requestId };
                reply.Errors = errorList.ToArray();
            }
            return reply;
        }


        public static class Queries
        {
            public static string GetFetchWorkRequest(string dbReference, string dbLink, int searchCriteria1Key, string searchCriteria1Value, int searchCriteria2Key, string searchCriteria2Value, int dateCriteriaKey, string startDate, string endDate, string wrStatus)
            {

                var paramCriteria1 = "";
                //establecemos los parámetros del criterio 1
                if (searchCriteria1Key == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.WORK_GROUP = '" + searchCriteria1Value + "'";
                else if (searchCriteria1Key == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.EQUIP_NO = '" + searchCriteria1Value + "'";//Falta buscar el equip ref //to do
                else if (searchCriteria1Key == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NP FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE EQ.PARENT_EQUIP = '" + searchCriteria1Value + "')";
                else if (searchCriteria1Key == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.CREATION_USER = '" + searchCriteria1Value + "'";
                else if (searchCriteria1Key == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.COMPLETED_BY = '" + searchCriteria1Value + "'";
                else if (searchCriteria1Key == SearchFieldCriteriaType.AssignedTo.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.ASSIGN_PERSON = '" + searchCriteria1Value + "'";
                else if (searchCriteria1Key == SearchFieldCriteriaType.RequestType.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = "WR.WORK_REQ_TYPE = '" + searchCriteria1Value + "'";
                else if (searchCriteria1Key == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                {
                    if (searchCriteria2Key == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                        paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria1Value + "' AND TRIM(LI.LIST_ID) = '" + searchCriteria2Value + "')";
                    else if (searchCriteria2Key != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteria2Value))
                        paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria1Value + "')";
                }
                else if (searchCriteria1Key == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                {
                    if (searchCriteria2Key == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                        paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria2Value + "' AND TRIM(LI.LIST_ID) = '" + searchCriteria1Value + "')";
                    else if (searchCriteria2Key != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteria2Value))
                        paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteria1Value + "')";
                }
                else if (searchCriteria1Key == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + "EQ WHERE EQ.EQUIP_GRP_ID = '" + searchCriteria1Value + "')";
                else if (searchCriteria1Key == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + "EQ WHERE EQ.EQUIP_CLASS = '" + searchCriteria1Value + "')";
                else if (searchCriteria1Key == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteria1Value).Select(g => g.Name).ToList(), ",", "'") + ")";
                else if (searchCriteria1Key == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                    paramCriteria1 = " AND WR.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteria1Value).Select(g => g.Name).ToList(), ",", "'") + ")";
                else
                    paramCriteria1 = " AND WR.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
                //

                var paramCriteria2 = "";
                //establecemos los parámetros del criterio 2
                if (searchCriteria2Key == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.WORK_GROUP = '" + searchCriteria2Value + "'";
                else if (searchCriteria2Key == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.EQUIP_NO = '" + searchCriteria2Value + "'";//Falta buscar el equip ref //to do
                else if (searchCriteria2Key == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE EQ.PARENT_EQUIP = '" + searchCriteria2Value + "')";
                else if (searchCriteria2Key == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.CREATION_USER = '" + searchCriteria2Value + "'";
                else if (searchCriteria2Key == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.COMPLETED_BY = '" + searchCriteria2Value + "'";
                else if (searchCriteria2Key == SearchFieldCriteriaType.AssignedTo.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.ASSIGN_PERSON = '" + searchCriteria2Value + "'";
                else if (searchCriteria2Key == SearchFieldCriteriaType.RequestType.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = "WR.WORK_REQ_TYPE = '" + searchCriteria2Value + "'";
                else if (searchCriteria2Key == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                {
                    if (searchCriteria1Key == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                        paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria2Value + "' AND TRIM(LI.LIST_ID) = '" + searchCriteria1Value + "')";
                    else if (searchCriteria1Key != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteria1Value))
                        paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria2Value + "')";
                }
                else if (searchCriteria2Key == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                {
                    if (searchCriteria1Key == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteria1Value))
                        paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteria1Value + "' AND TRIM(LI.LIST_ID) = '" + searchCriteria2Value + "')";
                    else if (searchCriteria1Key != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteria1Value))
                        paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteria2Value + "')";
                }
                else if (searchCriteria2Key == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + "EQ WHERE EQ.EQUIP_GRP_ID = '" + searchCriteria2Value + "')";
                else if (searchCriteria2Key == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.EQUIP_NO IN (SELECT EQ.EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + "EQ WHERE EQ.EQUIP_CLASS = '" + searchCriteria2Value + "')";
                else if (searchCriteria2Key == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteria2Value).Select(g => g.Name).ToList(), ",", "'") + ")";
                else if (searchCriteria2Key == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteria2Value))
                    paramCriteria2 = " AND WR.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteria2Value).Select(g => g.Name).ToList(), ",", "'") + ")";
                //




                //
                //establecemos los parámetros de estado
                string statusRequirement;
                if (string.IsNullOrEmpty(wrStatus))
                    statusRequirement = "";
                else if (wrStatus == WrStatusList.Uncompleted)
                    statusRequirement = " AND WR.REQUEST_STAT IN (" + MyUtilities.GetListInSeparator(WrStatusList.GetUncompletedStatusCodes(), ",", "'") + ")";
                else if (WrStatusList.GetStatusNames().Contains(wrStatus))
                    statusRequirement = " AND WR.REQUEST_STAT = '" + WrStatusList.GetStatusCode(wrStatus) + "'";
                else
                    statusRequirement = "";

                //establecemos los parámetros para el rango de fechas
                string paramDate;
                if (string.IsNullOrEmpty(startDate))
                    startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                if (string.IsNullOrEmpty(endDate))
                    endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                if (dateCriteriaKey == SearchDateCriteriaType.None.Key)
                    paramDate = "";
                if (dateCriteriaKey == SearchDateCriteriaType.Raised.Key)
                    paramDate = " AND WR.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.Closed.Key)
                    paramDate = " AND WR.CLOSED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.Modified.Key)
                    paramDate = " AND WR.LAST_MOD_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.Creation.Key)
                    paramDate = " AND WR.CREATION_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else if (dateCriteriaKey == SearchDateCriteriaType.Required.Key)
                    paramDate = " AND WR.REQUIRED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
                else
                    paramDate = " AND WR.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";

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
                            "   WR.SLA_WARN_HOURS" +
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
                            "" + paramCriteria1 +
                            "" + paramCriteria2 +
                            "" + statusRequirement +
                            "" + paramDate;

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
            public static string GetFetchWorkRequest(string dbReference, string dbLink, string requestId)
            {
                long defaultLong;
                if (long.TryParse(requestId, out defaultLong))
                    requestId = requestId.PadLeft(12, '0');

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
                            "   WR.SLA_WARN_HOURS" +
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

                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }
        public static class SearchFieldCriteriaType
        {
            public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
            public static KeyValuePair<int, string> WorkGroup = new KeyValuePair<int, string>(1, "WorkGroup");
            public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(2, "Equipment No");
            public static KeyValuePair<int, string> ProductiveUnit = new KeyValuePair<int, string>(3, "ProductiveUnit");
            public static KeyValuePair<int, string> Originator = new KeyValuePair<int, string>(4, "Originator");
            public static KeyValuePair<int, string> CompletedBy = new KeyValuePair<int, string>(5, "CompletedBy");
            public static KeyValuePair<int, string> AssignedTo = new KeyValuePair<int, string>(6, "AssignedTo");
            public static KeyValuePair<int, string> RequestType = new KeyValuePair<int, string>(7, "RequestType");
            public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(8, "ListType");
            public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(9, "ListId");
            public static KeyValuePair<int, string> Egi = new KeyValuePair<int, string>(10, "EGI");
            public static KeyValuePair<int, string> EquipmentClass = new KeyValuePair<int, string>(11, "Equipment Class");
            public static KeyValuePair<int, string> Area = new KeyValuePair<int, string>(12, "Area");
            public static KeyValuePair<int, string> Quartermaster = new KeyValuePair<int, string>(13, "SuperIntendencia");

            public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
            {
                var list = new List<KeyValuePair<int, string>> { None, WorkGroup, EquipmentReference, ProductiveUnit, Originator, CompletedBy, AssignedTo, RequestType, ListId, ListType, Egi, EquipmentClass, Area, Quartermaster };

                return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
            }
        }
        public static class SearchDateCriteriaType
        {
            public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
            public static KeyValuePair<int, string> Raised = new KeyValuePair<int, string>(1, "Raised");
            public static KeyValuePair<int, string> Closed = new KeyValuePair<int, string>(2, "Closed");
            public static KeyValuePair<int, string> Modified = new KeyValuePair<int, string>(3, "Modified");
            public static KeyValuePair<int, string> Creation = new KeyValuePair<int, string>(4, "Creation");
            public static KeyValuePair<int, string> Required = new KeyValuePair<int, string>(5, "Required");

            public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
            {
                var list = new List<KeyValuePair<int, string>> { None, Raised, Closed, Modified, Creation, Required };

                return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
            }
        }
    }

}

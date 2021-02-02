using System.Collections.Generic;
using System.Data;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;
using EllipseMaintSchedTaskClassLibrary.MstService;

namespace EllipseMaintSchedTaskClassLibrary
{
    public static class MstActions
    {
        public static List<MaintenanceScheduleTask> FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo, string schedIndicator)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo, schedIndicator);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<MaintenanceScheduleTask>();

            if (mstDataReader == null || mstDataReader.IsClosed)
                return list;
            while (mstDataReader.Read())
            {
                var mst = GetMstFromDataReader(mstDataReader);
                list.Add(mst);
            }

            return list;
        }

        
        public static MaintenanceScheduleTask FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.DbReference, ef.DbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            if (mstDataReader == null || mstDataReader.IsClosed || !mstDataReader.Read())
                return null;
            

            var mst = GetMstFromDataReader(mstDataReader);
            return mst;
        }
        private static MaintenanceScheduleTask GetMstFromDataReader(IDataRecord mstDataRecord)
        {
            // ReSharper disable once UseObjectOrCollectionInitializer
            var mst = new MaintenanceScheduleTask();

            mst.DistrictCode = "" + mstDataRecord["DSTRCT_CODE"].ToString().Trim();
            mst.WorkGroup = "" + mstDataRecord["WORK_GROUP"].ToString().Trim();
            mst.RecType = "" + mstDataRecord["REC_700_TYPE"].ToString().Trim();
            mst.EquipmentNo = mst.RecType == MstType.Equipment ? "" + mstDataRecord["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentGrpId = mst.RecType == MstType.Egi ? "" + mstDataRecord["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentDescription = "" + mstDataRecord["EQUIPMENT_DESC"].ToString().Trim();
            mst.CompCode = "" + mstDataRecord["COMP_CODE"].ToString().Trim();
            mst.CompModCode = "" + mstDataRecord["COMP_MOD_CODE"].ToString().Trim();
            mst.MaintenanceSchTask = "" + mstDataRecord["MAINT_SCH_TASK"].ToString().Trim();
            mst.JobDescCode = "" + mstDataRecord["JOB_DESC_CODE"].ToString().Trim();
            mst.SchedDescription1 = "" + mstDataRecord["SCHED_DESC_1"].ToString().Trim();
            mst.SchedDescription2 = "" + mstDataRecord["SCHED_DESC_2"].ToString().Trim();
            mst.AssignPerson = "" + mstDataRecord["ASSIGN_PERSON"].ToString().Trim();
            mst.StdJobNo = "" + mstDataRecord["STD_JOB_NO"].ToString().Trim();
            mst.AutoRequisitionInd = "" + mstDataRecord["AUTO_REQ_IND"].ToString().Trim();
            mst.MsHistFlag = "" + mstDataRecord["MS_HIST_FLG"].ToString().Trim();
            mst.SchedInd = "" + mstDataRecord["SCHED_IND_700"].ToString().Trim();
            mst.SchedFreq1 = "" + mstDataRecord["SCHED_FREQ_1"].ToString().Trim();
            mst.StatType1 = "" + mstDataRecord["STAT_TYPE_1"].ToString().Trim();
            mst.LastSchedStat1 = "" + mstDataRecord["LAST_SCH_ST_1"].ToString().Trim();
            mst.LastPerfStat1 = "" + mstDataRecord["LAST_PERF_ST_1"].ToString().Trim();
            mst.SchedFreq2 = "" + mstDataRecord["SCHED_FREQ_2"].ToString().Trim();
            mst.StatType2 = "" + mstDataRecord["STAT_TYPE_2"].ToString().Trim();
            mst.LastSchedStat2 = "" + mstDataRecord["LAST_SCH_ST_2"].ToString().Trim();
            mst.LastPerfStat2 = "" + mstDataRecord["LAST_PERF_ST_2"].ToString().Trim();
            mst.LastSchedDate = "" + mstDataRecord["LAST_SCH_DATE"].ToString().Trim();
            mst.LastPerfDate = "" + mstDataRecord["LAST_PERF_DATE"].ToString().Trim();
            mst.NextSchedDate = "" + mstDataRecord["NEXT_SCH_DATE"].ToString().Trim();
            mst.NextSchedStat = "" + mstDataRecord["NEXT_SCH_STAT"].ToString().Trim();
            mst.NextSchedValue = "" + mstDataRecord["NEXT_SCH_VALUE"].ToString().Trim();
            mst.ShutdownType = "" + mstDataRecord["SHUTDOWN_TYPE"].ToString().Trim();
            mst.ShutdownEquip = "" + mstDataRecord["SHUTDOWN_EQUIP"].ToString().Trim();
            mst.ShutdownNo = "" + mstDataRecord["SHUTDOWN_NO"].ToString().Trim();
            mst.CondMonPos = "" + mstDataRecord["COND_MON_POS"].ToString().Trim();
            mst.CondMonType = "" + mstDataRecord["COND_MON_TYPE"].ToString().Trim();
            mst.StatutoryFlg = "" + mstDataRecord["STATUTORY_FLG"].ToString().Trim();
            mst.OccurrenceType = "" + mstDataRecord["OCCURENCE_TYPE"].ToString().Trim();
            mst.DayOfWeek = "" + mstDataRecord["DAY_WEEK"].ToString().Trim();
            mst.DayOfMonth = "" + mstDataRecord["DAY_MONTH"].ToString().Trim();
            mst.StartYear = "" + mstDataRecord["START_YEAR"].ToString().Trim();
            mst.StartMonth = "" + mstDataRecord["START_MONTH"].ToString().Trim();

            return mst;
        }

        public static MSTServiceResult CreateMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {
            var service = new MSTService{ Url = urlService + "/MSTService" };

            if (string.IsNullOrWhiteSpace(mst?.IsInSeries) && mst?.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("8") && mst.MaintenanceSchTask.Length == 4)
                mst.IsInSeries = "Y";
            if (string.IsNullOrWhiteSpace(mst?.IsInSuppressionSeries) && mst?.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("9") && mst.MaintenanceSchTask.Length == 4)
                mst.IsInSuppressionSeries = "Y";

            return service.create(opContext, mst?.ToMstDto());
        }

        public static MSTServiceResult ModifyMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var service = new MSTService { Url = urlService + "/MSTService" };
            
            if (string.IsNullOrWhiteSpace(mst?.IsInSeries) && mst?.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("8") && mst.MaintenanceSchTask.Length == 4)
                mst.IsInSeries = "Y";
            if (string.IsNullOrWhiteSpace(mst?.IsInSuppressionSeries) && mst?.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("9") && mst.MaintenanceSchTask.Length == 4)
                mst.IsInSuppressionSeries = "Y";

            return service.update(opContext, mst?.ToMstDto());
        }
        public static ModifyNextScheduleDetailsServiceResult ModNextSchedMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var mstService = new MSTService{Url = urlService + "/MSTService"};
            var request = new ModifyNextScheduleDetailsDTO()
            {
                rec700Type = string.IsNullOrWhiteSpace(mst.RecType) ? "ES" : mst.RecType,
                equipNo = mst.EquipmentNo,
                equipGrpId = mst.EquipmentGrpId,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintSchTask = mst.MaintenanceSchTask,
                nextSchDate = string.IsNullOrWhiteSpace(mst.NextSchedDate) ? default : MyUtilities.ToDate(mst.NextSchedDate),
                nextSchDateSpecified = !string.IsNullOrWhiteSpace(mst.NextSchedDate),
                nextSchStat = mst.NextSchedStat,
                nextSchValue = string.IsNullOrWhiteSpace(mst.NextSchedValue) ? default : MyUtilities.ToDecimal(mst.NextSchedValue),
                nextSchValueSpecified = !string.IsNullOrWhiteSpace(mst.NextSchedValue),
                nextSchMeterValue = string.IsNullOrWhiteSpace(mst.NextSchedMeterValue) ? default : MyUtilities.ToDecimal(mst.NextSchedMeterValue),
                nextSchMeterValueSpecified = !string.IsNullOrWhiteSpace(mst.NextSchedMeterValue),
                occurenceType = mst.OccurrenceType,
                dayMonth = mst.DayOfMonth,
                dayWeek = mst.DayOfWeek,
                schedFreq1 = string.IsNullOrWhiteSpace(mst.SchedFreq1) ? default : MyUtilities.ToDecimal(mst.SchedFreq1),
                schedFreq1Specified = !string.IsNullOrWhiteSpace(mst.SchedFreq1),
                schedFreq2 = string.IsNullOrWhiteSpace(mst.SchedFreq2) ? default : MyUtilities.ToDecimal(mst.SchedFreq2),
                schedFreq2Specified = !string.IsNullOrWhiteSpace(mst.SchedFreq2),
                startMonth = mst.StartMonth,
                startYear = mst.StartYear,
            };

            return mstService.modifyNextScheduleDetails(opContext, request);
        }

        public static MSTServiceResult DeleteMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {
            var service = new MSTService{Url = urlService + "/MSTService" };

            var request = new MSTDTO
            {
                equipGrpId = mst.EquipmentGrpId,
                equipNo = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintSchTask = mst.MaintenanceSchTask
            };

            return service.delete(opContext, request);
        }
        
        public static OperationContext GetMstServiceOperationContext(string district, string position)
        {
            const int maxInstances = 100;
            var returnWarnings = Debugger.DebugWarnings;

            return GetMstServiceOperationContext(district, position, maxInstances, returnWarnings);
        }
        public static OperationContext GetMstServiceOperationContext(string district, string position, int maxInstances, bool returnWarnings)
        {
            var opContext = new OperationContext()
            {
                district = district,
                position = position,
                maxInstances = maxInstances,
                maxInstancesSpecified = true,
                returnWarnings = returnWarnings,
                returnWarningsSpecified = true
            };
            return opContext;
        }
        

        public static List<Mst> ForecastMaintenanceScheduleTask(string urlService, OperationContext opContext, MstForecast mstForecast)
        {
            var list = new List<Mst>();
            var service = new MSTService() { Url = urlService + "/MSTService" };

            
            var mstMwRestart = new MstService.MSTiMWPDTO();

            var result = service.forecast(opContext, mstForecast.ToDto(), mstMwRestart);

            foreach (var item in result)
            {
                var mstItem = new Mst(item.MSTiMWPDTO);
                list.Add(mstItem);
            }

            return list;
        }
        /*
        public static List<Mst> ForecastMaintenanceScheduleTaskPost(EllipseFunctions ef, MstForecast mstForecast)
        {
            var operationId1 = Util.GetNewOperationId();
            //var operationId2 = Util.GetNewOperationId();

            var responseDto = ef.InitiatePostConnection();
            if (responseDto.GotErrorMessages())
                throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "<actions>";
            requestXml = requestXml + "     <action>";
            requestXml = requestXml + "         <name>service</name>";
            requestXml = requestXml + "         <data>";
            requestXml = requestXml + "             <name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "             <operation>forecast</operation>";
            requestXml = requestXml + "             <className>mfui.actions.grid::ParentChildSearchAction</className>";
            requestXml = requestXml + "             <returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "             <dto uuid=\"" + Util.GetNewOperationId() + "\">";
            requestXml = requestXml + "                <maintSchTask>" + mstForecast.MaintSchTask + "</maintSchTask>";
            requestXml = requestXml + "                <equipNo>" + mstForecast.EquipNo + "</equipNo>";
            requestXml = requestXml + "                <nInstances>" + mstForecast.Ninstances + "</nInstances>";
            requestXml = requestXml + "                <showRelated>" + mstForecast.ShowRelated + "</showRelated>";
            requestXml = requestXml + "                <hideSuppressed>" + mstForecast.HideSuppressed + "</hideSuppressed>";
            requestXml = requestXml + "                <compCode>" + mstForecast.CompCode + "</compCode>";
            requestXml = requestXml + "                <compModCode>" + mstForecast.CompModCode+ "</compModCode>";
            requestXml = requestXml + "                <rec700Type>" + mstForecast.Rec700Type + "</rec700Type>";
            requestXml = requestXml + "             </dto>";
            //requestXml = requestXml + "             <requiredAttributes>";
            //requestXml = requestXml + "                <reference>true</reference>";
            //requestXml = requestXml + "                <plannedStartDate>true</plannedStartDate>";
            //requestXml = requestXml + "                <plannedStartTime>true</plannedStartTime>";
            //requestXml = requestXml + "                <woDesc>true</woDesc>";
            //requestXml = requestXml + "                <tolerancePC>true</tolerancePC>";
            //requestXml = requestXml + "                <toleranceDays>true</toleranceDays>";
            //requestXml = requestXml + "                <maintSchTask>true</maintSchTask>";
            //requestXml = requestXml + "                <equipNo>true</equipNo>";
            //requestXml = requestXml + "                <nInstances>true</nInstances>";
            //requestXml = requestXml + "                <showRelated>true</showRelated>";
            //requestXml = requestXml + "                <hideSuppressed>true</hideSuppressed>";
            //requestXml = requestXml + "                <compCode>true</compCode>";
            //requestXml = requestXml + "                <compModCode>true</compModCode>";
            //requestXml = requestXml + "                <rec700Type>true</rec700Type>";
            //requestXml = requestXml + "             </requiredAttributes>";
            //requestXml = requestXml + "             <maxInstances>50</maxInstances>";
            requestXml = requestXml + "         </data>";
            requestXml = requestXml + "         <id>" + operationId1 + "</id>";
            requestXml = requestXml + "     </action>";
            //requestXml = requestXml + "     <action>";
            //requestXml = requestXml + "     	<name>csbDirectiveForGrid</name>";
            //requestXml = requestXml + "     	<data>";
            //requestXml = requestXml + "     		<applicationName>msemst</applicationName>";
            //requestXml = requestXml + "     		<applicationPage>detail</applicationPage>";
            //requestXml = requestXml + "     		<gridId>forecastGrid</gridId>";
            //requestXml = requestXml + "     		<metadata>";
            //requestXml = requestXml + "     			<applicationName>msemst</applicationName>";
            //requestXml = requestXml + "     			<applicationPage>detail</applicationPage>";
            //requestXml = requestXml + "     			<applicationCustomName/>";
            //requestXml = requestXml + "     			<isCustomised>false</isCustomised>";
            //requestXml = requestXml + "     			<applicationCode/>";
            //requestXml = requestXml + "     		</metadata>";
            //requestXml = requestXml + "     	</data>";
            //requestXml = requestXml + "     	<id>" + operationId2 + "</id>";
            //requestXml = requestXml + "     </action>";
            requestXml = requestXml + "</actions>";
            requestXml = requestXml + "<chains>";
            //requestXml = requestXml + "	<chain>";
            //requestXml = requestXml + "		<fieldPairings/>";
            //requestXml = requestXml + "		<mapping/>";
            //requestXml = requestXml + "		<fromAction>" + operationId1 + "</fromAction>";
            //requestXml = requestXml + "		<toAction>" + operationId2 + "</toAction>";
            //requestXml = requestXml + "	</chain>";
            requestXml = requestXml + "</chains>";
            requestXml = requestXml + "<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "<application>msemst</application>";
            requestXml = requestXml + "<applicationPage>read</applicationPage>";
            requestXml = requestXml + "</interaction>";


            requestXml = requestXml.Replace("&", "&amp;");
            //requestXml = requestXml.Replace("\t", "");
            responseDto = ef.ExecutePostRequest(requestXml);

            
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);

            var mstList = new List<Mst>();
            var xPath = "/interaction/actions/action/data/result/dto";
            var xnList = MyUtilities.Xml.GetNodeList(responseDto.ResponseXML, xPath);
            foreach (XmlNode xn in xnList)
            {
                var myObject = new Mst();
                myObject = (Mst)MyUtilities.Xml.DeSerializeXmlNodeToObject(xn, myObject.GetType());
                mstList.Add(myObject);
            }
            return mstList;
        }
        */



    }



}

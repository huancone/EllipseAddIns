using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;
using EllipseMaintSchedTaskClassLibrary.MaintSchedTskService;
using System.Web.Services.Ellipse.Post;

namespace EllipseMaintSchedTaskClassLibrary
{
    public static class MstActions
    {
        public static List<MaintenanceScheduleTask> FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo, string schedIndicator)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo, schedIndicator);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<MaintenanceScheduleTask>();

            if (mstDataReader == null || mstDataReader.IsClosed || !mstDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
            while (mstDataReader.Read())
            {
                // ReSharper disable once UseObjectOrCollectionInitializer
                var mst = new MaintenanceScheduleTask();

                mst.DistrictCode = "" + mstDataReader["DSTRCT_CODE"].ToString().Trim();
                mst.WorkGroup = "" + mstDataReader["WORK_GROUP"].ToString().Trim();
                mst.RecType = "" + mstDataReader["REC_700_TYPE"].ToString().Trim();
                mst.EquipmentNo = mst.RecType == MstType.Equipment ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
                mst.EquipmentGrpId = mst.RecType == MstType.Egi ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
                mst.EquipmentDescription = "" + mstDataReader["EQUIPMENT_DESC"].ToString().Trim();
                mst.CompCode = "" + mstDataReader["COMP_CODE"].ToString().Trim();
                mst.CompModCode = "" + mstDataReader["COMP_MOD_CODE"].ToString().Trim();
                mst.MaintenanceSchTask = "" + mstDataReader["MAINT_SCH_TASK"].ToString().Trim();
                mst.JobDescCode = "" + mstDataReader["JOB_DESC_CODE"].ToString().Trim();
                mst.SchedDescription1 = "" + mstDataReader["SCHED_DESC_1"].ToString().Trim();
                mst.SchedDescription2 = "" + mstDataReader["SCHED_DESC_2"].ToString().Trim();
                mst.AssignPerson = "" + mstDataReader["ASSIGN_PERSON"].ToString().Trim();
                mst.StdJobNo = "" + mstDataReader["STD_JOB_NO"].ToString().Trim();
                mst.AutoRequisitionInd = "" + mstDataReader["AUTO_REQ_IND"].ToString().Trim();
                mst.MsHistFlag = "" + mstDataReader["MS_HIST_FLG"].ToString().Trim();
                mst.SchedInd = "" + mstDataReader["SCHED_IND_700"].ToString().Trim();
                mst.SchedFreq1 = "" + mstDataReader["SCHED_FREQ_1"].ToString().Trim();
                mst.StatType1 = "" + mstDataReader["STAT_TYPE_1"].ToString().Trim();
                mst.LastSchedStat1 = "" + mstDataReader["LAST_SCH_ST_1"].ToString().Trim();
                mst.LastPerfStat1 = "" + mstDataReader["LAST_PERF_ST_1"].ToString().Trim();
                mst.SchedFreq2 = "" + mstDataReader["SCHED_FREQ_2"].ToString().Trim();
                mst.StatType2 = "" + mstDataReader["STAT_TYPE_2"].ToString().Trim();
                mst.LastSchedStat2 = "" + mstDataReader["LAST_SCH_ST_2"].ToString().Trim();
                mst.LastPerfStat2 = "" + mstDataReader["LAST_PERF_ST_2"].ToString().Trim();
                mst.LastSchedDate = "" + mstDataReader["LAST_SCH_DATE"].ToString().Trim();
                mst.LastPerfDate = "" + mstDataReader["LAST_PERF_DATE"].ToString().Trim();
                mst.NextSchedDate = "" + mstDataReader["NEXT_SCH_DATE"].ToString().Trim();
                mst.NextSchedStat = "" + mstDataReader["NEXT_SCH_STAT"].ToString().Trim();
                mst.NextSchedValue = "" + mstDataReader["NEXT_SCH_VALUE"].ToString().Trim();
                mst.ShutdownType = "" + mstDataReader["SHUTDOWN_TYPE"].ToString().Trim();
                mst.ShutdownEquip = "" + mstDataReader["SHUTDOWN_EQUIP"].ToString().Trim();
                mst.ShutdownNo = "" + mstDataReader["SHUTDOWN_NO"].ToString().Trim();
                mst.CondMonPos = "" + mstDataReader["COND_MON_POS"].ToString().Trim();
                mst.CondMonType = "" + mstDataReader["COND_MON_TYPE"].ToString().Trim();
                mst.StatutoryFlg = "" + mstDataReader["STATUTORY_FLG"].ToString().Trim();
                mst.OccurrenceType = "" + mstDataReader["OCCURENCE_TYPE"].ToString().Trim();
                mst.DayOfWeek = "" + mstDataReader["DAY_WEEK"].ToString().Trim();
                mst.DayOfMonth = "" + mstDataReader["DAY_MONTH"].ToString().Trim();
                mst.StartYear = "" + mstDataReader["START_YEAR"].ToString().Trim();
                mst.StartMonth = "" + mstDataReader["START_MONTH"].ToString().Trim();
                list.Add(mst);
            }

            return list;
        }
        public static MaintenanceScheduleTask FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            if (mstDataReader == null || mstDataReader.IsClosed || !mstDataReader.HasRows || !mstDataReader.Read())
                return null;


            // ReSharper disable once UseObjectOrCollectionInitializer
            var mst = new MaintenanceScheduleTask();

            mst.DistrictCode = "" + mstDataReader["DSTRCT_CODE"].ToString().Trim();
            mst.WorkGroup = "" + mstDataReader["WORK_GROUP"].ToString().Trim();
            mst.RecType = "" + mstDataReader["REC_700_TYPE"].ToString().Trim();
            mst.EquipmentNo = mst.RecType == MstType.Equipment ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentGrpId = mst.RecType == MstType.Egi ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentDescription = "" + mstDataReader["EQUIPMENT_DESC"].ToString().Trim();
            mst.CompCode = "" + mstDataReader["COMP_CODE"].ToString().Trim();
            mst.CompModCode = "" + mstDataReader["COMP_MOD_CODE"].ToString().Trim();
            mst.MaintenanceSchTask = "" + mstDataReader["MAINT_SCH_TASK"].ToString().Trim();
            mst.JobDescCode = "" + mstDataReader["JOB_DESC_CODE"].ToString().Trim();
            mst.SchedDescription1 = "" + mstDataReader["SCHED_DESC_1"].ToString().Trim();
            mst.SchedDescription2 = "" + mstDataReader["SCHED_DESC_2"].ToString().Trim();
            mst.AssignPerson = "" + mstDataReader["ASSIGN_PERSON"].ToString().Trim();
            mst.StdJobNo = "" + mstDataReader["STD_JOB_NO"].ToString().Trim();
            mst.AutoRequisitionInd = "" + mstDataReader["AUTO_REQ_IND"].ToString().Trim();
            mst.MsHistFlag = "" + mstDataReader["MS_HIST_FLG"].ToString().Trim();
            mst.SchedInd = "" + mstDataReader["SCHED_IND_700"].ToString().Trim();
            mst.SchedFreq1 = "" + mstDataReader["SCHED_FREQ_1"].ToString().Trim();
            mst.StatType1 = "" + mstDataReader["STAT_TYPE_1"].ToString().Trim();
            mst.LastSchedStat1 = "" + mstDataReader["LAST_SCH_ST_1"].ToString().Trim();
            mst.LastPerfStat1 = "" + mstDataReader["LAST_PERF_ST_1"].ToString().Trim();
            mst.SchedFreq2 = "" + mstDataReader["SCHED_FREQ_2"].ToString().Trim();
            mst.StatType2 = "" + mstDataReader["STAT_TYPE_2"].ToString().Trim();
            mst.LastSchedStat2 = "" + mstDataReader["LAST_SCH_ST_2"].ToString().Trim();
            mst.LastPerfStat2 = "" + mstDataReader["LAST_PERF_ST_2"].ToString().Trim();
            mst.LastSchedDate = "" + mstDataReader["LAST_SCH_DATE"].ToString().Trim();
            mst.LastPerfDate = "" + mstDataReader["LAST_PERF_DATE"].ToString().Trim();
            mst.NextSchedDate = "" + mstDataReader["NEXT_SCH_DATE"].ToString().Trim();
            mst.NextSchedStat = "" + mstDataReader["NEXT_SCH_STAT"].ToString().Trim();
            mst.NextSchedValue = "" + mstDataReader["NEXT_SCH_VALUE"].ToString().Trim();
            mst.ShutdownType = "" + mstDataReader["SHUTDOWN_TYPE"].ToString().Trim();
            mst.ShutdownEquip = "" + mstDataReader["SHUTDOWN_EQUIP"].ToString().Trim();
            mst.ShutdownNo = "" + mstDataReader["SHUTDOWN_NO"].ToString().Trim();
            mst.CondMonPos = "" + mstDataReader["COND_MON_POS"].ToString().Trim();
            mst.CondMonType = "" + mstDataReader["COND_MON_TYPE"].ToString().Trim();
            mst.StatutoryFlg = "" + mstDataReader["STATUTORY_FLG"].ToString().Trim();
            mst.OccurrenceType = "" + mstDataReader["OCCURENCE_TYPE"].ToString().Trim();
            mst.DayOfWeek = "" + mstDataReader["DAY_WEEK"].ToString().Trim();
            mst.DayOfMonth = "" + mstDataReader["DAY_MONTH"].ToString().Trim();
            mst.StartYear = "" + mstDataReader["START_YEAR"].ToString().Trim();
            mst.StartMonth = "" + mstDataReader["START_MONTH"].ToString().Trim();
            return mst;
        }



        public static MaintSchedTskServiceCreateReplyDTO CreateMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService { Url = urlService + "/MaintSchedTskService" };
            var request = new MaintSchedTskServiceCreateRequestDTO
            {
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedDescription1 = mst.SchedDescription1,
                schedDescription2 = mst.SchedDescription2,
                workGroup = mst.WorkGroup,
                assignPerson = mst.AssignPerson,
                jobDescCode = mst.JobDescCode,
                stdJobNo = mst.StdJobNo,
                districtCode = mst.DistrictCode,
                autoRequisitionInd = MyUtilities.IsTrue(mst.AutoRequisitionInd),
                autoRequisitionIndSpecified = mst.AutoRequisitionInd != null,
                MSHistFlag = MyUtilities.IsTrue(mst.MsHistFlag),
                MSHistFlagSpecified = mst.MsHistFlag != null,
                schedInd = mst.SchedInd,
                statType1 = mst.StatType1,
                lastSchedStat1 = !string.IsNullOrWhiteSpace(mst.LastSchedStat1)
                    ? Convert.ToDecimal(mst.LastSchedStat1)
                    : 0,
                lastSchedStat1Specified = mst.LastSchedStat1 != null,
                schedFreq1 = !string.IsNullOrWhiteSpace(mst.SchedFreq1)
                    ? Convert.ToDecimal(mst.SchedFreq1)
                    : 0,
                schedFreq1Specified = mst.SchedFreq1 != null,
                lastPerfStat1 = !string.IsNullOrWhiteSpace(mst.LastPerfStat1)
                    ? Convert.ToDecimal(mst.LastPerfStat1)
                    : 0,
                lastPerfStat1Specified = mst.LastPerfStat1 != null,
                statType2 = mst.StatType2,
                lastSchedStat2 = !string.IsNullOrWhiteSpace(mst.LastSchedStat2)
                    ? Convert.ToDecimal(mst.LastSchedStat2)
                    : 0,
                lastSchedStat2Specified = mst.LastSchedStat2 != null,
                schedFreq2 = !string.IsNullOrWhiteSpace(mst.SchedFreq2)
                    ? Convert.ToDecimal(mst.SchedFreq2)
                    : 0,
                schedFreq2Specified = mst.SchedFreq2 != null,
                lastPerfStat2 = !string.IsNullOrWhiteSpace(mst.LastPerfStat2)
                    ? Convert.ToDecimal(mst.LastPerfStat2)
                    : 0,
                lastPerfStat2Specified = mst.LastPerfStat2 != null,
                lastSchedDate = mst.LastSchedDate,
                lastPerfDate = mst.LastPerfDate,
                statutoryFlg = MyUtilities.IsTrue(mst.StatutoryFlg),
                statutoryFlgSpecified = mst.StatutoryFlg != null,
                occurenceType = mst.OccurrenceType,
                dayOfWeek = mst.DayOfWeek,
                dayOfMonth = mst.DayOfMonth,
                startMonth = mst.StartMonth,
                startYear = mst.StartYear,
                conAstSegFrSpecified = true,
                conAstSegFr = 0,
                conAstSegToSpecified = true,
                conAstSegTo = 0

            };

            return proxyEquip.create(opContext, request);
        }

        public static MaintSchedTskServiceModifyReplyDTO ModifyMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceModifyRequestDTO
            {
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentNo = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedDescription1 = mst.SchedDescription1,
                schedDescription2 = mst.SchedDescription2,
                workGroup = mst.WorkGroup,
                assignPerson = mst.AssignPerson,
                jobDescCode = mst.JobDescCode,
                stdJobNo = mst.StdJobNo,
                districtCode = mst.DistrictCode,
                autoRequisitionInd = MyUtilities.IsTrue(mst.AutoRequisitionInd),
                autoRequisitionIndSpecified = mst.AutoRequisitionInd != null,
                MSHistFlag = MyUtilities.IsTrue(mst.MsHistFlag),
                MSHistFlagSpecified = mst.MsHistFlag != null,
                schedInd = mst.SchedInd,
                statType1 = mst.StatType1,
                lastSchedStat1 = !string.IsNullOrWhiteSpace(mst.LastSchedStat1) ? Convert.ToDecimal(mst.LastSchedStat1) : 0,
                lastSchedStat1Specified = mst.LastSchedStat1 != null,
                schedFreq1 = !string.IsNullOrWhiteSpace(mst.SchedFreq1) ? Convert.ToDecimal(mst.SchedFreq1) : 0,
                schedFreq1Specified = mst.SchedFreq1 != null,
                lastPerfStat1 = !string.IsNullOrWhiteSpace(mst.LastPerfStat1) ? Convert.ToDecimal(mst.LastPerfStat1) : 0,
                lastPerfStat1Specified = mst.LastPerfStat1 != null,
                statType2 = mst.StatType2,
                lastSchedStat2 = !string.IsNullOrWhiteSpace(mst.LastSchedStat2) ? Convert.ToDecimal(mst.LastSchedStat2) : 0,
                lastSchedStat2Specified = mst.LastSchedStat2 != null,
                schedFreq2 = !string.IsNullOrWhiteSpace(mst.SchedFreq2) ? Convert.ToDecimal(mst.SchedFreq2) : 0,
                schedFreq2Specified = mst.SchedFreq2 != null,
                lastPerfStat2 = !string.IsNullOrWhiteSpace(mst.LastPerfStat2) ? Convert.ToDecimal(mst.LastPerfStat2) : 0,
                lastPerfStat2Specified = mst.LastPerfStat2 != null,
                lastSchedDate = mst.LastSchedDate,
                lastPerfDate = mst.LastPerfDate,
                statutoryFlg = MyUtilities.IsTrue(mst.StatutoryFlg),
                statutoryFlgSpecified = mst.StatutoryFlg != null,
                occurenceType = mst.OccurrenceType,
                dayOfWeek = mst.DayOfWeek,
                dayOfMonth = mst.DayOfMonth,
                startMonth = mst.StartMonth,
                startYear = mst.StartYear,
                conAstSegFrSpecified = true,
                conAstSegFr = 1,
                conAstSegToSpecified = true,
                conAstSegTo = 1
            };

            proxyEquip.Url = urlService + "/MaintSchedTskService";
            return proxyEquip.modify(opContext, request);
        }

        public static void CreateMaintenanceScheduleTaskPost(EllipseFunctions ef, MaintenanceScheduleTask mst)
        {
            var responseDto = ef.InitiatePostConnection();
            if (responseDto.GotErrorMessages()) return;


            var requestXml = "";

            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "				<operation>create</operation>";
            requestXml = requestXml + "				<className>mfui.actions.detail::CreateAction</className>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<msHistFlg>Y</msHistFlg>";
            requestXml = requestXml + "					<occurenceType> </occurenceType>";
            requestXml = requestXml + "					<dayWeek>" + mst.DayOfWeek + "</dayWeek>";
            requestXml = requestXml + "					<startMonth>" + mst.StartMonth + "</startMonth>";
            requestXml = requestXml + "					<scheduleType> </scheduleType>";
            requestXml = requestXml + "					<nextSchInd> </nextSchInd>";
            requestXml = requestXml + "					<autoReqInd>N</autoReqInd>";
            requestXml = requestXml + "					<outputConAstSeg/>";
            requestXml = requestXml + "					<parentChildType/>";
            requestXml = requestXml + "					<equipNo/>";
            requestXml = requestXml + "					<outputMeasureType/>";
            requestXml = requestXml + "					<equipRef>" + mst.EquipmentNo + "</equipRef>";
            requestXml = requestXml + "					<equipNoD/>";
            requestXml = requestXml + "                 <equipGrpId>" + mst.EquipmentGrpId + "</equipGrpId>";
            requestXml = requestXml + "                 <compCode>" + mst.CompCode + "</compCode>";
            requestXml = requestXml + "                 <compModCode>" + mst.CompModCode + "</compModCode>";
            requestXml = requestXml + "                 <maintSchTask>" + mst.MaintenanceSchTask + "</maintSchTask>";
            requestXml = requestXml + "					<conAstSegFrNumeric/>";
            requestXml = requestXml + "					<conAstSegToNumeric/>";
            requestXml = requestXml + "					<segmentUnitOfMeasure/>";
            requestXml = requestXml + "					<fromInKilometers/>";
            requestXml = requestXml + "					<toInKilometers/>";
            requestXml = requestXml + "					<fromInMilesYards/>";
            requestXml = requestXml + "					<toInMilesYards/>";
            requestXml = requestXml + "					<fromInMilesChains/>";
            requestXml = requestXml + "					<toInMilesChains/>";
            requestXml = requestXml + "					<linkParent/>";
            requestXml = requestXml + "					<linkId/>";
            requestXml = requestXml + "					<displayInd>2</displayInd>";
            requestXml = requestXml + "					<rec700Type/>";
            requestXml = requestXml + "                 <schedDesc1>" + mst.SchedDescription1 + "</schedDesc1>";
            requestXml = requestXml + "                 <schedDesc2>" + mst.SchedDescription2 + " </schedDesc2>";
            requestXml = requestXml + "                 <workGroup>" + mst.WorkGroup + "</workGroup>";
            requestXml = requestXml + "                 <assignPerson>" + mst.AssignPerson + "</assignPerson>";
            requestXml = requestXml + "                 <stdJobNo>" + mst.StdJobNo + "</stdJobNo>";
            requestXml = requestXml + "					<unitsRequired/>";
            requestXml = requestXml + "					<stdUnitOfWork/>";
            requestXml = requestXml + "					<stdUnitsRequired/>";
            requestXml = requestXml + "					<unitsScale/>";
            requestXml = requestXml + "					<shutdownType/>";
            requestXml = requestXml + "					<jobDescCode>" + mst.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "					<dstrctCode>" + mst.DistrictCode + "</dstrctCode>";
            requestXml = requestXml + "					<schedInd700>" + mst.SchedInd + "</schedInd700>";
            requestXml = requestXml + "					<statType1>" + mst.StatType1 + "</statType1>";
            requestXml = requestXml + "					<lastSchStat1>" + (!string.IsNullOrWhiteSpace(mst.LastSchedStat1) ? Convert.ToString(mst.LastSchedStat1) : "") + "</lastSchStat1>";
            requestXml = requestXml + "					<lastPerfStat1>" + (!string.IsNullOrWhiteSpace(mst.LastPerfStat1) ? Convert.ToString(mst.LastPerfStat1) : "") + "</lastPerfStat1>";
            requestXml = requestXml + "					<dayMonth>" + mst.DayOfMonth + "</dayMonth>";
            requestXml = requestXml + "					<startYear>" + mst.StartYear + "</startYear>"; ;
            requestXml = requestXml + "					<schedFreq1>" + mst.SchedFreq1 + "</schedFreq1>";
            requestXml = requestXml + "					<statType2>" + mst.StatType2 + "</statType2>";
            requestXml = requestXml + "					<schedFreq1>" + mst.SchedFreq2 + "</schedFreq1>";
            requestXml = requestXml + "					<lastSchStat2>" + (!string.IsNullOrWhiteSpace(mst.LastSchedStat2) ? Convert.ToString(mst.LastSchedStat2) : "") + "</lastSchStat2>";
            requestXml = requestXml + "					<lastPerfStat2>" + (!string.IsNullOrWhiteSpace(mst.LastSchedStat2) ? Convert.ToString(mst.LastSchedStat2) : "") + "</lastPerfStat2>";
            requestXml = requestXml + "					<lastSchDate>" + (!string.IsNullOrWhiteSpace(mst.LastSchedDate) ? Convert.ToString(mst.LastSchedStat2) : "") + " </lastSchDate> ";
            requestXml = requestXml + "					<lastPerfDate>" + (!string.IsNullOrWhiteSpace(mst.LastPerfDate) ? Convert.ToString(mst.LastSchedStat2) : "") + " </LastPerfDate> ";
            requestXml = requestXml + "					<nextSchDate>" + (!string.IsNullOrWhiteSpace(mst.NextSchedDate) ? Convert.ToString(mst.LastSchedStat2) : "") + " </NextSchedDate> ";
            requestXml = requestXml + "					<nextSchStat>" + (!string.IsNullOrWhiteSpace(mst.NextSchedStat) ? Convert.ToString(mst.LastSchedStat2) : "") + " </NextSchedStat> ";
            requestXml = requestXml + "					<nextSchValue>" + (!string.IsNullOrWhiteSpace(mst.NextSchedValue) ? Convert.ToString(mst.LastSchedStat2) : "") + " </NextSchedValue> ";
            requestXml = requestXml + "					<nextSchMeterValue/>";
            requestXml = requestXml + "					<recallTimeHrs/>";
            requestXml = requestXml + "					<statutoryFlg>" + MyUtilities.IsTrue(mst.StatutoryFlg) + "</statutoryFlg>";
            requestXml = requestXml + "					<fixedScheduling>N</fixedScheduling>";
            requestXml = requestXml + "					<allowMultiple>Y</allowMultiple>";
            requestXml = requestXml + "					<stdJobDstrctCode/>";
            requestXml = requestXml + "					<stdJobDesc></stdJobDesc>";
            requestXml = requestXml + "					<stdJobWorkGroup></stdJobWorkGroup>";
            requestXml = requestXml + "					<stdJobWorkOrderType></stdJobWorkOrderType>";
            requestXml = requestXml + "					<stdJobMaintenanceType></stdJobMaintenanceType>";
            requestXml = requestXml + "					<stdJobOriginatorPriority></stdJobOriginatorPriority>";
            requestXml = requestXml + "					<stdJobCompCode/>";
            requestXml = requestXml + "					<stdJobCompModCode/>";
            requestXml = requestXml + "					<stdJobEstimatedDurationHours></stdJobEstimatedDurationHours>";
            requestXml = requestXml + "					<condMonPos/>";
            requestXml = requestXml + "					<condMonType/>";
            requestXml = requestXml + "					<isInSeries>N</isInSeries>";
            requestXml = requestXml + "					<isInSuppressionSeries>N</isInSuppressionSeries>";
            requestXml = requestXml + "					<hideSuppressed>Y</hideSuppressed>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msemst</application>";
            requestXml = requestXml + "	<applicationPage>create</applicationPage>";
            requestXml = requestXml + "	<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";


            requestXml = requestXml.Replace("&", "&amp;");
            requestXml = requestXml.Replace("\t", "");
            responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
        }

        public static void ModifyMaintenanceScheduleTaskPost(EllipseFunctions ef, MaintenanceScheduleTask mst)
        {
            var responseDto = ef.InitiatePostConnection();
            if (responseDto.GotErrorMessages()) return;

            var requestXml = "";

            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "				<operation>update</operation>";
            requestXml = requestXml + "				<className>mfui.actions.detail::UpdateAction</className>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto    uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<allowMultiple>Y</allowMultiple>";
            requestXml = requestXml + "					<conAstSegFr>0</conAstSegFr>";
            requestXml = requestXml + "					<conAstSegFrNumeric>" + mst.ConAstSegFr + "</conAstSegFrNumeric>";
            requestXml = requestXml + "					<conAstSegTo>0</conAstSegTo>";
            requestXml = requestXml + "					<conAstSegToNumeric>" + mst.ConAstSegTo + "</conAstSegToNumeric>";
            requestXml = requestXml + "					<dayMonth>" + mst.DayOfMonth + "</dayMonth>";
            requestXml = requestXml + "					<dayWeek> " + mst.DayOfWeek + " </dayWeek>";
            requestXml = requestXml + "					<dstrctCode>" + mst.DistrictCode + "</dstrctCode>";
            requestXml = requestXml + "					<equipEntity>" + mst.EquipmentNo + "</equipEntity>";
            requestXml = requestXml + "					<equipNo>" + mst.EquipmentNo + "</equipNo>";
            requestXml = requestXml + "					<equipRef>" + mst.EquipmentNo + "</equipRef>";
            requestXml = requestXml + "					<fixedScheduling>Y</fixedScheduling>";
            requestXml = requestXml + "					<isInSeries>Y</isInSeries>";
            requestXml = requestXml + "					<isInSuppressionSeries>Y</isInSuppressionSeries>";
            requestXml = requestXml + "					<jobDescCode>" + mst.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "					<lastPerfDate>" + mst.LastPerfDate + "</lastPerfDate>";
            requestXml = requestXml + "					<lastPerfStat1>" + mst.LastPerfStat1 + "</lastPerfStat1>";
            requestXml = requestXml + "					<lastSchDate>" + mst.LastSchedDate + "</lastSchDate>";
            requestXml = requestXml + "					<lastSchStat1>" + mst.LastSchedStat1 + "</lastSchStat1>";
            requestXml = requestXml + "					<linkedInd>N</linkedInd>";
            requestXml = requestXml + "					<maintSchTask>" + mst.MaintenanceSchTask + "</maintSchTask>";
            requestXml = requestXml + "					<msHistFlg>Y</msHistFlg>";
            requestXml = requestXml + "					<nextSchDate>" + mst.NextSchedDate + "</nextSchDate>";
            requestXml = requestXml + "					<rec700Type>" + mst.RecType + "</rec700Type>";
            requestXml = requestXml + "					<recallTimeHrs>0.00</recallTimeHrs>";
            requestXml = requestXml + "					<schedDesc1>" + mst.SchedDescription1 + "</schedDesc1>";
            requestXml = requestXml + "					<schedFreq1>" + mst.SchedFreq1 + "</schedFreq1>";
            requestXml = requestXml + "					<schedInd700>" + mst.SchedInd + "</schedInd700>";
            requestXml = requestXml + "					<startMonth>" + mst.StartMonth + "</startMonth>";
            requestXml = requestXml + "					<startYear>" + mst.StartYear + "</startYear>";
            requestXml = requestXml + "					<workGroup>" + mst.WorkGroup + "</workGroup>";
            requestXml = requestXml + "					<autoReqInd>N</autoReqInd>";
            requestXml = requestXml + "					<statType1>" + mst.StatType1 + "</statType1>";
            requestXml = requestXml + "					<statType2>" + mst.StatType2 + "</statType2>";
            requestXml = requestXml + "					<nextSchStat>" + mst.NextSchedStat + "</nextSchStat>";
            requestXml = requestXml + "					<nextSchValue>" + mst.NextSchedValue + "</nextSchValue>";
            requestXml = requestXml + "					<statutoryFlg>N</statutoryFlg>";
            requestXml = requestXml + "					<hideSuppressed>Y</hideSuppressed>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msemst</application>";
            requestXml = requestXml + "	<applicationPage>read</applicationPage>";
            requestXml = requestXml + "	<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");
            responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
        }

        public static MaintSchedTskServiceModNextSchedReplyDTO ModNextSchedMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceModNextSchedRequestDTO()
            {
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                nextSchedDate = mst.NextSchedDate,
                nextSchedValSpecified = string.IsNullOrWhiteSpace(mst.NextSchedValue),
                nextStat = mst.NextSchedStat,
                nextSchedVal = Convert.ToDecimal(mst.NextSchedValue)
            };

            proxyEquip.Url = urlService + "/MaintSchedTskService";
            return proxyEquip.modNextSched(opContext, request);
        }

        public static MaintSchedTskServiceDeleteReplyDTO DeleteMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {
            var proxyEquip = new MaintSchedTskService.MaintSchedTskService { Url = urlService + "/MaintSchedTskService" };

            //actualizamos primero el indicador y eliminamos la frecuencia
            var requestUpdate = new MaintSchedTskServiceModifyRequestDTO
            {
                workGroup = mst.WorkGroup,
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedFreq1 = 0,
                schedFreq2 = 0,
                schedInd = "9",
                schedFreq1Specified = true,
                schedFreq2Specified = true,
                statType1 = "",
                statType2 = ""
            };


            var request = new MaintSchedTskServiceDeleteRequestDTO
            {
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask
            };

            proxyEquip.modify(opContext, requestUpdate);
            return proxyEquip.delete(opContext, request);
        }

        public static class Queries
        {
            public static string GetFetchMstListQuery(string dbReference, string dbLink, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo, string schedIndicator = null)
            {
                if (!string.IsNullOrWhiteSpace(districtCode))
                    districtCode = " AND MST.DSTRCT_CODE = '" + districtCode + "'";
                if (!string.IsNullOrWhiteSpace(workGroup))
                    workGroup = " AND MST.WORK_GROUP = '" + workGroup + "'";
                if (!string.IsNullOrWhiteSpace(equipmentNo))
                    equipmentNo = " AND MST.EQUIP_NO = '" + equipmentNo + "'";
                if (!string.IsNullOrWhiteSpace(compCode))
                    compCode = " AND MST.COMP_CODE = '" + compCode + "'";
                if (!string.IsNullOrWhiteSpace(compModCode))
                    compModCode = " AND MST.COMP_MOD_CODE = '" + compModCode + "'";
                if (!string.IsNullOrWhiteSpace(taskNo))
                    taskNo = " AND MST.MAINT_SCH_TASK = '" + taskNo + "'";

                //establecemos los parámetros de estado de orden
                schedIndicator = MyUtilities.GetCodeValue(schedIndicator);
                string statusIndicator;
                if (string.IsNullOrEmpty(schedIndicator))
                    statusIndicator = "";
                else if (schedIndicator == MstIndicatorList.Active)
                    statusIndicator = " AND MST.SCHED_IND_700 IN (" + MyUtilities.GetListInSeparator(MstIndicatorList.GetActiveIndicatorCodes(), ",", "'") + ")";
                else if (MstIndicatorList.GetIndicatorNames().Contains(schedIndicator))
                    statusIndicator = " AND MST.SCHED_IND_700 = '" + MstIndicatorList.GetIndicatorCode(schedIndicator) + "'";
                else
                    statusIndicator = "";

                var query = "" +
                               " SELECT" +
                               "     MST.DSTRCT_CODE, MST.WORK_GROUP, MST.REC_700_TYPE, MST.EQUIP_NO, EQ.ITEM_NAME_1 EQUIPMENT_DESC, MST.COMP_CODE, MST.COMP_MOD_CODE, MST.MAINT_SCH_TASK," +
                               "     MST.JOB_DESC_CODE, MST.SCHED_DESC_1, MST.SCHED_DESC_2, MST.ASSIGN_PERSON, MST.STD_JOB_NO, MST.AUTO_REQ_IND, MST.MS_HIST_FLG, MST.SCHED_IND_700," +
                               "     MST.SCHED_FREQ_1, MST.STAT_TYPE_1, MST.LAST_SCH_ST_1, MST.LAST_PERF_ST_1," +
                               "     MST.SCHED_FREQ_2, MST.STAT_TYPE_2, MST.LAST_SCH_ST_2, MST.LAST_PERF_ST_2," +
                               "     MST.LAST_SCH_DATE, MST.LAST_PERF_DATE, MST.NEXT_SCH_DATE, MST.NEXT_SCH_STAT, MST.NEXT_SCH_VALUE," +
                               "     MST.OCCURENCE_TYPE, MST.DAY_WEEK, MST.DAY_MONTH, DECODE(TRIM(MST.LAST_SCH_DATE),NULL,'',SUBSTR(MST.LAST_SCH_DATE,1,4) ) START_YEAR, DECODE(TRIM(MST.LAST_SCH_DATE),NULL,'',SUBSTR(MST.LAST_SCH_DATE,5,2) )START_MONTH, " +
                               "     MST.SHUTDOWN_TYPE , MST.SHUTDOWN_EQUIP, MST.SHUTDOWN_NO, MST.COND_MON_POS, MST.COND_MON_TYPE, MST.STATUTORY_FLG" +
                               " FROM" +
                               "     " + dbReference + ".MSF700" + dbLink + " MST LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON MST.EQUIP_NO = EQ.EQUIP_NO" +
                               " WHERE" +
                               districtCode +
                               workGroup +
                               equipmentNo +
                               compCode +
                               compModCode +
                               taskNo +
                               statusIndicator +
                               " ORDER BY MST.MAINT_SCH_TASK DESC";
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }
    }



}

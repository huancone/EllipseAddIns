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
                equipmentNo = mst.EquipmentNo,
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
            };

            //
            request.conAstSegFr = default(decimal);
            request.conAstSegFrSpecified = false;
            request.conAstSegTo = default(decimal);
            request.conAstSegToSpecified = false;

            var attributeList = new MaintSchedTskService.Attribute[3];
            attributeList[0] = new MaintSchedTskService.Attribute
            {
                name = "conAstSegFrNumeric",
                value = "0.000000"
            };
            attributeList[1] = new MaintSchedTskService.Attribute
            {
                name = "conAstSegToNumeric",
                value = "0.000000"
            };
            attributeList[2] = new MaintSchedTskService.Attribute
            {
                name = "outputConAstSeg",
                value = "0.000000"
            };
            //
            request.customAttributes = attributeList;
            return proxyEquip.create(opContext, request);
        }

        [Obsolete("ModifyMaintenanceScheduleTask is deprecated, please use ModifyMaintenanceScheduleTaskPost instead.")]
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
                startYear = mst.StartYear
            };
            //schedInd700
            request.conAstSegFrSpecified = false;
            request.conAstSegTo = default(decimal);
            request.conAstSegToSpecified = false;

            var isInSeries = "N";
            var isInSupressionSeries = "N";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("8") && mst.MaintenanceSchTask.Length == 4)
                isInSeries = "Y";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("9") && mst.MaintenanceSchTask.Length == 4)
                isInSupressionSeries = "Y";
            var attributeList = new List<MaintSchedTskService.Attribute>();
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "conAstSegFrNumeric",
                value = "0.000000"
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "conAstSegToNumeric",
                value = "0.000000"
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "outputConAstSeg",
                value = "0.000000"
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "isInSeries",
                value = isInSeries
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "isInSuppressionSeries",
                value = isInSupressionSeries
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "schedInd700",
                value = mst.SchedInd
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "schedInd700_lang",
                value = mst.SchedInd
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "hideSuppressed",
                value = "Y"
            });
            attributeList.Add(new MaintSchedTskService.Attribute
            {
                name = "jobType",
                value = "M"
            });

            request.customAttributes = attributeList.ToArray();
            
            proxyEquip.Url = urlService + "/MaintSchedTskService";
            return proxyEquip.modify(opContext, request);
        }

        public static void CreateMaintenanceScheduleTaskPost(EllipseFunctions ef, MaintenanceScheduleTask mst)
        {
            var responseDto = ef.InitiatePostConnection();
            if (responseDto.GotErrorMessages())
                throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));


            var isInSeries = "N";
            var isInSupressionSeries = "N";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("8") && mst.MaintenanceSchTask.Length == 4)
                isInSeries = "Y";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("9") && mst.MaintenanceSchTask.Length == 4)
                isInSupressionSeries = "Y";

            var requestXml = "";
            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "<actions>";
            requestXml = requestXml + "     <action>";
            requestXml = requestXml + "         <name>service</name>";
            requestXml = requestXml + "         <data>";
            requestXml = requestXml + "             <name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "             <operation>create</operation>";
            requestXml = requestXml + "             <className>mfui.actions.detail::CreateAction</className>";
            requestXml = requestXml + "             <returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "             <dto uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "                 <msHistFlg>Y</msHistFlg>";
            requestXml = requestXml + "                 <occurenceType>" + mst.OccurrenceType + "</occurenceType>";
            requestXml = requestXml + "                 <dayWeek>" + mst.DayOfWeek + "</dayWeek>";
            requestXml = requestXml + "                 <startMonth>" + mst.StartMonth + "</startMonth>";
            requestXml = requestXml + "                 <autoReqInd>N</autoReqInd>";
            requestXml = requestXml + "                 <equipRef>" + mst.EquipmentNo + "</equipRef>";
            requestXml = requestXml + "                 <equipGrpId>" + mst.EquipmentGrpId + "</equipGrpId>";
            requestXml = requestXml + "                 <compCode>" + mst.CompCode + "</compCode>";
            requestXml = requestXml + "                 <compModCode>" + mst.CompModCode + "</compModCode>";
            requestXml = requestXml + "                 <maintSchTask>" + mst.MaintenanceSchTask + "</maintSchTask>";
            requestXml = requestXml + "                 <displayInd>2</displayInd>";
            requestXml = requestXml + "                 <schedDesc1>" + mst.SchedDescription1 + "</schedDesc1>";
            requestXml = requestXml + "                 <schedDesc2>" + mst.SchedDescription2 + "</schedDesc2>";
            requestXml = requestXml + "                 <workGroup>" + mst.WorkGroup + "</workGroup>";
            requestXml = requestXml + "                 <assignPerson>" + mst.AssignPerson + "</assignPerson>";
            requestXml = requestXml + "                 <stdJobNo>" + mst.StdJobNo + "</stdJobNo>";
            requestXml = requestXml + "                 <jobDescCode>" + mst.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "                 <dstrctCode>" + mst.DistrictCode + "</dstrctCode>";
            requestXml = requestXml + "                 <schedInd700>" + mst.SchedInd + "</schedInd700>";
            requestXml = requestXml + "                 <statType1>" + mst.StatType1 + "</statType1>";
            requestXml = requestXml + "                 <lastSchStat1>" + (!string.IsNullOrWhiteSpace(mst.LastSchedStat1) ? Convert.ToString(mst.LastSchedStat1) : "").Trim() + "</lastSchStat1>";
            requestXml = requestXml + "                 <lastPerfStat1>" + (!string.IsNullOrWhiteSpace(mst.LastPerfStat1) ? Convert.ToString(mst.LastPerfStat1) : "").Trim() + "</lastPerfStat1>";
            requestXml = requestXml + "                 <dayMonth>" + mst.DayOfMonth + "</dayMonth>";
            requestXml = requestXml + "                 <startYear>" + mst.StartYear + "</startYear>"; ;
            requestXml = requestXml + "                 <schedFreq1>" + mst.SchedFreq1 + "</schedFreq1>";
            requestXml = requestXml + "                 <statType2>" + mst.StatType2 + "</statType2>";
            requestXml = requestXml + "                 <schedFreq2>" + mst.SchedFreq2 + "</schedFreq2>";
            requestXml = requestXml + "                 <lastSchStat2>" + (!string.IsNullOrWhiteSpace(mst.LastSchedStat2) ? Convert.ToString(mst.LastSchedStat2).Trim() : "") + "</lastSchStat2>";
            requestXml = requestXml + "                 <lastPerfStat2>" + (!string.IsNullOrWhiteSpace(mst.LastPerfStat2) ? Convert.ToString(mst.LastPerfStat2).Trim() : "") + "</lastPerfStat2>";
            requestXml = requestXml + "                 <lastSchDate>" + (!string.IsNullOrWhiteSpace(mst.LastSchedDate) ? Convert.ToString(mst.LastSchedDate).Trim() : "") + "</lastSchDate> ";
            requestXml = requestXml + "                 <lastPerfDate>" + (!string.IsNullOrWhiteSpace(mst.LastPerfDate) ? Convert.ToString(mst.LastPerfDate).Trim() : "") + "</lastPerfDate> ";
            requestXml = requestXml + "                 <nextSchDate>" + (!string.IsNullOrWhiteSpace(mst.NextSchedDate) ? Convert.ToString(mst.NextSchedDate).Trim() : "") + "</nextSchDate> ";
            requestXml = requestXml + "                 <nextSchStat>" + (!string.IsNullOrWhiteSpace(mst.NextSchedStat) ? Convert.ToString(mst.NextSchedStat).Trim() : "") + "</nextSchStat> ";
            requestXml = requestXml + "                 <nextSchValue>" + (!string.IsNullOrWhiteSpace(mst.NextSchedValue) ? Convert.ToString(mst.NextSchedValue).Trim() : "") + "</nextSchValue> ";
            requestXml = requestXml + "                 <statutoryFlg>" + mst.StatutoryFlg + "</statutoryFlg>";
            requestXml = requestXml + "                 <fixedScheduling>N</fixedScheduling>";
            requestXml = requestXml + "                 <allowMultiple>" + mst.AllowMultiple + "</allowMultiple>";
            requestXml = requestXml + "                 <isInSeries>" + isInSeries + "</isInSeries>";
            requestXml = requestXml + "                 <isInSuppressionSeries>" + isInSupressionSeries + "</isInSuppressionSeries>";
            requestXml = requestXml + "                 <hideSuppressed>Y</hideSuppressed>";
            requestXml = requestXml + "                 </dto>";
            requestXml = requestXml + "         </data>";
            requestXml = requestXml + "         <id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "     </action>";
            requestXml = requestXml + "</actions>";
            requestXml = requestXml + "<chains/>";
            requestXml = requestXml + "<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "<application>msemst</application>";
            requestXml = requestXml + "<applicationPage>create</applicationPage>";
            requestXml = requestXml + "<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";


            requestXml = requestXml.Replace("&", "&amp;");
            //requestXml = requestXml.Replace("\t", "");
            responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
        }

        public static void ModifyMaintenanceScheduleTaskPost(EllipseFunctions ef, MaintenanceScheduleTask mst)
        {
            var responseDto = ef.InitiatePostConnection();
            if (responseDto.GotErrorMessages())
                throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

            var isInSeries = "N";
            var isInSupressionSeries = "N";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("8") && mst.MaintenanceSchTask.Length == 4)
                isInSeries = "Y";
            if (mst != null && mst.MaintenanceSchTask != null && mst.MaintenanceSchTask.StartsWith("9") && mst.MaintenanceSchTask.Length == 4)
                isInSupressionSeries = "Y";

            var requestXml = "";

            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "				<operation>update</operation>";
            requestXml = requestXml + "				<className>mfui.actions.detail::UpdateAction</className>";
            requestXml = requestXml + "				<returnWarnings>false</returnWarnings>";
            requestXml = requestXml + "				<dto>";//    uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<assignPerson>JRODRIG4</assignPerson>";
            requestXml = requestXml + "					<complTextCde/>";
            requestXml = requestXml + "					<completeInstr/>";
            requestXml = requestXml + (mst.ConAstSegFr != null ? "					<conAstSegFr>" + mst.ConAstSegFr + "</conAstSegFr>" : null);
            requestXml = requestXml + (mst.ConAstSegFr != null ? "					<conAstSegFrNumeric>" + mst.ConAstSegFr + "</conAstSegFrNumeric>" : null);
            requestXml = requestXml + (mst.ConAstSegTo != null ? "					<conAstSegTo>" + mst .ConAstSegTo + "</conAstSegTo>" : null);
            requestXml = requestXml + (mst.ConAstSegTo != null ? "					<conAstSegToNumeric>" + mst.ConAstSegTo + "</conAstSegToNumeric>" : null);
            requestXml = requestXml + (mst.DayOfWeek != null ? "					<dayWeek>" + mst.DayOfWeek + "</dayWeek>" : null);
            requestXml = requestXml + (mst.DistrictCode != null ? "					<dstrctCode>" + mst.DistrictCode + "</dstrctCode>" : null);
            requestXml = requestXml + (mst.EquipmentNo != null ? "					<equipEntity>" + mst.EquipmentNo + "</equipEntity>" : null);
            requestXml = requestXml + (mst.EquipmentNo != null ? "					<equipNo>" + mst.EquipmentNo + "</equipNo>" : null);
            //requestXml = requestXml + "					<equipNoD></equipNoD>";
            //requestXml = requestXml + "					<equipNoD2></equipNoD2>";
            //requestXml = requestXml (mst.EquipmentNo != null ? + "					<equipRef>" + mst.EquipmentNo + "</equipRef> : null)";
            requestXml = requestXml + "					<isInSeries>" + isInSeries + "</isInSeries>";
            requestXml = requestXml + (mst.JobDescCode != null ? "					<jobDescCode>" + mst .JobDescCode + "</jobDescCode>" : null);
            requestXml = requestXml + (mst.LastPerfDate != null ? "					<lastPerfDate>" + mst.LastPerfDate + "</lastPerfDate>" : null);
            requestXml = requestXml + (mst.LastPerfStat1 != null ? "					<lastPerfStat1>" + mst.LastPerfStat1 + "</lastPerfStat1>" : null);
            requestXml = requestXml + (mst.LastPerfStat2 != null ? "					<lastPerfStat2>" + mst.LastPerfStat2 + "</lastPerfStat2>" : null);
            requestXml = requestXml + (mst.LastSchedStat1 != null ? "					<lastSchStat1>" + mst.LastSchedStat1 + "</lastSchStat1>" : null);
            requestXml = requestXml + (mst.LastSchedStat2 != null ? "					<lastSchStat2>" + mst.LastSchedStat2 + "</lastSchStat2>" : null);
            requestXml = requestXml + "					<linkedInd>N</linkedInd>";
            requestXml = requestXml + (mst.MaintenanceSchTask != null ? "					<maintSchTask>" + mst.MaintenanceSchTask + "</maintSchTask>" : null);
            requestXml = requestXml + (mst.MsHistFlag != null ? "					<msHistFlg>" + mst.MsHistFlag + "</msHistFlg>" : null);
            requestXml = requestXml + "					<nextSchInd> </nextSchInd>";
            requestXml = requestXml + "					<occurenceType> </occurenceType>";
            requestXml = requestXml + (mst.RecType != null ? "					<rec700Type>" + mst.RecType + "</rec700Type>" : null);
            requestXml = requestXml + "					<recallTimeHrs>0.00</recallTimeHrs>";
            requestXml = requestXml + "					<safetyInstr/>";
            
            requestXml = requestXml + (mst.SchedDescription1 != null ? "                 <schedDesc1>" + mst.SchedDescription1 + "</schedDesc1>" : null);
            requestXml = requestXml + (mst.SchedDescription2 != null ? "                 <schedDesc2>" + mst.SchedDescription2 + "</schedDesc2>" : null);
            requestXml = requestXml + (mst.SchedFreq1 != null ? "					<schedFreq1>" + mst.SchedFreq1 + "</schedFreq1>" : null);
            requestXml = requestXml + (mst.SchedFreq2 != null ? "					<schedFreq2>" + mst.SchedFreq2 + "</schedFreq2>" : null);
            requestXml = requestXml + (mst.SchedInd != null ? "					<schedInd700>" + mst.SchedInd + "</schedInd700>" : null);
            requestXml = requestXml + "					<segmentUom/>";
            requestXml = requestXml + (mst.StatType1 != null ? "					<statType1>" + mst.StatType1 + "</statType1>" : null);
            requestXml = requestXml + (mst.StatType2 != null ? "					<statType2>" + mst.StatType2 + "</statType2>" : null);
           
            requestXml = requestXml + "					<tskDurHours>0</tskDurHours>";
            requestXml = requestXml + "					<unitsRequired>0.00</unitsRequired>";
            requestXml = requestXml + (mst.WorkGroup != null ? "					<workGroup>" + mst.WorkGroup + "</workGroup>" : null);
            requestXml = requestXml + "					<workOrder/>";
            requestXml = requestXml + "					<jobType>M</jobType>";

            //requestXml = requestXml + "					<dstrctCode_desc></dstrctCode_desc>";
            //requestXml = requestXml + "					<dstrctCode_lang></dstrctCode_lang>";
            //requestXml = requestXml + "					<schedInd700_desc></schedInd700_desc>";
            requestXml = requestXml + (mst.SchedInd != null ? "					<schedInd700_lang>" + mst.SchedInd + "</schedInd700_lang>" : null);
            requestXml = requestXml + "					<outputConAstSeg/>";
            requestXml = requestXml + "					<parentChildType/>";
            requestXml = requestXml + "					<outputMeasureType/>";
            requestXml = requestXml + (mst.EquipmentGrpId != null ? "					<equipGrpId>" + mst.EquipmentGrpId + "</equipGrpId>" : null);
            requestXml = requestXml + (mst.CompCode != null ? "					<compCode>" + mst.CompCode + "</compCode>" : null);
            requestXml = requestXml + (mst.CompModCode != null ? "					<compModCode>" + mst.CompModCode + "</compModCode>" : null);
            requestXml = requestXml + "					<segmentUnitOfMeasure/>";
            requestXml = requestXml + "					<fromInKilometers/>";
            requestXml = requestXml + "					<toInKilometers/>";
            requestXml = requestXml + "					<fromInMilesYards/>";
            requestXml = requestXml + "					<toInMilesYards/>";
            requestXml = requestXml + "					<fromInMilesChains/>";
            requestXml = requestXml + "					<toInMilesChains/>";
            requestXml = requestXml + "					<linkParent/>";
            requestXml = requestXml + "					<linkId/>";
            requestXml = requestXml + "					<displayInd/>";
            requestXml = requestXml + "					<stdUnitOfWork/>";
            requestXml = requestXml + "					<unitsScale/>";
            requestXml = requestXml + "					<shutdownType/>";
            requestXml = requestXml + (mst.AutoRequisitionInd != null ? "					<autoReqInd>" + mst.AutoRequisitionInd + "</autoReqInd>" : null);
            requestXml = requestXml + (mst.DayOfMonth != null ? "					<dayMonth>" + mst.DayOfMonth + "</dayMonth>" : null);
            requestXml = requestXml + (mst.StartMonth != null ? "					<startMonth>" + mst.StartMonth + "</startMonth>" : null);
            requestXml = requestXml + (mst.StartYear != null ? "					<startYear>" + mst.StartYear + "</startYear>" : null);
            requestXml = requestXml + (mst.LastSchedDate != null ? "					<lastSchDate>" + mst.LastSchedDate + "</lastSchDate>" : null);
            requestXml = requestXml + (mst.NextSchedDate != null ? "					<nextSchDate>" + mst.NextSchedDate + "</nextSchDate>" : null);
            requestXml = requestXml + "					<scheduleType> </scheduleType>";
            requestXml = requestXml + (mst.NextSchedStat != null ? "					<nextSchStat>" + mst.NextSchedStat + "</nextSchStat>" : null);
            requestXml = requestXml + (mst.NextSchedValue != null ? "					<nextSchValue>" + mst.NextSchedValue + "</nextSchValue>" : null);
            requestXml = requestXml + "					<nextSchMeterValue/>";
            requestXml = requestXml + (mst.StatutoryFlg != null ? "					<statutoryFlg>" + mst.StatutoryFlg + "</statutoryFlg>" : null);
            requestXml = requestXml + "					<fixedScheduling>N</fixedScheduling>";
            requestXml = requestXml + (mst.AllowMultiple != null ? "					<allowMultiple>" + mst.AllowMultiple  + "</allowMultiple>" : null);
            //requestXml = requestXml + "					<stdJobWorkGroup></stdJobWorkGroup>";
            //requestXml = requestXml + "					<stdJobWorkOrderType></stdJobWorkOrderType>";
            //requestXml = requestXml + "					<stdJobMaintenanceType></stdJobMaintenanceType>";
            //requestXml = requestXml + "					<stdJobOriginatorPriority></stdJobOriginatorPriority>";
            //requestXml = requestXml + "					<stdJobCompCode/>";
            //requestXml = requestXml + "					<stdJobCompModCode/>";
            //requestXml = requestXml + "					<stdJobEstimatedDurationHours></stdJobEstimatedDurationHours>";
            //requestXml = requestXml + "					<stdJobDesc></stdJobDesc>";
            //requestXml = requestXml + "					<stdJobDstrctCode_desc></stdJobDstrctCode_desc>";
            //requestXml = requestXml + "					<stdJobDstrctCode_lang></stdJobDstrctCode_lang>";
            requestXml = requestXml + (mst.StdJobNo != null ? "					<stdJobDstrctCode>" + mst.DistrictCode + "</stdJobDstrctCode>" : null);
            requestXml = requestXml + (mst.StdJobNo != null ? "					<stdJobNo>" + mst.StdJobNo + "</stdJobNo>" : null);
            requestXml = requestXml + "					<stdUnitsRequired/>";
            requestXml = requestXml + (mst.CondMonPos != null ? "					<condMonPos>" + mst.CondMonPos + "</condMonPos>" : null);
            requestXml = requestXml + (mst.CondMonType != null ? "					<condMonType>" + mst.CondMonType + "</condMonType>" : null);
            requestXml = requestXml + "					<isInSuppressionSeries>" + isInSupressionSeries + "</isInSuppressionSeries>";
            requestXml = requestXml + "					<hideSuppressed>Y</hideSuppressed>";
            requestXml = requestXml + "				</dto>";
            //requestXml = requestXml + "				<requiredAttributes>";
            //requestXml = requestXml + "					<outputConAstSeg>true</outputConAstSeg>";
            //requestXml = requestXml + "					<parentChildType>true</parentChildType>";
            //requestXml = requestXml + "					<equipNo>true</equipNo>";
            //requestXml = requestXml + "					<outputMeasureType>true</outputMeasureType>";
            //requestXml = requestXml + "					<equipRef>true</equipRef>";
            //requestXml = requestXml + "					<equipNoD>true</equipNoD>";
            //requestXml = requestXml + "					<equipGrpId>true</equipGrpId>";
            //requestXml = requestXml + "					<compCode>true</compCode>";
            //requestXml = requestXml + "					<compModCode>true</compModCode>";
            //requestXml = requestXml + "					<maintSchTask>true</maintSchTask>";
            //requestXml = requestXml + "					<conAstSegFrNumeric>true</conAstSegFrNumeric>";
            //requestXml = requestXml + "					<conAstSegToNumeric>true</conAstSegToNumeric>";
            //requestXml = requestXml + "					<segmentUnitOfMeasure>true</segmentUnitOfMeasure>";
            //requestXml = requestXml + "					<fromInKilometers>true</fromInKilometers>";
            //requestXml = requestXml + "					<toInKilometers>true</toInKilometers>";
            //requestXml = requestXml + "					<fromInMilesYards>true</fromInMilesYards>";
            //requestXml = requestXml + "					<toInMilesYards>true</toInMilesYards>";
            //requestXml = requestXml + "					<fromInMilesChains>true</fromInMilesChains>";
            //requestXml = requestXml + "					<toInMilesChains>true</toInMilesChains>";
            //requestXml = requestXml + "					<linkParent>true</linkParent>";
            //requestXml = requestXml + "					<linkId>true</linkId>";
            //requestXml = requestXml + "					<displayInd>true</displayInd>";
            //requestXml = requestXml + "					<rec700Type>true</rec700Type>";
            //requestXml = requestXml + "					<schedDesc1>" + (!string.IsNullOrWhiteSpace(mst.SchedDescription1)).ToString().ToLower() + "</schedDesc1>";
            //requestXml = requestXml + "					<schedDesc2>" + (!string.IsNullOrWhiteSpace(mst.SchedDescription2)).ToString().ToLower() + "</schedDesc2>";
            //requestXml = requestXml + "					<workGroup>true</workGroup>";
            //requestXml = requestXml + "					<assignPerson>true</assignPerson>";
            //requestXml = requestXml + "					<stdJobNo>true</stdJobNo>";
            //requestXml = requestXml + "					<unitsRequired>true</unitsRequired>";
            //requestXml = requestXml + "					<stdUnitOfWork>true</stdUnitOfWork>";
            //requestXml = requestXml + "					<stdUnitsRequired>true</stdUnitsRequired>";
            //requestXml = requestXml + "					<unitsScale>true</unitsScale>";
            //requestXml = requestXml + "					<shutdownType>true</shutdownType>";
            //requestXml = requestXml + "					<autoReqInd>true</autoReqInd>";
            //requestXml = requestXml + "					<jobDescCode>true</jobDescCode>";
            //requestXml = requestXml + "					<msHistFlg>true</msHistFlg>";
            //requestXml = requestXml + "					<dstrctCode>true</dstrctCode>";
            //requestXml = requestXml + "					<schedInd700>true</schedInd700>";
            //requestXml = requestXml + "					<statType1>true</statType1>";
            //requestXml = requestXml + "					<lastSchStat1>true</lastSchStat1>";
            //requestXml = requestXml + "					<lastPerfStat1>true</lastPerfStat1>";
            //requestXml = requestXml + "					<occurenceType>true</occurenceType>";
            //requestXml = requestXml + "					<dayWeek>true</dayWeek>";
            //requestXml = requestXml + "					<dayMonth>true</dayMonth>";
            //requestXml = requestXml + "					<startMonth>true</startMonth>";
            //requestXml = requestXml + "					<startYear>true</startYear>";
            //requestXml = requestXml + "					<schedFreq1>true</schedFreq1>";
            //requestXml = requestXml + "					<statType2>true</statType2>";
            //requestXml = requestXml + "					<schedFreq2>true</schedFreq2>";
            //requestXml = requestXml + "					<lastSchStat2>true</lastSchStat2>";
            //requestXml = requestXml + "					<lastPerfStat2>true</lastPerfStat2>";
            //requestXml = requestXml + "					<lastSchDate>true</lastSchDate>";
            //requestXml = requestXml + "					<lastPerfDate>true</lastPerfDate>";
            //requestXml = requestXml + "					<nextSchDate>true</nextSchDate>";
            //requestXml = requestXml + "					<scheduleType>true</scheduleType>";
            //requestXml = requestXml + "					<nextSchStat>true</nextSchStat>";
            //requestXml = requestXml + "					<nextSchValue>true</nextSchValue>";
            //requestXml = requestXml + "					<nextSchMeterValue>true</nextSchMeterValue>";
            //requestXml = requestXml + "					<nextSchInd>true</nextSchInd>";
            //requestXml = requestXml + "					<recallTimeHrs>true</recallTimeHrs>";
            //requestXml = requestXml + "					<statutoryFlg>true</statutoryFlg>";
            //requestXml = requestXml + "					<fixedScheduling>true</fixedScheduling>";
            //requestXml = requestXml + "					<allowMultiple>true</allowMultiple>";
            //requestXml = requestXml + "					<stdJobDstrctCode>true</stdJobDstrctCode>";
            //requestXml = requestXml + "					<stdJobDesc>true</stdJobDesc>";
            //requestXml = requestXml + "					<stdJobWorkGroup>true</stdJobWorkGroup>";
            //requestXml = requestXml + "					<stdJobWorkOrderType>true</stdJobWorkOrderType>";
            //requestXml = requestXml + "					<stdJobMaintenanceType>true</stdJobMaintenanceType>";
            //requestXml = requestXml + "					<stdJobOriginatorPriority>true</stdJobOriginatorPriority>";
            //requestXml = requestXml + "					<stdJobCompCode>true</stdJobCompCode>";
            //requestXml = requestXml + "					<stdJobCompModCode>true</stdJobCompModCode>";
            //requestXml = requestXml + "					<stdJobEstimatedDurationHours>true</stdJobEstimatedDurationHours>";
            //requestXml = requestXml + "					<condMonPos>true</condMonPos>";
            //requestXml = requestXml + "					<condMonType>true</condMonType>";
            //requestXml = requestXml + "					<isInSeries>true</isInSeries>";
            //requestXml = requestXml + "					<isInSuppressionSeries>true</isInSuppressionSeries>";
            //requestXml = requestXml + "					<hideSuppressed>true</hideSuppressed>";
            //requestXml = requestXml + "				</requiredAttributes>";
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

            var mstService = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceModNextSchedRequestDTO()
            {
                equipmentNo = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                nextSchedDate = mst.NextSchedDate,
                nextSchedValSpecified = string.IsNullOrWhiteSpace(mst.NextSchedValue),
                nextStat = mst.NextSchedStat,
                nextSchedVal = string.IsNullOrWhiteSpace(mst.NextSchedValue) ? 0 : Convert.ToDecimal(mst.NextSchedValue)
            };
            
            mstService.Url = urlService + "/MaintSchedTskService";
            return mstService.modNextSched(opContext, request);
        }

        public static MaintSchedTskServiceDeleteReplyDTO DeleteMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {
            var proxyEquip = new MaintSchedTskService.MaintSchedTskService { Url = urlService + "/MaintSchedTskService" };

            //actualizamos primero el indicador y eliminamos la frecuencia
            var requestUpdate = new MaintSchedTskServiceModifyRequestDTO
            {
                workGroup = mst.WorkGroup,
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentNo = mst.EquipmentNo,
                
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedFreq1 = 0,
                schedFreq2 = 0,
                schedInd = "9",
                schedFreq1Specified = true,
                schedFreq2Specified = true,
                statType1 = "",
                statType2 = "",
                conAstSegFrSpecified = false,
                conAstSegToSpecified = false
            };

            var attributeList = new MaintSchedTskService.Attribute[2];
            attributeList[0] = new MaintSchedTskService.Attribute
            {
                name = "conAstSegFrNumeric",
                value = "0"
            };
            attributeList[1] = new MaintSchedTskService.Attribute
            {
                name = "conAstSegToNumeric",
                value = "0"
            };

            requestUpdate.customAttributes = attributeList;

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

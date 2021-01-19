using SharedClassLibrary.Utilities;

namespace EllipseMaintSchedTaskClassLibrary
{
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

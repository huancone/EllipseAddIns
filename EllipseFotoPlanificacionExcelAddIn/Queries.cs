using SharedClassLibrary.Utilities;
using EllipseJobsClassLibrary;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public static partial class Queries
    {
        public static string InsertSigmanItemQuery(string dbReference, string dbLink, PlannerItem item)
        {
            //Se usa en la validación la funcíon NVL para comparación segura de nulls. Para Oracle NULL != NULL
            var query = " MERGE INTO SIGMAN.CUMPLIMIENTO_PLAN CP USING " +
                        " (SELECT " +
                        "    '" + item.Period + "' PERIOD, '" + item.WorkGroup + "' WORK_GROUP, '" + item.EquipNo + "' EQUIP_NO, '" + item.CompCode + "' COMP_CODE, '" + item.CompModCode + "' COMP_MOD_CODE, " +
                        "    '" + item.WorkOrder + "' WORK_ORDER, '" + item.MaintSchedTask + "' MAINT_SCH_TASK, '" + item.RaisedDate + "' RAISED_DATE, " +
                        "    '" + item.PlanDate + "' PLAN_STR_DATE, '" + item.NextSchedDate + "' NEXT_SCH_DATE, '" + item.LastPerfDate + "' LAST_PERF_DATE, '" + item.DurationHours + "' DURATION_HOURS, '" + item.LabourHours + "' LABOUR_HOURS, " +
                        "    '" + item.LastModUser + "' LAST_MOD_USER, '" + item.LastModItemDate + "' LAST_MOD_DATE, '" + item.RecordStatus + "' RECORD_STATUS FROM DUAL) REG ON (" +
                        "    NVL(TRIM(CP.PERIOD), ' ') = NVL(TRIM(REG.PERIOD), ' ') AND NVL(TRIM(CP.WORK_GROUP), ' ') = NVL(TRIM(REG.WORK_GROUP), ' ') " +
                        "    AND NVL(TRIM(CP.EQUIP_NO), ' ') = NVL(TRIM(REG.EQUIP_NO), ' ') AND NVL(TRIM(CP.COMP_CODE), ' ') = NVL(TRIM(REG.COMP_CODE), ' ') AND NVL(TRIM(CP.COMP_MOD_CODE), ' ') = NVL(TRIM(REG.COMP_MOD_CODE), ' ') " + "" +
                        "    AND NVL(TRIM(CP.WORK_ORDER), ' ') = NVL(TRIM(REG.WORK_ORDER), ' ') AND NVL(TRIM(CP.MAINT_SCH_TASK), ' ') = NVL(TRIM(REG.MAINT_SCH_TASK), ' ') " +
                        "    AND NVL(TRIM(CP.RAISED_DATE), ' ') = NVL(TRIM(REG.RAISED_DATE), ' ') AND NVL(TRIM(CP.RECORD_STATUS), '0') = NVL(TRIM(REG.RECORD_STATUS), '0')) " +
                        " WHEN MATCHED THEN UPDATE SET CP.PLAN_STR_DATE = TRIM(REG.PLAN_STR_DATE), CP.NEXT_SCH_DATE = TRIM(REG.NEXT_SCH_DATE), CP.LAST_PERF_DATE = TRIM(REG.LAST_PERF_DATE), CP.DURATION_HOURS = NVL(TRIM(REG.DURATION_HOURS), '0'), CP.LABOUR_HOURS = NVL(TRIM(REG.LABOUR_HOURS), '0')" +
                        " WHEN NOT MATCHED THEN INSERT " +
                        "   (PERIOD, WORK_GROUP, EQUIP_NO, COMP_CODE, COMP_MOD_CODE, " +
                        "   WORK_ORDER, MAINT_SCH_TASK, RAISED_DATE, " +
                        "   PLAN_STR_DATE, NEXT_SCH_DATE, LAST_PERF_DATE, DURATION_HOURS, LABOUR_HOURS, " +
                        "   LAST_MOD_USER, LAST_MOD_DATE, RECORD_STATUS)" +
                        " VALUES " +
                        "   (TRIM(REG.PERIOD), TRIM(REG.WORK_GROUP), TRIM(REG.EQUIP_NO), TRIM(REG.COMP_CODE), TRIM(REG.COMP_MOD_CODE), " +
                        "    TRIM(REG.WORK_ORDER), TRIM(REG.MAINT_SCH_TASK), TRIM(REG.RAISED_DATE), " +
                        "    TRIM(REG.PLAN_STR_DATE), TRIM(REG.NEXT_SCH_DATE), TRIM(REG.LAST_PERF_DATE), NVL(TRIM(REG.DURATION_HOURS), '0'), NVL(TRIM(REG.LABOUR_HOURS), '0'), " +
                        "    TRIM(REG.LAST_MOD_USER), TRIM(REG.LAST_MOD_DATE), NVL(TRIM(REG.RECORD_STATUS), '0'))";

                        query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetFetchSigmanPhotoQuery(string dbReference, string dbLink, string searchEntity, string dateType, string startDate, string finishDate, string [] workGroups)
        {
            var entityParam = "";
            switch (searchEntity)
            {
                case "Work Orders Only":
                    entityParam = " AND TRIM(PL.WORK_ORDER) IS NOT NULL";
                    break;
                case "MST Forecast Only":
                    entityParam = " AND TRIM(PL.WORK_ORDER) IS NULL";
                    break;
                case "Work Orders and MST Forecast":
                    entityParam = "";
                    break;
            }
            var workGroupParam = "";
            if (workGroups != null && workGroups.Length > 0)
            {
                if(workGroups.Length == 1)
                    workGroupParam = " AND PL.WORK_GROUP = '" + workGroups[0] + "' ";
                else
                {
                    workGroupParam = " AND PL.WORK_GROUP IN (";
                    foreach (var group in workGroups)
                        workGroupParam += "'" + group + "',";
                    workGroupParam = workGroupParam.Substring(0, workGroupParam.Length - 1) + ") ";
                }
            }

            var dateParam = "";
            if (dateType.Equals(SearchDateCriteriaType.Period.Value))
            {
                dateParam = " AND PL.PERIOD";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate.Substring(0, 6) + "' AND PL.PERIOD <= '" + finishDate.Substring(0, 6) + "'";
            }
            else
            {
                dateParam = " AND PL.PLAN_STR_DATE";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate + "' AND PL.PLAN_STR_DATE <= '" + finishDate + "'";

            }

            //escribimos el query
            var query = "" +
                        " SELECT" +
                        " 	PL.PERIOD, PL.WORK_GROUP," +
                        " 	PL.EQUIP_NO, PL.COMP_CODE, PL.COMP_MOD_CODE, PL.WORK_ORDER, PL.MAINT_SCH_TASK," +
                        " 	PL.RAISED_DATE, PL.PLAN_STR_DATE, PL.NEXT_SCH_DATE, PL.LAST_PERF_DATE, PL.DURATION_HOURS," +
                        " 	PL.LABOUR_HOURS, PL.LAST_MOD_USER, PL.LAST_MOD_DATE, PL.RECORD_STATUS" +
                        " FROM CUMPLIMIENTO_PLAN PL" +
                        " WHERE" +
                        " PL.RECORD_STATUS = '1'" +
                        entityParam +
                        dateParam +
                        workGroupParam;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetDeleteTaskQuery(string dbReference, string dbLink, string dateType, string startDate, string finishDate, string[] workGroups)
        {
            var workGroupParam = "";
            if (workGroups != null && workGroups.Length > 0)
            {
                if (workGroups.Length == 1)
                    workGroupParam = " AND WORK_GROUP = '" + workGroups[0] + "' ";
                else
                {
                    workGroupParam = " AND WORK_GROUP IN (";
                    foreach (var group in workGroups)
                        workGroupParam += "'" + group + "',";
                    workGroupParam = workGroupParam.Substring(0, workGroupParam.Length - 1) + ") ";
                }
            }

            var dateParam = "";
            if (dateType.Equals(SearchDateCriteriaType.Period.Value))
            {
                dateParam = " AND PERIOD";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate.Substring(0, 6) + "' AND PERIOD <= '" + finishDate.Substring(0, 6) + "'";
            }
            else
            {
                dateParam = " AND PLAN_STR_DATE";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate + "' AND PLAN_STR_DATE <= '" + finishDate + "'";

            }

            var query = "DELETE FROM SIGMAN.CUMPLIMIENTO_PLAN WHERE" +
                        dateParam +
                        workGroupParam;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            return query;
        }
        public static string GetDisableTaskStatusQuery(string dbReference, string dbLink, string dateType, string startDate, string finishDate, string[] workGroups)
        {
            var workGroupParam = "";
            if (workGroups != null && workGroups.Length > 0)
            {
                if (workGroups.Length == 1)
                    workGroupParam = " AND WORK_GROUP = '" + workGroups[0] + "' ";
                else
                {
                    workGroupParam = " AND WORK_GROUP IN (";
                    foreach (var group in workGroups)
                        workGroupParam += "'" + group + "',";
                    workGroupParam = workGroupParam.Substring(0, workGroupParam.Length - 1) + ") ";
                }
            }

            var dateParam = "";
            if (dateType.Equals(SearchDateCriteriaType.Period.Value))
            {
                dateParam = " AND PERIOD";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate.Substring(0, 6) + "' AND PERIOD <= '" + finishDate.Substring(0, 6) + "'";
            }
            else
            {
                dateParam = " AND PLAN_STR_DATE";
                if (string.IsNullOrWhiteSpace(finishDate))
                    dateParam = dateParam + " = '" + startDate.Substring(0, 6) + "'";
                else
                    dateParam = dateParam + " >= '" + startDate + "' AND PLAN_STR_DATE <= '" + finishDate + "'";

            }

            var query = "UPDATE SIGMAN.CUMPLIMIENTO_PLAN SET RECORD_STATUS = '0' WHERE" +
                        dateParam +
                        workGroupParam;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            return query;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;
using System.Threading;
using System.Windows.Forms.VisualStyles;
using EllipseJobsClassLibrary;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public static partial class Queries
    {
        public static string InsertSigmanItemQuery(string dbReference, string dbLink, PlannerItem item)
        {
            var query = "INSERT INTO SIGMAN.CUMPLIMIENTO_PLAN " +
                             "   (PERIOD, WORK_GROUP, EQUIP_NO, COMP_CODE, COMP_MOD_CODE, "+
                             "   WORK_ORDER, MAINT_SCH_TASK, RAISED_DATE, " +
                             "   PLAN_STR_DATE, NEXT_SCH_DATE, LAST_PERF_DATE, DURATION_HOURS, LABOUR_HOURS, " +
                             "   LAST_MOD_USER, LAST_MOD_DATE, RECORD_STATUS)" +
                             " VALUES "+
                             "   ('" + item.Period + "', '" + item.WorkGroup + "', '" + item.EquipNo + "', '" + item.CompCode + "', '" + item.CompModCode + "', " +
                             "    '" + item.WorkOrder + "', '" + item.MaintSchedTask + "', '" + item.RaisedDate + "', " +
                             "    '" + item.PlanDate + "', '" + item.NextSchedDate + "', '" + item.LastPerfDate + "', '" + item.DurationHours + "', '" + item.LabourHours + "', " +
                             "    '" + item.LastModUser + "', '" + item.LastModItemDate + "', '" + item.RecordStatus + "')";


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
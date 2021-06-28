using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;

namespace EllipseStandardJobsClassLibrary
{
    internal static class Queries
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
                "   STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10," +
                "   STD.PAPER_HIST" +
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
            string districtCodeParam = null;
            if (string.IsNullOrEmpty(districtCode))
                districtCodeParam = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCodeParam = " IN ('" + districtCode + "')";


            //establecemos los parámetrode de distrito
            if (string.IsNullOrEmpty(workGroup))
                workGroup = " IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
            else
                workGroup = " IN ('" + workGroup + "')";



            var query = "" +
                           " SELECT * FROM (WITH SOT AS (SELECT STD_JOB_NO, MAX(USO_OTS) USO_OTS, MAX(ULTIMO_USO) ULTIMO_USO, MAX(USO_MSTS) USO_MSTS FROM" +
                           "    (SELECT STD_JOB_NO, COUNT(*) USO_OTS, MAX(CREATION_DATE) ULTIMO_USO, 0 AS USO_MSTS FROM " + dbReference + ".MSF620" + dbLink + " WHERE DSTRCT_CODE " + districtCodeParam + " AND WORK_GROUP " + workGroup + " GROUP BY DSTRCT_CODE, STD_JOB_NO" +
                           "    UNION ALL SELECT STD_JOB_NO, 0 AS USO_OTS, MAX(LAST_SCH_DATE) AS ULTIMO_USO, COUNT(*) USO_MSTS FROM " + dbReference + ".MSF700" + dbLink + " WHERE DSTRCT_CODE " + districtCodeParam + " AND WORK_GROUP " + workGroup + " GROUP BY DSTRCT_CODE, STD_JOB_NO)" +
                           "    GROUP BY STD_JOB_NO)" +
                           "    SELECT" +
                           "    STD.DSTRCT_CODE, STD.WORK_GROUP, STD.STD_JOB_NO, STD.STD_JOB_DESC, STD.SJ_ACTIVE_STATUS, STD.ORIGINATOR_ID, STD.ORIG_PRIORITY," +
                           "    STD.WO_TYPE, STD.MAINT_TYPE, STD.ASSIGN_PERSON, STD.COMP_CODE, STD.COMP_MOD_CODE, STD.UNIT_OF_WORK, STD.UNITS_REQUIRED," +
                           "    STD.ACCOUNT_CODE, STD.REALL_ACCT_CDE, STD.PROJECT_NO," +
                           "    STD.CALC_DUR_HRS_SW, STD.EST_DUR_HRS, STD.RES_UPDATE_FLAG, STD.EST_LAB_HRS, STD.EST_LAB_COST, STD.MAT_UPDATE_FLAG, STD.EST_MAT_COST, STD.EQUIP_UPDATE_FLAG, STD.EST_EQUIP_COST, STD.EST_OTHER_COST," +
                           "    STD.CALC_LAB_HRS, STD.CALC_LAB_COST, STD.CALC_MAT_COST, STD.CALC_EQUIP_COST," +
                           "    STD.NO_OF_TASKS, SOT.USO_OTS, SOT.USO_MSTS, SOT.ULTIMO_USO," +
                           "    STD.WO_JOB_CODEX1, STD.WO_JOB_CODEX2, STD.WO_JOB_CODEX3, STD.WO_JOB_CODEX4, STD.WO_JOB_CODEX5," +
                           "    STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10," +
                           "    STD.PAPER_HIST" +
                           " FROM" +
                           "    " + dbReference + ".msf690" + dbLink + " STD LEFT JOIN SOT ON STD.STD_JOB_NO = SOT.STD_JOB_NO" +
                           " WHERE" +
                           " STD.WORK_GROUP " + workGroup +
                           " AND  STD.DSTRCT_CODE " + districtCodeParam +
                           " ORDER BY STD.WORK_GROUP, STD.STD_JOB_NO)";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchStandardQuery(string dbReference, string dbLink, string districtCode, string workGroup, string standardJob)
        {
            string districtCodeParam = null;
            if (string.IsNullOrEmpty(districtCode))
                districtCodeParam = " IN ('ICOR')";
            else
                districtCodeParam = " IN ('" + districtCode + "')";

            var query = "" +
                           " SELECT * FROM (WITH SOT AS (SELECT STD_JOB_NO, MAX(USO_OTS) USO_OTS, MAX(ULTIMO_USO) ULTIMO_USO, MAX(USO_MSTS) USO_MSTS FROM" +
                           "    (SELECT STD_JOB_NO, COUNT(*) USO_OTS, MAX(CREATION_DATE) ULTIMO_USO, 0 AS USO_MSTS FROM " + dbReference + ".MSF620" + dbLink + " WHERE DSTRCT_CODE " + districtCodeParam + " AND WORK_GROUP = '" + workGroup + "' GROUP BY DSTRCT_CODE, STD_JOB_NO" +
                           "    UNION ALL SELECT STD_JOB_NO, 0 AS USO_OTS, MAX(LAST_SCH_DATE) AS ULTIMO_USO, COUNT(*) USO_MSTS FROM " + dbReference + ".MSF700" + dbLink + " WHERE DSTRCT_CODE " + districtCodeParam + " AND WORK_GROUP = '" + workGroup + "' GROUP BY DSTRCT_CODE, STD_JOB_NO)" +
                           "    GROUP BY STD_JOB_NO)" +
                           "    SELECT" +
                           "    STD.DSTRCT_CODE, STD.WORK_GROUP, STD.STD_JOB_NO, STD.STD_JOB_DESC, STD.SJ_ACTIVE_STATUS, STD.ORIGINATOR_ID, STD.ORIG_PRIORITY," +
                           "    STD.WO_TYPE, STD.MAINT_TYPE, STD.ASSIGN_PERSON, STD.COMP_CODE, STD.COMP_MOD_CODE, STD.UNIT_OF_WORK, STD.UNITS_REQUIRED," +
                           "    STD.ACCOUNT_CODE, STD.REALL_ACCT_CDE, STD.PROJECT_NO," +
                           "    STD.CALC_DUR_HRS_SW, STD.EST_DUR_HRS, STD.RES_UPDATE_FLAG, STD.EST_LAB_HRS, STD.EST_LAB_COST, STD.MAT_UPDATE_FLAG, STD.EST_MAT_COST, STD.EQUIP_UPDATE_FLAG, STD.EST_EQUIP_COST, STD.EST_OTHER_COST," +
                           "    STD.CALC_LAB_HRS, STD.CALC_LAB_COST, STD.CALC_MAT_COST, STD.CALC_EQUIP_COST," +
                           "    STD.NO_OF_TASKS, SOT.USO_OTS, SOT.USO_MSTS, SOT.ULTIMO_USO," +
                           "    STD.WO_JOB_CODEX1, STD.WO_JOB_CODEX2, STD.WO_JOB_CODEX3, STD.WO_JOB_CODEX4, STD.WO_JOB_CODEX5," +
                           "    STD.WO_JOB_CODEX6, STD.WO_JOB_CODEX7, STD.WO_JOB_CODEX8, STD.WO_JOB_CODEX9, STD.WO_JOB_CODEX10," +
                           "   STD.PAPER_HIST" +
                           " FROM" +
                           "    " + dbReference + ".msf690" + dbLink + " STD LEFT JOIN SOT ON STD.STD_JOB_NO = SOT.STD_JOB_NO" +
                           " WHERE" +
                           " STD.WORK_GROUP = '" + workGroup + "'" +
                           " AND STD.DSTRCT_CODE " + districtCodeParam + "" +
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

}

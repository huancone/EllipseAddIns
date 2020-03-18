using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseWorkOrdersClassLibrary
{
    public static partial class Queries
    {

        public static string GetFetchWorkOrderTasksQuery(string dbReference, string dbLink, string districtCode, string workOrder, string woTaskNo)
        {
            var query = "" +
                        "SELECT " +
                        "	WO.DSTRCT_CODE, " +
                        "	WO.WORK_GROUP, " +
                        "	WO.WORK_ORDER, " +
                        "	WO.WO_DESC, " +
                        "	WT.WO_TASK_NO, " +
                        "	WT.WO_TASK_DESC, " +
                        "	WT.JOB_DESC_CODE, " +
                        "	WT.SAFETY_INSTR, " +
                        "	WT.COMPLETE_INSTR, " +
                        "	WT.COMPL_TEXT_CDE, " +
                        "	WT.ASSIGN_PERSON, " +
                        "	WT.EST_MACH_HRS, " +
                        "	WT.TSK_DUR_HOURS, " +
                        "	WT.PLAN_STR_DATE, " +
                        "	WT.PLAN_FIN_DATE, " +
                        "	WT.PLAN_STR_TIME, " +
                        "	WT.PLAN_FIN_TIME, " +
                        "	WT.EQUIP_GRP_ID, " +
                        "	WT.APL_TYPE, " +
                        "	WT.COMP_CODE, " +
                        "	WT.COMP_MOD_CODE, " +
                        "	WT.APL_SEQ_NO, " +
                        "	WT.TASK_STATUS_M, " +
                        "	WT.CLOSED_STATUS, " +
                        "	WT.COMPLETED_CODE, " +
                        "	WT.COMPLETED_BY, " +
                        "	WT.CLOSED_DT, " +
                        "	( " +
                        "		SELECT " +
                        "			COUNT(*) LABOR " +
                        "		FROM " +
                        "			ELLIPSE.MSF623 TSK " +
                        "			INNER JOIN ELLIPSE.MSF735 RS " +
                        "			ON RS.KEY_735_ID     = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                        "			   AND RS.REC_735_TYPE   = 'WT' " +
                        "		WHERE " +
                        "			TSK.WORK_ORDER = WO.WORK_ORDER " +
                        "			AND   TSK.WO_TASK_NO = WT.WO_TASK_NO " +
                        "	)NO_REC_LABOR, " +
                        "	( " +
                        "		SELECT " +
                        "			COUNT(*) MATER " +
                        "		FROM " +
                        "			ELLIPSE.MSF623 TSK " +
                        "			INNER JOIN ELLIPSE.MSF734 RS " +
                        "			ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                        "			   AND RS.CLASS_TYPE   = 'WT' " +
                        "		WHERE " +
                        "			TSK.WORK_ORDER = WO.WORK_ORDER " +
                        "			AND   TSK.WO_TASK_NO = WT.WO_TASK_NO " +
                        "	)NO_REC_MATERIAL " +
                        "FROM " +
                        "	" + dbReference + ".MSF620" + dbLink + " WO " +
                        "	INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT " +
                        "	ON WO.WORK_ORDER    = WT.WORK_ORDER " +
                        "	   AND WO.DSTRCT_CODE   = WT.DSTRCT_CODE " +
                        "	   AND WO.WORK_ORDER    = '" + workOrder + "'" +
                        "	   AND WO.DSTRCT_CODE   = '" + districtCode + "'";
            if (woTaskNo != "")
            {
                query = query + " AND WT.WO_TASK_NO   = " + woTaskNo + " ";
            }

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchWoTaskRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string reqType, string taskNo)
        {
            if (reqType.Equals(RequirementType.Labour.Key))
            {
                return GetFetchWoTaskLabourRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.Material.Key))
            {
                return GetFetchWoTaskMaterialRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.Equipment.Key))
            {
                return GetFetchWoTaskEquipmentRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.All.Key))
            {
                var labourSql = GetFetchWoTaskLabourRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
                var materialSql = GetFetchWoTaskMaterialRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
                var equipmentSql = GetFetchWoTaskEquipmentRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);

                return labourSql + " UNION ALL " + materialSql + " UNION ALL " + equipmentSql;
            }
            return null;
        }

        public static string GetFetchWoTaskLabourRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                        " SELECT " +
                        "	'" + RequirementType.Labour.Key + "' REQ_TYPE, " +
                        "	COALESCE(TSK.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                        "	COALESCE(TSK.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                        "	COALESCE(TSK.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                        "	COALESCE(TSK.WO_TASK_NO, TRR.WO_TASK_NO) WO_TASK_NO, " +
                        "	COALESCE(TSK.WO_TASK_DESC, TRR.WO_TASK_DESC) WO_TASK_DESC, " +
                        "	'N/A' SEQ_NO, " +
                        "	COALESCE(RS.RESOURCE_TYPE, TRR.RES_CODE) RES_CODE, " +
                        "	TO_NUMBER(RS.CREW_SIZE) EST_SIZE, " +
                        "	RS.EST_RESRCE_HRS UNITS_QTY, " +
                        "   TRR.ACT_RESRCE_HRS REAL_QTY, " + //Real Value
                        "	COALESCE(TT.TABLE_DESC, TRR.RES_DESC) RES_DESC, " +
                        "	'HR' UNITS, " +
                        "   1 SHARED_TASKS " +
                        " FROM " +
                        "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                        "	INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS " +
                        "	ON RS.KEY_735_ID     = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                        "	   AND RS.REC_735_TYPE   = 'WT' " +
                        "	INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT " +
                        "	ON TT.TABLE_CODE   = RS.RESOURCE_TYPE " +
                        "	   AND TT.TABLE_TYPE   = 'TT' " +
                        //Real Calculation
                        " FULL JOIN( " +
                        "   SELECT TR.DSTRCT_CODE, " +
                        "     WT.WORK_GROUP, " +
                        "     TX.WORK_ORDER, " +
                        "     TX.WO_TASK_NO, " +
                        "     WT.WO_TASK_DESC," + 
                        "     TR.RESOURCE_TYPE RES_CODE, " +
                        "     LTT.TABLE_DESC RES_DESC, " +
                        "     SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS " +
                        "   FROM " + dbReference + ".MSFX99" + dbLink + " TX " +
                        "     INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR " +
                        "       ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                        "       AND TR.WORK_ORDER = TX.WORK_ORDER " +
                        "       AND TR.USERNO = TX.USERNO " +
                        "       AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                        "       AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                        "       AND TR.REC900_TYPE = TX.REC900_TYPE " +
                        "       AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                        "       AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                        "     INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT " +
                        "       ON TX.DSTRCT_CODE = WT.DSTRCT_CODE " +
                        "       AND TX.WORK_ORDER = WT.WORK_ORDER " +
                        "       AND TX.WO_TASK_NO = WT.WO_TASK_NO " +
                        "     INNER JOIN ELLIPSE.MSF010 LTT " +
                        "       ON LTT.TABLE_CODE = TR.RESOURCE_TYPE " + 
                        "       AND LTT.TABLE_TYPE = 'TT' "+
                        "   WHERE TX.DSTRCT_CODE = '" + districtCode + "' " +
                        "     AND TX.WORK_ORDER = '" + workOrder + "' " +
                        "     AND TX.WO_TASK_NO = '" + taskNo + "' " +
                        "     AND TX.REC900_TYPE = 'L' " +
                        "   GROUP BY TR.DSTRCT_CODE, " +
                        "     WT.WORK_GROUP, " +
                        "     TX.WORK_ORDER, " +
                        "     TX.WO_TASK_NO, " +
                        "     WT.WO_TASK_DESC, " +
                        "     TR.RESOURCE_TYPE, " +
                        "     LTT.TABLE_DESC " +
                        " ) TRR ON " +
                        "     TSK.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                        "     AND TSK.WORK_GROUP = TRR.WORK_GROUP " +
                        "     AND TSK.WORK_ORDER = TRR.WORK_ORDER " +
                        "     AND TSK.WO_TASK_NO = TRR.WO_TASK_NO " +
                        "     AND RS.RESOURCE_TYPE = TRR.RES_CODE " +
                        //End Real Calculation
                        "WHERE " +
                        "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                        "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                        "	AND   TSK.WO_TASK_NO = '" + taskNo + "' ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchWoTaskMaterialRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                        "SELECT " +
                        "	'" + RequirementType.Material.Key + "' REQ_TYPE, " +
                        "	COALESCE(TSK.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                        "	COALESCE(TSK.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                        "	COALESCE(TSK.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                        "	TSK.WO_TASK_NO, " +
                        "	TSK.WO_TASK_DESC, " +
                        "	RS.SEQNCE_NO SEQ_NO, " +
                        "	COALESCE(RS.STOCK_CODE, TRR.RES_CODE) RES_CODE, " +
                        "   1 EST_SIZE, " +
                        "	RS.UNIT_QTY_REQD UNITS_QTY, " +
                        "   TRR.QTY_ISS REAL_QTY, " + //Real Value
                        "	COALESCE(SCT.DESC_LINEX1 || SCT.ITEM_NAME, TRR.RES_DESC) RES_DESC, " +
                        "	COALESCE(SCT.UNIT_OF_ISSUE, TRR.UNITS) UNITS, " +
                        //Se tomará el valor de SHARED_TASKS como cuántas tareas comparten el mismo stock code porque el valor real no se especifica por tarea (MSE140), si no por orden completa
                        "	(SELECT COUNT(*) FROM ELLIPSE.MSF734 SRS WHERE SRS.CLASS_KEY  LIKE TSK.DSTRCT_CODE || TSK.WORK_ORDER || '%' AND SRS.CLASS_TYPE   = 'WT' AND SRS.STOCK_CODE      = RS.STOCK_CODE) SHARED_TASKS " +
                        "FROM " +
                        "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                        "	INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS " +
                        "	ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                        "	AND RS.CLASS_TYPE   = 'WT' " +
                        "	LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT " +
                        "	ON RS.STOCK_CODE   = SCT.STOCK_CODE " +
                        //Real Calculation
                        " FULL JOIN(" +
                        "   SELECT" +
                        "      TX.DSTRCT_CODE," +
                        "      WO.WORK_GROUP," +
                        "      TX.WORK_ORDER," +
                        "      TR.STOCK_CODE AS RES_CODE," +
                        "      SUM(TR.QUANTITY_ISS) QTY_ISS," +
                        "      STT.DESC_LINEX1 || STT.ITEM_NAME RES_DESC, " +
                        "      STT.UNIT_OF_ISSUE UNITS " +
                        "   FROM" +
                        "      ELLIPSE.MSFX99 TX" +
                        "      INNER JOIN ELLIPSE.MSF900 TR" +
                        "       ON TR.FULL_PERIOD = TX.FULL_PERIOD" +
                        "       AND TR.WORK_ORDER = TX.WORK_ORDER" +
                        "       AND TR.USERNO = TX.USERNO" +
                        "       AND TR.TRANSACTION_NO = TX.TRANSACTION_NO" +
                        "       AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE" +
                        "       AND TR.REC900_TYPE = TX.REC900_TYPE" +
                        "       AND TR.PROCESS_DATE = TX.PROCESS_DATE" +
                        "       AND TR.DSTRCT_CODE = TX.DSTRCT_CODE" +
                        "      INNER JOIN ELLIPSE.MSF620 WO" +
                        "        ON WO.DSTRCT_CODE = TR.DSTRCT_CODE AND WO.WORK_ORDER = TR.WORK_ORDER" +
                        "      LEFT JOIN ELLIPSE.MSF100 STT " +
                        "        ON TR.STOCK_CODE = STT.STOCK_CODE " +
                        "   WHERE" +
                        "      TX.DSTRCT_CODE = '" + districtCode + "' " +
                        "      AND TX.WORK_ORDER = '" + workOrder + "' " +
                        "      AND TX.REC900_TYPE = 'S' " +
                        "   GROUP BY" +
                        "      TX.DSTRCT_CODE," +
                        "      WO.WORK_GROUP," +
                        "      TX.WORK_ORDER," +
                        "      TR.STOCK_CODE," +
                        "      STT.DESC_LINEX1, " +
                        "      STT.ITEM_NAME, " +
                        "      STT.UNIT_OF_ISSUE " +
                        " ) TRR ON" +
                        "     TSK.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                        "     AND TSK.WORK_GROUP = TRR.WORK_GROUP " +
                        "     AND TSK.WORK_ORDER = TRR.WORK_ORDER " +
                        "     AND RS.STOCK_CODE = TRR.RES_CODE " +
                        //End Real Calculation
                        "WHERE " +
                        "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                        "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                        "   AND   TSK.WO_TASK_NO = '" + taskNo + "' ";


            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchWoTaskEquipmentRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                        "SELECT " +
                        "	'" + RequirementType.Equipment.Key + "' REQ_TYPE, " +
                        "	COALESCE(TSK.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                        "	COALESCE(TSK.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                        "	COALESCE(TSK.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                        "	TSK.WO_TASK_NO, " +
                        "	TSK.WO_TASK_DESC, " +
                        "	RS.SEQNCE_NO SEQ_NO, " +
                        "	COALESCE(RS.EQPT_TYPE, TRR.RES_CODE) RES_CODE, " +
                        "	TO_NUMBER(RS.QTY_REQ) EST_SIZE, " +
                        "	RS.UNIT_QTY_REQD UNITS_QTY, " +
                        "   TRR.QTY_ISS REAL_QTY, " + 
                        "	COALESCE(EQT.TABLE_DESC, TRR.RES_DESC) RES_DESC, " +
                        "	COALESCE(DECODE(TRIM(RS.UOM), 'H5', 'HR', TRIM(RS.UOM)), TRIM(TRR.UNITS)) UNITS, " +
                        //Se tomará el valor de SHARED_TASKS como cuántas tareas comparten el mismo recurso de equipo porque el valor real no se refleja por tarea (MSO496), si no por orden completa
                        "	(SELECT COUNT(*) FROM ELLIPSE.MSF733 SRS WHERE SRS.CLASS_KEY  LIKE TSK.DSTRCT_CODE || TSK.WORK_ORDER || '%' AND SRS.CLASS_TYPE   = 'WT' AND SRS.EQPT_TYPE      = RS.EQPT_TYPE) SHARED_TASKS " +
                        "FROM " +
                        "	" + dbReference + ".MSF623" + dbLink + " TSK " +
                        "	INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS " +
                        "	ON RS.CLASS_KEY    = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                        "	   AND RS.CLASS_TYPE   = 'WT' " +
                        "	INNER JOIN " + dbReference + ".MSF010" + dbLink + " EQT " +
                        "	ON RS.EQPT_TYPE   = EQT.TABLE_CODE " +
                        "	   AND TABLE_TYPE     = 'ET' " +
                        //Real Calculation
                        " FULL JOIN(" +
                        "   SELECT " +
                        "     TX.DSTRCT_CODE, " +
                        "     WO.WORK_GROUP, " +
                        "     TX.WORK_ORDER, " +
                        "     ETT.TABLE_CODE AS RES_CODE, " +
                        "     SUM(TR.STAT_VALUE) QTY_ISS, " +
                        "     ETT.TABLE_DESC RES_DESC, " +
                        "     TR.STAT_TYPE UNITS " +
                        "     FROM ELLIPSE.MSFX99 TX " +
                        "     INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR " +
                        "       ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                        "       AND TR.WORK_ORDER = TX.WORK_ORDER " +
                        "       AND TR.USERNO = TX.USERNO " +
                        "       AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                        "       AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                        "       AND TR.REC900_TYPE = TX.REC900_TYPE " +
                        "       AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                        "       AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                        "     INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO " +
                        "       ON TX.DSTRCT_CODE = WO.DSTRCT_CODE AND TX.WORK_ORDER = WO.WORK_ORDER " +
                        "     LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ " +
                        "       ON TR.MEMO_EQUIP = EQ.EQUIP_NO " +
                        "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " ETT " +
                        "       ON EQ.EQPT_TYPE = ETT.TABLE_CODE " +
                        "     WHERE TX.DSTRCT_CODE = '" + districtCode + "' " +
                        "       AND TX.WORK_ORDER = '" + workOrder + "' " +
                        "       AND TX.REC900_TYPE = 'E' " +
                        "     GROUP BY " +
                        "       TX.DSTRCT_CODE, " +
                        "       WO.WORK_GROUP, " +
                        "       TX.WORK_ORDER, " +
                        "       ETT.TABLE_CODE, " +
                        "       ETT.TABLE_DESC,  " +
                        "       TR.STAT_TYPE " +
                        " ) TRR ON" +
                        "     TSK.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                        "     AND TSK.WORK_GROUP = TRR.WORK_GROUP " +
                        "     AND TSK.WORK_ORDER = TRR.WORK_ORDER " +
                        "     AND RS.EQPT_TYPE = TRR.RES_CODE " +
                        "	  AND DECODE(TRIM(RS.UOM), 'H5', 'HR', TRIM(RS.UOM)) = TRIM(TRR.UNITS) " +
                        //End Real Calculation
                        "WHERE " +
                        "	TSK.DSTRCT_CODE = '" + districtCode + "' " +
                        "	AND   TSK.WORK_ORDER = '" + workOrder + "' " +
                        "	AND   TSK.WO_TASK_NO = '" + taskNo + "'";


            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        /*
        public static string GetFetchWoTaskRealRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string reqType, string taskNo = null)
        {
            var query = "";
            if (!string.IsNullOrWhiteSpace(reqType) && reqType.Equals(RequirementType.Labour.Key))
            {
                query = "WITH RES_REAL AS ( ";
                query += "    SELECT ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WT.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.WO_TASK_NO, ";
                query += "        WT.WO_TASK_DESC, ";
                query += "        TR.RESOURCE_TYPE RES_CODE, ";
                query += "        TT.TABLE_DESC RES_DESC, ";
                query += "        SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                query += "        ON TT.TABLE_CODE = TR.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                query += "        INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT ";
                query += "        ON WT.DSTRCT_CODE = TR.DSTRCT_CODE AND WT.WORK_ORDER = TR.WORK_ORDER AND WT.WO_TASK_NO = TR.WO_TASK_NO ";
                query += "    WHERE ";
                query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null)
                    query += "AND   TR.WO_TASK_NO = '" + taskNo + "' ";
                query += "    GROUP BY ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WT.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.WO_TASK_NO, ";
                query += "        WT.WO_TASK_DESC, ";
                query += "        TR.RESOURCE_TYPE, ";
                query += "        TT.TABLE_DESC ";
                query += "),RES_EST AS ( ";
                query += "    SELECT ";
                query += "        TSK.DSTRCT_CODE, ";
                query += "        TSK.WORK_GROUP, ";
                query += "        TSK.WORK_ORDER, ";
                query += "        TSK.WO_TASK_NO, ";
                query += "        TSK.WO_TASK_DESC, ";
                query += "        RS.RESOURCE_TYPE RES_CODE, ";
                query += "        TT.TABLE_DESC RES_DESC, ";
                query += "        TO_NUMBER(RS.CREW_SIZE) QTY_REQ, ";
                query += "        RS.EST_RESRCE_HRS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                query += "        INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS ";
                query += "        ON RS.KEY_735_ID = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.REC_735_TYPE = 'WT' ";
                query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                query += "        ON TT.TABLE_CODE = RS.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                query += "    WHERE ";
                query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null)
                    query += "      AND   TSK.WO_TASK_NO = '" + taskNo + "' ";
                query += "),TABLA_REC AS ( ";
                query += "    SELECT ";
                query += "        DECODE(RES_EST.DSTRCT_CODE,NULL,RES_REAL.DSTRCT_CODE,RES_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                query += "        DECODE(RES_EST.WORK_GROUP,NULL,RES_REAL.WORK_GROUP,RES_EST.WORK_GROUP) WORK_GROUP, ";
                query += "        DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, ";
                query += "        DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, ";
                query += "        DECODE(RES_EST.WO_TASK_DESC,NULL,RES_REAL.WO_TASK_DESC,RES_EST.WO_TASK_DESC) WO_TASK_DESC, ";
                query += "        DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, ";
                query += "        DECODE(RES_EST.RES_DESC,NULL,RES_REAL.RES_DESC,RES_EST.RES_DESC) RES_DESC, ";
                query += "        RES_EST.QTY_REQ, ";
                query += "        RES_REAL.ACT_RESRCE_HRS, ";
                query += "        RES_EST.EST_RESRCE_HRS ";
                query += "    FROM ";
                query += "        RES_REAL ";
                query += "        FULL JOIN RES_EST ";
                query += "        ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO AND RES_REAL.RES_CODE = RES_EST.RES_CODE ";
                query += ") SELECT ";
                query += "    " + RequirementType.Labour.Key + " REQ_TYPE, ";
                query += "    TABLA_REC.DSTRCT_CODE, ";
                query += "    TABLA_REC.WORK_GROUP, ";
                query += "    TABLA_REC.WORK_ORDER, ";
                query += "    TABLA_REC.WO_TASK_NO, ";
                query += "    TABLA_REC.WO_TASK_DESC, ";
                query += "    '' SEQ_NO, ";
                query += "    TABLA_REC.RES_CODE, ";
                query += "    TABLA_REC.RES_DESC, ";
                query += "    '' UNITS, ";
                query += "    TABLA_REC.QTY_REQ, ";
                query += "    NULL QTY_ISS, ";
                query += "    DECODE(TABLA_REC.EST_RESRCE_HRS, NULL, 0, TABLA_REC.EST_RESRCE_HRS) EST_RESRCE_HRS, ";
                query += "    DECODE(TABLA_REC.ACT_RESRCE_HRS, NULL, 0, TABLA_REC.ACT_RESRCE_HRS) ACT_RESRCE_HRS ";
                query += "FROM ";
                query += "    TABLA_REC ";
            }
            else if (!string.IsNullOrWhiteSpace(reqType) && reqType.Equals(RequirementType.Material.Key))
            {
                query = "WITH MAT_REAL AS ( ";
                query += "    SELECT ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WO.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.STOCK_CODE AS RES_CODE, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                query += "        SCT.UNIT_OF_ISSUE UNITS, ";
                query += "        SUM(TR.QUANTITY_ISS) QTY_ISS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                query += "        LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                query += "        ON TR.STOCK_CODE = SCT.STOCK_CODE ";
                query += "        INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO ";
                query += "        ON WO.DSTRCT_CODE = TR.DSTRCT_CODE AND WO.WORK_ORDER = TR.WORK_ORDER ";
                query += "    WHERE ";
                query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' AND   TX.REC900_TYPE = 'S' ";
                query += "    GROUP BY ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WO.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.STOCK_CODE, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME, ";
                query += "        SCT.UNIT_OF_ISSUE ";
                query += "),MAT_EST AS ( ";
                query += "    SELECT ";
                query += "        TSK.DSTRCT_CODE, ";
                query += "        TSK.WORK_GROUP, ";
                query += "        TSK.WORK_ORDER, ";
                query += "        TSK.WO_TASK_NO, ";
                query += "        TSK.WO_TASK_DESC, ";
                query += "        RS.SEQNCE_NO SEQ_NO, ";
                query += "        RS.STOCK_CODE RES_CODE, ";
                query += "        RS.UNIT_QTY_REQD QTY_REQ, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                query += "        SCT.UNIT_OF_ISSUE UNITS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                query += "        INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS ";
                query += "        ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO ";
                query += "        INNER JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                query += "        ON RS.STOCK_CODE = SCT.STOCK_CODE AND RS.CLASS_TYPE = 'WT' ";
                query += "    WHERE ";
                query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null) query += "           AND  TSK.WO_TASK_NO = '" + taskNo + "' ";
                query += "),TABLA_MAT AS ( ";
                query += "    SELECT ";
                query += "        DECODE(MAT_EST.DSTRCT_CODE,NULL,MAT_REAL.DSTRCT_CODE,MAT_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                query += "        DECODE(MAT_EST.WORK_GROUP,NULL,MAT_REAL.WORK_GROUP,MAT_EST.WORK_GROUP) WORK_GROUP, ";
                query += "        DECODE(MAT_EST.WORK_ORDER,NULL,MAT_REAL.WORK_ORDER,MAT_EST.WORK_ORDER) WORK_ORDER, ";
                query += "        MAT_EST.WO_TASK_NO, ";
                query += "        MAT_EST.WO_TASK_DESC, ";
                query += "        MAT_EST.SEQ_NO, ";
                query += "        DECODE(MAT_EST.RES_CODE,NULL,MAT_REAL.RES_CODE,MAT_EST.RES_CODE) RES_CODE, ";
                query += "        DECODE(MAT_EST.RES_DESC,NULL,MAT_REAL.RES_DESC,MAT_EST.RES_DESC) RES_DESC, ";
                query += "        DECODE(MAT_EST.UNITS,NULL,MAT_REAL.UNITS,MAT_EST.UNITS) UNITS, ";
                query += "        MAT_EST.QTY_REQ, ";
                query += "        MAT_REAL.QTY_ISS ";
                query += "    FROM ";
                query += "        MAT_REAL ";
                query += "        FULL JOIN MAT_EST ";
                query += "        ON MAT_REAL.DSTRCT_CODE = MAT_EST.DSTRCT_CODE AND MAT_REAL.WORK_ORDER = MAT_EST.WORK_ORDER AND MAT_REAL.RES_CODE = MAT_EST.RES_CODE ";
                query += ")SELECT ";
                query += "    " + RequirementType.Material.Key + " REQ_TYPE, ";
                query += "    TABLA_MAT.DSTRCT_CODE, ";
                query += "    TABLA_MAT.WORK_GROUP, ";
                query += "    TABLA_MAT.WORK_ORDER, ";
                query += "    TABLA_MAT.WO_TASK_NO, ";
                query += "    TABLA_MAT.WO_TASK_DESC, ";
                query += "    TABLA_MAT.SEQ_NO, ";
                query += "    TABLA_MAT.RES_CODE, ";
                query += "    TABLA_MAT.RES_DESC, ";
                query += "    DECODE(TABLA_MAT.UNITS, NULL, '', TABLA_MAT.UNITS) UNITS, ";
                query += "    TABLA_MAT.QTY_REQ, ";
                query += "    DECODE(TABLA_MAT.QTY_ISS, NULL, 0,TABLA_MAT.QTY_ISS) QTY_ISS, ";
                query += "    0 EST_RESRCE_HRS, ";
                query += "    0 ACT_RESRCE_HRS ";
                query += "  FROM ";
                query += "    TABLA_MAT ";
            }
            else
            {
                query = "WITH MAT_REAL AS ( ";
                query += "    SELECT ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WO.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.STOCK_CODE AS RES_CODE, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                query += "        SCT.UNIT_OF_ISSUE UNITS, ";
                query += "        SUM(TR.QUANTITY_ISS) QTY_ISS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                query += "        LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                query += "        ON TR.STOCK_CODE = SCT.STOCK_CODE ";
                query += "        INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO ";
                query += "        ON WO.DSTRCT_CODE = TR.DSTRCT_CODE AND WO.WORK_ORDER = TR.WORK_ORDER ";
                query += "    WHERE ";
                query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' AND   TX.REC900_TYPE = 'S' ";
                query += "    GROUP BY ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WO.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.STOCK_CODE, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME, ";
                query += "        SCT.UNIT_OF_ISSUE ";
                query += "),MAT_EST AS ( ";
                query += "    SELECT ";
                query += "        TSK.DSTRCT_CODE, ";
                query += "        TSK.WORK_GROUP, ";
                query += "        TSK.WORK_ORDER, ";
                query += "        TSK.WO_TASK_NO, ";
                query += "        TSK.WO_TASK_DESC, ";
                query += "        RS.SEQNCE_NO SEQ_NO, ";
                query += "        RS.STOCK_CODE RES_CODE, ";
                query += "        RS.UNIT_QTY_REQD QTY_REQ, ";
                query += "        SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, ";
                query += "        SCT.UNIT_OF_ISSUE UNITS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                query += "        INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS ";
                query += "        ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO ";
                query += "        INNER JOIN " + dbReference + ".MSF100" + dbLink + " SCT ";
                query += "        ON RS.STOCK_CODE = SCT.STOCK_CODE AND RS.CLASS_TYPE = 'WT' ";
                query += "    WHERE ";
                query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null) query += "           AND  TSK.WO_TASK_NO = '" + taskNo + "' ";
                query += "),TABLA_MAT AS ( ";
                query += "    SELECT ";
                query += "        DECODE(MAT_EST.DSTRCT_CODE,NULL,MAT_REAL.DSTRCT_CODE,MAT_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                query += "        DECODE(MAT_EST.WORK_GROUP,NULL,MAT_REAL.WORK_GROUP,MAT_EST.WORK_GROUP) WORK_GROUP, ";
                query += "        DECODE(MAT_EST.WORK_ORDER,NULL,MAT_REAL.WORK_ORDER,MAT_EST.WORK_ORDER) WORK_ORDER, ";
                query += "        MAT_EST.WO_TASK_NO, ";
                query += "        MAT_EST.WO_TASK_DESC, ";
                query += "        MAT_EST.SEQ_NO, ";
                query += "        DECODE(MAT_EST.RES_CODE,NULL,MAT_REAL.RES_CODE,MAT_EST.RES_CODE) RES_CODE, ";
                query += "        DECODE(MAT_EST.RES_DESC,NULL,MAT_REAL.RES_DESC,MAT_EST.RES_DESC) RES_DESC, ";
                query += "        DECODE(MAT_EST.UNITS,NULL,MAT_REAL.UNITS,MAT_EST.UNITS) UNITS, ";
                query += "        MAT_EST.QTY_REQ, ";
                query += "        MAT_REAL.QTY_ISS ";
                query += "    FROM ";
                query += "        MAT_REAL ";
                query += "        FULL JOIN MAT_EST ";
                query += "        ON MAT_REAL.DSTRCT_CODE = MAT_EST.DSTRCT_CODE AND MAT_REAL.WORK_ORDER = MAT_EST.WORK_ORDER AND MAT_REAL.RES_CODE = MAT_EST.RES_CODE ";
                query += "),RES_REAL AS ( ";
                query += "    SELECT ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WT.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.WO_TASK_NO, ";
                query += "        WT.WO_TASK_DESC, ";
                query += "        TR.RESOURCE_TYPE RES_CODE, ";
                query += "        TT.TABLE_DESC RES_DESC, ";
                query += "        SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSFX99" + dbLink + " TX ";
                query += "        INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR ";
                query += "        ON TR.FULL_PERIOD = TX.FULL_PERIOD AND TR.WORK_ORDER = TX.WORK_ORDER AND TR.USERNO = TX.USERNO AND TR.TRANSACTION_NO = TX.TRANSACTION_NO AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE AND TR.REC900_TYPE = TX.REC900_TYPE AND TR.PROCESS_DATE = TX.PROCESS_DATE AND TR.DSTRCT_CODE = TX.DSTRCT_CODE ";
                query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                query += "        ON TT.TABLE_CODE = TR.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                query += "        INNER JOIN " + dbReference + ".MSF623" + dbLink + " WT ";
                query += "        ON WT.DSTRCT_CODE = TR.DSTRCT_CODE AND WT.WORK_ORDER = TR.WORK_ORDER AND WT.WO_TASK_NO = TR.WO_TASK_NO ";
                query += "    WHERE ";
                query += "        TR.DSTRCT_CODE = '" + districtCode + "' AND   TR.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null)
                    query += "AND   TR.WO_TASK_NO = '" + taskNo + "' ";
                query += "    GROUP BY ";
                query += "        TR.DSTRCT_CODE, ";
                query += "        WT.WORK_GROUP, ";
                query += "        TR.WORK_ORDER, ";
                query += "        TR.WO_TASK_NO, ";
                query += "        WT.WO_TASK_DESC, ";
                query += "        TR.RESOURCE_TYPE, ";
                query += "        TT.TABLE_DESC ";
                query += "),RES_EST AS ( ";
                query += "    SELECT ";
                query += "        TSK.DSTRCT_CODE, ";
                query += "        TSK.WORK_GROUP, ";
                query += "        TSK.WORK_ORDER, ";
                query += "        TSK.WO_TASK_NO, ";
                query += "        TSK.WO_TASK_DESC, ";
                query += "        RS.RESOURCE_TYPE RES_CODE, ";
                query += "        TT.TABLE_DESC RES_DESC, ";
                query += "        TO_NUMBER(RS.CREW_SIZE) QTY_REQ, ";
                query += "        RS.EST_RESRCE_HRS ";
                query += "    FROM ";
                query += "        " + dbReference + ".MSF623" + dbLink + " TSK ";
                query += "        INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS ";
                query += "        ON RS.KEY_735_ID = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.REC_735_TYPE = 'WT' ";
                query += "        INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT ";
                query += "        ON TT.TABLE_CODE = RS.RESOURCE_TYPE AND TT.TABLE_TYPE = 'TT' ";
                query += "    WHERE ";
                query += "        TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null)
                    query += "      AND   TSK.WO_TASK_NO = '" + taskNo + "' ";
                query += "),TABLA_REC AS ( ";
                query += "    SELECT ";
                query += "        DECODE(RES_EST.DSTRCT_CODE,NULL,RES_REAL.DSTRCT_CODE,RES_EST.DSTRCT_CODE) DSTRCT_CODE, ";
                query += "        DECODE(RES_EST.WORK_GROUP,NULL,RES_REAL.WORK_GROUP,RES_EST.WORK_GROUP) WORK_GROUP, ";
                query += "        DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, ";
                query += "        DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, ";
                query += "        DECODE(RES_EST.WO_TASK_DESC,NULL,RES_REAL.WO_TASK_DESC,RES_EST.WO_TASK_DESC) WO_TASK_DESC, ";
                query += "        DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, ";
                query += "        DECODE(RES_EST.RES_DESC,NULL,RES_REAL.RES_DESC,RES_EST.RES_DESC) RES_DESC, ";
                query += "        RES_EST.QTY_REQ, ";
                query += "        RES_REAL.ACT_RESRCE_HRS, ";
                query += "        RES_EST.EST_RESRCE_HRS ";
                query += "    FROM ";
                query += "        RES_REAL ";
                query += "        FULL JOIN RES_EST ";
                query += "        ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO AND RES_REAL.RES_CODE = RES_EST.RES_CODE ";
                query += ") SELECT ";
                query += "    " + RequirementType.Material.Key + " REQ_TYPE, ";
                query += "    TABLA_MAT.DSTRCT_CODE, ";
                query += "    TABLA_MAT.WORK_GROUP, ";
                query += "    TABLA_MAT.WORK_ORDER, ";
                query += "    TABLA_MAT.WO_TASK_NO, ";
                query += "    TABLA_MAT.WO_TASK_DESC, ";
                query += "    TABLA_MAT.SEQ_NO, ";
                query += "    TABLA_MAT.RES_CODE, ";
                query += "    TABLA_MAT.RES_DESC, ";
                query += "    DECODE(TABLA_MAT.UNITS, NULL, '', TABLA_MAT.UNITS) UNITS, ";
                query += "    TABLA_MAT.QTY_REQ, ";
                query += "    DECODE(TABLA_MAT.QTY_ISS, NULL, 0,TABLA_MAT.QTY_ISS) QTY_ISS, ";
                query += "    0 EST_RESRCE_HRS, ";
                query += "    0 ACT_RESRCE_HRS ";
                query += "  FROM ";
                query += "    TABLA_MAT ";
                query += "UNION ALL ";
                query += "SELECT ";
                query += "    " + RequirementType.Labour.Key + " REQ_TYPE, ";
                query += "    TABLA_REC.DSTRCT_CODE, ";
                query += "    TABLA_REC.WORK_GROUP, ";
                query += "    TABLA_REC.WORK_ORDER, ";
                query += "    TABLA_REC.WO_TASK_NO, ";
                query += "    TABLA_REC.WO_TASK_DESC, ";
                query += "    '' SEQ_NO, ";
                query += "    TABLA_REC.RES_CODE, ";
                query += "    TABLA_REC.RES_DESC, ";
                query += "    '' UNITS, ";
                query += "    TABLA_REC.QTY_REQ, ";
                query += "    NULL QTY_ISS, ";
                query += "    DECODE(TABLA_REC.EST_RESRCE_HRS, NULL, 0, TABLA_REC.EST_RESRCE_HRS) EST_RESRCE_HRS, ";
                query += "    DECODE(TABLA_REC.ACT_RESRCE_HRS, NULL, 0, TABLA_REC.ACT_RESRCE_HRS) ACT_RESRCE_HRS ";
                query += "FROM ";
                query += "    TABLA_REC ";
                query += "UNION ALL ";
                query += "SELECT ";
                query += "    " + RequirementType.Equipment.Key + " REQ_TYPE, ";
                query += "    TSK.DSTRCT_CODE, ";
                query += "    TSK.WORK_GROUP, ";
                query += "    TSK.WORK_ORDER, ";
                query += "    TSK.WO_TASK_NO, ";
                query += "    TSK.WO_TASK_DESC, ";
                query += "    RS.SEQNCE_NO SEQ_NO, ";
                query += "    RS.EQPT_TYPE RES_CODE, ";
                query += "    ET.TABLE_DESC RES_DESC, ";
                query += "    RS.UOM UNITS, ";
                query += "    RS.QTY_REQ, ";
                query += "    0 QTY_ISS, ";
                query += "    DECODE(RS.UNIT_QTY_REQD, NULL, 0, RS.UNIT_QTY_REQD) EST_RESRCE_HRS, ";
                query += "    0 ACT_RESRCE_HRS ";
                query += "FROM ";
                query += "    " + dbReference + ".MSF623" + dbLink + " TSK ";
                query += "    INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS ";
                query += "    ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO AND RS.CLASS_TYPE = 'WT' ";
                query += "    INNER JOIN " + dbReference + ".MSF010" + dbLink + " ET ";
                query += "    ON RS.EQPT_TYPE = ET.TABLE_CODE AND TABLE_TYPE = 'ET' ";
                query += "WHERE ";
                query += "    TSK.DSTRCT_CODE = '" + districtCode + "' AND   TSK.WORK_ORDER = '" + workOrder + "' ";
                if (taskNo != null)
                    query += "AND   TSK.WO_TASK_NO = '" + taskNo + "'";
            }

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            return query;

        }}*/
    }
}

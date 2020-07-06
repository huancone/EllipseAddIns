using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Constants;

namespace EllipseJobsClassLibrary
{
    public static partial class Queries
    {
        public static string GetEllipseResourcesQuery(string dbReference, string dbLink, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var groupList = new List<string>();

            if (primakeryKey == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList.Add(primaryValue);
            else if (primakeryKey == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == primaryValue).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == primaryValue).Select(g => g.Name).ToList();

            var query = "WITH CTE_DATES ( CTE_DATE ) AS ( " +
                        "    SELECT CAST(TO_DATE('" + startDate + "','YYYYMMDD') AS DATE) CTE_DATE FROM DUAL " +
                        "    UNION ALL " +
                        "    SELECT CAST( (CTE_DATE + 1) AS DATE) CTE_DATE FROM CTE_DATES WHERE TRUNC(CTE_DATE) + 1 <= TO_DATE('" + endDate + "','YYYYMMDD') " +
                        "),FECHAS AS ( " +
                        "    SELECT TO_CHAR(CTE_DATE,'YYYYMMDD') FECHA FROM CTE_DATES " +
                        ") SELECT " +
                        "    ELL.WORK_GROUP GRUPO, " +
                        "    FECHAS.FECHA FECHA, " +
                        "    ELL.RESOURCE_TYPE RECURSO, " +
                        "    ELL.REQ_RESRC_NO CANTIDAD, " +
                        "    ROUND( (TO_DATE(FECHAS.FECHA || ' ' || DEF_STOP_TIME,'YYYYMMDD HH24MISS') - TO_DATE(FECHAS.FECHA || ' ' || DEF_STR_TIME,'YYYYMMDD HH24MISS') ) * 24 * ELL.REQ_RESRC_NO * (1 - ( (WG.BDOWN_ALLOW_PC + ASSIGN_OTH_PC) / 100) ),2) HORAS " +
                        "  FROM " +
                        "    " + dbReference + ".MSF730_RESRC_REQ" + dbLink + " ELL " +
                        "    INNER JOIN " + dbReference + ".MSF720" + dbLink + " WG " +
                        "    ON ELL.WORK_GROUP = WG.WORK_GROUP, " +
                        "    FECHAS " +
                        "  WHERE " +
                        "    ELL.WORK_GROUP IN (" + groupList.Aggregate("", (current, g) => current + "'" + g + "'") + ") ";

            return query;
        }

        public static string GetPsoftResourcesQuery(string dbReference, string dbLink, string district, int primakeryKey, string primaryValue, string startDate, string endDate)
        {
            var groupList = new List<string>();

            if (primakeryKey == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList.Add(primaryValue);
            else if (primakeryKey == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(primaryValue))
                groupList = Groups.GetWorkGroupList().Where(g => g.Area == primaryValue).Select(g => g.Name).ToList();
            else
                groupList = Groups.GetWorkGroupList().Where(g => g.Details == primaryValue).Select(g => g.Name).ToList();

            var query = "WITH CTE_DATES ( CTE_DATE ) AS ( " +
                        "    SELECT CAST(TO_DATE('" + startDate + "','YYYYMMDD') AS DATE) CTE_DATE FROM DUAL " +
                        "    UNION ALL " +
                        "    SELECT CAST( (CTE_DATE + 1) AS DATE) CTE_DATE FROM CTE_DATES WHERE TRUNC(CTE_DATE) + 1 <= TO_DATE('" + endDate + "','YYYYMMDD') " +
                        "),FECHAS AS ( " +
                        "    SELECT TO_CHAR(CTE_DATE,'YYYYMMDD') FECHA FROM CTE_DATES " +
                        ") SELECT " +
                        "    WE.WORK_GROUP GRUPO, " +
                        "    FECHAS.FECHA, " +
                        "    EMP.RESOURCE_TYPE RECURSO, " +
                        "    TURNOS.CEDULA, " +
                        "    TRIM(EMP.FIRST_NAME) || ' ' || TRIM(EMP.SURNAME) NOMBRE, " +
                        "    ROUND(TURNOS.HORAS,2) HORAS " +
                        "  FROM " +
                        "    " + dbReference + ".MSF810" + dbLink + " EMP " +
                        "    INNER JOIN " + dbReference + ".MSF723" + dbLink + " WE " +
                        "    ON EMP.EMPLOYEE_ID = WE.EMPLOYEE_ID " +
                        "    AND   WE.STOP_DT_REVSD = '00000000' " +
                        "    AND WE.WORK_GROUP IN (" + groupList.Aggregate("", (current, g) => current + "'" + g + "'") + ") " +
                        "    LEFT JOIN SIGMAN.ASISTENCIA TURNOS " +
                        "    ON LPAD(EMP.EMPLOYEE_ID,11,'0') = LPAD(TURNOS.CEDULA,11,'0') " +
                        "    INNER JOIN FECHAS " +
                        "    ON   TURNOS.FECHAP = FECHAS.FECHA " +
                        "  WHERE " +
                        "    TRIM(EMP.RESOURCE_TYPE) IS NOT NULL  " +
                        "    AND   TRIM(EMP.RESOURCE_TYPE) NOT IN ('SMPT','SSUP') " +
                        "ORDER BY WE.WORK_GROUP, " +
                        "    FECHAS.FECHA, " +
                        "    EMP.RESOURCE_TYPE, " +
                        "    TURNOS.CEDULA ";
            return query;
        }

        public static string SaveResourcesQuery(string dbReference, LabourResources l)
        {
            var query = "MERGE INTO SIGMDC.RECURSOS_PROGRAMACION T USING " +
                         "(SELECT " +
                         " '" + l.WorkGroup + "' GRUPO, " +
                         " '" + l.ResourceCode + "' RECURSO, " +
                         " '" + l.Date + "' FECHA, " +
                         " '" + l.EstimatedLabourHours + "' HORAS_PRO, " +
                         " '" + l.AvailableLabourHours + "' HORAS_DISPO " +
                         " FROM DUAL)S ON ( " +
                         " T.GRUPO = S.GRUPO " +
                         " AND T.RECURSO = S.RECURSO " +
                         " AND T.FECHA = S.FECHA " +
                         ") " +
                         "WHEN MATCHED THEN UPDATE SET T.HORAS_PRO = S.HORAS_PRO, T.HORAS_DISPO = S.HORAS_DISPO " +
                         "WHEN NOT MATCHED THEN INSERT(GRUPO, RECURSO, FECHA, HORAS_PRO, HORAS_DISPO) " +
                         "VALUES(S.GRUPO, S.RECURSO, S.FECHA, S.HORAS_PRO, S.HORAS_DISPO) ";

            return query;
        }

        public static string SaveTaskQuery(string dbReference, JobTask t)
        {
            var query = "MERGE INTO SIGMDC.SEG_PROGRAMACION T USING " +
                         "(SELECT " +
                         " '" + t.WorkGroup + "' WORK_GROUP, " +
                         " '" + t.PlanStrDate + "' FECHA, " +
                         " '" + t.WorkOrder + "' WORK_ORDER, " +
                         " '" + t.WoTaskNo + "' WO_TASK_NO " +
                         " FROM DUAL)S ON ( " +
                         " T.WORK_GROUP = S.WORK_GROUP " +
                         " AND T.FECHA = S.FECHA " +
                         " AND T.WORK_ORDER = S.WORK_ORDER " +
                         " AND T.WO_TASK_NO = S.WO_TASK_NO " +
                         ") " +
                         "WHEN NOT MATCHED THEN INSERT(WORK_GROUP, FECHA, WORK_ORDER, WO_TASK_NO) " +
                         "VALUES(S.WORK_GROUP, S.FECHA, S.WORK_ORDER, S.WO_TASK_NO) ";
            return query;
        }

        public static string GetEllipseSingleTaskQuery(string dbReference, string dbLink, string district, string reference, string referenceTask, string referenceStartDate, string referenceStartHour, string referenceFinDate, string referenceFinHour, string startDate, string finDate, string resourceCode)
        {
            var query = "WITH CTE_DATES ( " +
                        "     STARTDATE, " +
                        "     ENDDATE " +
                        " ) AS ( " +
                        "     SELECT " +
                        "         CAST(TO_DATE('" + startDate + " 060000','YYYYMMDD HH24MISS') AS DATE) STARTDATE, " +
                        "         CAST(TO_DATE('" + startDate + " 180000','YYYYMMDD HH24MISS') AS DATE) ENDDATE " +
                        "     FROM " +
                        "         DUAL " +
                        "     UNION ALL " +
                        "     SELECT " +
                        "         CAST( (CTE_DATES.STARTDATE + 0.5) AS DATE) STARTDATE, " +
                        "         CAST( (CTE_DATES.ENDDATE + 0.5) AS DATE) ENDDATE " +
                        "     FROM " +
                        "         CTE_DATES " +
                        "     WHERE " +
                        "         TRUNC(CTE_DATES.ENDDATE) + 0.5 <= TO_DATE('" + finDate + " 180000','YYYYMMDD HH24MISS') " +
                        " ),TASKS AS ( " +
                        "     SELECT " +
                        "         'WT' TASK_TYPE, " +
                        "         WT.DSTRCT_CODE, " +
                        "         WT.WORK_GROUP, " +
                        "         WT.WORK_ORDER, " +
                        "         WT.WO_TASK_NO, " +
                        "         WT.WO_TASK_DESC, " +
                        "         TO_DATE(WT.PLAN_STR_DATE || WT.PLAN_STR_TIME,'YYYYMMDD HH24MISS') PLAN_STR_DATE, " +
                        "         TO_DATE(WT.PLAN_STR_DATE || WT.PLAN_STR_TIME,'YYYYMMDD HH24MISS') + WT.TSK_DUR_HOURS / 24 PLAN_FIN_DATE, " +
                        "         WT.TSK_DUR_HOURS, " +
                        "         WT.CALC_LAB_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSF623 WT " +
                        "     WHERE " +
                        "         WT.DSTRCT_CODE = 'ICOR' " +
                        "         AND WT.WORK_ORDER = '" + reference + "' " +
                        "         AND WT.WO_TASK_NO = '" + referenceTask + "' " +

                        "         AND TRIM(WT.TSK_DUR_HOURS) IS NOT NULL " +
                        "         AND TRIM(WT.PLAN_STR_DATE) IS NOT NULL " +
                        "         AND TRIM(WT.PLAN_STR_DATE) <> '00000000' " +
                        "     UNION ALL " +
                        "     SELECT " +
                        "         'ST' TASK_TYPE, " +
                        "         ST.DSTRCT_CODE, " +
                        "         ST.WORK_GROUP, " +
                        "         ST.STD_JOB_NO, " +
                        "         ST.STD_JOB_TASK, " +
                        "         ST.SJ_TASK_DESC, " +
                        "         TO_DATE('" + referenceStartDate + "' || '" + referenceStartHour + "','YYYYMMDD HH24MISS') PLAN_STR_DATE, " +
                        "         TO_DATE('" + referenceStartDate + "' || '" + referenceStartHour + "','YYYYMMDD HH24MISS') + ST.TSK_DUR_HOURS / 24 PLAN_FIN_DATE, " +
                        "         ST.TSK_DUR_HOURS, " +
                        "         ST.CALC_LAB_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSF693 ST " +
                        "     WHERE " +
                        "         ST.DSTRCT_CODE = 'ICOR' " +
                        "         AND ST.STD_JOB_NO = '" + reference + "' " +
                        "         AND ST.STD_JOB_TASK = '" + referenceTask + "' " +
                        " ),SHIFT_TASKS AS ( " +
                        "     SELECT " +
                        "         TASKS.DSTRCT_CODE, " +
                        "         TASKS.WORK_GROUP, " +
                        "         TASKS.WORK_ORDER, " +
                        "         TASKS.WO_TASK_NO, " +
                        "         TASKS.WO_TASK_DESC, " +
                        "         CTE_DATES.STARTDATE   SHIFT, " +
                        "         CASE " +
                        "             WHEN TASKS.PLAN_STR_DATE >= CTE_DATES.STARTDATE THEN " +
                        "                 TASKS.PLAN_STR_DATE " +
                        "             ELSE " +
                        "                 CTE_DATES.STARTDATE " +
                        "         END PLAN_STR_DATE, " +
                        "         CASE " +
                        "             WHEN TASKS.PLAN_FIN_DATE <= CTE_DATES.ENDDATE THEN " +
                        "                 TASKS.PLAN_FIN_DATE " +
                        "             ELSE " +
                        "                 CTE_DATES.ENDDATE " +
                        "         END PLAN_FIN_DATE, " +
                        "         TASKS.TSK_DUR_HOURS " +
                        "     FROM " +
                        "         TASKS " +
                        "         INNER JOIN CTE_DATES " +
                        "         ON TASKS.PLAN_STR_DATE < CTE_DATES.ENDDATE " +
                        "            AND TASKS.PLAN_FIN_DATE > CTE_DATES.STARTDATE " +
                        " ),RES_REAL AS ( " +
                        "     SELECT " +
                        "         TR.DSTRCT_CODE, " +
                        "         TR.WORK_ORDER, " +
                        "         TR.WO_TASK_NO, " +
                        "         TR.RESOURCE_TYPE   RES_CODE, " +
                        "         SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS " +
                        "     FROM " +
                        "         ELLIPSE.MSFX99 TX " +
                        "         INNER JOIN ELLIPSE.MSF900 TR " +
                        "         ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                        "            AND TR.WORK_ORDER = TX.WORK_ORDER " +
                        "            AND TR.USERNO = TX.USERNO " +
                        "            AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                        "            AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                        "            AND TR.REC900_TYPE = TX.REC900_TYPE " +
                        "            AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                        "            AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                        "            AND TR.DSTRCT_CODE = 'ICOR' " +
                        "            AND TR.WORK_ORDER = '" + reference + "' " +
                        "            AND TR.WO_TASK_NO = '" + referenceTask + "' " +
                        "            AND TR.RESOURCE_TYPE = '" + resourceCode + "' " +
                        "     GROUP BY " +
                        "         TR.DSTRCT_CODE, " +
                        "         TR.WORK_ORDER, " +
                        "         TR.WO_TASK_NO, " +
                        "         TR.RESOURCE_TYPE " +
                        " ),RES_EST AS ( " +
                        "     SELECT " +
                        "         TASKS.DSTRCT_CODE, " +
                        "         TASKS.WORK_ORDER, " +
                        "         TASKS.WO_TASK_NO, " +
                        "         RS.RESOURCE_TYPE   RES_CODE, " +
                        "         TT.TABLE_DESC      RES_DESC, " +
                        "         TO_NUMBER(RS.CREW_SIZE) QTY_REQ, " +
                        "         RS.EST_RESRCE_HRS " +
                        "     FROM " +
                        "         TASKS " +
                        "         INNER JOIN ELLIPSE.MSF735 RS " +
                        "         ON RS.KEY_735_ID LIKE 'ICOR" + reference + referenceTask + "%' " +
                        "         INNER JOIN ELLIPSE.MSF010 TT " +
                        "         ON TT.TABLE_CODE = RS.RESOURCE_TYPE " +
                        "            AND TT.TABLE_TYPE = 'TT' " +
                        "     WHERE " +
                        "         TASKS.DSTRCT_CODE = 'ICOR' " +
                        "         AND RS.RESOURCE_TYPE = '" + resourceCode + "' " +
                        "         AND RS.REC_735_TYPE IN ('WT', 'ST') " +
                        " ),TABLA_REC AS ( " +
                        "     SELECT " +
                        "         RES_EST.DSTRCT_CODE, " +
                        "         DECODE(RES_EST.WORK_ORDER,NULL,RES_REAL.WORK_ORDER,RES_EST.WORK_ORDER) WORK_ORDER, " +
                        "         DECODE(RES_EST.WO_TASK_NO,NULL,RES_REAL.WO_TASK_NO,RES_EST.WO_TASK_NO) WO_TASK_NO, " +
                        "         DECODE(RES_EST.RES_CODE,NULL,RES_REAL.RES_CODE,RES_EST.RES_CODE) RES_CODE, " +
                        "         RES_EST.QTY_REQ, " +
                        "         DECODE(RES_EST.EST_RESRCE_HRS,NULL,0,RES_EST.EST_RESRCE_HRS) EST_RESRCE_HRS, " +
                        "         DECODE(RES_REAL.ACT_RESRCE_HRS,NULL,0,RES_REAL.ACT_RESRCE_HRS) ACT_RESRCE_HRS " +
                        "     FROM " +
                        "         RES_REAL " +
                        "         FULL JOIN RES_EST " +
                        "         ON RES_REAL.DSTRCT_CODE = RES_EST.DSTRCT_CODE " +
                        "            AND RES_REAL.WORK_ORDER = RES_EST.WORK_ORDER " +
                        "            AND RES_REAL.WO_TASK_NO = RES_EST.WO_TASK_NO " +
                        "            AND RES_REAL.RES_CODE = RES_EST.RES_CODE " +
                        " )SELECT " +
                        "     SHIFT_TASKS.WORK_GROUP, " +
                        "     SHIFT_TASKS.WORK_ORDER, " +
                        "     SHIFT_TASKS.WO_TASK_NO, " +
                        "     SHIFT_TASKS.WO_TASK_DESC, " +
                        "     SHIFT_TASKS.SHIFT, " +
                        "     SHIFT_TASKS.PLAN_STR_DATE, " +
                        "     SHIFT_TASKS.PLAN_FIN_DATE, " +
                        "     SHIFT_TASKS.TSK_DUR_HOURS, " +
                        "     ROUND(24 * ( SHIFT_TASKS.PLAN_FIN_DATE - SHIFT_TASKS.PLAN_STR_DATE ),2) SHIFT_TSK_DUR_HOURS, " +
                        "     TABLA_REC.RES_CODE, " +
                        "     TABLA_REC.QTY_REQ, " +
                        "     TABLA_REC.EST_RESRCE_HRS, " +
                        "     TABLA_REC.ACT_RESRCE_HRS, " +
                        "     DECODE(SHIFT_TASKS.TSK_DUR_HOURS, 0, 0, ROUND(TABLA_REC.EST_RESRCE_HRS * ( 24 * ( SHIFT_TASKS.PLAN_FIN_DATE - SHIFT_TASKS.PLAN_STR_DATE ) / SHIFT_TASKS.TSK_DUR_HOURS ),2)) SHIFT_LAB_HOURS " +
                        " FROM " +
                        "     SHIFT_TASKS " +
                        "     INNER JOIN TABLA_REC " +
                        "     ON SHIFT_TASKS.WORK_ORDER = TABLA_REC.WORK_ORDER " +
                        "        AND SHIFT_TASKS.WO_TASK_NO = TABLA_REC.WO_TASK_NO " +
                        "        AND SHIFT_TASKS.DSTRCT_CODE = TABLA_REC.DSTRCT_CODE ";

            return query;
        }
        public static string GetJobTaskAdditionalQuery(string dbReference, string dbLink, JobTask task)
        {
            string headerQuery = "";

            if (!string.IsNullOrWhiteSpace(task.WorkOrder))
            {
                headerQuery = " WITH ORDERS AS (" +
                        "   SELECT " +
                        "     WO.DSTRCT_CODE," +
                        "     WO.WORK_GROUP," +
                        "     WO.EQUIP_NO," +
                        "     WO.COMP_CODE," +
                        "     WO.COMP_MOD_CODE," +
                        "     WO.MAINT_SCH_TASK," +
                        "     WO.STD_JOB_NO," +
                        "     WO.WORK_ORDER," +
                        "     WOT.WO_TASK_NO," +
                        "     WO.PLAN_STR_DATE," +
                        "     COALESCE(TRIM(WO.ORIG_SCHED_DATE), WO.PLAN_STR_DATE) ORIG_SCHED_DATE," +
                        "     WO.REQ_START_DATE," +
                        "     WO.REQ_BY_DATE," +
                        "     WO.COMPLETED_CODE," +
                        "     WO.ASSIGN_PERSON WO_ASSIGN_PERSON," +
                        "     WOT.ASSIGN_PERSON," +
                        "     WO.MAINT_TYPE," +
                        "     WO.WO_TYPE," +
                        "     (SELECT" +
                        "       CASE" +
                        "         WHEN PRIMARY_STAT_1 = 'Y' THEN STAT_TYPE_1" +
                        "         WHEN PRIMARY_STAT_2 = 'Y' THEN STAT_TYPE_2" +
                        "         WHEN PRIMARY_STAT_3 = 'Y' THEN STAT_TYPE_3" +
                        "         WHEN PRIMARY_STAT_4 = 'Y' THEN STAT_TYPE_4" +
                        "         WHEN PRIMARY_STAT_5 = 'Y' THEN STAT_TYPE_5" +
                        "         WHEN PRIMARY_STAT_6 = 'Y' THEN STAT_TYPE_6" +
                        "         WHEN PRIMARY_STAT_7 = 'Y' THEN STAT_TYPE_7" +
                        "         WHEN PRIMARY_STAT_8 = 'Y' THEN STAT_TYPE_8" +
                        "         WHEN PRIMARY_STAT_9 = 'Y' THEN STAT_TYPE_9" +
                        "         WHEN PRIMARY_STAT_10 = 'Y' THEN STAT_TYPE_10" +
                        "         WHEN PRIMARY_STAT_11 = 'Y' THEN STAT_TYPE_11" +
                        "         WHEN PRIMARY_STAT_12 = 'Y' THEN STAT_TYPE_12" +
                        "         WHEN PRIMARY_STAT_13 = 'Y' THEN STAT_TYPE_13" +
                        "         WHEN PRIMARY_STAT_14 = 'Y' THEN STAT_TYPE_14" +
                        "         WHEN PRIMARY_STAT_15 = 'Y' THEN STAT_TYPE_15" +
                        "         WHEN PRIMARY_STAT_16 = 'Y' THEN STAT_TYPE_16" +
                        "         WHEN PRIMARY_STAT_17 = 'Y' THEN STAT_TYPE_17" +
                        "         WHEN PRIMARY_STAT_18 = 'Y' THEN STAT_TYPE_18" +
                        "         WHEN PRIMARY_STAT_19 = 'Y' THEN STAT_TYPE_19" +
                        "         WHEN PRIMARY_STAT_20 = 'Y' THEN STAT_TYPE_20" +
                        "       END STAT_TYPE_PR" +
                        "     FROM " + dbReference + ".MSF617_OP_STATS" + dbLink + " OPS WHERE EGI_REC_TYPE = 'E' AND EQUIP_GRP_ID = WO.EQUIP_NO) E_STAT_TYPE_PR," +
                        "         (SELECT" +
                        "       CASE" +
                        "         WHEN PRIMARY_STAT_1 = 'Y' THEN STAT_TYPE_1" +
                        "         WHEN PRIMARY_STAT_2 = 'Y' THEN STAT_TYPE_2" +
                        "         WHEN PRIMARY_STAT_3 = 'Y' THEN STAT_TYPE_3" +
                        "         WHEN PRIMARY_STAT_4 = 'Y' THEN STAT_TYPE_4" +
                        "         WHEN PRIMARY_STAT_5 = 'Y' THEN STAT_TYPE_5" +
                        "         WHEN PRIMARY_STAT_6 = 'Y' THEN STAT_TYPE_6" +
                        "         WHEN PRIMARY_STAT_7 = 'Y' THEN STAT_TYPE_7" +
                        "         WHEN PRIMARY_STAT_8 = 'Y' THEN STAT_TYPE_8" +
                        "         WHEN PRIMARY_STAT_9 = 'Y' THEN STAT_TYPE_9" +
                        "         WHEN PRIMARY_STAT_10 = 'Y' THEN STAT_TYPE_10" +
                        "         WHEN PRIMARY_STAT_11 = 'Y' THEN STAT_TYPE_11" +
                        "         WHEN PRIMARY_STAT_12 = 'Y' THEN STAT_TYPE_12" +
                        "         WHEN PRIMARY_STAT_13 = 'Y' THEN STAT_TYPE_13" +
                        "         WHEN PRIMARY_STAT_14 = 'Y' THEN STAT_TYPE_14" +
                        "         WHEN PRIMARY_STAT_15 = 'Y' THEN STAT_TYPE_15" +
                        "         WHEN PRIMARY_STAT_16 = 'Y' THEN STAT_TYPE_16" +
                        "         WHEN PRIMARY_STAT_17 = 'Y' THEN STAT_TYPE_17" +
                        "         WHEN PRIMARY_STAT_18 = 'Y' THEN STAT_TYPE_18" +
                        "         WHEN PRIMARY_STAT_19 = 'Y' THEN STAT_TYPE_19" +
                        "         WHEN PRIMARY_STAT_20 = 'Y' THEN STAT_TYPE_20" +
                        "       END STAT_TYPE_PR" +
                        "     FROM " + dbReference + ".MSF617_OP_STATS" + dbLink + " OPS WHERE EGI_REC_TYPE = 'G' AND EQUIP_GRP_ID = (SELECT EQUIP_GRP_ID FROM " + dbReference + ".MSF600" + dbLink + " WHERE EQUIP_NO = WO.EQUIP_NO)) G_STAT_TYPE_PR" +
                        "   FROM " + dbReference + ".MSF623" + dbLink + " WOT JOIN " + dbReference + ".MSF620" + dbLink + " WO ON WOT.DSTRCT_CODE = WO.DSTRCT_CODE AND WOT.WORK_ORDER = WO.WORK_ORDER AND WOT.WORK_GROUP = WO.WORK_GROUP" +
                        "   WHERE WOT.WORK_ORDER = '" + task.WorkOrder + "' AND WOT.DSTRCT_CODE = '" + task.DstrctCode + "' AND WOT.WO_TASK_NO = '" + task.WoTaskNo.PadLeft(3, '0') + "'" +
                        " ),";
                
            }
            else if (!string.IsNullOrWhiteSpace(task.MaintSchTask))
            {
                headerQuery = "" +
                        " WITH TASK AS(" +
                        "   SELECT " +
                        "     '" + task.DstrctCode + "' DSTRCT_CODE," +
                        "     '" + task.WorkGroup + "' WORK_GROUP," +
                        "     '" + task.EquipNo + "' EQUIP_NO," +
                        "     '" + task.CompCode + "' COMP_CODE," +
                        "     '" + task.CompModCode + "' COMP_MOD_CODE," +
                        "     '" + task.MaintSchTask + "' MAINT_SCH_TASK," +
                        "     '" + task.StdJobNo + "' STD_JOB_NO," +
                        "     '" + task.WorkOrder + "' WORK_ORDER," +
                        "     '" + task.StdJobTask + "'  STD_JOB_TASK," +
                        "     '" + task.PlanStrDate + "' PLAN_STR_DATE," +
                        "     '" + task.OriginalPlannedStartDate + "' ORIG_SCHED_DATE," +
                        "     '" + task.PlanStrDate + "' REQ_START_DATE," +
                        "     '" + task.PlanStrDate + "' REQ_BY_DATE," +
                        "     '' COMPLETED_CODE" +
                        "   FROM DUAL" +
                        " )," +
                        " ORDERS AS (" +
                        "   SELECT " +
                        "     TASK.DSTRCT_CODE," +
                        "     TASK.WORK_GROUP," +
                        "     TASK.EQUIP_NO," +
                        "     TASK.COMP_CODE," +
                        "     TASK.COMP_MOD_CODE," +
                        "     TASK.MAINT_SCH_TASK," +
                        "     SJ.STD_JOB_NO," +
                        "     TASK.WORK_ORDER," +
                        "     SJT.STD_JOB_TASK WO_TASK_NO," +
                        "     TASK.PLAN_STR_DATE," +
                        "     TASK.ORIG_SCHED_DATE," +
                        "     TASK.REQ_START_DATE," +
                        "     TASK.REQ_BY_DATE," +
                        "     TASK.COMPLETED_CODE," +
                        "     SJ.ASSIGN_PERSON WO_ASSIGN_PERSON," +
                        "     SJT.ASSIGN_PERSON," +
                        "     SJ.MAINT_TYPE," +
                        "     SJ.WO_TYPE," +
                        "     (SELECT" +
                        "       CASE" +
                        "         WHEN PRIMARY_STAT_1 = 'Y' THEN STAT_TYPE_1" +
                        "         WHEN PRIMARY_STAT_2 = 'Y' THEN STAT_TYPE_2" +
                        "         WHEN PRIMARY_STAT_3 = 'Y' THEN STAT_TYPE_3" +
                        "         WHEN PRIMARY_STAT_4 = 'Y' THEN STAT_TYPE_4" +
                        "         WHEN PRIMARY_STAT_5 = 'Y' THEN STAT_TYPE_5" +
                        "         WHEN PRIMARY_STAT_6 = 'Y' THEN STAT_TYPE_6" +
                        "         WHEN PRIMARY_STAT_7 = 'Y' THEN STAT_TYPE_7" +
                        "         WHEN PRIMARY_STAT_8 = 'Y' THEN STAT_TYPE_8" +
                        "         WHEN PRIMARY_STAT_9 = 'Y' THEN STAT_TYPE_9" +
                        "         WHEN PRIMARY_STAT_10 = 'Y' THEN STAT_TYPE_10" +
                        "         WHEN PRIMARY_STAT_11 = 'Y' THEN STAT_TYPE_11" +
                        "         WHEN PRIMARY_STAT_12 = 'Y' THEN STAT_TYPE_12" +
                        "         WHEN PRIMARY_STAT_13 = 'Y' THEN STAT_TYPE_13" +
                        "         WHEN PRIMARY_STAT_14 = 'Y' THEN STAT_TYPE_14" +
                        "         WHEN PRIMARY_STAT_15 = 'Y' THEN STAT_TYPE_15" +
                        "         WHEN PRIMARY_STAT_16 = 'Y' THEN STAT_TYPE_16" +
                        "         WHEN PRIMARY_STAT_17 = 'Y' THEN STAT_TYPE_17" +
                        "         WHEN PRIMARY_STAT_18 = 'Y' THEN STAT_TYPE_18" +
                        "         WHEN PRIMARY_STAT_19 = 'Y' THEN STAT_TYPE_19" +
                        "         WHEN PRIMARY_STAT_20 = 'Y' THEN STAT_TYPE_20" +
                        "       END STAT_TYPE_PR" +
                        "     FROM " + dbReference + ".MSF617_OP_STATS" + dbLink + " OPS WHERE EGI_REC_TYPE = 'E' AND EQUIP_GRP_ID = TASK.EQUIP_NO) E_STAT_TYPE_PR," +
                        "         (SELECT" +
                        "       CASE" +
                        "         WHEN PRIMARY_STAT_1 = 'Y' THEN STAT_TYPE_1" +
                        "         WHEN PRIMARY_STAT_2 = 'Y' THEN STAT_TYPE_2" +
                        "         WHEN PRIMARY_STAT_3 = 'Y' THEN STAT_TYPE_3" +
                        "         WHEN PRIMARY_STAT_4 = 'Y' THEN STAT_TYPE_4" +
                        "         WHEN PRIMARY_STAT_5 = 'Y' THEN STAT_TYPE_5" +
                        "         WHEN PRIMARY_STAT_6 = 'Y' THEN STAT_TYPE_6" +
                        "         WHEN PRIMARY_STAT_7 = 'Y' THEN STAT_TYPE_7" +
                        "         WHEN PRIMARY_STAT_8 = 'Y' THEN STAT_TYPE_8" +
                        "         WHEN PRIMARY_STAT_9 = 'Y' THEN STAT_TYPE_9" +
                        "         WHEN PRIMARY_STAT_10 = 'Y' THEN STAT_TYPE_10" +
                        "         WHEN PRIMARY_STAT_11 = 'Y' THEN STAT_TYPE_11" +
                        "         WHEN PRIMARY_STAT_12 = 'Y' THEN STAT_TYPE_12" +
                        "         WHEN PRIMARY_STAT_13 = 'Y' THEN STAT_TYPE_13" +
                        "         WHEN PRIMARY_STAT_14 = 'Y' THEN STAT_TYPE_14" +
                        "         WHEN PRIMARY_STAT_15 = 'Y' THEN STAT_TYPE_15" +
                        "         WHEN PRIMARY_STAT_16 = 'Y' THEN STAT_TYPE_16" +
                        "         WHEN PRIMARY_STAT_17 = 'Y' THEN STAT_TYPE_17" +
                        "         WHEN PRIMARY_STAT_18 = 'Y' THEN STAT_TYPE_18" +
                        "         WHEN PRIMARY_STAT_19 = 'Y' THEN STAT_TYPE_19" +
                        "         WHEN PRIMARY_STAT_20 = 'Y' THEN STAT_TYPE_20" +
                        "       END STAT_TYPE_PR" +
                        "     FROM " + dbReference + ".MSF617_OP_STATS" + dbLink + " OPS WHERE EGI_REC_TYPE = 'G' AND EQUIP_GRP_ID = (SELECT EQUIP_GRP_ID FROM " + dbReference + ".MSF600" + dbLink + " WHERE EQUIP_NO = TASK.EQUIP_NO)) G_STAT_TYPE_PR" +
                        "   FROM TASK JOIN " + dbReference + ".MSF693" + dbLink + " SJT ON " +
                        "     TASK.DSTRCT_CODE = SJT.DSTRCT_CODE AND TASK.STD_JOB_NO = SJT.STD_JOB_NO AND TASK.STD_JOB_TASK = SJT.STD_JOB_TASK" +
                        "   JOIN " + dbReference + ".MSF690" + dbLink + " SJ ON " +
                        "     SJT.DSTRCT_CODE = SJ.DSTRCT_CODE AND SJT.STD_JOB_NO = SJ.STD_JOB_NO AND SJT.WORK_GROUP = SJ.WORK_GROUP" +
                        " ),";
            }
            else
                throw new ArgumentException("No work order or maintenance schedule task received", "WorkOrder, MaintSchTask");

            var query = headerQuery +
                        " MSTFULL AS (" +
                        "   SELECT MST.WORK_GROUP," +
                        "     MST.EQUIP_NO," +
                        "     MST.COMP_CODE," +
                        "     MST.COMP_MOD_CODE," +
                        "     MST.MAINT_SCH_TASK," +
                        "     MST.STD_JOB_NO," +
                        "     MST.JOB_DESC_CODE," +
                        "     MST.SCHED_IND_700," +
                        "     MST.STAT_TYPE_1," +
                        "     CASE" +
                        "       WHEN MST.SCHED_IND_700 IN ('7','8')" +
                        "         THEN MST.SCHED_FREQ_1 * 30" +
                        "       ELSE" +
                        "         CASE" +
                        "           WHEN SUBSTR(MST.MAINT_SCH_TASK, 1, 1) = '8' THEN SUM(MST.SCHED_FREQ_1) OVER(" +
                        "               PARTITION BY MST.WORK_GROUP, MST.EQUIP_NO, MST.COMP_CODE, MST.COMP_MOD_CODE, SUBSTR(MST.MAINT_SCH_TASK, 1, 2)" +
                        "               ORDER BY" +
                        "                   MST.MAINT_SCH_TASK" +
                        "               RANGE BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW" +
                        "           )" +
                        "           ELSE MST.SCHED_FREQ_1" +
                        "         END" +
                        "       END AS SCHED_FREQ_1," +
                        "     CASE" +
                        "       WHEN SUBSTR(MST.MAINT_SCH_TASK, 1, 1) IN ('9','8') AND MST.JOB_DESC_CODE = 'Z9'" +
                        "         THEN" +
                        "           CASE WHEN MST.SCHED_IND_700 IN ('7','8') " +
                        "             THEN MIN(MST.SCHED_FREQ_1 * 30) OVER(PARTITION BY MST.WORK_GROUP, MST.EQUIP_NO, MST.COMP_CODE, MST.COMP_MOD_CODE, SUBSTR(MST.MAINT_SCH_TASK, 1, 2), MST.JOB_DESC_CODE)" +
                        "           ELSE" +
                        "             MIN(MST.SCHED_FREQ_1) OVER(PARTITION BY MST.WORK_GROUP, MST.EQUIP_NO, MST.COMP_CODE, MST.COMP_MOD_CODE, SUBSTR(MST.MAINT_SCH_TASK, 1, 2), MST.JOB_DESC_CODE)" +
                        "           END" +
                        "       WHEN MST.SCHED_IND_700 IN ('7','8') THEN MST.SCHED_FREQ_1 * 30" +
                        "       ELSE MST.SCHED_FREQ_1" +
                        "     END AS MIN_FREQ" +
                        "   FROM " + dbReference + ".MSF700" + dbLink + " MST" +
                        "   WHERE MST.EQUIP_NO IN (SELECT EQUIP_NO FROM ORDERS)" +
                        "   AND MST.SCHED_IND_700 <> '9'" +
                        " )," +
                        " WOMST AS" +
                        " (" +
                        "   SELECT WO.*, " +
                        "     MST.JOB_DESC_CODE," +
                        "     MST.SCHED_IND_700," +
                        "     MST.STAT_TYPE_1," +
                        "     MST.SCHED_FREQ_1," +
                        "     MST.MIN_FREQ," +
                        "     (COALESCE(MST.STAT_TYPE_1, COALESCE(WO.E_STAT_TYPE_PR, WO.G_STAT_TYPE_PR))) EQ_STAT_TYPE_PR," +
                        "     (" +
                        "       SELECT" +
                        "         MAX(STAT.CUM_VALUE)" +
                        "       FROM" +
                        "         " + dbReference + ".MSF400" + dbLink + " STAT" +
                        "       WHERE" +
                        "         STAT.EQUIP_NO = WO.EQUIP_NO" +
                        "         AND STAT.STAT_TYPE  = (COALESCE(MST.STAT_TYPE_1, COALESCE(WO.E_STAT_TYPE_PR, WO.G_STAT_TYPE_PR)))" +
                        "         AND STAT.STAT_DATE  = (" +
                        "           SELECT" +
                        "             MAX(" + dbReference + ".MSF400.STAT_DATE)" +
                        "           FROM" +
                        "             " + dbReference + ".MSF400" + dbLink + "" +
                        "           WHERE" +
                        "             " + dbReference + ".MSF400.EQUIP_NO = WO.EQUIP_NO" +
                        "             AND " + dbReference + ".MSF400.STAT_TYPE     = (COALESCE(MST.STAT_TYPE_1, COALESCE(WO.E_STAT_TYPE_PR, WO.G_STAT_TYPE_PR)))" +
                        "             AND " + dbReference + ".MSF400.STAT_DATE <= WO.ORIG_SCHED_DATE" +
                        "             AND " + dbReference + ".MSF400.KEY_400_TYPE  = 'E'" +
                        "         )" +
                        "     ) AS SCHED_STAT_VALUE," +
                        "     (" +
                        "       SELECT" +
                        "         MAX(STAT.CUM_VALUE)" +
                        "       FROM" +
                        "         " + dbReference + ".MSF400" + dbLink + " STAT" +
                        "       WHERE" +
                        "         STAT.EQUIP_NO = WO.EQUIP_NO" +
                        "         AND STAT.STAT_TYPE  = (COALESCE(MST.STAT_TYPE_1, COALESCE(WO.E_STAT_TYPE_PR, WO.G_STAT_TYPE_PR)))" +
                        "         AND STAT.STAT_DATE  = (" +
                        "           SELECT" +
                        "             MAX(" + dbReference + ".MSF400.STAT_DATE)" +
                        "           FROM" +
                        "             " + dbReference + ".MSF400" + dbLink + "" +
                        "           WHERE" +
                        "             " + dbReference + ".MSF400.EQUIP_NO = WO.EQUIP_NO" +
                        "             AND " + dbReference + ".MSF400.STAT_TYPE     = (COALESCE(MST.STAT_TYPE_1, COALESCE(WO.E_STAT_TYPE_PR, WO.G_STAT_TYPE_PR)))" +
                        "             AND " + dbReference + ".MSF400.STAT_DATE <= TO_CHAR(SYSDATE, 'YYYYMMDD')" +
                        "             AND " + dbReference + ".MSF400.KEY_400_TYPE  = 'E'" +
                        "         )" +
                        "     ) AS ACTUAL_STAT_VALUE" +
                        "   FROM ORDERS WO LEFT JOIN MSTFULL MST " +
                        "     ON  NVL(TRIM(WO.EQUIP_NO), ' ') = NVL(TRIM(MST.EQUIP_NO), ' ')             " +
                        "     AND NVL(TRIM(WO.COMP_CODE), ' ') =  NVL(TRIM(MST.COMP_CODE), ' ')          " +
                        "     AND NVL(TRIM(WO.COMP_MOD_CODE), ' ') = NVL(TRIM(MST.COMP_MOD_CODE), ' ')   " +
                        "     AND NVL(TRIM(WO.MAINT_SCH_TASK), ' ') = NVL(TRIM(MST.MAINT_SCH_TASK), ' ') " +
                        " )," +
                        " WOMST_DATES AS (" +
                        "   SELECT" +
                        "     WM.*, 'DATE' STYPE," +
                        "     CASE" +
                        "       WHEN TRIM(WM.SCHED_FREQ_1) IS NULL THEN TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') - 7), 'DD'), 'YYYYMMDD')" +
                        "       WHEN WM.SCHED_FREQ_1 * 0.10 >= WM.MIN_FREQ THEN TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') - WM.MIN_FREQ), 'DD'), 'YYYYMMDD')" +
                        "       ELSE TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') - FLOOR(WM.SCHED_FREQ_1 * 0.10 + 1)), 'DD'), 'YYYYMMDD')" +
                        "     END AS MIN_SCHED_DT," +
                        "     CASE" +
                        "       WHEN TRIM(WM.SCHED_FREQ_1) IS NULL THEN TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') + 7), 'DD'), 'YYYYMMDD')" +
                        "       WHEN WM.SCHED_FREQ_1 * 0.10 >= WM.MIN_FREQ THEN TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') + WM.MIN_FREQ), 'DD'), 'YYYYMMDD')" +
                        "       ELSE TO_CHAR(TRUNC((TO_DATE(WM.ORIG_SCHED_DATE, 'YYYYMMDD') + FLOOR(WM.SCHED_FREQ_1 * 0.10 + 1)), 'DD'), 'YYYYMMDD')" +
                        "     END AS MAX_SCHED_DT," +
                        "     0 AS MIN_SCH_STAT," +
                        "     0 AS MAX_SCH_STAT" +
                        "   FROM" +
                        "   WOMST WM" +
                        "   WHERE" +
                        "     TRIM(WM.STAT_TYPE_1) IS NULL" +
                        " ), WOMST_STATS AS (" +
                        "     SELECT" +
                        "         WM.*, 'STAT' STYPE," +
                        "         NULL AS MIN_SCHED_DT," +
                        "         NULL AS MAX_SCHED_DT," +
                        "         CASE" +
                        "             WHEN TRIM(WM.STAT_TYPE_1) IS NULL THEN 0" +
                        "             WHEN WM.SCHED_FREQ_1 * 0.10 >= WM.MIN_FREQ THEN ( WM.SCHED_STAT_VALUE - FLOOR((WM.MIN_FREQ) + 1) )" +
                        "             ELSE ( WM.SCHED_STAT_VALUE - FLOOR((WM.SCHED_FREQ_1 * 0.10) + 1) )" +
                        "         END AS MIN_SCH_STAT," +
                        "         CASE" +
                        "             WHEN TRIM(WM.STAT_TYPE_1) IS NULL THEN 0" +
                        "             WHEN WM.SCHED_FREQ_1 * 0.10 >= WM.MIN_FREQ THEN ( WM.SCHED_STAT_VALUE + FLOOR((WM.MIN_FREQ) + 1) )" +
                        "             ELSE ( WM.SCHED_STAT_VALUE + FLOOR((WM.SCHED_FREQ_1 * 0.10) + 1) )" +
                        "         END AS MAX_SCH_STAT" +
                        "     FROM" +
                        "         WOMST WM" +
                        "     WHERE" +
                        "         TRIM(WM.STAT_TYPE_1) IS NOT NULL" +
                        " )" +
                        " SELECT * FROM WOMST_DATES UNION ALL SELECT * FROM WOMST_STATS";

            return query;
        }
    }
}

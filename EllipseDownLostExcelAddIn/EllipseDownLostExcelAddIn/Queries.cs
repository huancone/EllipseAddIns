using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseDownLostExcelAddIn
{
    public class Queries
    {
        /// <summary>
        /// Obtiene el query para el listado de Down de un equipo dado
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <param name="equipmentNo"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns>string: Query para el listado de Down de un equipo dado</returns>
        public static string GetEquipmentDownQuery(string dbreference, string dblink, string equipmentNo, string startDate, string endDate)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Down, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Down por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Down
            var query = "" +
                " SELECT" +
                "     SUBSTR(DW.REC_EQUIP_420,2) EQUIP_NO, DW.COMP_CODE, DW.COMP_MOD_CODE," +
                "     (99999999 - DW.REV_STAT_DATE) START_DATE, DW.STOP_TIME START_TIME," +
                "     (99999999 - DW.REV_STAT_DATE) FINISH_DATE, DW.START_TIME FINISH_TIME, DW.ELAPSED_HOURS, DW.SHIFT, 'DOWN' EVENT_TYPE," +
                "     DW.DOWN_TIME_CODE EVENT_CODE, COD.TABLE_DESC DESCRIPTION, (DW.WORK_ORDER || ' - ' || WO.WO_DESC) WO_COMMENT," +
                "     DW.SEQUENCE_NO, DW.SHIFT_SEQ_NO, WO.WO_STATUS_M" +
                "   FROM" +
                "     " + dbreference + ".MSF420" + dblink + " DW" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(DW.DOWN_TIME_CODE) = TRIM(COD.TABLE_CODE)" +
                "   LEFT JOIN ELLIPSE.MSF620 WO ON DW.WORK_ORDER = WO.WORK_ORDER" +
                "   WHERE" +
                "     DW.REC_EQUIP_420 = 'E' ||'" + equipmentNo + "'" +
                "     AND (99999999 - DW.REV_STAT_DATE) BETWEEN '" + startDate + "' AND '" + endDate + "'" +
                "     AND COD.TABLE_TYPE = 'DT'" +
                "   ORDER BY EQUIP_NO, COMP_CODE, COMP_MOD_CODE, START_DATE, START_TIME";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        /// <summary>
        /// Obtiene el query para el listado de Lost Production de un equipo dado
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <param name="equipmentNo"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns>string: Query para el listado de Lost de un equipo dado</returns>
        public static string GetEquipmentLostQuery(string dbreference, string dblink, string equipmentNo, string startDate, string endDate)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Lost, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Lost por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Lost
            var query = "" +
                " SELECT" +
                " LS.EQUIP_NO, '' COMP_CODE, '' COMP_MOD_CODE," +
                " (99999999 - LS.REV_STAT_DATE) START_DATE, LS.STOP_TIME START_TIME," +
                " (99999999 - LS.REV_STAT_DATE) FINISH_DATE, LS.START_TIME FINISH_TIME, LS.ELAPSED_HOURS, LS.SHIFT, 'LOST' EVENT_TYPE," +
                " LS.LOST_PROD_CODE EVENT_CODE, COD.TABLE_DESC DESCRIPTION, " +
                " (SELECT TRIM(LPTEXT.STD_MEDIUM_1 || LPTEXT.STD_MEDIUM_2 || LPTEXT.STD_MEDIUM_3 || LPTEXT.STD_MEDIUM_4 || LPTEXT.STD_MEDIUM_5) FROM ELLIPSE.MSF096_STD_MEDIUM LPTEXT WHERE LPTEXT.STD_TEXT_CODE = 'LP' AND ROWNUM = 1 AND LPTEXT.STD_KEY = RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) WO_COMMENT," +
                " LS.SEQUENCE_NO, LS.SHIFT_SEQ_NO" +
                "   FROM" +
                "     " + dbreference + ".MSF470" + dblink + " LS" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(LS.LOST_PROD_CODE) = TRIM(COD.TABLE_CODE)" +
                "   WHERE" +
                "     LS.EQUIP_NO = '" + equipmentNo + "'" +
                "     AND (99999999 - LS.REV_STAT_DATE) BETWEEN '" + startDate + "' AND '" + endDate + "'" +
                "     AND COD.TABLE_TYPE = 'LP'" +
                "   ORDER BY EQUIP_NO, COMP_CODE, COMP_MOD_CODE, START_DATE, START_TIME";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        /// <summary>
        /// Obtiene el query listado de Códigos Down del sistema
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <returns>string: Query para el listado de Códigos Down del Sistema</returns>
        public static string GetDownTimeCodeListQuery(string dbreference, string dblink)
        {
            var query = "" +
            " SELECT" +
            "   CASE COD.TABLE_TYPE WHEN 'LP' THEN 'LOST' WHEN 'DT' THEN 'DOWN' END EVENT_TYPE," +
            "   TRIM(COD.TABLE_CODE) CODE, " +
            "   TRIM(COD.TABLE_DESC) DESCRIPTION" +
            " FROM " + dbreference + ".MSF010" + dblink + " COD" +
            " WHERE TABLE_TYPE = 'DT'" +
            " ORDER BY TABLE_TYPE, TABLE_CODE, TABLE_DESC";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        /// <summary>
        /// Obtiene el query listado de Códigos Lost Production del sistema
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <returns>string: Query para el listado de Códigos Lost del Sistema</returns>
        public static string GetLostProdCodeListQuery(string dbreference, string dblink)
        {
            var query = "" +
            " SELECT" +
            "   CASE COD.TABLE_TYPE WHEN 'LP' THEN 'LOST' WHEN 'DT' THEN 'DOWN' END EVENT_TYPE," +
            "   TRIM(COD.TABLE_CODE) CODE, " +
            "   TRIM(COD.TABLE_DESC) DESCRIPTION" +
            " FROM " + dbreference + ".MSF010" + dblink + " COD" +
            " WHERE TABLE_TYPE = 'LP'" +
            " ORDER BY TABLE_TYPE, TABLE_CODE, TABLE_DESC";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        [Obsolete("Utilizado para los métodos que utilizan los servicios de Down directamente. Se marca obsoleto porque el sistema no establece comunicación con el servicio")]
        public static string GetSingleDownQuery(string dbreference, string dblink, string equipmentNo, string downCode, string startDate, string shiftCode, string startTime, string endTime)
        {
            var query = "" +
                "   SELECT" +
                "     SUBSTR(DW.REC_EQUIP_420,2) EQUIP_NO, DW.COMP_CODE, DW.COMP_MOD_CODE," +
                "     DW.SEQUENCE_NO, DW.SHIFT_SEQ_NO, " +
                "     (99999999 - DW.REV_STAT_DATE) START_DATE, DW.STOP_TIME START_TIME," +
                "     (99999999 - DW.REV_STAT_DATE) FINISH_DATE, DW.START_TIME FINISH_TIME, DW.ELAPSED_HOURS, DW.SHIFT, 'DOWN' EVENT_TYPE," +
                "     DW.DOWN_TIME_CODE EVENT_CODE, DW.WORK_ORDER WO_COMMENT" +
                "   FROM" +
                "     " + dbreference + ".MSF420" + dblink + " DW" +
                "   WHERE" +
                "     DW.REC_EQUIP_420 = 'E' ||'" + equipmentNo + "'" +
                "     AND DW.DOWN_TIME_CODE = '" + downCode + "'" +
                "     AND (99999999 - DW.REV_STAT_DATE) = '" + startDate + "'" +
                "     AND DW.SHIFT = '" + shiftCode + "'" +
                "     AND DW.STOP_TIME = LPAD('" + startTime + "', 4, '0')" +
                "     AND DW.START_TIME = LPAD('" + endTime + "', 4, '0')";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetSingleLostQuery(string dbreference, string dblink, string equipmentNo, string lostCode, string startDate, string shiftCode, string startTime, string endTime)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Lost, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Lost por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Lost
            var query = "" +
                " SELECT" +
                " LS.EQUIP_NO, '' COMP_CODE, '' COMP_MOD_CODE," +
                " LS.SEQUENCE_NO, LS.SHIFT_SEQ_NO," +
                " (99999999 - LS.REV_STAT_DATE) START_DATE, LS.STOP_TIME START_TIME," +
                " (99999999 - LS.REV_STAT_DATE) FINISH_DATE, LS.START_TIME FINISH_TIME, LS.ELAPSED_HOURS, LS.SHIFT, 'LOST' EVENT_TYPE," +
                " LS.LOST_PROD_CODE EVENT_CODE, " +
                " (SELECT TRIM(LPTEXT.STD_MEDIUM_1 || LPTEXT.STD_MEDIUM_2 || LPTEXT.STD_MEDIUM_3 || LPTEXT.STD_MEDIUM_4 || LPTEXT.STD_MEDIUM_5) FROM ELLIPSE.MSF096_STD_MEDIUM LPTEXT WHERE LPTEXT.STD_TEXT_CODE = 'LP' AND LPTEXT.STD_KEY = RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) WO_COMMENT," +
                " (RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) STD_KEY " +
                "   FROM" +
                "     " + dbreference + ".MSF470" + dblink + " LS" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(LS.LOST_PROD_CODE) = TRIM(COD.TABLE_CODE)" +
                "   WHERE" +
                "     LS.EQUIP_NO = '" + equipmentNo + "'" +
                "     AND COD.TABLE_TYPE = 'LP'" +
                "     AND LS.LOST_PROD_CODE = '" + lostCode + "'" +
                "     AND (99999999 - LS.REV_STAT_DATE) = '" + startDate + "'" +
                "     AND LS.SHIFT = '" + shiftCode + "'" +
                "     AND LS.STOP_TIME = LPAD('" + startTime + "', 4, '0')" +
                "     AND LS.START_TIME = LPAD('" + endTime + "', 4, '0')";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        /// <summary>
        /// Consulta los eventos de la red industrial de PBV
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static string GetDownLostPbv(string startDate, string endDate)
        {
            var query = "" +
            "WITH " +
            " SHIFT AS " +
            " ( " +
            "   SELECT " +
            "     CONVERT( DATETIME, '" + startDate + " 06:00:00', 20 ) STARTTIME, " +
            "     CONVERT( DATETIME, '" + endDate + " 06:00:00', 20 ) ENDTIME " +
            " ) " +
            " , " +
            " SC AS " +
            " ( " +
            "   SELECT " +
            "     CASE SCC.ID " +
            "       WHEN 13 " +
            "       THEN 'LOST' " +
            "       WHEN 15 " +
            "       THEN 'LOST' " +
            "       WHEN 14 " +
            "       THEN 'DOWN' " +
            "       WHEN 23 " +
            "       THEN 'DOWN' " +
            "       ELSE 'DOWN' " +
            "     END TIPO, " +
            "     SCC.ID, " +
            "     SCC.DESCRIPTION " +
            "   FROM " +
            "     SCADARDB.DBO.STATUSCHANGECAUSE SCC " +
            "   WHERE " +
            "     SCC.PARENTID IS NULL " +
            "   UNION ALL " +
            "   SELECT " +
            "     SC.TIPO, " +
            "     SCC.ID, " +
            "     SCC.DESCRIPTION " +
            "   FROM " +
            "     SC " +
            "   INNER JOIN SCADARDB.DBO.STATUSCHANGECAUSE SCC " +
            "   ON " +
            "     SC.ID = SCC.PARENTID " +
            " )" +
            " , " +
            " PUSH AS " +
            " ( " +
            "   SELECT " +
            "     ROW_NUMBER( ) OVER( PARTITION BY PU.ASSETID ORDER BY PUSH.TIMESTAMP ASC ) ROWNUMBER, " +
            "     PUSH.TIMESTAMP, " +
            "     PU.ASSETID, " +
            "     PUSH.PRODUCTIVEUNITSTATUSTAGVALUE PUS, " +
            "     EVENTSEQUENCEID " +
            "   FROM " +
            "     PRODUCTIVEUNITSSTATUSHISTORY PUSH " +
            "   INNER JOIN PRODUCTIVEUNITS PU " +
            "   ON " +
            "     PU.PRODUCTIVEUNITSID = PUSH.PRODUCTIVEUNITID " +
            "   INNER JOIN PRODUCTIVEUNITSFUNCTIONSLIST PUFL " +
            "   ON " +
            "     PUFL.PRODUCTIVEUNITSID = PUSH.PRODUCTIVEUNITID " +
            "   INNER JOIN DBO.IFIXSTATUSCODESHISTORY SCH " +
            "   ON " +
            "     PUFL.STATUSCODE_TAGNAME = SCH.FUNCTIONSTATUSTAGNAME " +
            "   AND PUSH.TIMESTAMP        = SCH.TIMESTAMP" +
            " )" +
            " , " +
            " PU_SEL AS " +
            " ( " +
            "   SELECT " +
            "     PUSH.TIMESTAMP   STARTTIME, " +
            "     CASE WHEN PUSH_1.TIMESTAMP IS NULL THEN GETDATE()  ELSE PUSH_1.TIMESTAMP END ENDTIME, " +
            "     PUSH.ASSETID     PUASSETID, " +
            "     PUSH.EVENTSEQUENCEID " +
            "   FROM " +
            "     PUSH " +
            "   LEFT JOIN PUSH PUSH_1 " +
            "   ON " +
            "     PUSH.ASSETID         = PUSH_1.ASSETID " +
            "   AND PUSH.ROWNUMBER + 1 = PUSH_1.ROWNUMBER" +
            "   WHERE " +
            "     PUSH.PUS                = 60000 " +
            "   OR( PUSH.EVENTSEQUENCEID IS NOT NULL " +
            "   AND PUSH.PUS              = 30000 )" +
            " )" +
            " , " +
            " PU_EVENT AS " +
            " ( " +
            "   SELECT " +
            "     PUASSET.ASSETDESC EQUIPMENT, " +
            "     REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) COMPONENT, " +
            "     CASE " +
            "       WHEN PU_SEL.STARTTIME < SHIFT.STARTTIME " +
            "       THEN SHIFT.STARTTIME " +
            "       ELSE PU_SEL.STARTTIME " +
            "     END STARTTIME, " +
            "     CASE " +
            "       WHEN PU_SEL.ENDTIME > SHIFT.ENDTIME " +
            "       THEN SHIFT.ENDTIME " +
            "       ELSE PU_SEL.ENDTIME " +
            "     END ENDTIME, " +
            "     CASE SC.TIPO " +
            "       WHEN 'DOWN' " +
            "       THEN 'F' + REPLICATE( '0', 3 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '    ' ) " +
            "       WHEN 'LOST' " +
            "       THEN 'L' + REPLICATE( '0', 3 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '    ' ) " +
            "       ELSE 'DW' " +
            "     END FAILURE, " +
            "     CASE " +
            "       WHEN SYMPTOMS.SYMPTOMID IS NULL " +
            "       THEN NULL " +
            "       ELSE 'S' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, SYMPTOMS.SYMPTOMID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SYMPTOMS.SYMPTOMID ), '0' ) " +
            "     END SYMPTOMID, " +
            "     CASE " +
            "       WHEN FAILUREASSET.ASSETTYPEID IS NULL " +
            "       THEN NULL " +
            "       ELSE 'P' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) " +
            "     END ASSETTYPEID, " +
            "     CASE " +
            "       WHEN SC.ID IS NULL " +
            "       THEN NULL " +
            "       ELSE 'C' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) " +
            "     END STATUSCHANGEID, " +
            "     CASE WHEN SC.TIPO = 'DOWN' " +
            "       THEN 'EP' + REPLICATE( '0', 6 - LEN( ISNULL( CONVERT( VARCHAR, EH.EVENTSEQUENCEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, EH.EVENTSEQUENCEID ), '0' ) " +
            "       ELSE NULL " +
            "     END EVENT, " +
            "     SC.TIPO, " +
            "     SC.DESCRIPTION, " +
            "     EH.COMMENT " +
            "   FROM " +
            "     PU_SEL " +
            "   INNER JOIN SCADARDB.DBO.ASSETS PUASSET " +
            "   ON " +
            "     PU_SEL.PUASSETID = PUASSET.ASSETID " +
            "   LEFT JOIN SCADARDB.DBO.EVENTSHISTORY EH " +
            "   ON " +
            "     PU_SEL.EVENTSEQUENCEID = EH.EVENTSEQUENCEID " +
            "   LEFT JOIN SCADARDB.DBO.ASSETS FAILUREASSET " +
            "   ON " +
            "     EH.FAILEDASSETID = FAILUREASSET.ASSETID " +
            "   LEFT JOIN SC " +
            "   ON " +
            "     EH.FAILEDASSETFAILUREMODEID = SC.ID " +
            "   LEFT JOIN SYMPTOMS " +
            "   ON " +
            "     EH.SYMPTOMID = SYMPTOMS.SYMPTOMID" +
            "   INNER JOIN SHIFT " +
            "   ON " +
            "     SHIFT.STARTTIME <= PU_SEL.ENDTIME " +
            "   AND SHIFT.ENDTIME >= PU_SEL.STARTTIME  " +
            " )" +
            "SELECT " +
            " PU_EVENT.EQUIPMENT, " +
            " PU_EVENT.COMPONENT, " +
            " '' COMP_MOD_CODE, " +
            " CONVERT( VARCHAR, PU_EVENT.STARTTIME, 112 ) STAR_DATE, " +
            " REPLACE( CONVERT( VARCHAR( 5 ), PU_EVENT.STARTTIME, 108 ), ':', '' ) STAR_TIME, " +
            " CONVERT( VARCHAR, PU_EVENT.ENDTIME, 112 ) FINISH_DATE, " +
            " REPLACE( CONVERT( VARCHAR( 5 ), PU_EVENT.ENDTIME, 108 ), ':', '' ) FINISH_TIME, " +
            " '' ELAPSED, " +
            " '' COLLECTION, " +
            " 'A' SHIFT, " +
            " PU_EVENT.TIPO EVENT_TYPE, " +
            " PU_EVENT.FAILURE EVENT_CODE, " +
            " PU_EVENT.DESCRIPTION EVENT_DESC, " +
            " PU_EVENT.COMMENT, " +
            " PU_EVENT.EVENT, " +
            " PU_EVENT.SYMPTOMID, " +
            " PU_EVENT.ASSETTYPEID, " +
            " PU_EVENT.STATUSCHANGEID " +
            "FROM " +
            " PU_EVENT";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

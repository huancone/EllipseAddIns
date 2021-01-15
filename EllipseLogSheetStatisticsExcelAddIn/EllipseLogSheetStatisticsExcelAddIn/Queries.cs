using SharedClassLibrary.Utilities;

namespace EllipseLogSheetStatisticsExcelAddIn
{
    public static class Queries
    {
        public static string GetQueryDefaultModelData(string modelCode, string dbReference, string dbLink)
        {
            var query = "" +
                    " WITH" +
                    "     EGI AS" +
                    "     (" +
                    "     SELECT" +
                    "         LSG.MODEL_CODE," +
                    "         LSG.MODEL_SEQ_NO," +
                    "         LSG.ENTRY_GRP" +
                    "     FROM" +
                    "         " + dbReference + ".MSF460" + dbLink + " LSG" +
                    "     WHERE" +
                    "         LSG.MODEL_CODE      = '" + modelCode + "'" +
                    "     AND LSG.REC_460_TYPE   = 'L'" +
                    "     AND LSG.ENTRY_460_TYPE = 'G'" +
                    "     )" +
                    "     ," +
                    "     LSER AS" +
                    "     (" +
                    "     SELECT" +
                    "         LSE.MODEL_CODE," +
                    "         '' AS EGI," +
                    "         LSE.MODEL_SEQ_NO," +
                    "         TRIM(LSE.ENTRY_GRP) ENTRY_GRP" +
                    "     FROM" +
                    "         " + dbReference + ".MSF460" + dbLink + " LSE" +
                    "     WHERE" +
                    "         LSE.MODEL_CODE      = '" + modelCode + "'" +
                    "     AND LSE.REC_460_TYPE   = 'L'" +
                    "     AND LSE.ENTRY_460_TYPE = 'E'" +
                    "     ORDER BY" +
                    "         MODEL_SEQ_NO ASC" +
                    "     )" +
                    "     ," +
                    "     LSGR AS" +
                    "     (" +
                    "     SELECT" +
                    "         EGI.MODEL_CODE," +
                    "         TRIM(EGI.ENTRY_GRP) AS EGI," +
                    "         EGI.MODEL_SEQ_NO," +
                    "         TRIM(EQS.EQUIP_NO) ENTRY_GRP" +
                    "     FROM" +
                    "         " + dbReference + ".MSF600" + dbLink + " EQS" +
                    "     JOIN EGI" +
                    "     ON" +
                    "         TRIM(EQS.EQUIP_GRP_ID) = TRIM(EGI.ENTRY_GRP)" +
                    "     )" +
                    "     ," +
                    "     MODEL_EQUIP AS" +
                    "     (" +
                    "     SELECT" +
                    "         LE.MODEL_CODE, LE.EGI, LE.MODEL_SEQ_NO, LE.ENTRY_GRP," +
                    "         DEFV.OPERATOR_FLG, DEFV.OPERATOR_ID," +
                    "         DEFV.ACCOUNT_FLG, DEFV.ACCOUNT_CODE," +
                    "         DEFV.WORK_ORDER_FLG, DEFV.WORK_ORDER," +
                    "         DEFV.SOURCE_LOC_FLG, DEFV.SOURCE_LOC," +
                    "         DEFV.DEST_LOC_FLG, DEFV.DEST_LOC," +
                    "         DEFV.MATERIAL_FLG, DEFV.MATERIAL_CODE," +
                    "         DEFV.STAT_VALUE_1, DEFV.STAT_IO_FLG_1," +
                    "         DEFV.STAT_VALUE_2, DEFV.STAT_IO_FLG_2," +
                    "         DEFV.STAT_VALUE_3, DEFV.STAT_IO_FLG_3," +
                    "         DEFV.STAT_VALUE_4, DEFV.STAT_IO_FLG_4," +
                    "         DEFV.STAT_VALUE_5, DEFV.STAT_IO_FLG_5," +
                    "         DEFV.STAT_VALUE_6, DEFV.STAT_IO_FLG_6," +
                    "         DEFV.STAT_VALUE_7, DEFV.STAT_IO_FLG_7," +
                    "         DEFV.STAT_VALUE_8, DEFV.STAT_IO_FLG_8," +
                    "         DEFV.STAT_VALUE_9, DEFV.STAT_IO_FLG_9," +
                    "         DEFV.STAT_VALUE_10, DEFV.STAT_IO_FLG_10" +
                    "     FROM" +
                    "         LSER LE" +
                    "     JOIN " + dbReference + ".MSF615" + dbLink + " DEFV" +
                    "     ON" +
                    "         LE.ENTRY_GRP = TRIM(DEFV.EQUIP_NO)" +
                    "     WHERE" +
                    "         DEFV.EGI_REC_TYPE = 'E'" +
                    "     UNION" +
                    "     SELECT" +
                    "         LE.MODEL_CODE, LE.EGI, LE.MODEL_SEQ_NO, LE.ENTRY_GRP," +
                    "         DEFV.OPERATOR_FLG, DEFV.OPERATOR_ID," +
                    "         DEFV.ACCOUNT_FLG, DEFV.ACCOUNT_CODE," +
                    "         DEFV.WORK_ORDER_FLG, DEFV.WORK_ORDER," +
                    "         DEFV.SOURCE_LOC_FLG, DEFV.SOURCE_LOC," +
                    "         DEFV.DEST_LOC_FLG, DEFV.DEST_LOC," +
                    "         DEFV.MATERIAL_FLG, DEFV.MATERIAL_CODE," +
                    "         DEFV.STAT_VALUE_1, DEFV.STAT_IO_FLG_1," +
                    "         DEFV.STAT_VALUE_2, DEFV.STAT_IO_FLG_2," +
                    "         DEFV.STAT_VALUE_3, DEFV.STAT_IO_FLG_3," +
                    "         DEFV.STAT_VALUE_4, DEFV.STAT_IO_FLG_4," +
                    "         DEFV.STAT_VALUE_5, DEFV.STAT_IO_FLG_5," +
                    "         DEFV.STAT_VALUE_6, DEFV.STAT_IO_FLG_6," +
                    "         DEFV.STAT_VALUE_7, DEFV.STAT_IO_FLG_7," +
                    "         DEFV.STAT_VALUE_8, DEFV.STAT_IO_FLG_8," +
                    "         DEFV.STAT_VALUE_9, DEFV.STAT_IO_FLG_9," +
                    "         DEFV.STAT_VALUE_10, DEFV.STAT_IO_FLG_10" +
                    "     FROM" +
                    "         LSGR LE" +
                    "     JOIN " + dbReference + ".MSF615" + dbLink + " DEFV" +
                    "     ON" +
                    "         TRIM(LE.EGI) = TRIM(DEFV.EQUIP_NO)" +
                    "     WHERE" +
                    "         DEFV.EGI_REC_TYPE = 'G'" +
                    "     UNION" +
                    "     SELECT" +
                    "         LE.MODEL_CODE, LE.EGI, LE.MODEL_SEQ_NO, LE.ENTRY_GRP," +
                    "         DEFV.OPERATOR_FLG, DEFV.OPERATOR_ID," +
                    "         DEFV.ACCOUNT_FLG, DEFV.ACCOUNT_CODE," +
                    "         DEFV.WORK_ORDER_FLG, DEFV.WORK_ORDER," +
                    "         DEFV.SOURCE_LOC_FLG, DEFV.SOURCE_LOC," +
                    "         DEFV.DEST_LOC_FLG, DEFV.DEST_LOC," +
                    "         DEFV.MATERIAL_FLG, DEFV.MATERIAL_CODE," +
                    "         DEFV.STAT_VALUE_1, DEFV.STAT_IO_FLG_1," +
                    "         DEFV.STAT_VALUE_2, DEFV.STAT_IO_FLG_2," +
                    "         DEFV.STAT_VALUE_3, DEFV.STAT_IO_FLG_3," +
                    "         DEFV.STAT_VALUE_4, DEFV.STAT_IO_FLG_4," +
                    "         DEFV.STAT_VALUE_5, DEFV.STAT_IO_FLG_5," +
                    "         DEFV.STAT_VALUE_6, DEFV.STAT_IO_FLG_6," +
                    "         DEFV.STAT_VALUE_7, DEFV.STAT_IO_FLG_7," +
                    "         DEFV.STAT_VALUE_8, DEFV.STAT_IO_FLG_8," +
                    "         DEFV.STAT_VALUE_9, DEFV.STAT_IO_FLG_9," +
                    "         DEFV.STAT_VALUE_10, DEFV.STAT_IO_FLG_10" +
                    "       FROM" +
                    "         LSGR LE" +
                    "         LEFT JOIN " + dbReference + ".MSF615" + dbLink + " DEFV" +
                    "         ON" +
                    "         TRIM(LE.EGI) = TRIM(DEFV.EQUIP_NO)" +
                    "       WHERE DEFV.EQUIP_NO IS NULL" +
                    "     )      " +
                    "     SELECT" +
                    "       ME.*, TRIM(EQ.PLANT_NO) EQ_REFERENCE" +
                    "     FROM" +
                    "       MODEL_EQUIP ME LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON TRIM(ME.ENTRY_GRP) = TRIM(EQ.EQUIP_NO)" +
                    "     ORDER BY" +
                    "       ME.MODEL_CODE," +
                    "       ME.MODEL_SEQ_NO," +
                    "       ME.ENTRY_GRP";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetDefaultHeaderData(string modelCode, string dbReference, string dbLink)
        {
            var query = "" +
                    " WITH MODEL_COL AS" +
                    "   (SELECT MD.MODEL_CODE," +
                    "     MD.COLUMN_HEAD_1," +
                    "     MD.PROD_DT_TY_1," +
                    "     MD.COLUMN_HEAD_2," +
                    "     MD.PROD_DT_TY_2," +
                    "     MD.COLUMN_HEAD_3," +
                    "     MD.PROD_DT_TY_3," +
                    "     MD.COLUMN_HEAD_4," +
                    "     MD.PROD_DT_TY_4," +
                    "     MD.COLUMN_HEAD_5," +
                    "     MD.PROD_DT_TY_5," +
                    "     MD.COLUMN_HEAD_6," +
                    "     MD.PROD_DT_TY_6," +
                    "     MD.COLUMN_HEAD_7," +
                    "     MD.PROD_DT_TY_7," +
                    "     MD.COLUMN_HEAD_8," +
                    "     MD.PROD_DT_TY_8," +
                    "     MD.COLUMN_HEAD_9," +
                    "     MD.PROD_DT_TY_9," +
                    "     MD.COLUMN_HEAD_10," +
                    "     MD.PROD_DT_TY_10" +
                    "   FROM " + dbReference + ".MSF430" + dbLink + " MD" +
                    "   WHERE MD.MODEL_CODE = '" + modelCode + "'" +
                    "   ) ," +
                    "   COLUMN_NAME AS" +
                    "   (SELECT ROWNUM AS INDICE," +
                    "     MODEL_CODE," +
                    "     COLUMNAS," +
                    "     HEADER_NAME" +
                    "   FROM MODEL_COL UNPIVOT (HEADER_NAME FOR COLUMNAS IN ( COLUMN_HEAD_1, COLUMN_HEAD_2, COLUMN_HEAD_3, COLUMN_HEAD_4, COLUMN_HEAD_5, COLUMN_HEAD_6, COLUMN_HEAD_7, COLUMN_HEAD_8, COLUMN_HEAD_9, COLUMN_HEAD_10) )" +
                    "   )," +
                    "   COLUMN_TYPE AS" +
                    "   (SELECT ROWNUM AS INDICE," +
                    "     MODEL_CODE," +
                    "     COLUMNAS," +
                    "     VALUE_TYPE" +
                    "   FROM MODEL_COL UNPIVOT ( VALUE_TYPE FOR COLUMNAS IN (PROD_DT_TY_1, PROD_DT_TY_2, PROD_DT_TY_3, PROD_DT_TY_4, PROD_DT_TY_5, PROD_DT_TY_6, PROD_DT_TY_7, PROD_DT_TY_8, PROD_DT_TY_9, PROD_DT_TY_10) )" +
                    "   )," +
                    "   HEADER_DEFAULT_VALUES AS" +
                    "   (SELECT CN.MODEL_CODE," +
                    "     CN.INDICE," +
                    "     CN.HEADER_NAME," +
                    "     CT.VALUE_TYPE" +
                    "   FROM COLUMN_NAME CN" +
                    "   JOIN COLUMN_TYPE CT" +
                    "   ON CN.INDICE = CT.INDICE" +
                    "   )" +
                    " SELECT * FROM HEADER_DEFAULT_VALUES HDV";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

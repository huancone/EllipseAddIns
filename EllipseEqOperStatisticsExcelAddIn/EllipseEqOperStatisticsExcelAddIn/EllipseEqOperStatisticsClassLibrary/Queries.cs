using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Utilities;

namespace EllipseEqOperStatisticsExcelAddIn.EllipseEqOperStatisticsClassLibrary
{
    public static class Queries
    {
        public static string GetEquipmentLastMeterValueQuery(string dbReference, string dbLink, string equipNo, string statType, string statDate)
        {
            var sqlQuery = "" +
                           " SELECT" +
                           "   EQUIP_NO," +
                           "   STAT_DATE," +
                           "   STAT_TYPE," +
                           "   SHIFT," +
                           "   SHIFT_SEQ_NO," +
                           "   TRC_SEQ_NO," +
                           "   CUM_VALUE," +
                           "   REC400_TYPE," +
                           "   STAT_VALUE," +
                           "   STAT_DATE_SQ," +
                           "   METER_VALUE" +
                           " FROM" +
                           "   (" +
                           "     SELECT" +
                           "       METER_VALUE," +
                           "        EQUIP_NO," +
                           "        STAT_DATE," +
                           "        STAT_TYPE," +
                           "        SHIFT," +
                           "        SHIFT_SEQ_NO," +
                           "        TRC_SEQ_NO," +
                           "        CUM_VALUE," +
                           "        REC400_TYPE," +
                           "        STAT_VALUE," +
                           "       STAT_DATE || SHIFT_SEQ_NO STAT_DATE_SQ," +
                           "       MAX(STAT_DATE || SHIFT_SEQ_NO) OVER(PARTITION BY EQUIP_NO) MAX_FECHA" +
                           "     FROM" +
                           "       " + dbReference + ".MSF400" + dbLink +
                           "     WHERE" +
                           "       STAT_TYPE = '" + statType + "'" +
                           "     AND KEY_400_TYPE = 'E'" +
                           "     AND EQUIP_NO = '" + equipNo + "' AND STAT_DATE <= '" + statDate + "'" +
                           "   )" +
                           " WHERE" +
                           "   STAT_DATE_SQ = MAX_FECHA";
            
            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }

        public static string GetEquipmentOperStatisticsQuery(string dbReference, string dbLink, string equipNo, string statType, string startDate, string finishDate)
        {

            var statParam = "";
            if (!string.IsNullOrWhiteSpace(statType))
            {
                statParam = "   AND ST.STAT_TYPE = '" + statType + "'";
            }

            var startDateParam = "";
            if (!string.IsNullOrWhiteSpace(startDate))
            {
                startDateParam = "   AND ST.STAT_DATE   >= '" + startDate + "'";
            }

            var finishDateParam = "";
            if (!string.IsNullOrWhiteSpace(finishDate))
            {
                finishDateParam = "   AND ST.STAT_DATE   <= '" + finishDate + "'";
            }

            var sqlQuery = "" +
                           " SELECT " +
                           "    ST.EQUIP_NO," +
                           "    EQ.ITEM_NAME_1," +
                           "    EQ.ITEM_NAME_2," +
                           "    ST.STAT_DATE," +
                           "    ST.STAT_TYPE," +
                           "    ST.SHIFT," +
                           "    ST.SHIFT_SEQ_NO," +
                           "    ST.STAT_DATE || ST.SHIFT_SEQ_NO STAT_DATE_SQ," +
                           "    ST.TRC_SEQ_NO," +
                           "    ST.REC400_TYPE," +
                           "    ST.METER_VALUE," +
                           "    ST.CUM_VALUE," +
                           "    ST.STAT_VALUE" +
                           "   FROM ELLIPSE.MSF400 ST JOIN ELLIPSE.MSF600 EQ ON ST.EQUIP_NO = EQ.EQUIP_NO" +
                           " WHERE " +
                           statParam +
                           "   AND ST.KEY_400_TYPE = 'E'" +
                           "   AND ST.EQUIP_NO = '" + equipNo + "'" +
                           startDateParam +
                           finishDateParam;


            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseBulkMaterialExcelAddIn
{
    public static class Queries
    {
        public static string GetBulkAccountCode(string equipNo, string dbReference, string dbLink)
        {
            var query = "" +
                "WITH " +
                "  REFERENCE AS " +
                "  ( " +
                "    SELECT " +
                "      RC.REF_NO, " +
                "      RC.SCREEN_LITERAL, " +
                "      RCD.ENTITY_VALUE EQUIP_NO, " +
                "      RCD.REF_CODE BULK_ACCOUNT, " +
                "      RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM FECHA, " +
                "      MAX ( RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM ) OVER ( PARTITION BY RCD.REF_NO, RCD.ENTITY_VALUE ) MAX_FECHA " +
                "    FROM " +
                "      " + dbReference + ".MSF071" + dbLink + " RCD " +
                "    INNER JOIN " + dbReference + ".MSF070" + dbLink + " RC " +
                "    ON " +
                "      RCD.ENTITY_TYPE = RC.ENTITY_TYPE " +
                "    AND RC.REF_NO = RCD.REF_NO " +
                "    WHERE " +
                "      RCD.ENTITY_TYPE = 'EQP' " +
                "    AND RCD.REF_NO = '003' " +
                "  ) " +
                "SELECT " +
                "  EQ.EQUIP_NO, EQ.EQUIP_CLASS, EQ.EQUIP_CLASSIFX19," +
                "  EQ.ACCOUNT_CODE, RF.BULK_ACCOUNT  " +
                "FROM " +
                "  " + dbReference + ".MSF600" + dbLink + " EQ LEFT JOIN REFERENCE RF ON EQ.EQUIP_NO = RF.EQUIP_NO AND RF.FECHA = RF.MAX_FECHA  " +
                "WHERE " +
                "  EQ.EQUIP_NO = '" + equipNo + "'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFuelCapacity(string equipNo, string dbReference, string dbLink)
        {
            var query = "" +
                "WITH   " +
                "  EQUIPO AS   " +
                "  (   " +
                "    SELECT   " +
                "      EQ.EQUIP_NO   " +
                "    FROM   " +
                "      " + dbReference + ".MSF600" + dbLink + " EQ   " +
                "    WHERE   " +
                "      EQ.EQUIP_NO = '" + equipNo + "'   " +
                "  )   " +
                "  ,   " +
                "  BASE AS   " +
                "  (   " +
                "    SELECT   " +
                "      1 PESO,   " +
                "      PROFILES.EQUIP_GRP_ID,   " +
                "      PROFILES.FUEL_OIL_TYPE,   " +
                "      PROFILES.FUEL_CAPACITY   " +
                "    FROM   " +
                "      " + dbReference + ".MSF617_GENERAL" + dbLink + "  PROFILES   " +
                "    WHERE   " +
                "      PROFILES.EGI_REC_TYPE = 'E'   " +
                "    AND TRIM ( PROFILES.FUEL_OIL_TYPE ) IS NOT NULL   " +
                "    UNION ALL   " +
                "    SELECT   " +
                "      0 PESO,   " +
                "      PROFILES.EQUIP_GRP_ID,   " +
                "      PROFILES.FUEL_OIL_TYPE,   " +
                "      PROFILES.FUEL_CAPACITY   " +
                "    FROM   " +
                "      " + dbReference + ".MSF617_GENERAL" + dbLink + "  PROFILES   " +
                "    WHERE   " +
                "      PROFILES.EGI_REC_TYPE = 'G'   " +
                "    AND TRIM ( PROFILES.FUEL_OIL_TYPE ) IS NOT NULL   " +
                "  )   " +
                "  ,   " +
                "  EQUIPOS AS   " +
                "  (   " +
                "    SELECT   " +
                "      BASE.PESO,   " +
                "      EQ.EQUIP_NO,   " +
                "      EQ.EQUIP_GRP_ID,   " +
                "      BASE.FUEL_OIL_TYPE,   " +
                "      BASE.FUEL_CAPACITY   " +
                "    FROM   " +
                "      " + dbReference + ".MSF600" + dbLink + "  EQ   " +
                "    LEFT JOIN BASE   " +
                "    ON   " +
                "      EQ.EQUIP_NO = BASE.EQUIP_GRP_ID   " +
                "    OR EQ.EQUIP_GRP_ID = BASE.EQUIP_GRP_ID   " +
                "    WHERE   " +
                "      EQ.DSTRCT_CODE = 'ICOR'   " +
                "  )   " +
                "  ,   " +
                "  PROFILES AS   " +
                "  (   " +
                "    SELECT   " +
                "      EQUIPOS.PESO,   " +
                "      MAX ( EQUIPOS.PESO ) OVER ( PARTITION BY EQUIPOS.EQUIP_NO, EQUIPOS.EQUIP_GRP_ID ) MAX_PESO,   " +
                "      EQUIPOS.EQUIP_NO,   " +
                "      EQUIPOS.EQUIP_GRP_ID,   " +
                "      EQUIPOS.FUEL_OIL_TYPE,   " +
                "      EQUIPOS.FUEL_CAPACITY   " +
                "    FROM   " +
                "      EQUIPOS   " +
                "  )   " +
                "SELECT   " +
                "  EQUIPO.EQUIP_NO,   " +
                "  DECODE ( PROFILES.EQUIP_GRP_ID, NULL, 'NO TIENE', TRIM(PROFILES.EQUIP_GRP_ID) ) EQUIP_GRP_ID,   " +
                "  DECODE ( PROFILES.FUEL_OIL_TYPE, NULL, 'NO TIENE', TRIM(PROFILES.FUEL_OIL_TYPE) ) FUEL_OIL_TYPE,   " +
                "  DECODE ( PROFILES.FUEL_CAPACITY, NULL, 0, PROFILES.FUEL_CAPACITY ) FUEL_CAPACITY   " +
                "FROM   " +
                "  EQUIPO   " +
                "LEFT JOIN PROFILES   " +
                "ON   " +
                "  EQUIPO.EQUIP_NO = PROFILES.EQUIP_NO   " +
                "AND PROFILES.PESO = PROFILES.MAX_PESO   " +
                "ORDER BY   " +
                "  PROFILES.PESO   ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetLastStatistic(string equipNo, string statType, string statDate, string dbReference, string dbLink)
        {
            var query = "" +
                "SELECT " +
                "  STAT.EQUIP_NO, " +
                "  STAT.ITEM_NAME_1, " +
                "  STAT.METER_VALUE, " +
                "  STAT.STAT_TYPE, " +
                "  STAT.STAT_DATE " +
                "FROM " +
                "  ( " +
                "    SELECT " +
                "      EQ.EQUIP_NO, " +
                "      EQ.ITEM_NAME_1, " +
                "      STAT.STAT_TYPE, " +
                "      STAT.STAT_DATE, " +
                "      STAT.STAT_VALUE, " +
                "      STAT.CUM_VALUE, " +
                "      STAT.METER_VALUE, " +
                "      MAX ( STAT.STAT_DATE ) OVER ( PARTITION BY STAT.EQUIP_NO, STAT.STAT_TYPE ) MAX_DATE, " +
                "      EQ.DSTRCT_CODE " +
                "    FROM " +
                "      " + dbReference + ".MSF600" + dbLink + " EQ " +
                "    LEFT JOIN " + dbReference + ".MSF400" + dbLink + " STAT " +
                "    ON " +
                "      EQ.EQUIP_NO = STAT.EQUIP_NO " +
                "    WHERE " +
                "      EQ.EQUIP_NO = '" + equipNo + "' " +
                "    AND EQ.DSTRCT_CODE = 'ICOR' " +
                "    AND STAT.STAT_TYPE = '" + statType + "' " +
                "    AND STAT_DATE <= '" + statDate + "' " +
                "  ) " +
                "  STAT " +
                "WHERE " +
                "  STAT.MAX_DATE = STAT.STAT_DATE " +
                "OR STAT.STAT_DATE IS NULL ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetListIdList(string dbReference, string dbLink, string listType)
        {
            var query = "" +
                        "SELECT EQL.LIST_TYP, EQL.LIST_ID FROM " + dbReference + ".MSF606" + dbLink + " EQL " +
                        "WHERE EQL.LIST_TYP = '" + listType + "'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }


}

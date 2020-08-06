using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseEqOperStatisticsExcelAddIn.EllipseEqOperStatisticsClassLibrary
{
    public static class EqOperStatisticsActions
    {
        /// <summary>
        /// Obtiene la descripción del equipo a partir del número de equipo
        /// </summary>
        /// <param name="eFunctions"></param>
        /// <param name="equipNo">string: EquipmentNo para obtener la descripción</param>
        /// <returns>string: EquipmentNo. Null si el equipo no existe</returns>
        public static string GetEquipmentDescription(EllipseFunctions eFunctions, string equipNo)
        {
            var dbReference = eFunctions.DbReference;
            var dbLink = eFunctions.DbLink;

            var query = "SELECT EQ.* FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE TRIM(EQ.EQUIP_NO) = '" + equipNo + "'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drEquipments = eFunctions.GetQueryResult(query);

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return null;

            while (drEquipments.Read())
                return ("" + drEquipments["ITEM_NAME_1"]).Trim() + " " + ("" + drEquipments["ITEM_NAME_2"]).Trim();
            return null;
        }

        /// <summary>
        /// Obtiene la descripción del equipo a partir del número de equipo
        /// </summary>
        /// <param name="eFunctions"></param>
        /// <param name="equipNo">string: EquipmentNo para obtener la descripción</param>
        /// <param name="statType">string: Tipo de estadística a obtener</param>
        /// <param name="statDate">string: Fecha digitada en formato YYYYMMDD</param>
        /// <returns>string[2]: {fecha, medidor}. Null si no existe</returns>
        public static StatRegister GetEquipmentLastStat(EllipseFunctions eFunctions, string equipNo, string statType, string statDate)
        {
            var dbReference = eFunctions.DbReference;
            var dbLink = eFunctions.DbLink;

            var sqlQuery = Queries.GetEquipmentLastMeterValueQuery(dbReference, dbLink, equipNo, statType, statDate);

            var dr = eFunctions.GetQueryResult(sqlQuery);

            if (dr == null || dr.IsClosed || !dr.HasRows) return null;

            dr.Read();

            var stat = new StatRegister();
            stat.StatDate = ("" + dr["STAT_DATE"]).Trim();
            stat.MeterValue = ("" + dr["METER_VALUE"]).Trim();

            stat.EquipNo = ("" + dr["EQUIP_NO"]).Trim();
            stat.StatDate = ("" + dr["STAT_DATE"]).Trim();
            stat.StatType = ("" + dr["STAT_TYPE"]).Trim();
            stat.Shift = ("" + dr["SHIFT"]).Trim();
            stat.ShiftSeqNo = ("" + dr["SHIFT_SEQ_NO"]).Trim();
            stat.StatDateSeq = ("" + dr["STAT_DATE_SQ"]).Trim();
            stat.TransactSeqNo = ("" + dr["TRC_SEQ_NO"]).Trim();
            stat.EntryType = ("" + dr["REC400_TYPE"]).Trim();
            stat.CumValue = ("" + dr["CUM_VALUE"]).Trim();
            stat.StatValue = ("" + dr["STAT_VALUE"]).Trim();
            stat.MeterValue = ("" + dr["METER_VALUE"]).Trim();


            return stat;
        }

        public static List<StatRegister> ReviewEquipmentOperStatistics(EllipseFunctions eFunctions, string equipNo, string statType, string startDate, string finishDate)
        {
            var dbReference = eFunctions.DbReference;
            var dbLink = eFunctions.DbLink;

            var sqlQuery = Queries.GetEquipmentOperStatisticsQuery(dbReference, dbLink, equipNo, statType, startDate, finishDate);

            var dr = eFunctions.GetQueryResult(sqlQuery);
            var list = new List<StatRegister>();
            if (dr == null || dr.IsClosed || !dr.HasRows) return list;

            while (dr.Read())
            {
                var stat = new StatRegister();

                stat.StatDate = ("" + dr["STAT_DATE"]).Trim();
                stat.MeterValue = ("" + dr["METER_VALUE"]).Trim();
                stat.EquipNo = ("" + dr["EQUIP_NO"]).Trim();
                stat.EquipDesc1 = ("" + dr["ITEM_NAME_1"]).Trim();
                stat.EquipDesc2 = ("" + dr["ITEM_NAME_2"]).Trim();
                stat.StatDate = ("" + dr["STAT_DATE"]).Trim();
                stat.StatType = ("" + dr["STAT_TYPE"]).Trim();
                stat.Shift = ("" + dr["SHIFT"]).Trim();
                stat.ShiftSeqNo = ("" + dr["SHIFT_SEQ_NO"]).Trim();
                stat.StatDateSeq = ("" + dr["STAT_DATE_SQ"]).Trim();
                stat.TransactSeqNo = ("" + dr["TRC_SEQ_NO"]).Trim();
                stat.EntryType = ("" + dr["REC400_TYPE"]).Trim();
                stat.CumValue = ("" + dr["CUM_VALUE"]).Trim();
                stat.StatValue = ("" + dr["STAT_VALUE"]).Trim();
                stat.MeterValue = ("" + dr["METER_VALUE"]).Trim();

                list.Add(stat);
            }

            return list;
        }
    }
}

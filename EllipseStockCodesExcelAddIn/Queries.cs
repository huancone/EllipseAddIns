using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Utilities;

namespace EllipseStockCodesExcelAddIn
{
    internal static class Queries
    {
        public static string GetFetchInventoryStockCodeQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, bool validOnly, bool preferedOnly)
        {

            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))
                paramDistrict = " AND (PN.DSTRCT_CODE = '" + districtCode + "' OR TRIM(PN.DSTRCT_CODE) IS NULL)";
            string paramSearch;
            if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
                paramSearch = " AND SC.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            else if (searchCriteriaKey.Equals(SearchCriteriaType.PartNumber.Value))
                paramSearch = " AND TRIM(PN.PART_NO) = '" + searchCriteriaValue + "'";
            else
                paramSearch = " AND SC.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE PREF_PART_IND = MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           "WITH SCINV AS(" +
                           "    SELECT SC.STOCK_CODE, PN.PART_NO, PN.MNEMONIC, PN.DSTRCT_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, SC.CLASS STOCK_CLASS, SC.STOCK_TYPE," +
                           "        INV.CREATION_DATE, INV.LAST_MOD_DATE, INV.CLASS, INV.RAF, INV.INVENT_COST_PR AS PRICE, INV.HOME_WHOUSE, ELLIPSE.GET_SOH('" + districtCode + "', SC.STOCK_CODE) AS OWNED_SOH, ELLIPSE.GET_CONSIGN_SOH('" + districtCode + "', SC.STOCK_CODE) AS CONSIGN_SOH," +
                           "        INV.IN_TRANSIT, INV.DUES_IN, INV.DUES_OUT, INV.RESERVED, INV.ROP, INV.ROQ, INV.REORDER_QTY, INV.EXP_ELEMENT, INV.RESTRICT_RULE, INV.DIRECT_ORDER_IND, INV.PURCH_OFFICER, INV.INVT_CONTROLLR," +
                           "        PN.PREF_PART_IND, PN.STATUS_CODES," +
                           "        MIN(PN.PREF_PART_IND) OVER (PARTITION BY SC.STOCK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY SC.STOCK_CODE ORDER BY SC.STOCK_CODE, PN.PREF_PART_IND ASC) ROWPPI" +
                           "    FROM ELLIPSE.MSF100 SC" +
                           "        LEFT JOIN ELLIPSE.MSF110 PN ON SC.STOCK_CODE = PN.STOCK_CODE" +
                           "        LEFT JOIN ELLIPSE.MSF170 INV ON SC.STOCK_CODE = INV.STOCK_CODE" +
                           "    WHERE " +
                           " " + paramValidOnly +
                           " " + paramDistrict +
                           " " + paramSearch +
                           ")" +
                           "SELECT * FROM SCINV" +
                           " " + paramPreferedOnly;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetFetchRequisitionQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, string dateCriteria, string startDate, string finishDate, bool validOnly, bool preferedOnly)
        {
            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                paramDistrict = " AND RQ.DSTRCT_CODE = '" + districtCode + "'";// + " AND PN.DSTRCT_CODE = '" + districtCode + "'";
            if (dateCriteria.Equals(SearchDateCriteriaType.Raised.Value))
            {
                if (!string.IsNullOrWhiteSpace(startDate))
                    startDate = " AND RQ.CREATION_DATE >= " + startDate;
                if (!string.IsNullOrWhiteSpace(finishDate))
                    finishDate = " AND RQ.CREATION_DATE <= " + finishDate;
            }

            var paramReqNo = "";
            var paramStockCode = "";
            if (searchCriteriaKey.Equals(SearchCriteriaType.ItemCode.Value))
            {
                paramReqNo = " AND RQI.IREQ_NO = '" + searchCriteriaValue + "'";
            }
            else if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
            {
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }
            else
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE PREF_PART_IND = MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           "WITH REQSC AS (" +
                           " SELECT " +
                           "   RQI.DSTRCT_CODE, RQI.IREQ_NO, RQ.IREQ_TYPE, RQ.ISS_TRAN_TYPE, RQI.STOCK_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, PN.PART_NO, PN.MNEMONIC, RQI.IREQ_ITEM," +
                           "   RQ.AUTHSD_STATUS, RQ.HDR_140_STATUS, RQI.ITEM_141_STAT," +
                           "   RQ.PRIORITY_CODE, RQI.WHOUSE_ID, RQ.REQUESTED_BY, RQ.CREATION_DATE, RQ.REQ_BY_DATE, RQ.DELIV_INSTR_A, RQ.DELIV_INSTR_B," +
                           "   RQI.QTY_REQ, RQI.QTY_ISSUED, RQI.PO_ITEM_NO," +
                           "   PN.PREF_PART_IND, PN.STATUS_CODES," +
                           "   MIN(PN.PREF_PART_IND) OVER (PARTITION BY RQI.STOCK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY RQI.IREQ_NO, RQI.IREQ_ITEM, RQI.STOCK_CODE ORDER BY RQI.STOCK_CODE, PN.PREF_PART_IND ASC) ROWPPI" +
                           " FROM ELLIPSE.MSF141 RQI" +
                           " JOIN ELLIPSE.MSF140 RQ ON RQI.IREQ_NO = RQ.IREQ_NO" +
                           " LEFT JOIN ELLIPSE.MSF100 SC ON RQI.STOCK_CODE = SC.STOCK_CODE" +
                           " LEFT JOIN ELLIPSE.MSF110 PN ON RQI.STOCK_CODE = PN.STOCK_CODE" +
                           " WHERE" +
                           " " + paramValidOnly +
                           " " + paramReqNo +
                           " " + paramStockCode +
                           " " + paramDistrict +
                           " " + startDate +
                           " " + finishDate +
                           ")" +
                           "SELECT * FROM REQSC" +
                           " " + paramPreferedOnly;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchRequisitioneDetailedQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, string dateCriteria, string startDate, string finishDate, bool validOnly, bool preferedOnly)
        {
            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                paramDistrict = " AND RQ.DSTRCT_CODE = '" + districtCode + "'";// + " AND PN.DSTRCT_CODE = '" + districtCode + "'";
            if (dateCriteria.Equals(SearchDateCriteriaType.Raised.Value))
            {
                if (!string.IsNullOrWhiteSpace(startDate))
                    startDate = " AND RQ.CREATION_DATE >= " + startDate;
                if (!string.IsNullOrWhiteSpace(finishDate))
                    finishDate = " AND RQ.CREATION_DATE <= " + finishDate;
            }

            var paramReqNo = "";
            var paramStockCode = "";
            if (searchCriteriaKey.Equals(SearchCriteriaType.ItemCode.Value))
            {
                paramReqNo = " AND RQI.IREQ_NO = '" + searchCriteriaValue + "'";
            }
            else if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
            {
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }
            else
                paramStockCode = " AND RQI.STOCK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE PREF_PART_IND = MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           "WITH REQSC AS (" +
                           " SELECT " +
                           "   RQI.DSTRCT_CODE, RQR.EQUIP_NO, EQ.EQUIP_GRP_ID, RQR.GL_ACCOUNT, RQR.WORK_ORDER, WO.WO_DESC," +
                           "   RQI.IREQ_NO, RQ.IREQ_TYPE, RQ.ISS_TRAN_TYPE, RQI.STOCK_CODE, SC.ITEM_NAME, SC.STK_DESC, SC.UNIT_OF_ISSUE, PN.PART_NO, PN.MNEMONIC, RQI.IREQ_ITEM," +
                           "   RQ.AUTHSD_STATUS, RQ.HDR_140_STATUS, RQI.ITEM_141_STAT," +
                           "   RQ.PRIORITY_CODE, RQI.WHOUSE_ID, RQ.REQUESTED_BY, RQ.CREATION_DATE, RQ.REQ_BY_DATE, RQ.DELIV_INSTR_A, RQ.DELIV_INSTR_B," +
                           "   RQI.QTY_REQ, RQI.QTY_ISSUED, RQI.PO_ITEM_NO," +
                           "   PN.PREF_PART_IND, PN.STATUS_CODES," +
                           "   MIN(PN.PREF_PART_IND) OVER (PARTITION BY RQI.STOCK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY RQI.IREQ_NO, RQI.IREQ_ITEM, RQI.STOCK_CODE ORDER BY RQI.STOCK_CODE, PN.PREF_PART_IND ASC) ROWPPI" +
                           " FROM ELLIPSE.MSF141 RQI" +
                           " JOIN ELLIPSE.MSF140 RQ ON RQI.IREQ_NO = RQ.IREQ_NO" +
                           " LEFT JOIN ELLIPSE.MSF232 RQR ON RQ.IREQ_NO||' '||' 0000' = RQR.REQUISITION_NO" +
                           " LEFT JOIN ELLIPSE.MSF600 EQ ON RQR.EQUIP_NO = EQ.EQUIP_NO" +
                           " LEFT JOIN ELLIPSE.MSF620 WO ON (RQR.WORK_ORDER = WO.WORK_ORDER AND RQ.DSTRCT_CODE = WO.DSTRCT_CODE)" +
                           " LEFT JOIN ELLIPSE.MSF100 SC ON RQI.STOCK_CODE = SC.STOCK_CODE" +
                           " LEFT JOIN ELLIPSE.MSF110 PN ON RQI.STOCK_CODE = PN.STOCK_CODE" +
                           " WHERE" +
                           " " + paramValidOnly +
                           " " + paramReqNo +
                           " " + paramStockCode +
                           " " + paramDistrict +
                           " " + startDate +
                           " " + finishDate +
                           " )" +
                           " SELECT " +
                           "   DSTRCT_CODE DISTRITO, EQUIP_NO EQUIP, EQUIP_GRP_ID FLOTA_EGI, GL_ACCOUNT CENTRO_COSTO, WORK_ORDER OT, WO_DESC DESC_OT," +
                           "   IREQ_NO NRO_VALE, IREQ_ITEM NRO_ITEM, QTY_REQ CANT_REQ, QTY_ISSUED CANT_DESP, IREQ_TYPE TIPO_VALE, ISS_TRAN_TYPE TIPO_TRAN, STOCK_CODE, " +
                           "   ITEM_NAME, STK_DESC, UNIT_OF_ISSUE, PART_NO, MNEMONIC," +
                           "   PRIORITY_CODE COD_PRIORIDAD, WHOUSE_ID BODEGA," +
                           "   REQUESTED_BY REQUERIDO_POR, CREATION_DATE FECHA_CREACION, REQ_BY_DATE FECHA_REQUERIDO, DELIV_INSTR_A||DELIV_INSTR_B INSTRUCCIONES_ENTREGA" +
                           " FROM REQSC" +
                           " " + paramPreferedOnly;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchPurchaseOrderQuery(string dbReference, string dbLink, string districtCode, string searchCriteriaKey, string searchCriteriaValue, string dateCriteria, string startDate, string finishDate, bool validOnly, bool preferedOnly)
        {
            var paramDistrict = "";
            if (!string.IsNullOrWhiteSpace(districtCode))//muchos stockcodes no tienen registrado distrito en los parte número
                paramDistrict = " AND PO.DSTRCT_CODE = '" + districtCode + "'";
            if (dateCriteria.Equals(SearchDateCriteriaType.Raised.Value))
            {
                if (!string.IsNullOrWhiteSpace(startDate))
                    startDate = " AND PO.CREATION_DATE >= " + startDate;
                if (!string.IsNullOrWhiteSpace(finishDate))
                    finishDate = " AND PO.CREATION_DATE <= " + finishDate;
            }

            var paramPurchaseOrder = "";
            var paramStockCode = "";
            if (searchCriteriaKey.Equals(SearchCriteriaType.ItemCode.Value))
            {
                paramPurchaseOrder = " PO.PO_NO = '" + searchCriteriaValue + "'";
            }
            else if (searchCriteriaKey.Equals(SearchCriteriaType.StockCode.Value))
            {
                paramStockCode = " POI.PREQ_STK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }
            else
            {
                paramStockCode = " POI.PREQ_STK_CODE = '" + searchCriteriaValue.PadLeft(9, '0') + "'";
            }

            var paramValidOnly = validOnly ? " AND PN.STATUS_CODES = 'V'" : "";
            var paramPreferedOnly = preferedOnly ? " WHERE POITEMS.PREF_PART_IND = POITEMS.MINPPI AND ROWPPI = 1" : "";

            var query = "" +
                           " WITH POITEMS AS(" +
                           "  SELECT" +
                           "    POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE, SC.ITEM_NAME, SC.DESC_LINEX1, SC.DESC_LINEX2, SC.DESC_LINEX3, SC.DESC_LINEX4, PN.PART_NO, PN.MNEMONIC, " +
                           "    POI.GROSS_PRICE_P, POI.UNIT_OF_PURCH, POI.CONV_FACTOR, " +
                           "    PN.PREF_PART_IND, MIN(PN.PREF_PART_IND) OVER (PARTITION BY POI.PREQ_STK_CODE) MINPPI, ROW_NUMBER() OVER (PARTITION BY POI.PO_NO, POI.PO_ITEM_NO, POI.PREQ_STK_CODE ORDER BY POI.PREQ_STK_CODE, PN.PREF_PART_IND ASC) ROWPPI, " +
                           "    PO.STATUS_220, PO.CREATION_DATE, PO.ORDER_DATE, POI.ORIG_DUE_DATE, POI.ORIG_NET_PR_I, POI.CURR_NET_PR_I, POI.ORIG_QTY_I, POI.CURR_QTY_I, POI.QTY_RCV_OFST_I, POI.OFST_RCPT_DATE, POI.QTY_RCV_DIR_I, POI.ONST_RCPT_DATE, PO.FREIGHT_CODE, PO.DELIV_LOCATION, POI.EXPEDITE_CODE, PO.SUPPLIER_NO, SUP.SUPPLIER_NAME, PO.PO_MEDIUM_IND, PO.ORIGIN_CODE, PO.PURCH_OFFICER, PO.TEAM_ID" +
                           "  FROM" +
                           "    ELLIPSE.MSF220 PO JOIN ELLIPSE.MSF221 POI ON PO.PO_NO = POI.PO_NO LEFT JOIN ELLIPSE.MSF100 SC ON POI.PREQ_STK_CODE = SC.STOCK_CODE LEFT JOIN ELLIPSE.MSF110 PN ON POI.PREQ_STK_CODE = PN.STOCK_CODE LEFT JOIN ELLIPSE.MSF200 SUP ON PO.SUPPLIER_NO = SUP.SUPPLIER_NO" +
                           "  WHERE" +
                           " " + paramPurchaseOrder +
                           " " + paramStockCode +
                           " " + paramDistrict +
                           " " + startDate +
                           " " + finishDate +
                           " " + paramValidOnly +
                           "  ORDER BY POI.PO_ITEM_NO" +
                           "  )," +
                           " SCSTAT AS(" +
                           " SELECT STAT.DSTRCT_CODE, SC.STOCK_CODE, STAT.CREATION_DATE, STAT.LAST_MOD_DATE, SC.STK_DESC, SC.UNIT_OF_ISSUE, STAT.CLASS, STAT.RAF AS ALGORITMO, STAT.INVENT_COST_PR AS PRICE, STAT.HOME_WHOUSE AS BODEGA_PRINCIPAL, ELLIPSE.GET_SOH('" + districtCode + "',SC.STOCK_CODE) AS OWNED_SOH, ELLIPSE.GET_CONSIGN_SOH('" + districtCode + "', SC.STOCK_CODE) AS CONSIGN_SOH," +
                           "  STAT.IN_TRANSIT, STAT.DUES_IN, STAT.DUES_OUT, STAT.RESERVED, STAT.ROP, STAT.REORDER_QTY ROQ, STAT.EXP_ELEMENT AS DETALLE_GASTO, STAT.RESTRICT_RULE AS RESTR, STAT.DIRECT_ORDER_IND AS DO_IND, STAT.PURCH_OFFICER AS PURCHASER, " +
                           "  (SELECT SUM(UNSCHED_USAGE) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE " +
                           "  AND FULL_ACCT_PER BETWEEN (SELECT TO_CHAR(TO_DATE(MAX(FULL_ACCT_PER),'YYYYMM')-365,'YYYYMM') FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) AND (SELECT MAX(FULL_ACCT_PER) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) " +
                           "  AND TRIM(WHOUSE_ID) IS NOT NULL ) AS USO12_UNSCH, " +
                           "  (SELECT SUM(UNSCHED_USAGE) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE " +
                           "  AND FULL_ACCT_PER=(SELECT MAX(FULL_ACCT_PER) FROM ELLIPSE.MSF175 WHERE DSTRCT_CODE=STAT.DSTRCT_CODE AND STOCK_CODE=SC.STOCK_CODE) " +
                           "  AND TRIM(WHOUSE_ID) IS NOT NULL ) AS CURRENT_UNSCH, " +
                           "  STAT.INVT_CONTROLLR AS ADI FROM ELLIPSE.MSF100 SC LEFT JOIN ELLIPSE.MSF170 STAT ON SC.STOCK_CODE = STAT.STOCK_CODE)" +
                           "  SELECT * FROM POITEMS LEFT JOIN SCSTAT ON POITEMS.PREQ_STK_CODE = SCSTAT.STOCK_CODE AND SCSTAT.DSTRCT_CODE = 'ICOR'" +
                           " " + paramPreferedOnly;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

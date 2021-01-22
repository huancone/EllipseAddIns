using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStandardJobsExcelAddIn
{
    internal static class Queries
    {
        public static string GetAplRequirementsQuery(string dbReference, string dbLink, string aplEgi, string aplType, string aplCompCode, string aplCompModCode, string seqNo)
        {
            if (string.IsNullOrWhiteSpace(aplCompCode))
                aplCompCode = " IS NULL";
            else
                aplCompCode = " = '" + aplCompCode + "'";

            if (string.IsNullOrWhiteSpace(aplCompModCode))
                aplCompModCode = " IS NULL";
            else
                aplCompModCode = " = '" + aplCompModCode + "'";

            var sqlQuery = "" +
                " SELECT" +
                "   AST.EQUIP_GRP_ID , AST.APL_TYPE, AST.COMP_CODE, AST.COMP_MOD_CODE, AST.APL_SEQ_NO, AST.APL_ITEM_NUM, AST.PART_NO, AST.MNEMONIC, AST.STOCK_CODE, AST.ITEM_DESC, AST.QTY_REQUIRED, AST.QTY_INSTALLED" +
                " FROM" +
                "   " + dbReference + ".MSF131" + dbLink + "  AST" +
                " WHERE" +
                "   TRIM(AST.EQUIP_GRP_ID) = '" + aplEgi + "' AND AST.APL_SEQ_NO = '" + seqNo + "' AND AST.APL_TYPE = '" + aplType + "' AND TRIM(AST.COMP_CODE) " + aplCompCode + " AND TRIM(AST.COMP_MOD_CODE) " + aplCompModCode + "";

            return sqlQuery;
        }

        public static string FetchReferenceCodeItems(string dbReference, string dbLink, string entityType, string entityValue, string refNo, string seqNum = null)
        {
            if (!string.IsNullOrWhiteSpace(refNo))
                refNo = " AND RC.REF_NO = '" + refNo.PadLeft(3, '0') + "'";
            if (!string.IsNullOrWhiteSpace(seqNum))
                seqNum = " AND RC.SEQ_NUM = '" + seqNum.PadLeft(3, '0') + "'";
            var query = "" +
                        " SELECT RC.ENTITY_TYPE, " +
                        "   RC.ENTITY_VALUE, " +
                        "   RC.REF_NO, " +
                        "   RC.SEQ_NUM, " +
                        "   RC.REF_CODE, " +
                        "   RCE.FIELD_TYPE, " +
                        "   RCE.SHORT_NAMES, " +
                        "   RCE.SCREEN_LITERAL, " +
                        "   RC.STD_TXT_KEY, " +
                        "   RCE.STD_TEXT_FLAG " +
                        " FROM " +
                        "     " + dbReference + ".MSF071" + dbLink + " RC LEFT JOIN " + dbReference + ".MSF070" + dbLink + " RCE " +
                        "         ON (RC.ENTITY_TYPE = RCE.ENTITY_TYPE AND RC.REF_NO = RCE.REF_NO) " +
                        " WHERE RCE.ENTITY_TYPE = '" + entityType + "' " +
                        " AND RC.ENTITY_VALUE = '" + entityValue + "' " +
                        " " + refNo +
                        " " + seqNum;
            return query;
        }
    }

}

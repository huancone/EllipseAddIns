using SharedClassLibrary.Utilities;

namespace EllipseReferenceCodesClassLibrary
{
    internal static class Queries
    {
        public static string FetchReferenceCodeEntities(string dbReference, string dbLink, string entityType)
        {
            //escribimos el query
            var query = "" +
                        " SELECT RCE.ENTITY_TYPE," +
                        "     RCE.REF_NO," +
                        "     RCE.REPEAT_COUNT," +
                        "     RCE.FIELD_TYPE," +
                        "     RCE.SHORT_NAMES," +
                        "     RCE.SCREEN_LITERAL," +
                        "     RCE.LENGTH," +
                        "     RCE.STD_TEXT_FLAG" +
                        " FROM" +
                        "   " + dbReference + ".MSF070" + dbLink + " RCE" +
                        " WHERE RCE.ENTITY_TYPE = '" + entityType + "'";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
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
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }


    }
}

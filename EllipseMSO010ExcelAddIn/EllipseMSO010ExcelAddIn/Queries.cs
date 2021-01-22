using SharedClassLibrary.Utilities;

namespace EllipseMSO010ExcelAddIn
{
    internal static class Queries
    {
        public static string GetItemCodeList(string dbReference, string dbLink, int searchCriteriaKey1, int typeCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, int typeCriteriaKey2, string searchCriteriaValue2, int statusCriteriaKey)
        {
            string typeCriteria1;
            //establecemos los tipos de búsqueda 1
            if (typeCriteriaKey1 == SearchTypeCriteriaType.EqualsTo.Key)
                typeCriteria1 = "= '" + searchCriteriaValue1 + "'";
            else if (typeCriteriaKey1 == SearchTypeCriteriaType.StartsWith.Key)
                typeCriteria1 = "LIKE '" + searchCriteriaValue1 + "%'";
            else if (typeCriteriaKey1 == SearchTypeCriteriaType.EndsWith.Key)
                typeCriteria1 = "LIKE '%" + searchCriteriaValue1 + "'";
            else if (typeCriteriaKey1 == SearchTypeCriteriaType.Contains.Key)
                typeCriteria1 = "LIKE '%" + searchCriteriaValue1 + "%'";
            else
                typeCriteria1 = "= '" + searchCriteriaValue1 + "'";

            var queryCriteria1 = "";
            //establecemos los parámetros del criterio 1
            if (searchCriteriaKey1 == SearchFieldCriteriaType.Type.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = " AND CO.TABLE_TYPE " + typeCriteria1 + "";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Code.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = " AND CO.TABLE_CODE " + typeCriteria1 + "";

            string typeCriteria2;
            //establecemos los tipos de búsqueda 2
            if (typeCriteriaKey2 == SearchTypeCriteriaType.EqualsTo.Key)
                typeCriteria2 = "= '" + searchCriteriaValue2 + "'";
            else if (typeCriteriaKey2 == SearchTypeCriteriaType.StartsWith.Key)
                typeCriteria2 = "LIKE '" + searchCriteriaValue2 + "%'";
            else if (typeCriteriaKey2 == SearchTypeCriteriaType.EndsWith.Key)
                typeCriteria2 = "LIKE '%" + searchCriteriaValue2 + "'";
            else if (typeCriteriaKey2 == SearchTypeCriteriaType.Contains.Key)
                typeCriteria2 = "LIKE '%" + searchCriteriaValue2 + "%'";
            else
                typeCriteria2 = "= '" + searchCriteriaValue2 + "'";

            var queryCriteria2 = "";
            //establecemos los parámetros del criterio 2
            if (searchCriteriaKey2 == SearchFieldCriteriaType.Type.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = " AND CO.TABLE_TYPE " + typeCriteria2 + "";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Code.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = " AND CO.TABLE_CODE " + typeCriteria2 + "";

            string statusCriteria;
            if (statusCriteriaKey == StatusCode.Active.Key)
                statusCriteria = " AND CO.ACTIVE_FLAG = 'Y'";
            else if (statusCriteriaKey == StatusCode.Inactive.Key)
                statusCriteria = " AND CO.ACTIVE_FLAG = 'N'";
            else
                statusCriteria = " AND CO.ACTIVE_FLAG = 'Y'";


            var query = " SELECT CO.TABLE_TYPE," +
                           "   CO.TABLE_CODE," +
                           "   CO.TABLE_DESC," +
                           "   CO.ACTIVE_FLAG," +
                           "   TY.TABLE_DESC TYPE_DESC," +
                           "   CO.ASSOC_REC" +
                           " FROM ELLIPSE.MSF010 CO" +
                           " LEFT JOIN ELLIPSE.MSF010 TY" +
                           " ON CO.TABLE_TYPE  = TY.TABLE_CODE" +
                           " AND TY.TABLE_TYPE = 'XX'" +
                           " WHERE " +
                           " " + queryCriteria1 +
                           " " + queryCriteria2 +
                           " " + statusCriteria +
                           " ORDER BY CO.TABLE_CODE ASC";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

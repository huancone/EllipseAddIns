
using System.Collections.Generic;
using EllipseCommonsClassLibrary;

namespace EllipseEquipmentClassLibrary
{
    public static class EquipmentListActions
    {
        public static List<EquipListItem> FetchListEquipmentsList(EllipseFunctions ef, int primakeryKey, string primaryValue, int secondarykey, string secondaryValue)
        {
            var sqlQuery = Queries.GetFetchListEquipmentsListQuery(ef.dbReference, ef.dbLink, primakeryKey, primaryValue, secondarykey, secondaryValue);
            var drItem = ef.GetQueryResult(sqlQuery);
            var list = new List<EquipListItem>();

            if (drItem == null || drItem.IsClosed || !drItem.HasRows) return list;
            while (drItem.Read())
            {
                var item = new EquipListItem()
                {
                    EquipNo = drItem["EQUIP_NO"].ToString().Trim(),
                    EquipDescription = drItem["ITEM_NAME_1"].ToString().Trim() + " " + drItem["ITEM_NAME_2"].ToString().Trim(),
                    ListType = drItem["LIST_TYP"].ToString().Trim(),
                    ListId = drItem["LIST_ID"].ToString().Trim(),
                    ListNumber = drItem["LIST_NUMBER"].ToString().Trim(),
                    ListDescription = drItem["LIST_DESCR_1"].ToString().Trim() + " " + drItem["LIST_DESCR_2"].ToString().Trim(),
                    ListReference = drItem["LIST_REF"].ToString().Trim(),
                    ListOwner = drItem["LIST_OWN_EMPL"].ToString().Trim(),
                    ListOwnerPosition = drItem["LIST_OWN_POSN"].ToString().Trim(),
                };
                list.Add(item);
            }

            return list;
        }
        
        public static class Queries
        {
            public static string GetFetchListEquipmentsListQuery(string dbReference, string dbLink, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2)
            {
                string queryCriteria1;
                //establecemos los parámetros del criterio 1
                if (searchCriteriaKey1 == EquipListSearchFieldCriteria.EquipmentNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND ELI.EQUIP_NO = '" + searchCriteriaValue1 + "'";
                //else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                //    queryCriteria1 = " AND EQ.EQUIP_NO = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_TYP = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_ID = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_NUMBER = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListDescription.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND TRIM(EL.LIST_DESCR_1)||' '||TRIM(EL.LIST_DESCR_2) LIKE '%" + searchCriteriaValue1 + "%'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_REF = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListOwner.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_OWN_EMPL = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListOwnerPosition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_OWN_POSN = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListRaisedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_RAISED_BY = '" + searchCriteriaValue1 + "'";
                //else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.MemberEquipNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                //    queryCriteria1 = " AND EQ.LIST_RAISED_BY = '" + searchCriteriaValue1 + "'";
                else
                    queryCriteria1 = " AND ELI.EQUIP_NO = '" + searchCriteriaValue1 + "'";

                var queryCriteria2 = "";

                var query = " SELECT" +
                            "   ELI.MEM_EQUIP_GRP EQUIP_NO," +
                            "   EQ.ITEM_NAME_1," +
                            "   EQ.ITEM_NAME_2," +
                            "   EL.LIST_TYP," +
                            "   EL.LIST_ID," +
                            "   EL.LIST_NUMBER," +
                            "   EL.LIST_DESCR_1," +
                            "   EL.LIST_DESCR_2," +
                            "   EL.LIST_REF," +
                            "   EL.LIST_OWN_EMPL," +
                            "   EL.LIST_OWN_POSN," +
                            "   EL.LIST_RAISED_BY" +
                            " FROM" +
                            "   ELLIPSE.MSF607 ELI" +
                            "   LEFT JOIN ELLIPSE.MSF606 EL ON ELI.LIST_TYP  = EL.LIST_TYP AND ELI.LIST_ID = LI.LIST_ID" +
                            "   LEFT JOIN ELLIPSE.MSF600 EQ ON ELI.MEM_EQUIP_GRP = EQ.EQUIP_NO" +
                            " WHERE" +
                            " " + queryCriteria1 +
                            " " + queryCriteria2;

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }
    }
}

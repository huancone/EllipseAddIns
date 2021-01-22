using System.Collections.Generic;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;
using EllipseEquipmentClassLibrary.EquipmentListService;

namespace EllipseEquipmentClassLibrary
{
    public static class ListActions
    {
        public static List<EquipListItem> FetchListEquipmentsList(EllipseFunctions ef, int primakeryKey, string primaryValue, int secondarykey, string secondaryValue, string statusValue)
        {
            var sqlQuery = Queries.GetFetchListEquipmentsListQuery(ef.DbReference, ef.DbLink, primakeryKey, primaryValue, secondarykey, secondaryValue, statusValue);
            var drItem = ef.GetQueryResult(sqlQuery);
            var list = new List<EquipListItem>();

            if (drItem == null || drItem.IsClosed) 
                return list;

            while (drItem.Read())
            {
                var item = new EquipListItem()
                {
                    EquipNo = drItem["EQUIP_NO"].ToString().Trim(),
                    EquipDescription = drItem["ITEM_NAME_1"].ToString().Trim() + " " + drItem["ITEM_NAME_2"].ToString().Trim(),
                    ListType = drItem["LIST_TYP"].ToString().Trim(),
                    ListId = drItem["LIST_ID"].ToString().Trim(),
                    Status = drItem["EQUIP_STATUS"].ToString().Trim(),
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

       public static EquipmentListServiceCreateEquipItemReplyDTO AddEquipmentToList(OperationContext operationContext, string urlService, EquipListItem equipListItem)
        {
            var proxyEquip = new EquipmentListService.EquipmentListService();
            var request = new EquipmentListServiceCreateEquipItemRequestDTO()
            {
                memEquipmentNo = equipListItem.EquipNo,
                listType = equipListItem.ListType,
                listId = equipListItem.ListId,
            };
            proxyEquip.Url = urlService + "/EquipmentList";
            return proxyEquip.createEquipItem(operationContext, request);
        }
        public static EquipmentListServiceDelEquipItemReplyDTO DeleteEquipmentFromList(OperationContext operationContext, string urlService, EquipListItem equipListItem)
        {
            var proxyEquip = new EquipmentListService.EquipmentListService();
            var request = new EquipmentListServiceDelEquipItemRequestDTO()
            {
                memEquipmentNo = equipListItem.EquipNo,
                listType = equipListItem.ListType,
                listId = equipListItem.ListId,
            };
            proxyEquip.Url = urlService + "/EquipmentList";
            return proxyEquip.delEquipItem(operationContext, request);
        }
        internal static class Queries
        {
            public static string GetFetchListEquipmentsListQuery(string dbReference, string dbLink, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, string eqStatus)
            {
                
                //establecemos los parámetros del criterio 1
                string queryCriteria1;
                if (searchCriteriaKey1 == EquipListSearchFieldCriteria.EquipmentNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.EQUIP_NO = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    var equipParamsQuery = EquipmentActions.Queries.GetEquipReferencesQuery(dbReference, dbLink, null, searchCriteriaValue1);
                    queryCriteria1 = " AND EQ.EQUIP_NO IN (" + equipParamsQuery + ")";
                }
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND ELI.LIST_TYP = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND ELI.LIST_ID = '" + searchCriteriaValue1 + "'";
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
                else if (searchCriteriaKey1 == EquipListSearchFieldCriteria.MemberEquipNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EL.LIST_TYP||EL.LIST_ID IN (SELECT DISTINCT LIST_TYP||LIST_ID FROM ELLIPSE.MSF607 WHERE MEM_EQUIP_GRP = '" + searchCriteriaValue1 + "')";
                else
                    queryCriteria1 = " AND ELI.EQUIP_NO = '" + searchCriteriaValue1 + "'";

                //establecemos los parámetros del criterio 1
                var queryCriteria2 = "";
                if (searchCriteriaKey2 == EquipListSearchFieldCriteria.EquipmentNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.EQUIP_NO = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    var equipParamsQuery = EquipmentActions.Queries.GetEquipReferencesQuery(dbReference, dbLink, null, searchCriteriaValue2);
                    queryCriteria2 = " AND EQ.EQUIP_NO IN (" + equipParamsQuery + ")";
                }
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND ELI.LIST_TYP = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND ELI.LIST_ID = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_NUMBER = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListDescription.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND TRIM(EL.LIST_DESCR_1)||' '||TRIM(EL.LIST_DESCR_2) LIKE '%" + searchCriteriaValue2 + "%'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_REF = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListOwner.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_OWN_EMPL = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListOwnerPosition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_OWN_POSN = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.ListRaisedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_RAISED_BY = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == EquipListSearchFieldCriteria.MemberEquipNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EL.LIST_TYP||EL.LIST_ID IN (SELECT DISTINCT LIST_TYP||LIST_ID FROM ELLIPSE.MSF607 WHERE MEM_EQUIP_GRP = '" + searchCriteriaValue2 + "')";

                //establecemos los parámetros de estado
                string statusRequirement;
                if (string.IsNullOrEmpty(eqStatus))
                    statusRequirement = "";
                else
                    statusRequirement = " AND EQ.EQUIP_STATUS = '" + MyUtilities.GetCodeKey(eqStatus) + "'";

                var query = " SELECT" +
                            "   ELI.MEM_EQUIP_GRP EQUIP_NO, " +
                            "   EQ.ITEM_NAME_1, " +
                            "   EQ.ITEM_NAME_2, " +
                            "   EL.LIST_TYP, " +
                            "   EL.LIST_ID, " +
                            "   EQ.EQUIP_STATUS, " +
                            "   EL.LIST_NUMBER, " +
                            "   EL.LIST_DESCR_1, " +
                            "   EL.LIST_DESCR_2, " +
                            "   EL.LIST_REF, " +
                            "   EL.LIST_OWN_EMPL, " +
                            "   EL.LIST_OWN_POSN, " +
                            "   EL.LIST_RAISED_BY " +
                            " FROM " +
                            "   " + dbReference + ".MSF607" + dbLink + " ELI " +
                            "   LEFT JOIN " + dbReference + ".MSF606" + dbLink + " EL ON ELI.LIST_TYP  = EL.LIST_TYP AND ELI.LIST_ID = EL.LIST_ID " +
                            "   LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON ELI.MEM_EQUIP_GRP = EQ.EQUIP_NO " +
                            " WHERE " +
                            " ELI.MEM_TYPE = 'E' " +
                            " " + queryCriteria1 +
                            " " + queryCriteria2 +
                            " " + statusRequirement;
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }
    }
}

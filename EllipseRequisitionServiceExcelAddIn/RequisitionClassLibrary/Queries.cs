using System;
using System.Linq;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;
using EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.SearchConstants;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary
{
    static class Queries
    {
        public static string GetItemUnitOfIssue(string stockCode)
        {
            var query = "SELECT UNIT_OF_ISSUE FROM ELLIPSE.MSF100 SC WHERE SC.STOCK_CODE = '" + stockCode + "' ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetItemDirectOrder(string stockCode)
        {
            var query = "SELECT SCI.DIRECT_ORDER_IND FROM ELLIPSE.MSF170 SCI WHERE STOCK_CODE = '" + stockCode + "' ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetRequisitionListQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, int dateCriteriaKey, string startDate, string endDate, string reqStatus)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrWhiteSpace(districtCode))
                districtCode = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            var queryCriteria1 = "";
            //establecemos los parámetros del criterio 1
            if (searchCriteriaKey1 == SearchFieldCriteriaType.Requisition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.IREQ_NO = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.CreatedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.CREATED_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.RequestedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.REQUESTED_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.RequestedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.REQ_BY_POS = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AuthorizedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.AUTHSD_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AuthorizedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.AUTHSD_POSITION = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.GL_ACCOUNT = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.WORK_ORDER = '" + searchCriteriaValue1 + "'";
            //else if (searchCriteriaKey1 == SearchFieldCriteriaType.StockCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            //    queryCriteria1 = "AND  = '" + searchCriteriaValue1 + "'";
            //else if (searchCriteriaKey1 == SearchFieldCriteriaType.PartNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            //    queryCriteria1 = "AND  = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.EQUIP_NO = '" + searchCriteriaValue1 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.PARENT_WO = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            //


            var queryCriteria2 = "";
            //establecemos los parámetros del criterio 2
            if (searchCriteriaKey2 == SearchFieldCriteriaType.Requisition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.IREQ_NO = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.CreatedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.CREATED_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.RequestedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.REQUESTED_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.RequestedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.REQ_BY_POS = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AuthorizedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.AUTHSD_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AuthorizedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.AUTHSD_POSITION = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.GL_ACCOUNT = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.WORK_ORDER = '" + searchCriteriaValue2 + "'";
            //else if (searchCriteriaKey2 == SearchFieldCriteriaType.StockCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            //    queryCriteria2 = "AND  = '" + searchCriteriaValue2 + "'";
            //else if (searchCriteriaKey2 == SearchFieldCriteriaType.PartNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            //    queryCriteria2 = "AND  = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.EQUIP_NO = '" + searchCriteriaValue2 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.PARENT_WO = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "')";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            //


            //establecemos los parámetros de estado de orden
            var statusRequirement = "";
            if (string.IsNullOrWhiteSpace(reqStatus))
                statusRequirement = "";
            else if (reqStatus == RequisitionStatus.Uncompleted.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9'";
            else if (reqStatus == RequisitionStatus.Pending.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> 'P'";
            else if (reqStatus == RequisitionStatus.Unauthorized.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9' AND (TRIM(RQ.AUTHSD_STATUS) IS NULL OR TRIM(RQ.AUTHSD_STATUS) = 'U') ";
            else if (reqStatus == RequisitionStatus.Authorized.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9' AND TRIM(RQ.AUTHSD_STATUS) = 'A'";
            else if (reqStatus == RequisitionStatus.Awaiting.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS IN ('1', '2', '3') AND TRIM(RQ.AUTHSD_STATUS) = 'A'";

            //establecemos los parámetros para el rango de fechas
            string dateParameters;
            if (string.IsNullOrWhiteSpace(startDate))
                startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
            if (string.IsNullOrWhiteSpace(endDate))
                endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                dateParameters = "";
            else if (dateCriteriaKey == SearchDateCriteriaType.Creation.Key)
                dateParameters = " AND RQ.CREATION_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.Required.Key)
                dateParameters = " AND RQ.REQ_BY_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.WoRaisedDate.Key)
                dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.WoPlanStartDate.Key)
                dateParameters = " AND WO.PLAN_STR_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.IgnoreDate.Key)
                dateParameters = " ";
            else
                dateParameters = " AND RQ.CREATION_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";

            var query = "" +
                        " SELECT RQ.DSTRCT_CODE," +
                        "   WO.WORK_GROUP," +
                        "   RQR.EQUIP_NO," +
                        "   RQR.WORK_ORDER," +
                        "   WO.WO_DESC," +
                        "   RQR.PROJECT_NO," +
                        "   RQR.GL_ACCOUNT," +
                        "   RQ.IREQ_NO," +
                        "   RQ.NUM_OF_ITEMS," +
                        "   WO.RAISED_DATE WO_RAISED_DATE," +
                        "   WO.PLAN_STR_DATE WO_PLAN_STR_DATE," +
                        "   RQ.CREATION_DATE," +
                        "   RQ.REQ_BY_DATE," +
                        "   RQ.AUTHSD_DATE," +
                        "   RQ.CREATED_BY," +
                        "   RQ.REQUESTED_BY," +
                        "   RQ.REQ_BY_POS," +
                        "   RQ.AUTHSD_BY," +
                        "   RQ.AUTHSD_POSITION," +
                        "   DECODE(RQ.HDR_140_STATUS," +
                        "       'P', 'PENDING'," +
                        "       '0', 'NOT PRINTED'," +
                        "       '1', 'PRINT REQUESTED'," +
                        "       '2', 'PARTIALLY ACQUITTED'," +
                        "       '3', 'IDR COMPLETED'," +
                        "       '9', 'COMPLETE'," +
                        "       RQ.HDR_140_STATUS) REQ_STATUS," +
                        "   DECODE(RQ.AUTHSD_STATUS," +
                        "       'A', 'AUTHORIZED'," +
                        "       'U', 'UNAUTHORIZED'," +
                        "       RQ.AUTHSD_STATUS) AUTHSD_STATUS," +
                        "   RQ.IREQ_TYPE," +
                        "   RQ.ISS_TRAN_TYPE," +
                        "   RQ.ORIG_WHOUSE_ID," +
                        "   RQ.PRIORITY_CODE," +
                        "   EQ.EQUIP_CLASS," +
                        "   EQ.EQUIP_GRP_ID," +
                        "   EQ.PARENT_EQUIP" +
                        " FROM" +
                        "   ELLIPSE.MSF232 RQR JOIN ELLIPSE.MSF140 RQ" +
                        "     ON RQ.DSTRCT_CODE = RQR.DSTRCT_CODE" +
                        "     AND SUBSTR(RQR.REQUISITION_NO, 1, 6) = RQ.IREQ_NO" +
                        "     AND RQR.DSTRCT_CODE " + districtCode +
                        "     AND RQR.REQ_232_TYPE = 'I'" +
                        "   LEFT JOIN ELLIPSE.MSF620 WO" +
                        "     ON RQR.WORK_ORDER = WO.WORK_ORDER" +
                        "     AND RQR.DSTRCT_CODE = WO.DSTRCT_CODE" +
                        "   LEFT JOIN ELLIPSE.MSF600 EQ" +
                        "     ON WO.EQUIP_NO = EQ.EQUIP_NO" +
                        " WHERE" +
                        " " + queryCriteria1 +
                        " " + queryCriteria2 +
                        " " + statusRequirement +
                        dateParameters;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetRequisitionControlListQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, int dateCriteriaKey, string startDate, string endDate, string reqStatus)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrWhiteSpace(districtCode))
                districtCode = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            var queryCriteria1 = "";
            //establecemos los parámetros del criterio 1
            if (searchCriteriaKey1 == SearchFieldCriteriaType.Requisition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.IREQ_NO = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.CreatedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.CREATED_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.RequestedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.REQUESTED_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.RequestedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.REQ_BY_POS = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AuthorizedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.AUTHSD_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AuthorizedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQ.AUTHSD_POSITION = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.GL_ACCOUNT = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.WORK_ORDER = '" + searchCriteriaValue1 + "'";
            //else if (searchCriteriaKey1 == SearchFieldCriteriaType.StockCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            //    queryCriteria1 = "AND  = '" + searchCriteriaValue1 + "'";
            //else if (searchCriteriaKey1 == SearchFieldCriteriaType.PartNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            //    queryCriteria1 = "AND  = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND RQR.EQUIP_NO = '" + searchCriteriaValue1 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.PARENT_WO = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            //


            var queryCriteria2 = "";
            //establecemos los parámetros del criterio 2
            if (searchCriteriaKey2 == SearchFieldCriteriaType.Requisition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.IREQ_NO = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.CreatedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.CREATED_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.RequestedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.REQUESTED_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.RequestedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.REQ_BY_POS = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AuthorizedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.AUTHSD_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AuthorizedPos.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQ.AUTHSD_POSITION = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.GL_ACCOUNT = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.WORK_ORDER = '" + searchCriteriaValue2 + "'";
            //else if (searchCriteriaKey2 == SearchFieldCriteriaType.StockCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            //    queryCriteria2 = "AND  = '" + searchCriteriaValue2 + "'";
            //else if (searchCriteriaKey2 == SearchFieldCriteriaType.PartNumber.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            //    queryCriteria2 = "AND  = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND RQR.EQUIP_NO = '" + searchCriteriaValue2 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.PARENT_WO = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "')";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = "AND RQR.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            //


            //establecemos los parámetros de estado de orden
            var statusRequirement = "";
            if (string.IsNullOrWhiteSpace(reqStatus))
                statusRequirement = "";
            else if (reqStatus == RequisitionStatus.Uncompleted.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9'";
            else if (reqStatus == RequisitionStatus.Pending.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> 'P'";
            else if (reqStatus == RequisitionStatus.Unauthorized.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9' AND (TRIM(RQ.AUTHSD_STATUS) IS NULL OR TRIM(RQ.AUTHSD_STATUS) = 'U') ";
            else if (reqStatus == RequisitionStatus.Authorized.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS <> '9' AND TRIM(RQ.AUTHSD_STATUS) = 'A'";
            else if (reqStatus == RequisitionStatus.Awaiting.Value)
                statusRequirement = " AND RQ.HDR_140_STATUS IN ('1', '2', '3') AND TRIM(RQ.AUTHSD_STATUS) = 'A'";

            //establecemos los parámetros para el rango de fechas
            string dateParameters;
            if (string.IsNullOrWhiteSpace(startDate))
                startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
            if (string.IsNullOrWhiteSpace(endDate))
                endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                dateParameters = "";
            else if (dateCriteriaKey == SearchDateCriteriaType.Creation.Key)
                dateParameters = " AND RQ.CREATION_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.Required.Key)
                dateParameters = " AND RQ.REQ_BY_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.WoRaisedDate.Key)
                dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.WoPlanStartDate.Key)
                dateParameters = " AND WO.PLAN_STR_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.IgnoreDate.Key)
                dateParameters = " ";
            else
                dateParameters = " AND RQ.CREATION_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";

            var query = "WITH RQWO AS (" +
                        "   SELECT RQ.DSTRCT_CODE," +
                        "     WO.WORK_GROUP," +
                        "     RQR.EQUIP_NO," +
                        "     RQR.WORK_ORDER," +
                        "     WO.WO_DESC," +
                        "     RQR.PROJECT_NO," +
                        "     RQR.GL_ACCOUNT," +
                        "     RQ.IREQ_NO," +
                        "     RQ.NUM_OF_ITEMS," +
                        "     WO.RAISED_DATE WO_RAISED_DATE," +
                        "     WO.PLAN_STR_DATE WO_PLAN_STR_DATE," +
                        "     RQ.CREATION_DATE," +
                        "     RQ.REQ_BY_DATE," +
                        "     RQ.AUTHSD_DATE," +
                        "     RQ.CREATED_BY," +
                        "     RQ.REQUESTED_BY," +
                        "     RQ.REQ_BY_POS," +
                        "     RQ.AUTHSD_BY," +
                        "     RQ.AUTHSD_POSITION," +
                        "     DECODE(RQ.HDR_140_STATUS," +
                        "         'P', 'PENDING'," +
                        "         '0', 'NOT PRINTED'," +
                        "         '1', 'PRINT REQUESTED'," +
                        "         '2', 'PARTIALLY ACQUITTED'," +
                        "         '3', 'IDR COMPLETED'," +
                        "         '9', 'COMPLETE'," +
                        "         RQ.HDR_140_STATUS) REQ_STATUS," +
                        "     DECODE(RQ.AUTHSD_STATUS," +
                        "         'A', 'AUTHORIZED'," +
                        "         'U', 'UNAUTHORIZED'," +
                        "         RQ.AUTHSD_STATUS) AUTHSD_STATUS," +
                        "     RQ.IREQ_TYPE," +
                        "     RQ.ISS_TRAN_TYPE," +
                        "     RQ.ORIG_WHOUSE_ID," +
                        "     RQ.PRIORITY_CODE," +
                        "     EQ.EQUIP_CLASS," +
                        "     EQ.EQUIP_GRP_ID," +
                        "     EQ.PARENT_EQUIP" +
                        "   FROM" +
                        "     ELLIPSE.MSF232 RQR JOIN ELLIPSE.MSF140 RQ" +
                        "       ON RQ.DSTRCT_CODE = RQR.DSTRCT_CODE" +
                        "       AND SUBSTR(RQR.REQUISITION_NO, 1, 6) = RQ.IREQ_NO" +
                        "       AND RQR.DSTRCT_CODE " + districtCode +
                        "       AND RQR.REQ_232_TYPE = 'I'" +
                        "     LEFT JOIN ELLIPSE.MSF620 WO" +
                        "       ON RQR.WORK_ORDER = WO.WORK_ORDER" +
                        "       AND RQR.DSTRCT_CODE = WO.DSTRCT_CODE" +
                        "     LEFT JOIN ELLIPSE.MSF600 EQ" +
                        "       ON WO.EQUIP_NO = EQ.EQUIP_NO" +
                        "   WHERE" +
                        "   " + queryCriteria1 +
                        "   " + queryCriteria2 +
                        "   " + statusRequirement +
                        dateParameters +
                        " )," +
                        " REALREQ AS (" +
                        "   SELECT RQWO.DSTRCT_CODE," +
                        "     RQWO.WORK_GROUP," +
                        "     RQWO.EQUIP_NO," +
                        "     RQWO.WORK_ORDER," +
                        "     RQWO.WO_DESC," +
                        "     RQWO.PROJECT_NO," +
                        "     RQWO.GL_ACCOUNT," +
                        "     RQWO.WO_PLAN_STR_DATE," +
                        "     LISTAGG(RQWO.IREQ_NO, ',') WITHIN GROUP(ORDER BY RQWO.IREQ_NO) AS IREQS_NO," +
                        "     RQI.STOCK_CODE," +
                        "     RQWO.REQ_BY_DATE," +
                        "     SUM(RQI.QTY_REQ) QTY_REQUIRED," +
                        "     SUM(RQI.QTY_ISSUED) QTY_ISSUED," +
                        "     SUM(RQI.QTY_RECEIVED) QTY_RECEIVED," +
                        "     RQI.ITEM_PRICE," + //Precio de Compra
                        "     SCINV.INVENT_COST_PR INV_PRICE_PR," + // Precio de Inventario Primario
                        "     SCINV.INVENT_PRICE_S INV_PRICE_SC" +// Precio de Inventario Secundario
                        "   FROM " +
                        "     RQWO LEFT JOIN ELLIPSE.MSF141 RQI" +
                        "       ON RQWO.DSTRCT_CODE = RQI.DSTRCT_CODE" +
                        "       AND RQWO.IREQ_NO = RQI.IREQ_NO" +
                        "     LEFT JOIN ELLIPSE.MSF170 SCINV" +
                        "       ON RQI.STOCK_CODE = SCINV.STOCK_CODE AND RQWO.DSTRCT_CODE = SCINV.DSTRCT_CODE" +
                        "   GROUP BY " +
                        "     RQWO.DSTRCT_CODE," +
                        "     RQWO.WORK_GROUP," +
                        "     RQWO.EQUIP_NO," +
                        "     RQWO.WORK_ORDER," +
                        "     RQWO.WO_DESC," +
                        "     RQWO.PROJECT_NO," +
                        "     RQWO.GL_ACCOUNT," +
                        "     RQWO.WO_PLAN_STR_DATE," +
                        "     RQI.STOCK_CODE," +
                        "     RQWO.REQ_BY_DATE," +
                        "     RQI.ITEM_PRICE," +
                        "     SCINV.INVENT_COST_PR," +
                        "     SCINV.INVENT_PRICE_S" +
                        " )," +
                        " PLANREQ AS(" +
                        "   SELECT " +
                        "     RQWO.DSTRCT_CODE," +
                        "     RQWO.WORK_GROUP," +
                        "     RQWO.EQUIP_NO," +
                        "     RQWO.WORK_ORDER," +
                        "     RQWO.WO_DESC," +
                        "     RQWO.PROJECT_NO," +
                        "     RQWO.GL_ACCOUNT," +
                        "     RQWO.WO_PLAN_STR_DATE," +
                        "     WORP.STOCK_CODE," +
                        "     RQWO.REQ_BY_DATE," +
                        "     SUM(WORP.UNIT_QTY_REQD) QTY_PLAN_REQUIRED," +
                        "     SUM(WORP.QTY_REQUISITIONED) QTY_PLAN_REQUISITIONED," +
                        "     SCINV.INVENT_COST_PR INV_PRICE_PR," +
                        "     SCINV.INVENT_PRICE_S INV_PRICE_SC" +
                        "   FROM " +
                        "     RQWO JOIN ELLIPSE.MSF623 WOT" +
                        "       ON RQWO.WORK_ORDER   = WOT.WORK_ORDER" +
                        "       AND RQWO.DSTRCT_CODE = WOT.DSTRCT_CODE" +
                        "     JOIN ELLIPSE.MSF734 WORP" +
                        "       ON RQWO.DSTRCT_CODE || RQWO.WORK_ORDER || WOT.WO_TASK_NO = WORP.CLASS_KEY" +
                        "       AND WORP.CLASS_TYPE = 'WT'" +
                        "     LEFT JOIN ELLIPSE.MSF170 SCINV" +
                        "       ON WORP.STOCK_CODE = SCINV.STOCK_CODE  AND RQWO.DSTRCT_CODE = SCINV.DSTRCT_CODE" +
                        "   GROUP BY " +
                        "     RQWO.DSTRCT_CODE," +
                        "     RQWO.WORK_GROUP," +
                        "     RQWO.EQUIP_NO," +
                        "     RQWO.WORK_ORDER," +
                        "     RQWO.WO_DESC," +
                        "     RQWO.PROJECT_NO," +
                        "     RQWO.GL_ACCOUNT," +
                        "     RQWO.WO_PLAN_STR_DATE," +
                        "     WORP.STOCK_CODE," +
                        "     RQWO.REQ_BY_DATE," +
                        "     SCINV.INVENT_COST_PR," +
                        "     SCINV.INVENT_PRICE_S" +
                        " )" +
                        " SELECT " +
                        "   COALESCE(RQ.DSTRCT_CODE, PQ.DSTRCT_CODE) DSTRCT_CODE," +
                        "   COALESCE(RQ.WORK_GROUP, PQ.WORK_GROUP) WORK_GROUP," +
                        "   COALESCE(RQ.EQUIP_NO, PQ.EQUIP_NO) EQUIP_NO," +
                        "   COALESCE(RQ.WORK_ORDER, PQ.WORK_ORDER) WORK_ORDER," +
                        "   COALESCE(RQ.WO_DESC, PQ.WO_DESC) WO_DESC," +
                        "   COALESCE(RQ.PROJECT_NO, PQ.PROJECT_NO) PROJECT_NO," +
                        "   COALESCE(RQ.GL_ACCOUNT, PQ.GL_ACCOUNT) GL_ACCOUNT," +
                        "   COALESCE(RQ.WO_PLAN_STR_DATE, PQ.WO_PLAN_STR_DATE) WO_PLAN_STAR_DATE," +
                        "   RQ.IREQS_NO," +
                        "   COALESCE(RQ.STOCK_CODE, PQ.STOCK_CODE) STOCK_CODE," +
                        "   COALESCE(RQ.REQ_BY_DATE, PQ.REQ_BY_DATE) REQ_BY_DATE," +
                        "   NVL(PQ.QTY_PLAN_REQUIRED, 0) QTY_PLAN_REQUIRED," +
                        "   NVL(PQ.QTY_PLAN_REQUISITIONED, 0) QTY_PLAN_REQUISITIONED," +
                        "   NVL(RQ.QTY_REQUIRED, 0) QTY_REQUIRED," +
                        "   NVL(RQ.QTY_ISSUED, 0) QTY_ISSUED," +
                        "   NVL(RQ.QTY_RECEIVED, 0) QTY_RECEIVED," +
                        "   RQ.ITEM_PRICE," +
                        "   COALESCE(RQ.INV_PRICE_PR, PQ.INV_PRICE_PR) INV_PRICE_PR," +
                        "   COALESCE(RQ.INV_PRICE_SC, PQ.INV_PRICE_SC) INV_PRICE_SC" +
                        " FROM " +
                        "   REALREQ RQ FULL JOIN PLANREQ PQ" +
                        "     ON RQ.DSTRCT_CODE = PQ.DSTRCT_CODE" +
                        "     AND RQ.WORK_GROUP = PQ.WORK_GROUP" +
                        "     AND RQ.EQUIP_NO = PQ.EQUIP_NO" +
                        "     AND RQ.WORK_ORDER = PQ.WORK_ORDER" +
                        "     AND RQ.WO_DESC = PQ.WO_DESC" +
                        "     AND RQ.PROJECT_NO = PQ.PROJECT_NO" +
                        "     AND RQ.GL_ACCOUNT = PQ.GL_ACCOUNT" +
                        "     AND RQ.WO_PLAN_STR_DATE = PQ.WO_PLAN_STR_DATE" +
                        "     AND RQ.STOCK_CODE = PQ.STOCK_CODE";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

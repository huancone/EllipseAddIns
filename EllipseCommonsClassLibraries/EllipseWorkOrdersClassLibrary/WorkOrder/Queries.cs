using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;

namespace EllipseWorkOrdersClassLibrary
{
    public static partial class Queries
    {
        /// <summary>
        /// Obtiene el Query para la consulta de una o más órdenes de trabajo
        /// </summary>
        /// <param name="dbReference">string: Referencia de base de datos (Ej: MIMSPROD, ELLIPSE) </param>
        /// <param name="dbLink">string: link de conexión de base de datos (Ej: @MLDBMIMS)</param>
        /// <param name="districtCode">string: distrito de consulta. Si es nulo se consulta para todos los distritos</param>
        /// <param name="searchCriteriaKey1">int: Indica el tipo de búsqueda según la clase SearchFieldCriteriaType.Type.Key. Valor por defecto (0 - None). (Ej: SearchFieldCriteriaType.WorkGroup.Key) </param>
        /// <param name="searchCriteriaValue1"></param>
        /// <param name="searchCriteriaKey2"></param>
        /// <param name="searchCriteriaValue2"></param>
        /// <param name="dateCriteriaKey"></param>
        /// <param name="startDate">string: fecha en format yyyyMMdd para parámetro de fecha inicial. Predeterminado inicio del año</param>
        /// <param name="endDate">string: fecha en format yyyyMMdd para parámetro de fecha final. Predeterminado fecha de hoy</param>
        /// <param name="woStatus">string: especifica qué estado de la orden se va a consultar WoStatusList.StatusName. Si es nulo se consulta cualquier estado></param>
        /// <returns>string: Query de consulta para ejecución</returns>
        public static string GetFetchWoQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, int dateCriteriaKey, string startDate, string endDate, string woStatus)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrWhiteSpace(districtCode))
                districtCode = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            var queryCriteria1 = "";
            //establecemos los parámetros del criterio 1
            if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.EQUIP_NO = '" + searchCriteriaValue1 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.ORIGINATOR_ID = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.COMPLETED_BY = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND TRIM(SUBSTR(WO.DSTRCT_ACCT_CODE, 5)) = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.WorkRequest.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.REQUEST_ID = '" + searchCriteriaValue1.PadLeft(12, '0') + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.PARENT_WO = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
            {
                if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria1 = "AND WO.EQUIP_NO IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
            }
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND EQ.EQUIP_CLASS = '" + searchCriteriaValue1 + "'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue1).Select(g => g.Name).ToList(), ",", "'") + ")";
            else
                queryCriteria1 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ",", "'") + ")";
            //

            var queryCriteria2 = "";
            //establecemos los parámetros del criterio 2
            if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.EQUIP_NO = '" + searchCriteriaValue2 + "'"; //Falta buscar el equip ref //to do
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.PARENT_EQUIP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Originator.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.ORIGINATOR_ID = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.CompletedBy.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria1 = "AND WO.COMPLETED_BY = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND TRIM(SUBSTR(WO.DSTRCT_ACCT_CODE, 5)) = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkRequest.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.REQUEST_ID = '" + searchCriteriaValue2.PadLeft(12, '0') + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ParentWorkOrder.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.PARENT_WO = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "'";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria2 = "AND TRIM(WO.EQUIP_NO) IN (SELECT DISTINCT TRIM(LI.MEM_EQUIP_GRP) EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "'";
            }
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND EQ.EQUIP_CLASS = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + MyUtilities.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            //

            //establecemos los parámetros de estado de orden
            string statusRequirement;
            if (string.IsNullOrWhiteSpace(woStatus))
                statusRequirement = "";
            else if (woStatus == WoStatusList.Uncompleted)
                statusRequirement = " AND WO.WO_STATUS_M IN (" + MyUtilities.GetListInSeparator(WoStatusList.GetUncompletedStatusCodes(), ",", "'") + ")";
            else if (WoStatusList.GetStatusNames().Contains(woStatus))
                statusRequirement = " AND WO.WO_STATUS_M = '" + WoStatusList.GetStatusCode(woStatus) + "'";
            else
                statusRequirement = "";

            //establecemos los parámetros para el rango de fechas
            string dateParameters;
            if (string.IsNullOrWhiteSpace(startDate))
                startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
            if (string.IsNullOrWhiteSpace(endDate))
                endDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

            if (dateCriteriaKey == SearchDateCriteriaType.Raised.Key)
                dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.Closed.Key)
                dateParameters = " AND WO.CLOSED_DT BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.RequiredBy.Key)
                dateParameters = " AND WO.REQ_BY_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.RequiredStart.Key)
                dateParameters = " AND WO.REQ_START_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.PlannedStart.Key)
                dateParameters = " AND WO.PLAN_STR_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.PlannedFinnish.Key)
                dateParameters = " AND WO.PLAN_FIN_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.LastModified.Key)
                dateParameters = " AND WO.LAST_MOD_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            else if (dateCriteriaKey == SearchDateCriteriaType.NotFinalized.Key)
                dateParameters = " AND WO.CLOSED_DT BETWEEN '" + startDate + "' AND '" + endDate + "' AND WO.FINAL_COSTS <> 'Y'";
            else
                dateParameters = " AND WO.RAISED_DATE BETWEEN '" + startDate + "' AND '" + endDate + "'";
            //escribimos el query
            var query = "" +
                        " SELECT" +
                        " WO.DSTRCT_CODE, WO.WORK_GROUP, WO.WORK_ORDER, WO.WO_STATUS_M, WO.WO_DESC, " +
                        " WO.EQUIP_NO, WO.COMP_CODE, WO.COMP_MOD_CODE, WO.LOCATION, WO.RAISED_DATE, WO.RAISED_TIME," +
                        " WO.ORIGINATOR_ID, WO.ORIG_PRIORITY, WO.ORIG_DOC_TYPE, WO.ORIG_DOC_NO, WO.REQUEST_ID, WO.MSSS_STATUS_IND," +
                        " WO.WO_TYPE, WO.MAINT_TYPE, WO.WO_STATUS_U, WO.STD_JOB_NO, WO.MAINT_SCH_TASK, WO.AUTO_REQ_IND, WO.ASSIGN_PERSON, WO.PLAN_PRIORITY, WO.CLOSED_COMMIT_DT, WO.UNIT_OF_WORK, WO.UNITS_REQUIRED, FAILURE_PART, PC_COMPLETE, UNITS_COMPLETE, WO.RELATED_WO," +
                        " WO.REQ_START_DATE, WO.REQ_START_TIME, WO.REQ_BY_DATE, WO.REQ_BY_TIME, WO.PLAN_STR_DATE, WO.PLAN_STR_TIME, WO.PLAN_FIN_DATE, WO.PLAN_FIN_TIME," +
                        " SUBSTR(WO.DSTRCT_ACCT_CODE, 5) DSTRCT_ACCT_CODE, WO.PROJECT_NO, WO.PARENT_WO," +
                        " WO.WO_JOB_CODEX1, WO.WO_JOB_CODEX2, WO.WO_JOB_CODEX3, WO.WO_JOB_CODEX4, WO.WO_JOB_CODEX5, WO.WO_JOB_CODEX6, WO.WO_JOB_CODEX7, WO.WO_JOB_CODEX8, WO.WO_JOB_CODEX9, WO.WO_JOB_CODEX10," +
                        " CASE WHEN TRIM(WO.WO_JOB_CODEX1||WO.WO_JOB_CODEX2||WO.WO_JOB_CODEX3||WO.WO_JOB_CODEX4||WO.WO_JOB_CODEX5||WO.WO_JOB_CODEX6||WO.WO_JOB_CODEX7||WO.WO_JOB_CODEX8||WO.WO_JOB_CODEX9||WO.WO_JOB_CODEX10) IS NULL THEN 'N' ELSE 'Y' END JOB_CODES," +
                        " WO.COMPLETED_CODE, WO.COMPLETED_BY," +
                        " CASE WHEN WO.DSTRCT_CODE || WO.WORK_ORDER NOT IN (SELECT STV.STD_KEY FROM " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " STV WHERE STV.STD_TEXT_CODE=('CW')) THEN 'N' ELSE 'Y' END COMPLETE_TEXT_FLAG," +
                        " WO.CLOSED_DT," +
                        " WOEST.CALC_DUR_HRS_SW, WOEST.EST_DUR_HRS, WOEST.ACT_DUR_HRS," +
                        " WOEST.RES_UPDATE_FLAG, WOEST.EST_LAB_HRS, WOEST.CALC_LAB_HRS, WOEST.ACT_LAB_HRS, WOEST.EST_LAB_COST, WOEST.CALC_LAB_COST, WOEST.ACT_LAB_COST," +
                        " WOEST.MAT_UPDATE_FLAG, WOEST.EST_MAT_COST, WOEST.CALC_MAT_COST, WOEST.ACT_MAT_COST," +
                        " WOEST.EQUIP_UPDATE_FLAG, WOEST.EST_EQUIP_COST, WOEST.CALC_EQUIP_COST, WOEST.ACT_EQUIP_COST," +
                        " WOEST.EST_OTHER_COST, WOEST.ACT_OTHER_COST," +
                        " WO.LOCATION_FR, WO.LOCATION, WO.NOTICE_LOCN," +
                        " WO.LAST_MOD_DATE, WO.FINAL_COSTS," +
                        " EQ.EQUIP_CLASS, EQ.EQUIP_GRP_ID, EQ.PARENT_EQUIP" +
                        " FROM" +
                        " " + dbReference + ".MSF620" + dbLink + " WO LEFT JOIN " + dbReference + ".MSF621" + dbLink + " WOEST ON (WO.WORK_ORDER    = WOEST.WORK_ORDER AND WO.DSTRCT_CODE = WOEST.DSTRCT_CODE)" +
                        " " + "LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON WO.EQUIP_NO = EQ.EQUIP_NO" +
                        " WHERE" +
                        " " + queryCriteria1 +
                        " " + queryCriteria2 +
                        " " + statusRequirement +
                        " AND WO.DSTRCT_CODE " + districtCode +
                        dateParameters;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        /// <summary>
        /// Obtiene el Query para la consulta de una o más órdenes de trabajo
        /// </summary>
        /// <param name="dbReference">string: Referencia de base de datos (Ej: MIMSPROD, ELLIPSE) </param>
        /// <param name="dbLink">string: link de conexión de base de datos (Ej: @MLDBMIMS)</param>
        /// <param name="districtCode">string: distrito de consulta. Si es nulo se consulta para todos los distritos</param>
        /// <param name="workOrder">string: número de la orden de trabajo</param>
        /// <returns>string: Query de consulta para ejecución</returns>
        public static string GetFetchWoQuery(string dbReference, string dbLink, string districtCode, string workOrder)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrWhiteSpace(districtCode))
                districtCode = " IN (" + MyUtilities.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            //escribimos el query
            var query = "" +
                        " SELECT" +
                        " WO.DSTRCT_CODE, WO.WORK_GROUP, WO.WORK_ORDER, WO.WO_STATUS_M, WO.WO_DESC, " +
                        " WO.EQUIP_NO, WO.COMP_CODE, WO.COMP_MOD_CODE, WO.LOCATION, WO.RAISED_DATE, WO.RAISED_TIME," +
                        " WO.ORIGINATOR_ID, WO.ORIG_PRIORITY, WO.ORIG_DOC_TYPE, WO.ORIG_DOC_NO, WO.REQUEST_ID, WO.MSSS_STATUS_IND," +
                        " WO.WO_TYPE, WO.MAINT_TYPE, WO.WO_STATUS_U, WO.STD_JOB_NO, WO.MAINT_SCH_TASK, WO.AUTO_REQ_IND, WO.ASSIGN_PERSON, WO.PLAN_PRIORITY, WO.CLOSED_COMMIT_DT, WO.UNIT_OF_WORK, WO.UNITS_REQUIRED, FAILURE_PART, PC_COMPLETE, UNITS_COMPLETE, WO.RELATED_WO," +
                        " WO.REQ_START_DATE, WO.REQ_START_TIME, WO.REQ_BY_DATE, WO.REQ_BY_TIME, WO.PLAN_STR_DATE, WO.PLAN_STR_TIME, WO.PLAN_FIN_DATE, WO.PLAN_FIN_TIME," +
                        " SUBSTR(WO.DSTRCT_ACCT_CODE, 5) DSTRCT_ACCT_CODE, WO.PROJECT_NO, WO.PARENT_WO," +
                        " WO.WO_JOB_CODEX1, WO.WO_JOB_CODEX2, WO.WO_JOB_CODEX3, WO.WO_JOB_CODEX4, WO.WO_JOB_CODEX5, WO.WO_JOB_CODEX6, WO.WO_JOB_CODEX7, WO.WO_JOB_CODEX8, WO.WO_JOB_CODEX9, WO.WO_JOB_CODEX10," +
                        " CASE WHEN TRIM(WO.WO_JOB_CODEX1||WO.WO_JOB_CODEX2||WO.WO_JOB_CODEX3||WO.WO_JOB_CODEX4||WO.WO_JOB_CODEX5||WO.WO_JOB_CODEX6||WO.WO_JOB_CODEX7||WO.WO_JOB_CODEX8||WO.WO_JOB_CODEX9||WO.WO_JOB_CODEX10) IS NULL THEN 'N' ELSE 'Y' END JOB_CODES," +
                        " WO.COMPLETED_CODE, WO.COMPLETED_BY," +
                        " CASE WHEN WO.DSTRCT_CODE || WO.WORK_ORDER NOT IN (SELECT STV.STD_KEY FROM " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " STV WHERE STV.STD_TEXT_CODE=('CW')) THEN 'N' ELSE 'Y' END COMPLETE_TEXT_FLAG," +
                        " WO.CLOSED_DT," +
                        " WOEST.CALC_DUR_HRS_SW, WOEST.EST_DUR_HRS, WOEST.ACT_DUR_HRS," +
                        " WOEST.RES_UPDATE_FLAG, WOEST.EST_LAB_HRS, WOEST.CALC_LAB_HRS, WOEST.ACT_LAB_HRS, WOEST.EST_LAB_COST, WOEST.CALC_LAB_COST, WOEST.ACT_LAB_COST," +
                        " WOEST.MAT_UPDATE_FLAG, WOEST.EST_MAT_COST, WOEST.CALC_MAT_COST, WOEST.ACT_MAT_COST," +
                        " WOEST.EQUIP_UPDATE_FLAG, WOEST.EST_EQUIP_COST, WOEST.CALC_EQUIP_COST, WOEST.ACT_EQUIP_COST," +
                        " WOEST.EST_OTHER_COST, WOEST.ACT_OTHER_COST," +
                        " WO.LOCATION_FR, WO.LOCATION, WO.NOTICE_LOCN," +
                        " WO.LAST_MOD_DATE, WO.FINAL_COSTS" +
                        " FROM" +
                        " " + dbReference + ".MSF620" + dbLink + " WO LEFT JOIN " + dbReference + ".MSF621" + dbLink + " WOEST ON (WO.WORK_ORDER    = WOEST.WORK_ORDER AND WO.DSTRCT_CODE = WOEST.DSTRCT_CODE)" +
                        " WHERE" +
                        " WO.WORK_ORDER = '" + workOrder + "'" +
                        " AND WO.DSTRCT_CODE " + districtCode;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }


        public static string GetFetchOrigDocNo(string dbReference, string dbLink, string districtCode, string workGroup, string origDocType, string origDocNo)
        {
            var query = "";
            query += "SELECT ";
            query += "    WORK_ORDER ";
            query += "FROM ";
            query += "    " + dbReference + ".MSF620" + dbLink + " WO ";
            query += "WHERE ";
            query += "    WO.DSTRCT_CODE = '" + districtCode + "' ";
            query += "    AND WO.WORK_GROUP = '" + workGroup + "' ";
            query += "    AND WO.ORIG_DOC_TYPE = 'OT' ";
            query += "    AND WO.ORIG_DOC_NO = '" + origDocNo + "' ";
            return query;
        }

        public static string GetFetchWoRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string reqType, string taskNo)
        {
            if (reqType.Equals(RequirementType.Labour.Key))
            {
                return GetFetchWoLabourRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.Material.Key))
            {
                return GetFetchWoMaterialRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.Equipment.Key))
            {
                return GetFetchWoEquipmentRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
            }
            else if (reqType.Equals(RequirementType.All.Key))
            {
                var labourSql = GetFetchWoLabourRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
                var materialSql = GetFetchWoMaterialRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);
                var equipmentSql = GetFetchWoEquipmentRequirementsQuery(dbReference, dbLink, districtCode, workOrder, taskNo);

                return labourSql + " UNION ALL " + materialSql + " UNION ALL " + equipmentSql;
            }
            return null;
        }


        public static string GetFetchWoLabourRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                " SELECT '" + RequirementType.Labour.Key + "' REQ_TYPE, " +
                "   COALESCE(WOR.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                "   COALESCE(WOR.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                "   COALESCE(WOR.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                "   COALESCE(WOR.WO_DESC, TRR.WO_DESC) WO_TASK_DESC, " + //Aunque es WO_DESC por compatibilidad estructural de consulta con tareas se coloca así
                "   'ORDER' WO_TASK_NO, " + //Compatibilidad estructural
                "   'N/A' SEQ_NO, " +
                "   COALESCE(WOR.RESOURCE_TYPE, TRR.RES_CODE) RES_CODE, " +
                "   WOR.EST_SIZE, " +
                "   WOR.UNITS_QTY, " +
                "   TRR.ACT_RESRCE_HRS REAL_QTY, " +
                "   COALESCE(WOR.RES_DESC, TRR.RES_DESC) RES_DESC, " +
                "   'HR' UNITS, " +
                "   WOR.SHARED_TASKS " +
                " FROM( " +
                "   SELECT WO.DSTRCT_CODE, " +
                "     WO.WORK_GROUP, " +
                "     WO.WORK_ORDER, " +
                "     WO_DESC, " +
                "     RS.RESOURCE_TYPE, " +
                "     SUM(TO_NUMBER(RS.CREW_SIZE)) EST_SIZE, " +
                "     SUM(RS.EST_RESRCE_HRS) UNITS_QTY, " +
                "     TT.TABLE_DESC RES_DESC, " +
                "     COUNT(TSK.WO_TASK_NO) SHARED_TASKS " +
                "   FROM " +
                "     " + dbReference + ".MSF620" + dbLink + " WO INNER JOIN " + dbReference + ".MSF623" + dbLink + " TSK " +
                "       ON WO.DSTRCT_CODE = TSK.DSTRCT_CODE " +
                "       AND WO.WORK_ORDER = TSK.WORK_ORDER " +
                "     INNER JOIN " + dbReference + ".MSF735" + dbLink + " RS " +
                "       ON RS.KEY_735_ID = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                "       AND RS.REC_735_TYPE = 'WT' " +
                "     INNER JOIN " + dbReference + ".MSF010" + dbLink + " TT " +
                "       ON TT.TABLE_CODE = RS.RESOURCE_TYPE " +
                "       AND TT.TABLE_TYPE = 'TT' " +
                "   WHERE WO.DSTRCT_CODE = '" + districtCode + "' " +
                "     AND WO.WORK_ORDER = '" + workOrder + "' " +
                "   GROUP BY WO.DSTRCT_CODE, " +
                "     WO.WORK_GROUP, " +
                "     WO.WORK_ORDER, " +
                "     WO.WO_DESC, " +
                "     RS.RESOURCE_TYPE, " +
                "     TT.TABLE_DESC) WOR " +
                //Real Calculation
                "   FULL JOIN " +
                "   ( " +
                "     SELECT TR.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       TX.WORK_ORDER, " +
                "       WO.WO_DESC, " +
                "       TR.RESOURCE_TYPE RES_CODE, " +
                "       LTT.TABLE_DESC RES_DESC, " +
                "       SUM(TR.NO_OF_HOURS) ACT_RESRCE_HRS " +
                "     FROM " + dbReference + ".MSFX99" + dbLink + " TX " +
                "     INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR " +
                "     ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                "     AND TR.WORK_ORDER = TX.WORK_ORDER " +
                "     AND TR.USERNO = TX.USERNO " +
                "     AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                "     AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                "     AND TR.REC900_TYPE = TX.REC900_TYPE " +
                "     AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                "     AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                "     INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO " +
                "     ON TX.DSTRCT_CODE = WO.DSTRCT_CODE " +
                "     AND TX.WORK_ORDER = WO.WORK_ORDER " +
                "     INNER JOIN " + dbReference + ".MSF010" + dbLink + " LTT " +
                "     ON LTT.TABLE_CODE = TR.RESOURCE_TYPE " +
                "     AND LTT.TABLE_TYPE = 'TT' " +
                "     WHERE TX.DSTRCT_CODE = '" + districtCode + "' " +
                "     AND TX.WORK_ORDER = '" + workOrder + "' " +
                "     AND TX.REC900_TYPE = 'L' " +
                "     GROUP BY TR.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       TX.WORK_ORDER, " +
                "       WO.WO_DESC, " +
                "       TR.RESOURCE_TYPE, " +
                "       LTT.TABLE_DESC " +
                "   ) TRR ON WOR.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                "     AND WOR.WORK_GROUP = TRR.WORK_GROUP " +
                "     AND WOR.WORK_ORDER = TRR.WORK_ORDER " +
                "     AND WOR.RESOURCE_TYPE = TRR.RES_CODE ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchWoMaterialRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                " SELECT 'MAT' REQ_TYPE, " +
                "     COALESCE(WOR.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                "     COALESCE(WOR.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                "     COALESCE(WOR.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                "     'ORDER' WO_TASK_NO, " +
                "     COALESCE(WOR.WO_TASK_DESC, TRR.WO_TASK_DESC) WO_TASK_DESC, " +
                "     'N/A' SEQ_NO, " +
                "     COALESCE(WOR.RES_CODE, TRR.RES_CODE) RES_CODE, " +
                "     COALESCE(WOR.SHARED_TASKS, 1) EST_SIZE, " +
                "     WOR.UNITS_QTY, " +
                "     TRR.QTY_ISS REAL_QTY, " +
                "     COALESCE(WOR.RES_DESC, TRR.RES_DESC) RES_DESC, " +
                "     COALESCE(WOR.UNITS, TRR.UNITS) UNITS, " +
                "     WOR.SHARED_TASKS " +
                " FROM( " +
                "     SELECT " +
                "     TSK.DSTRCT_CODE, " +
                "     TSK.WORK_GROUP, " +
                "     TSK.WORK_ORDER, " +
                "     WO.WO_DESC WO_TASK_DESC, " +
                "     RS.STOCK_CODE RES_CODE, " +
                "     SUM(RS.UNIT_QTY_REQD) UNITS_QTY, " +
                "     SCT.DESC_LINEX1 || SCT.ITEM_NAME RES_DESC, " +
                "     SCT.UNIT_OF_ISSUE UNITS, " +
                "     (SELECT COUNT(*) FROM " + dbReference + ".MSF734" + dbLink + " SRS " +
                "       WHERE SRS.CLASS_KEY LIKE TSK.DSTRCT_CODE || TSK.WORK_ORDER || '%' " +
                "       AND SRS.CLASS_TYPE = 'WT' " +
                "       AND SRS.STOCK_CODE = RS.STOCK_CODE " +
                "     ) SHARED_TASKS " +
                "     FROM " +
                "     " + dbReference + ".MSF620" + dbLink + " WO " +
                "     JOIN " + dbReference + ".MSF623" + dbLink + " TSK " +
                "     ON WO.DSTRCT_CODE = TSK.DSTRCT_CODE AND WO.WORK_ORDER = TSK.WORK_ORDER " +
                "     INNER JOIN " + dbReference + ".MSF734" + dbLink + " RS " +
                "     ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                "     AND RS.CLASS_TYPE = 'WT' " +
                "     LEFT JOIN " + dbReference + ".MSF100" + dbLink + " SCT " +
                "     ON RS.STOCK_CODE = SCT.STOCK_CODE " +
                "     WHERE TSK.DSTRCT_CODE = '" + districtCode + "' " +
                "     AND TSK.WORK_ORDER = '" + workOrder + "' " +
                "     GROUP BY " +
                "     TSK.DSTRCT_CODE, " +
                "     TSK.WORK_GROUP, " +
                "     TSK.WORK_ORDER, " +
                "     WO.WO_DESC, " +
                "     RS.STOCK_CODE, " +
                "     SCT.DESC_LINEX1, " +
                "     SCT.ITEM_NAME, " +
                "     SCT.UNIT_OF_ISSUE " +
                " ) WOR " +
                " FULL JOIN( " +
                "     SELECT TX.DSTRCT_CODE, " +
                "         WO.WORK_GROUP, " +
                "         TX.WORK_ORDER, " +
                "         WO.WO_DESC WO_TASK_DESC, " +
                "         TR.STOCK_CODE AS RES_CODE, " +
                "         SUM(TR.QUANTITY_ISS) QTY_ISS, " +
                "         STT.DESC_LINEX1 || STT.ITEM_NAME RES_DESC, " +
                "         STT.UNIT_OF_ISSUE UNITS " +
                "     FROM " + dbReference + ".MSFX99" + dbLink + " TX " +
                "     INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR " +
                "     ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                "     AND TR.WORK_ORDER = TX.WORK_ORDER " +
                "     AND TR.USERNO = TX.USERNO " +
                "     AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                "     AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                "     AND TR.REC900_TYPE = TX.REC900_TYPE " +
                "     AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                "     AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                "     INNER JOIN " + dbReference + ".MSF620" + dbLink + " WO " +
                "     ON WO.DSTRCT_CODE = TR.DSTRCT_CODE " +
                "     AND WO.WORK_ORDER = TR.WORK_ORDER " +
                "     LEFT JOIN " + dbReference + ".MSF100" + dbLink + " STT " +
                "     ON TR.STOCK_CODE = STT.STOCK_CODE " +
                "     WHERE TX.DSTRCT_CODE = '" + districtCode + "' " +
                "     AND TX.WORK_ORDER = '" + workOrder + "' " +
                "     AND TX.REC900_TYPE = 'S' " +
                "     GROUP BY TX.DSTRCT_CODE, " +
                "         WO.WORK_GROUP, " +
                "         TX.WORK_ORDER, " +
                "         WO.WO_DESC, " +
                "         TR.STOCK_CODE, " +
                "         STT.DESC_LINEX1, " +
                "         STT.ITEM_NAME, " +
                "         STT.UNIT_OF_ISSUE " +
                "     ) TRR ON WOR.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                " AND WOR.WORK_GROUP = TRR.WORK_GROUP " +
                " AND WOR.WORK_ORDER = TRR.WORK_ORDER " +
                " AND WOR.RES_CODE = TRR.RES_CODE ";


            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

        public static string GetFetchWoEquipmentRequirementsQuery(string dbReference, string dbLink, string districtCode, string workOrder, string taskNo)
        {
            var query = "" +
                " SELECT 'EQP' REQ_TYPE, " +
                "   COALESCE(WOR.DSTRCT_CODE, TRR.DSTRCT_CODE) DSTRCT_CODE, " +
                "   COALESCE(WOR.WORK_GROUP, TRR.WORK_GROUP) WORK_GROUP, " +
                "   COALESCE(WOR.WORK_ORDER, TRR.WORK_ORDER) WORK_ORDER, " +
                "   'ORDER' WO_TASK_NO, " +
                "   WOR.WO_TASK_DESC, " +
                "   'N/A' SEQ_NO, " +
                "   COALESCE(WOR.RES_CODE, TRR.RES_CODE) RES_CODE, " +
                "   WOR.EST_SIZE, " +
                "   WOR.UNITS_QTY, " +
                "   TRR.QTY_ISS REAL_QTY, " +
                "   COALESCE(WOR.RES_DESC, TRR.RES_DESC) RES_DESC, " +
                "   COALESCE(TRIM(WOR.UNITS), TRIM(TRR.UNITS)) UNITS, " +
                "   WOR.SHARED_TASKS " +
                "   FROM( " +
                "     SELECT " +
                "       WO.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       WO.WORK_ORDER, " +
                "       WO.WO_DESC WO_TASK_DESC, " +
                "       RS.EQPT_TYPE RES_CODE, " +
                "       SUM(TO_NUMBER(RS.QTY_REQ)) EST_SIZE, " +
                "       SUM(RS.UNIT_QTY_REQD) UNITS_QTY, " +
                "       EQT.TABLE_DESC RES_DESC, " +
                "       DECODE(TRIM(RS.UOM), 'H5', 'HR', TRIM(RS.UOM)) UNITS, " +
                "       (SELECT COUNT(*) FROM " + dbReference + ".MSF733" + dbLink + " SRS WHERE SRS.CLASS_KEY  LIKE WO.DSTRCT_CODE || WO.WORK_ORDER || '%' AND SRS.CLASS_TYPE = 'WT' AND SRS.EQPT_TYPE = RS.EQPT_TYPE) SHARED_TASKS " +
                " FROM " +
                "       " + dbReference + ".MSF620" + dbLink + " WO INNER JOIN " + dbReference + ".MSF623" + dbLink + " TSK " +
                "         ON WO.DSTRCT_CODE = TSK.DSTRCT_CODE AND WO.WORK_ORDER = TSK.WORK_ORDER " +
                "       INNER JOIN " + dbReference + ".MSF733" + dbLink + " RS " +
                "         ON RS.CLASS_KEY = TSK.DSTRCT_CODE || TSK.WORK_ORDER || TSK.WO_TASK_NO " +
                "         AND RS.CLASS_TYPE = 'WT' " +
                "       INNER JOIN " + dbReference + ".MSF010" + dbLink + " EQT " +
                "         ON RS.EQPT_TYPE = EQT.TABLE_CODE AND TABLE_TYPE = 'ET' " +
                "     WHERE " +
                "       WO.DSTRCT_CODE = '" + districtCode + "' " +
                "       AND WO.WORK_ORDER = '" + workOrder + "' " +
                "     GROUP BY " +
                "       WO.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       WO.WORK_ORDER, " +
                "       WO.WO_DESC, " +
                "       RS.EQPT_TYPE, " +
                "       EQT.TABLE_DESC, " +
                "       DECODE(TRIM(RS.UOM), 'H5', 'HR', TRIM(RS.UOM)) " +
                "   ) WOR " +
                " FULL JOIN " +
                "   ( " +
                "     SELECT TX.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       TX.WORK_ORDER, " +
                "       WO.WO_DESC WO_TASK_DESC, " +
                "       ETT.TABLE_CODE AS RES_CODE, " +
                "       SUM(TR.STAT_VALUE) QTY_ISS, " +
                "       ETT.TABLE_DESC RES_DESC, " +
                "       TR.STAT_TYPE UNITS " +
                "     FROM " +
                "       " + dbReference + ".MSF620" + dbLink + " WO " +
                "       INNER JOIN " + dbReference + ".MSFX99" + dbLink + " TX " +
                "         ON WO.DSTRCT_CODE = TX.DSTRCT_CODE " +
                "         AND WO.WORK_ORDER = TX.WORK_ORDER " +
                "       INNER JOIN " + dbReference + ".MSF900" + dbLink + " TR " +
                "         ON TR.FULL_PERIOD = TX.FULL_PERIOD " +
                "         AND TR.WORK_ORDER = TX.WORK_ORDER " +
                "         AND TR.USERNO = TX.USERNO " +
                "         AND TR.TRANSACTION_NO = TX.TRANSACTION_NO " +
                "         AND TR.ACCOUNT_CODE = TX.ACCOUNT_CODE " +
                "         AND TR.REC900_TYPE = TX.REC900_TYPE " +
                "         AND TR.PROCESS_DATE = TX.PROCESS_DATE " +
                "         AND TR.DSTRCT_CODE = TX.DSTRCT_CODE " +
                "       LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ " +
                "         ON TR.MEMO_EQUIP = EQ.EQUIP_NO " +
                "       LEFT JOIN " + dbReference + ".MSF010" + dbLink + " ETT " +
                "         ON EQ.EQPT_TYPE = ETT.TABLE_CODE " +
                "     WHERE WO.DSTRCT_CODE = '" + districtCode + "' " +
                "       AND WO.WORK_ORDER = '" + workOrder + "' " +
                "       AND TX.REC900_TYPE = 'E' " +
                "     GROUP BY TX.DSTRCT_CODE, " +
                "       WO.WORK_GROUP, " +
                "       TX.WORK_ORDER, " +
                "       WO.WO_DESC, " +
                "       ETT.TABLE_CODE, " +
                "       ETT.TABLE_DESC, " +
                "       TR.STAT_TYPE " +
                "   ) TRR " +
                "   ON WOR.DSTRCT_CODE = TRR.DSTRCT_CODE " +
                "   AND WOR.WORK_GROUP = TRR.WORK_GROUP " +
                "   AND WOR.WORK_ORDER = TRR.WORK_ORDER " +
                "   AND WOR.RES_CODE = TRR.RES_CODE " +
                "   AND TRIM(WOR.UNITS) = TRIM(TRR.UNITS) ";


            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }

    }
}

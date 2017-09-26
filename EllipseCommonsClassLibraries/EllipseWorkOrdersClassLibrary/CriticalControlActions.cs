using System;
using System.Collections.Generic;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Constants;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using System.Text;
using System.Linq;

namespace EllipseWorkOrdersClassLibrary
{
    public class CriticalControlActions
    {
        public static List<CriticalControl> FetchCriticalControl(EllipseFunctions ef, string urlService, OperationContext opContext, string district, int primakeryKey, string primaryValue)
        {
            var sqlQuery = Queries.GetFetchCriticalControlsQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, 0, "", 0, "", "", "");
            var drCriticalControl = ef.GetQueryResult(sqlQuery);
            var list = new List<CriticalControl>();

            var newef = new EllipseFunctions(ef);
            var stOpContext = StdText.GetCustomOpContext(district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            newef.SetConnectionPoolingType(false);

            if (drCriticalControl == null || drCriticalControl.IsClosed || !drCriticalControl.HasRows) return list;
            while (drCriticalControl.Read())
            {
                var control = new CriticalControl
                {
                    WorkOrder = drCriticalControl["WORK_ORDER"].ToString().Trim(),
                    TaskNo = drCriticalControl["WO_TASK_NO"].ToString().Trim(),
                    TaskDescription = drCriticalControl["WO_TASK_DESC"].ToString().Trim(),
                    WorkOrderDescription = drCriticalControl["WO_DESC"].ToString().Trim(),
                    CriticalCode = drCriticalControl["JOB_DESC_CODE"].ToString().Trim(),
                    CriticalDescription = drCriticalControl["JOBD_CODE_DESC"].ToString().Trim(),
                    EquipmentNo = drCriticalControl["EQUIP_NO"].ToString().Trim(),
                    AssignPerson = drCriticalControl["ASSIGN_PERSON"].ToString().Trim(),
                    Department = drCriticalControl["DEPARTMENT"].ToString().Trim(),
                    Quartermaster = drCriticalControl["QUARTERMASTER"].ToString().Trim(),
                    PlanStartDate = drCriticalControl["PLAN_STR_DATE"].ToString().Trim(),
                    RaisedDate = drCriticalControl["RAISED_DATE"].ToString().Trim(),
                    MaintSchTask = drCriticalControl["MAINT_SCH_TASK"].ToString().Trim(),
                    StdJobNo = drCriticalControl["STD_JOB_NO"].ToString().Trim(),
                    Status = drCriticalControl["STATUS"].ToString().Trim(),
                    CompletedCode = drCriticalControl["COMPLETED_CODE"].ToString().Trim(),
                    CompletedBy = drCriticalControl["COMPLETED_BY"].ToString().Trim(),
                    CompletedDate = drCriticalControl["CLOSED_DT"].ToString().Trim(),
                    InstructionsCode = drCriticalControl["JINSTCODE"].ToString().Trim(),
                    FrequencyText = drCriticalControl["FREQUENCY"].ToString().Trim()
                };

                control.InstructionsText = StdText.GetText(urlService, stOpContext, control.InstructionsCode);
                list.Add(control);
            }

            return list;
        }
        public static CriticalControl FetchCriticalControl(EllipseFunctions ef, string urlService, OperationContext opContext, string district, string workOrder)
        {
            var sqlQuery = Queries.GetFetchCriticalControlsQuery(ef.dbReference, ef.dbLink, district, workOrder);
            var drCriticalControl = ef.GetQueryResult(sqlQuery);
            CriticalControl control = new CriticalControl();

            var newef = new EllipseFunctions(ef);
            var stOpContext = StdText.GetCustomOpContext(district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            newef.SetConnectionPoolingType(false);

            if (drCriticalControl == null || drCriticalControl.IsClosed || !drCriticalControl.HasRows) return control;
            while (drCriticalControl.Read())
            {
                control = new CriticalControl
                {
                    WorkOrder = drCriticalControl["WORK_ORDER"].ToString().Trim(),
                    TaskNo = drCriticalControl["WO_TASK_NO"].ToString().Trim(),
                    TaskDescription = drCriticalControl["WO_TASK_DESC"].ToString().Trim(),
                    WorkOrderDescription = drCriticalControl["WO_DESC"].ToString().Trim(),
                    CriticalCode = drCriticalControl["JOB_DESC_CODE"].ToString().Trim(),
                    CriticalDescription = drCriticalControl["JOBD_CODE_DESC"].ToString().Trim(),
                    EquipmentNo = drCriticalControl["EQUIP_NO"].ToString().Trim(),
                    AssignPerson = drCriticalControl["ASSIGN_PERSON"].ToString().Trim(),
                    Department = drCriticalControl["DEPARTMENT"].ToString().Trim(),
                    Quartermaster = drCriticalControl["QUARTERMASTER"].ToString().Trim(),
                    PlanStartDate = drCriticalControl["PLAN_STR_DATE"].ToString().Trim(),
                    RaisedDate = drCriticalControl["RAISED_DATE"].ToString().Trim(),
                    MaintSchTask = drCriticalControl["MAINT_SCH_TASK"].ToString().Trim(),
                    StdJobNo = drCriticalControl["STD_JOB_NO"].ToString().Trim(),
                    Status = drCriticalControl["STATUS"].ToString().Trim(),
                    CompletedCode = drCriticalControl["COMPLETED_CODE"].ToString().Trim(),
                    CompletedBy = drCriticalControl["COMPLETED_BY"].ToString().Trim(),
                    CompletedDate = drCriticalControl["CLOSED_DT"].ToString().Trim(),
                    InstructionsCode = drCriticalControl["JINSTCODE"].ToString().Trim(),
                    FrequencyText = drCriticalControl["FREQUENCY"].ToString().Trim()
                };

                control.InstructionsText = StdText.GetText(urlService, stOpContext, control.InstructionsCode);
            }
            newef.CloseConnection();
            return control;
        }

        public static string GetStringForExport(List<CriticalControl> criticalControlsLis, CriticalControlDefaultExport exportOptions)
        {

            //Creamos la instancia de StrBuilder para adicionar al RTF
            var stringRtf = new StringBuilder();
            //Inicio del rtf
            stringRtf.Append(@"{\rtf1\deff2 {\colortbl\red0\green0\blue0;\red255\green255\blue0;\red0\green77\blue187;\red0\green77\blue187;\red255\green00\blue0;\red0\green176\blue80;\red255\green192\blue0;}");

            foreach (var cc in criticalControlsLis)
            {
                stringRtf.Append(@"\line \b " + (exportOptions.StdJobNo ? cc.StdJobNo + " - " : "") + (exportOptions.WorkOrderDescription ? cc.WorkOrderDescription : "") + @"\b0");
                stringRtf.Append(@"\line Tarea " + (exportOptions.TaskNo ? cc.TaskNo + " - " : "") + @" \i " + (exportOptions.TaskDescription ? cc.TaskDescription : "") + @"\i0");
                stringRtf.Append(@"\trowd \cellx3000 \cellx10000");
                if (exportOptions.EquipmentNo)
                    stringRtf.Append(@"\intbl \i Equipo \i0 \cell " + cc.EquipmentNo + @"\cell \row");

                if (exportOptions.MaintSchTask)
                    stringRtf.Append(@"\intbl \i Mst \i0 \cell " + cc.MaintSchTask + @"\cell \row");

                if (exportOptions.WorkOrder)
                    stringRtf.Append(@"\intbl \i Orden de Trabajo \i0 \cell " + cc.WorkOrder + @"\cell \row");

                if (exportOptions.PlanStartDate)
                    stringRtf.Append(@"\intbl \i Fecha Planeada \i0 \cell " + cc.PlanStartDate + @"\cell \row");

                if (exportOptions.RaisedDate)
                    stringRtf.Append(@"\intbl \i Fecha Origen \i0 \cell " + cc.RaisedDate + @"\cell \row");

                if (exportOptions.FrequencyText)
                    stringRtf.Append(@"\intbl \i Frecuencia \i0 \cell " + cc.FrequencyText + @"\cell \row");

                if (exportOptions.AssignPerson)
                    stringRtf.Append(@"\intbl \i Responsable \i0 \cell " + cc.AssignPerson + @"\cell \row");

                if (exportOptions.CriticalDescription)
                    stringRtf.Append(@"\intbl \i Criticidad \i0 \cell " + cc.CriticalDescription + @"\cell \row");

                var statusColor = "";
                if (cc.Status.Equals("VENCIDA"))
                    statusColor = @"\cf4";
                else if (cc.Status.Equals("NO REALIZADA"))
                    statusColor = @"\cf4";
                else if (cc.Status.Equals("COMPLETADA"))
                    statusColor = @"\cf5";
                else if (cc.Status.Equals("CANCELADA"))
                    statusColor = @"\cf6";
                else if (cc.Status.Equals("OTRO"))
                    statusColor = @"\cf6";

                if (exportOptions.Status)
                    stringRtf.Append(@"\intbl \i Estado \i0 \cell " + statusColor + " " + cc.Status + @"\cf0 \cell \row");

                if (exportOptions.InstructionsText)
                    stringRtf.Append(@"\intbl \i Detalles \i0 \cell " + cc.InstructionsText + @"\cell \row");
                stringRtf.Append(@"\pard");
            }

            stringRtf.Append(@"}");

            return stringRtf.ToString();

        }
    }
    public class Queries
    {
        public static string GetFetchCriticalControlsQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, int dateCriteriaKey, string startDate, string endDate, string woStatus)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrEmpty(districtCode))
                districtCode = " IN (" + Utils.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            string queryCriteria1;
            //establecemos los parámetros del criterio 1
            if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1) && searchCriteriaValue1.Trim().Equals("PUERTO BOLIVAR"))
                queryCriteria1 = " AND WO.WORK_GROUP = 'PTOSEG'";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1) && searchCriteriaValue1.Trim().Equals("PLANTAS DE CARBON"))
                queryCriteria1 = " AND (WO.WORK_GROUP = 'PCSERVI' AND WO.EQUIP_NO IN ( 'LABMINA', 'PLANTACARBON', '2000000', '2220000','2150000', '3000000','2050605','2020000'))";
            else if (searchCriteriaKey1 == SearchFieldCriteriaType.Quartermaster.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1) && searchCriteriaValue1.Trim().Equals("FERROCARRIL"))
                queryCriteria1 = " AND (WO.WORK_GROUP IN ('SEGFFCC','VIASM','VIASP','MTOLOC','EQAUXV','CTC','ICARROS') AND WO.EQUIP_NO = 'FERROCARRIL')";
            else
                queryCriteria1 = " AND ((WO.WORK_GROUP = 'PTOSEG') OR " +
                                " (WO.WORK_GROUP = 'PCSERVI' AND WO.EQUIP_NO IN ( 'LABMINA', 'PLANTACARBON', '2000000', '2220000','2150000', '3000000','2050605','2020000')) OR" +
                                " (WO.WORK_GROUP IN ('SEGFFCC','VIASM','VIASP','MTOLOC','EQAUXV','CTC','ICARROS') AND WO.EQUIP_NO = 'FERROCARRIL'))";

            var queryCriteria2 = "";
            //establecemos los parámetros del criterio 2
            if (searchCriteriaKey2 == SearchFieldCriteriaType.WorkGroup.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP = '" + searchCriteriaValue2 + "'";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.EQUIP_NO = '" + searchCriteriaValue2 + "'";//Falta buscar el equip ref //to do
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
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Details == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            else if (searchCriteriaKey2 == SearchFieldCriteriaType.Area.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                queryCriteria2 = "AND WO.WORK_GROUP IN (" + Utils.GetListInSeparator(Groups.GetWorkGroupList().Where(g => g.Area == searchCriteriaValue2).Select(g => g.Name).ToList(), ",", "'") + ")";
            //

            //establecemos los parámetros de estado de orden
            string statusRequirement;
            if (string.IsNullOrEmpty(woStatus))
                statusRequirement = " AND TRIM(WO.COMPLETED_CODE) IS NULL";
            else if (woStatus == WoStatusList.Uncompleted)
                statusRequirement = " AND TRIM(WO.COMPLETED_CODE) IS NULL";
            else if (WoStatusList.GetStatusNames().Contains(woStatus))
                statusRequirement = " AND WO.COMPLETED_CODE = '" + WoStatusList.GetStatusCode(woStatus) + "'";
            else
                statusRequirement = " AND TRIM(WO.COMPLETED_CODE) IS NULL";

            //establecemos los parámetros para el rango de fechas
            string dateParameters;
            if (string.IsNullOrEmpty(startDate))
                startDate = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
            if (string.IsNullOrEmpty(endDate))
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
                dateParameters = "";
            //escribimos el query
            var query = "" +
                    " SELECT WO.WORK_ORDER," +
                    "   WOT.WO_TASK_NO," +
                    "   WOT.WO_TASK_DESC," +
                    "   DECODE(TRIM(WOT.JOB_DESC_CODE), NULL, MST.JOB_DESC_CODE, WOT.JOB_DESC_CODE) JOB_DESC_CODE," +
                    "   (SELECT TABLE_DESC FROM ELLIPSE.MSF010 WHERE TABLE_TYPE='JD' AND TRIM(TABLE_CODE)  = DECODE(TRIM(WOT.JOB_DESC_CODE), NULL, TRIM(MST.JOB_DESC_CODE), TRIM(WOT.JOB_DESC_CODE)) ) JOBD_CODE_DESC," +
                    "   WO.WO_DESC," +
                    "   WO.EQUIP_NO," +
                    "   WO.ASSIGN_PERSON," +
                    "   TRIM((SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE REF_NO = '030' AND ENTITY_TYPE = 'WKO' AND SEQ_NUM = '001' AND entity_type  ='WKO' AND entity_value = '1'|| 'ICOR'||WO.WORK_ORDER)) DEPARTMENT," +
                    "   CASE WHEN TRIM(WO.EQUIP_NO) = 'FERROCARRIL' THEN 'FERROCARRIL' ELSE (DECODE(TRIM(WO.WORK_GROUP), 'PTOSEG', 'PUERTO BOLIVAR', 'PCSERVI', 'PLANTAS DE CARBON', 'FERROCARRIL', 'FERROCARRIL')) END QUARTERMASTER," +
                    "   WO.RAISED_DATE," +
                    "   WO.PLAN_STR_DATE," +
                    "   WO.MAINT_SCH_TASK," +
                    "   CASE " +
                    "     WHEN TRIM(WO.COMPLETED_CODE) IN ('01', '02', '06') THEN 'COMPLETADA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) = '07' THEN 'NO REALIZADA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) = 'CN' THEN 'CANCELADA'" +
                    "     WHEN (ROW_NUMBER() OVER (PARTITION BY WO.MAINT_SCH_TASK, WO.STD_JOB_NO, WOT.WO_TASK_NO, WOT.ASSIGN_PERSON ORDER BY WO.PLAN_STR_DATE DESC)) > 1 THEN 'VENCIDA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) IS NULL THEN 'VIGENTE'" +
                    "     ELSE 'OTRO'" +
                    "   END STATUS," +
                    "   WO.STD_JOB_NO," +
                    "   WO.COMPLETED_CODE," +
                    "   WO.COMPLETED_BY," +
                    "   WO.CLOSED_DT," +
                    "   ('WI'||WO.DSTRCT_CODE||WO.WORK_ORDER||WOT.WO_TASK_NO) JINSTCODE," +
                    "   CASE" +
                    "     WHEN MST.SCHED_IND_700 = '1' THEN MST.SCHED_FREQ_1 || ' Days/Last Sched. Date'" +
                    "     WHEN MST.SCHED_IND_700 = '2' THEN MST.SCHED_FREQ_1 || ' ' || MST.STAT_TYPE_1 || '/Last Sched. Stat'" +
                    "     WHEN MST.SCHED_IND_700 = '3' THEN MST.SCHED_FREQ_1 || 'Days/Last Perf. Date'" +
                    "     WHEN MST.SCHED_IND_700 = '4' THEN MST.SCHED_FREQ_1 || ' ' || MST.STAT_TYPE_1 || '/Last Perf. Stat'" +
                    "     WHEN MST.SCHED_IND_700 = '7' THEN 'Day ' || MST.DAY_MONTH || '/' || MST.SCHED_FREQ_1 || ' Months'" +
                    "     WHEN MST.SCHED_IND_700 = '8' THEN TO_CHAR(TO_DATE(MST.OCCURENCE_TYPE, 'J'), 'fmJth') || ' ' || TO_CHAR(TO_DATE(MST.DAY_WEEK, 'J'), 'Day') || '/' || MST.SCHED_FREQ_1 || ' Months'" +
                    "     ELSE 'INACTIVE'" +
                    "   END FREQUENCY" +
                    " FROM ELLIPSE.MSF620 WO" +
                    " LEFT JOIN ELLIPSE.MSF623 WOT" +
                    " ON WO.WORK_ORDER             = WOT.WORK_ORDER" +
                    " LEFT JOIN ELLIPSE.MSF700 MST" +
                    " ON WO.STD_JOB_NO = MST.STD_JOB_NO" +
                    " WHERE" +
                    " " + "AND (WOT.JOB_DESC_CODE IN ( 'A', 'F','C','CP') OR MST.JOB_DESC_CODE IN ( 'A', 'F','C','CP'))  " +
                    " " + "AND TRIM(WO.MAINT_SCH_TASK) IS NOT NULL" +
                    " " + queryCriteria1 +
                    " " + queryCriteria2 +
                    " " + statusRequirement +
                    " AND WO.DSTRCT_CODE " + districtCode +
                    " "  + dateParameters +
                    " ORDER BY ASSIGN_PERSON, PLAN_STR_DATE ASC";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
        public static string GetFetchCriticalControlsQuery(string dbReference, string dbLink, string districtCode, string workOrder)
        {
            //establecemos los parámetrode de distrito
            if (string.IsNullOrEmpty(districtCode))
                districtCode = " IN (" + Utils.GetListInSeparator(Districts.GetDistrictList(), ",", "'") + ")";
            else
                districtCode = " = '" + districtCode + "'";

            //escribimos el query
            var query = "" +
                    " SELECT WO.WORK_ORDER," +
                    "   WOT.WO_TASK_NO," +
                    "   WOT.WO_TASK_DESC," +
                    "   DECODE(TRIM(WOT.JOB_DESC_CODE), NULL, MST.JOB_DESC_CODE, WOT.JOB_DESC_CODE) JOB_DESC_CODE," +
                    "   (SELECT TABLE_DESC FROM ELLIPSE.MSF010 WHERE TABLE_TYPE='JD' AND TRIM(TABLE_CODE)  = DECODE(TRIM(WOT.JOB_DESC_CODE), NULL, TRIM(MST.JOB_DESC_CODE), TRIM(WOT.JOB_DESC_CODE)) ) JOBD_CODE_DESC," +
                    "   WO.WO_DESC," +
                    "   WO.EQUIP_NO," +
                    "   WO.ASSIGN_PERSON," +
                    "   TRIM((SELECT REF_CODE FROM ELLIPSE.MSF071 WHERE REF_NO = '030' AND ENTITY_TYPE = 'WKO' AND SEQ_NUM = '001' AND entity_type  ='WKO' AND entity_value = '1'|| 'ICOR'||WO.WORK_ORDER)) DEPARTMENT," +
                    "   CASE WHEN TRIM(WO.EQUIP_NO) = 'FERROCARRIL' THEN 'FERROCARRIL' ELSE (DECODE(TRIM(WO.WORK_GROUP), 'PTOSEG', 'PUERTO BOLIVAR', 'PCSERVI', 'PLANTAS DE CARBON', 'FERROCARRIL', 'FERROCARRIL')) END QUARTERMASTER," +
                    "   WO.RAISED_DATE," +
                    "   WO.PLAN_STR_DATE," +
                    "   WO.MAINT_SCH_TASK," +
                    "   CASE " +
                    "     WHEN TRIM(WO.COMPLETED_CODE) IN ('01', '02', '06') THEN 'COMPLETADA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) = '07' THEN 'NO REALIZADA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) = 'CN' THEN 'CANCELADA'" +
                    "     WHEN (ROW_NUMBER() OVER (PARTITION BY WO.MAINT_SCH_TASK, WO.STD_JOB_NO, WOT.WO_TASK_NO, WOT.ASSIGN_PERSON ORDER BY WO.PLAN_STR_DATE DESC)) > 1 THEN 'VENCIDA'" +
                    "     WHEN TRIM(WO.COMPLETED_CODE) IS NULL THEN 'VIGENTE'" +
                    "     ELSE 'OTRO'" +
                    "   END STATUS," +
                    "   WO.STD_JOB_NO," +
                    "   WO.COMPLETED_CODE," +
                    "   WO.COMPLETED_BY," +
                    "   WO.CLOSED_DT," +
                    "   ('WI'||WO.DSTRCT_CODE||WO.WORK_ORDER||WOT.WO_TASK_NO) JINSTCODE," +
                    "   CASE" +
                    "     WHEN MST.SCHED_IND_700 = '1' THEN MST.SCHED_FREQ_1 || ' Days/Last Sched. Date'" +
                    "     WHEN MST.SCHED_IND_700 = '2' THEN MST.SCHED_FREQ_1 || ' ' || MST.STAT_TYPE_1 || '/Last Sched. Stat'" +
                    "     WHEN MST.SCHED_IND_700 = '3' THEN MST.SCHED_FREQ_1 || 'Days/Last Perf. Date'" +
                    "     WHEN MST.SCHED_IND_700 = '4' THEN MST.SCHED_FREQ_1 || ' ' || MST.STAT_TYPE_1 || '/Last Perf. Stat'" +
                    "     WHEN MST.SCHED_IND_700 = '7' THEN 'Day ' || MST.DAY_MONTH || '/' || MST.SCHED_FREQ_1 || ' Months'" +
                    "     WHEN MST.SCHED_IND_700 = '8' THEN TO_CHAR(TO_DATE(MST.OCCURENCE_TYPE, 'J'), 'fmJth') || ' ' || TO_CHAR(TO_DATE(MST.DAY_WEEK, 'J'), 'Day') || '/' || MST.SCHED_FREQ_1 || ' Months'" +
                    "     ELSE 'INACTIVE'" +
                    "   END FREQUENCY" +
                    " FROM ELLIPSE.MSF620 WO" +
                    " LEFT JOIN ELLIPSE.MSF623 WOT" +
                    " ON WO.WORK_ORDER             = WOT.WORK_ORDER" +
                    " LEFT JOIN ELLIPSE.MSF700 MST" +
                    " ON WO.STD_JOB_NO = MST.STD_JOB_NO" +
                    " WHERE" +
                    " " + "WO.WORK_ORDER = '" + workOrder + "'" +
                    " AND WO.DSTRCT_CODE " + districtCode +
                    " ORDER BY ASSIGN_PERSON, PLAN_STR_DATE ASC";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }

    
}

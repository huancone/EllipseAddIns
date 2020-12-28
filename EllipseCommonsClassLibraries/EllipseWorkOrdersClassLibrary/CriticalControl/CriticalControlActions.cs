using System.Collections.Generic;
using SharedClassLibrary.Ellipse;
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
            var sqlQuery = Queries.GetFetchCriticalControlsQuery(ef.DbReference, ef.DbLink, district, primakeryKey, primaryValue, 0, "", 0, "", "", "");
            var drCriticalControl = ef.GetQueryResult(sqlQuery);
            var list = new List<CriticalControl>();

            var newef = new EllipseFunctions(ef);
            var stOpContext = StdText.GetCustomOpContext(district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            newef.SetConnectionPoolingType(false);

            if (drCriticalControl == null || drCriticalControl.IsClosed) return list;
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
        public static CriticalControl FetchCriticalControl(EllipseFunctions ef, string urlService, OperationContext opContext, string district, string workOrder, string woTask)
        {
            var sqlQuery = Queries.GetFetchCriticalControlsQuery(ef.DbReference, ef.DbLink, district, workOrder, woTask);
            var drCriticalControl = ef.GetQueryResult(sqlQuery);
            CriticalControl control = new CriticalControl();

            var newef = new EllipseFunctions(ef);
            var stOpContext = StdText.GetCustomOpContext(district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            newef.SetConnectionPoolingType(false);

            if (drCriticalControl == null || drCriticalControl.IsClosed) return control;
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
    
    
}

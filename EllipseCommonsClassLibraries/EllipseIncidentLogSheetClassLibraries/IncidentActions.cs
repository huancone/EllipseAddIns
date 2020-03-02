using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary.Classes;

namespace EllipseIncidentLogSheetClassLibraries
{
    //TO DO
    //20200201 Para todos los procesos (creación y eliminación) no se han revisado los procesos para funcionamientos múltiples
    //El comportamiento del módulo es compatible con la agregación de a un elemento. Para la eliminación si hay múltiples coincidencias, solo eliminará el primer registro coincidente
    public static class IncidentActions
    {
        public static Screen.ScreenDTO CreateIncident(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, string workGroup, string date, string shift, IncidentItem item)
        {
            var itemList = new List<IncidentItem>();
            itemList.Add(item);
            return CreateIncident(eFunctions, opContext, urlService, workGroup, date, shift, itemList);
        }
        public static Screen.ScreenDTO CreateIncident(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, string workGroup, string date, string shift, List<IncidentItem> itemList)
        {
            var service = new Screen.ScreenService();
            var request = new Screen.ScreenSubmitRequestDTO();

            service.Url = urlService + "/ScreenService";

            eFunctions.RevertOperation(opContext, service);
            var reply = service.executeScreen(opContext, "MSO627");

            if (eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);
            if (reply != null && reply.mapName != "MSM627A")
                throw new Exception("No se ha podido ingresar al programa MSO627");

            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("WORK_GROUP1I", workGroup);
            arrayFields.Add("RAISED_DATE1I", date);
            arrayFields.Add("SHIFT1I", shift);
            request.screenFields = arrayFields.ToArray();

            request.screenKey = "1";
            reply = service.submit(opContext, request);

            if (eFunctions.CheckReplyWarning(reply))
                reply = service.submit(opContext, request);

            if (eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            if (reply != null && reply.mapName != "MSM627B")
                throw new Exception("Se ha producido un error al ingresar al MSM627B");

            
            arrayFields = new ArrayScreenNameValue();
            //arrayFields.Add("WO_PREFIX2I", item.WorkOrderPrefix);

            var i = 1;

            foreach (var item in itemList)
            {
                arrayFields.Add("RAISED_TIME2I" + i, item.RaisedTime);
                arrayFields.Add("INCIDENT_DESC2I" + i, item.IncidentDescription);
                arrayFields.Add("MAINT_TYPE2I" + i, item.MaintenanceType);
                arrayFields.Add("ORIGINATOR_ID2I" + i, item.Originator);
                arrayFields.Add("JOB_DUR_FINISH2I" + i, item.JobDurationFinish);
                arrayFields.Add("INCID_STATUS2I" + i, item.IncidentStatus);
                arrayFields.Add("EQUIP_REF2I" + i, item.EquipmentReference);
                arrayFields.Add("COMP_CODE2I" + i, item.ComponentCode);
                arrayFields.Add("COMP_MOD_CODE2I" + i, item.ModifierCode);
                arrayFields.Add("JOB_DUR_CODE2I" + i, item.JobDurationCode);
                arrayFields.Add("JOB_DUR_HOURS2I" + i, item.DurationHours);
                arrayFields.Add("STD_JOB_NO2I" + i, item.StandardJob);
                arrayFields.Add("CORRECT_DESC2I" + i, item.CorrectiveDescription);
                arrayFields.Add("WORK_ORDER2I" + i, item.WorkOrder);
                request.screenFields = arrayFields.ToArray();

                request.screenKey = "1";
                reply = service.submit(opContext, request);

                while (eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm")
                    reply = service.submit(opContext, request);

                if (eFunctions.CheckReplyError(reply))
                    throw new Exception(reply.message);

                i++;
                if (i > 4 && reply != null && reply.mapName.Equals("MSM627B"))
                    i = 1;
            }

            if (reply != null && reply.mapName != "MSM627A")
                throw new Exception("Ha ocurrido un error y no se ha podido finalizar el proceso");
            return reply;
        }
        public static Screen.ScreenDTO DeleteIncident(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, string workGroup, string date, string shift, string originator, string equipmentReference, string incidentStatus, IncidentItem item)
        {
            var service = new Screen.ScreenService();
            var request = new Screen.ScreenSubmitRequestDTO();

            service.Url = urlService + "/ScreenService";

            eFunctions.RevertOperation(opContext, service);
            var reply = service.executeScreen(opContext, "MSO627");

            if (eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);
            if (reply != null && reply.mapName != "MSM627A")
                throw new Exception("No se ha podido ingresar al programa MSO627");

            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "3");
            arrayFields.Add("WORK_GROUP1I", workGroup);
            arrayFields.Add("RAISED_DATE1I", date);
            arrayFields.Add("SHIFT1I", shift);
            arrayFields.Add("ORIGINATOR_ID1I", originator);
            arrayFields.Add("EQUIP_REF1I", equipmentReference);
            arrayFields.Add("INCID_STATUS1I", incidentStatus);//C - Closed, O - Open
            request.screenFields = arrayFields.ToArray();

            request.screenKey = "1";
            reply = service.submit(opContext, request);

            if (eFunctions.CheckReplyWarning(reply))
                reply = service.submit(opContext, request);

            if (eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            if (reply != null && reply.mapName != "MSM627B")
                throw new Exception("Se ha producido un error al ingresar al MSM627B");

            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            var i = 1;
            var currentReplyItem = GetIncidentItemFromDtoFields(reply.screenFields, i);

            while (reply != null && reply.mapName == "MSM627B" && !currentReplyItem.Equals(item))
            {
                i++;
                currentReplyItem = GetIncidentItemFromDtoFields(reply.screenFields, i);
                if (i <= 4) continue;
                i = 1;
                //envíe a la siguiente pantalla
                request = new Screen.ScreenSubmitRequestDTO();
                request.screenKey = "1";
                reply = service.submit(opContext, request);
                replyFields = new ArrayScreenNameValue(reply.screenFields);
                currentReplyItem = GetIncidentItemFromDtoFields(reply.screenFields, i);
            }

            if (reply != null && reply.mapName != "MSM627B")
                throw new ArgumentException("No se ha encontrado el registro");

            arrayFields = new ArrayScreenNameValue();
            //arrayFields.Add("WO_PREFIX2I", item.WorkOrderPrefix);

            arrayFields.Add("ACTION2I" + i, "D");
            request = new Screen.ScreenSubmitRequestDTO();
            request.screenFields = arrayFields.ToArray();
            request.screenKey = "1";
            reply = service.submit(opContext, request);
            while (eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm")
            {
                request.screenKey = "1";
                reply = service.submit(opContext, request);
            }

            if (eFunctions.CheckReplyError(reply))
                throw new Exception(reply.message);

            //Si hay pantallas con items todavía después de eliminar el registro
            replyFields = new ArrayScreenNameValue(reply.screenFields);
            while (reply != null && reply.mapName == "MSM627B" && !string.IsNullOrWhiteSpace(replyFields.GetField("RAISED_TIME2I1").value))
            {
                request = new Screen.ScreenSubmitRequestDTO();
                request.screenKey = "1";
                reply = service.submit(opContext, request);
                replyFields = new ArrayScreenNameValue(reply.screenFields);

                if (eFunctions.CheckReplyError(reply))
                    throw new Exception(reply.message);
            }
            //

            if (reply != null && reply.mapName != "MSM627A")
                throw new Exception("Ha ocurrido un error y no se ha podido finalizar el proceso");
            return reply;
        }

        public static IncidentItem GetIncidentItemFromDtoFields(Screen.ScreenFieldDTO[] screenFields, int index)
        {
            var replyFields = new ArrayScreenNameValue(screenFields);
            var item = new IncidentItem();
            var i = index;

            item.RaisedTime = replyFields.GetField("RAISED_TIME2I" + i) != null ? replyFields.GetField("RAISED_TIME2I" + i).value : null;
            item.IncidentDescription = replyFields.GetField("INCIDENT_DESC2I" + i) != null ? replyFields.GetField("INCIDENT_DESC2I" + i).value : null;
            item.MaintenanceType = replyFields.GetField("MAINT_TYPE2I" + i) != null ? replyFields.GetField("MAINT_TYPE2I" + i).value : null;
            item.Originator = replyFields.GetField("ORIGINATOR_ID2I" + i) != null ? replyFields.GetField("ORIGINATOR_ID2I" + i).value : null;
            item.JobDurationFinish = replyFields.GetField("JOB_DUR_FINISH2I" + i) != null ? replyFields.GetField("JOB_DUR_FINISH2I" + i).value : null;
            item.IncidentStatus = replyFields.GetField("INCID_STATUS2I" + i) != null ? replyFields.GetField("INCID_STATUS2I" + i).value : null;
            item.EquipmentReference = replyFields.GetField("EQUIP_REF2I" + i) != null ? replyFields.GetField("EQUIP_REF2I" + i).value : null;
            item.ComponentCode = replyFields.GetField("COMP_CODE2I" + i) != null ? replyFields.GetField("COMP_CODE2I" + i).value : null;
            item.ModifierCode = replyFields.GetField("COMP_MOD_CODE2I" + i) != null ? replyFields.GetField("COMP_MOD_CODE2I" + i).value : null;
            item.JobDurationCode = replyFields.GetField("JOB_DUR_CODE2I" + i) != null ? replyFields.GetField("JOB_DUR_CODE2I" + i).value : null;
            item.DurationHours = replyFields.GetField("JOB_DUR_HOURS2I" + i) != null ? replyFields.GetField("JOB_DUR_HOURS2I" + i).value : null;
            item.StandardJob = replyFields.GetField("STD_JOB_NO2I" + i) != null ? replyFields.GetField("STD_JOB_NO2I" + i).value : null;
            item.CorrectiveDescription = replyFields.GetField("CORRECT_DESC2I" + i) != null ? replyFields.GetField("CORRECT_DESC2I" + i).value : null;
            item.WorkOrder = replyFields.GetField("WORK_ORDER2I" + i) != null ? replyFields.GetField("WORK_ORDER2I" + i).value : null;

            if (item.RaisedTime != null)
                item.RaisedTime = item.RaisedTime.Replace(":", "");

            return item;
        }
    }
}

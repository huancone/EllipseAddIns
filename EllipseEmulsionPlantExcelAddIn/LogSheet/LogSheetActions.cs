using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary.Classes;

namespace EllipseEmulsionPlantExcelAddIn.LogSheet
{
    public static class LogSheetActions
    {
        public static string CreateLogSheet(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, string modelCode, string modelDate, string modelShift, List<LogSheetEquipmentInputItem> inputItems)
        {
            var logSheet = new LogSheetItem(modelCode, modelDate, modelShift, inputItems);

            return CreateLogSheet(eFunctions, opContext, urlService, logSheet);
        }
        public static string CreateLogSheet(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, LogSheetItem logSheet)
        {
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            //Proceso del screen
            var screenService = new Screen.ScreenService();
            screenService.Url = urlService;
            //Aseguro que no esté en alguna pantalla antigua
            eFunctions.RevertOperation(opContext, screenService);
            //ejecutamos el programa
            var replySheet = screenService.executeScreen(opContext, "MSO435");

            //validamos el ingreso al programa
            if (replySheet.mapName != "MSM435A")
                throw new Exception("No se pudo establecer comunicación con el servicio");

            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("MODEL_CODE1I", logSheet.ModelName);
            arrayFields.Add("STAT_DATE1I", logSheet.Date);
            arrayFields.Add("SHIFT1I", logSheet.ShiftCode);
            //arrayFields.Add("MODEL_MODE1I",""); //no usado
            //arrayFields.Add("RUN_ID1I", ""); //no usado

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";

            replySheet = screenService.submit(opContext, requestSheet);

            eFunctions.CheckReplyWarning(replySheet);//si hay debug activo muestra el warning de lo contrario depende del proceso del OP


            if (replySheet == null)
                throw new Exception("No se puede establecer conexión con el programa MSM435B");
            if (eFunctions.CheckReplyError(replySheet) || replySheet.message.StartsWith("X2"))
                throw new Exception("Se ha producido un error. " + replySheet.message);
            if (replySheet.mapName != "MSM435B")
                throw new Exception("No se ha podido acceder al programa MSM435B");

            //Creamos la nueva pantalla de envío reutilizando las declaraciones anteriores
            requestSheet = new Screen.ScreenSubmitRequestDTO();
            arrayFields = new ArrayScreenNameValue();

            //ingresamos los elementos (name, value) para los campos a enviar   
            arrayFields.Add("STAT_DATE2I", logSheet.Date);
            arrayFields.Add("SHIFT2I", logSheet.ShiftCode);

            var screenIndex = 1;
            foreach (var item in logSheet.InputItems)
            {
                if (screenIndex > 7)
                {
                    //enviar Screen
                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = screenService.submit(opContext, requestSheet);
                    arrayFields = new ArrayScreenNameValue();
                    //
                    if (replySheet != null && replySheet.mapName != "MSM435B")
                        break;
                    screenIndex = 1;
                }

                //eS(screenIndex) = fv
                arrayFields.Add("ACTION2I" + screenIndex, item.Action);
                arrayFields.Add(item.Operator.Id + screenIndex, item.Operator.Value);
                arrayFields.Add(item.PlantNo.Id + screenIndex, item.PlantNo.Value);
                arrayFields.Add(item.AccountCode.Id + screenIndex, item.AccountCode.Value);
                arrayFields.Add(item.WorkOrder.Id + screenIndex, item.WorkOrder.Value);
                arrayFields.Add(item.Input1.Id + screenIndex, item.Input1.Value);
                arrayFields.Add(item.Input2.Id + screenIndex, item.Input2.Value);
                arrayFields.Add(item.Input3.Id + screenIndex, item.Input3.Value);
                arrayFields.Add(item.Input4.Id + screenIndex, item.Input4.Value);
                arrayFields.Add(item.Input5.Id + screenIndex, item.Input5.Value);
                arrayFields.Add(item.Input6.Id + screenIndex, item.Input6.Value);
                arrayFields.Add(item.Input7.Id + screenIndex, item.Input7.Value);
                arrayFields.Add(item.Input8.Id + screenIndex, item.Input8.Value);
                arrayFields.Add(item.Input9.Id + screenIndex, item.Input9.Value);
                //arrayFields.Add(item.Input10.Id + screenIndex, item.Input10.Value);

                if (item.Action == "I")
                {
                    //enviar Screen
                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = screenService.submit(opContext, requestSheet);
                    var field = arrayFields.GetField("ACTION2I" + screenIndex);//Se reinicia el valor para que al enviar no vuelva a hacer insert, sino simplemente continúe con el screen
                    field.value = "";
                    //
                    if (replySheet != null && replySheet.mapName != "MSM435B")
                        break;

                    if (screenIndex >= 7) //es una condición especial cuando se añade estando en el último registro, porque el sistema envía y cambia el screen de una vez
                    {
                        screenIndex = 0; //se iguala a cero porque al terminar el bucle exterior sube el index a 1, que es lo que se necesitaría para la siguiente iteración
                        arrayFields = new ArrayScreenNameValue();
                    }
                }
                screenIndex++;
            }
            requestSheet = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };

            replySheet = screenService.submit(opContext, requestSheet);
            //si hay debug activo muestra el warning de lo contrario depende del proceso del OP
            eFunctions.CheckReplyWarning(replySheet);

            if (replySheet == null || replySheet.mapName != "MSM435A")
                throw new Exception("Se ha producido un error al enviar la solicitud de creación");

            if (eFunctions.CheckReplyError(replySheet) || replySheet.message.StartsWith("X2"))
                throw new Exception("Se ha producido un error. " + replySheet.message);

            if (replySheet != null && !eFunctions.CheckReplyError(replySheet) && replySheet.mapName == "MSM435A")
                return "SUCCESS:" + "Se han cargado exitosamente los datos";

            return "WARNING: No se ha recibido una respuesta del servicio. Por favor valide que los datos fueron cargados";

            //---fin proceso del screen
        }
        public static string DeleteLogSheet(EllipseFunctions eFunctions, Screen.OperationContext opContext, string urlService, LogSheetItem logSheet)
        {
            var requestSheet = new Screen.ScreenSubmitRequestDTO();

            //Proceso del screen
            var screenService = new Screen.ScreenService();
            screenService.Url = urlService;
            //Aseguro que no esté en alguna pantalla antigua
            eFunctions.RevertOperation(opContext, screenService);
            //ejecutamos el programa
            var replySheet = screenService.executeScreen(opContext, "MSO435");

            //validamos el ingreso al programa
            if (replySheet.mapName != "MSM435A")
                throw new Exception("No se pudo establecer comunicación con el servicio");

            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "3");
            arrayFields.Add("MODEL_CODE1I", logSheet.ModelName);
            arrayFields.Add("STAT_DATE1I", logSheet.Date);
            arrayFields.Add("SHIFT1I", logSheet.ShiftCode);
            //arrayFields.Add("MODEL_MODE1I",""); //no usado
            //arrayFields.Add("RUN_ID1I", ""); //no usado

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";

            replySheet = screenService.submit(opContext, requestSheet);

            eFunctions.CheckReplyWarning(replySheet);//si hay debug activo muestra el warning de lo contrario depende del proceso del OP


            if (replySheet == null)
                throw new Exception("No se puede establecer conexión con el programa MSM435B");
            if (eFunctions.CheckReplyError(replySheet) || replySheet.message.StartsWith("X2"))
                throw new Exception("Se ha producido un error. " + replySheet.message);
            if (replySheet.mapName != "MSM435B")
                throw new Exception("No se ha podido acceder al programa MSM435B");

            //Creamos la nueva pantalla de envío reutilizando las declaraciones anteriores
            requestSheet = new Screen.ScreenSubmitRequestDTO();
            arrayFields = new ArrayScreenNameValue();

            //ingresamos los elementos (name, value) para los campos a enviar   
            arrayFields.Add("STAT_DATE2I", logSheet.Date);
            arrayFields.Add("SHIFT2I", logSheet.ShiftCode);
            arrayFields.Add("DELETE2I", "Y");

            requestSheet = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };

            replySheet = screenService.submit(opContext, requestSheet);
            //si hay debug activo muestra el warning de lo contrario depende del proceso del OP
            eFunctions.CheckReplyWarning(replySheet);

            if (replySheet == null || replySheet.mapName != "MSM435A")
                throw new Exception("Se ha producido un error al enviar la solicitud de eliminación");

            if(replySheet != null && replySheet.message.Contains("LOGSHEET HAS BEEN FLAGGED FOR DELETION"))
                return "SUCCESS:" + replySheet.message;

            if (eFunctions.CheckReplyError(replySheet) || replySheet.message.StartsWith("X2"))
                throw new Exception(replySheet.message);
            
            if (replySheet != null && !eFunctions.CheckReplyError(replySheet) && replySheet.mapName == "MSM435A")
                return "SUCCESS:" + "Se han cargado exitosamente los datos";

            return "WARNING: No se ha recibido una respuesta del servicio. Por favor valide que los datos fueron cargados";

            //---fin proceso del screen
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseLabourCostingExcelAddIn.LabourCostingTransService;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseLabourCostingExcelAddIn
{
    public static class LabourEmployeeActions
    {

        public static LabourCostingTransServiceResult CreateEmployeeMse(string urlService, OperationContext opContext, LabourEmployee labourEmployee, bool replaceExisting = true)
        {
            var proxyLt = new LabourCostingTransService.LabourCostingTransService { Url = urlService + "/LabourCostingTrans" };

            var requestLt = new LabourCostingTransDTO
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                labourCostingHours = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingHours) ? Convert.ToDecimal(labourEmployee.LabourCostingHours) : default(decimal),
                labourCostingHoursSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingHours),
                labourCostingValue = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingValue) ? Convert.ToDecimal(labourEmployee.LabourCostingValue) : default(decimal),
                labourCostingValueSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingValue),
                interDistrictCode = labourEmployee.InterDistrictCode,
                accountCode = labourEmployee.AccountCode,
                postingStatus = labourEmployee.PostingStatus,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask,
                employee = labourEmployee.Employee,
                equipmentNo = labourEmployee.EquipmentNo,
                equipmentReference = labourEmployee.EquipmentRef,
                percentComplete = !string.IsNullOrWhiteSpace(labourEmployee.PercentComplete) ? Convert.ToDecimal(labourEmployee.PercentComplete) : default(decimal),
                percentCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.PercentComplete),
                earnCode = labourEmployee.EarnCode,
                labourClass = labourEmployee.LabourClass,
                overtimeInd = labourEmployee.OvertimeInd,
                overtimeIndSpecified = true,
                unitsComplete = !string.IsNullOrWhiteSpace(labourEmployee.UnitsComplete) ? Convert.ToDecimal(labourEmployee.UnitsComplete) : default(decimal),
                unitsCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.UnitsComplete),
                completedCode = labourEmployee.CompletedCode
            };

            LabourCostingTransServiceResult createResult;
            //Search Existing
            if (replaceExisting)
            {
                var requestSearch = new LabourCostingTransSearchParam
                {
                    transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)),
                        int.Parse(labourEmployee.TransactionDate.Substring(4, 2)),
                        int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                    transactionDateSpecified = true,
                    employee = labourEmployee.Employee,
                    project = labourEmployee.Project,
                    workOrder = labourEmployee.WorkOrder,
                    workOrderTask = labourEmployee.WorkOrderTask
                };
                var searchRestartDto = new LabourCostingTransDTO();
                //La búsqueda solo toma en cuenta la fecha de transacción y el employee id
                var replySearch = proxyLt.search(opContext, requestSearch, searchRestartDto);
                //Existe un elemento
                if (replySearch != null && replySearch.Length >= 1)
                {
                    foreach (var replyItem in replySearch)
                    {
                        //Las comparaciones deben hacerse con LPAD para poder establecer bien las comparaciones númericas que trae ellipse en su información
                        var equalTranDate =
                            replyItem.labourCostingTransDTO.transactionDate.Equals(requestSearch.transactionDate);
                        var equalEmployee = replyItem.labourCostingTransDTO.employee.PadLeft(20, '0')
                            .Equals(requestSearch.employee.PadLeft(20, '0'));
                        //posibles nulos de WorkOrders y/o Projects
                        var equalWo = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrder) && replyItem
                                          .labourCostingTransDTO.workOrder.PadLeft(20, '0')
                                          .Equals(requestSearch.workOrder.PadLeft(20, '0'));

                        string itemTaskNo = string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrderTask) ? "000" : replyItem.labourCostingTransDTO.workOrderTask;
                        var equalTask = itemTaskNo.Equals(!string.IsNullOrWhiteSpace(requestSearch.workOrderTask) ? requestSearch.workOrderTask.PadLeft(3, '0') : "000");
                        var equalProject = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.project) &&
                                           replyItem.labourCostingTransDTO.project.PadLeft(20, '0')
                                               .Equals(requestSearch.project.PadLeft(20, '0'));

                        if (!equalTranDate || !equalEmployee || ((!equalWo || !equalTask) && !equalProject)) continue;
                        var delResult = proxyLt.delete(opContext, replySearch[0].labourCostingTransDTO);
                        if (delResult.errors != null && delResult.errors.Length > 0)
                        {
                            var errorMessage = "";
                            foreach (var error in delResult.errors)
                                errorMessage += error.messageText + ". ";
                            throw new Exception(errorMessage);
                        }
                        break;
                    }
                }

            }
            //se envía la acción
            //return proxyLt.multipleCreate(opContext, multipleRequestLt);
            createResult = proxyLt.create(opContext, requestLt);

            if (createResult.errors != null && createResult.errors.Length > 0)
            {
                var errorMessage = "";
                foreach (var error in createResult.errors)
                    errorMessage += error.messageText + ". ";
                throw new Exception(errorMessage);
            }
            return createResult;
        }

        public static LabourCostingTransServiceResult DeleteEmployeeMse(string urlService, OperationContext opContext, LabourEmployee labourEmployee)
        {
            var proxyLt = new LabourCostingTransService.LabourCostingTransService { Url = urlService + "/LabourCostingTrans" };

            var requestLt = new LabourCostingTransDTO
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                labourCostingHours = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingHours) ? Convert.ToDecimal(labourEmployee.LabourCostingHours) : default(decimal),
                labourCostingHoursSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingHours),
                labourCostingValue = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingValue) ? Convert.ToDecimal(labourEmployee.LabourCostingValue) : default(decimal),
                labourCostingValueSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingValue),
                interDistrictCode = labourEmployee.InterDistrictCode,
                accountCode = labourEmployee.AccountCode,
                postingStatus = labourEmployee.PostingStatus,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask,
                employee = labourEmployee.Employee,
                equipmentNo = labourEmployee.EquipmentNo,
                equipmentReference = labourEmployee.EquipmentRef,
                percentComplete = !string.IsNullOrWhiteSpace(labourEmployee.PercentComplete) ? Convert.ToDecimal(labourEmployee.PercentComplete) : default(decimal),
                percentCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.PercentComplete),
                earnCode = labourEmployee.EarnCode,
                labourClass = labourEmployee.LabourClass,
                overtimeInd = labourEmployee.OvertimeInd,
                overtimeIndSpecified = true,
                unitsComplete = !string.IsNullOrWhiteSpace(labourEmployee.UnitsComplete) ? Convert.ToDecimal(labourEmployee.UnitsComplete) : default(decimal),
                unitsCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.UnitsComplete),
                completedCode = labourEmployee.CompletedCode
            };

            //Search Existing
            var requestSearch = new LabourCostingTransSearchParam
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)),
                    int.Parse(labourEmployee.TransactionDate.Substring(4, 2)),
                    int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                employee = labourEmployee.Employee,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask
            };
            var searchRestartDto = new LabourCostingTransDTO();
            //La búsqueda solo toma en cuenta la fecha de transacción y el employee id
            var replySearch = proxyLt.search(opContext, requestSearch, searchRestartDto);
            //Existe un elemento
            if (replySearch == null || replySearch.Length < 1) return proxyLt.create(opContext, requestLt);
            foreach (var replyItem in replySearch)
            {
                //Las comparaciones deben hacerse con LPAD para poder establecer bien las comparaciones númericas que trae ellipse en su información
                var equalTranDate =
                    replyItem.labourCostingTransDTO.transactionDate.Equals(requestSearch.transactionDate);
                var equalEmployee = replyItem.labourCostingTransDTO.employee.PadLeft(20, '0')
                    .Equals(requestSearch.employee.PadLeft(20, '0'));
                //posibles nulos de WorkOrders y/o Projects
                var equalWo = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrder) && replyItem
                                    .labourCostingTransDTO.workOrder.PadLeft(20, '0')
                                    .Equals(requestSearch.workOrder.PadLeft(20, '0'));

                string itemTaskNo = string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrderTask) ? "000" : replyItem.labourCostingTransDTO.workOrderTask;
                var equalTask = itemTaskNo.Equals(!string.IsNullOrWhiteSpace(requestSearch.workOrderTask) ? requestSearch.workOrderTask.PadLeft(3, '0') : "000");
                var equalProject = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.project) &&
                                    replyItem.labourCostingTransDTO.project.PadLeft(20, '0')
                                        .Equals(requestSearch.project.PadLeft(20, '0'));

                if (!equalTranDate || !equalEmployee || ((!equalWo || !equalTask) && !equalProject)) continue;
                var delResult = proxyLt.delete(opContext, replySearch[0].labourCostingTransDTO);
                if (delResult.errors != null && delResult.errors.Length > 0)
                {
                    var errorMessage = "";
                    foreach (var error in delResult.errors)
                        errorMessage += error.messageText + ". ";
                    throw new Exception(errorMessage);
                }
                return delResult;
            }
            return null;
        }

        /// <summary>
        /// Carga un registro de labor para el empleado asignado
        /// </summary>
        /// <param name="opSheet"></param>
        /// <param name="labourEmployee"></param>
        /// <param name="replaceExisting">bool: Si es true y ya existe un registro con la misma ot-tarea se modificará por el nuevo registro. Si es false, se ignorará el registro y siempre se adicionará uno nuevo</param>
        [Obsolete("Not used anymore", true)]//Deprecated
        public static void CreateEmployeeMso(EllipseFunctions eFunctions, string urlService, Screen.OperationContext opSheet, LabourEmployee labourEmployee, bool autoTaskAssigment, bool replaceExisting = true)
        {
            //Proceso del screen
            var proxySheet = new Screen.ScreenService();

            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var arrayFields = new ArrayScreenNameValue();

            //Selección de ambiente
            proxySheet.Url = urlService + "/ScreenService";
            //Aseguro que no esté en alguna pantalla antigua
            eFunctions.RevertOperation(opSheet, proxySheet);
            //ejecutamos el programa
            var replySheet = proxySheet.executeScreen(opSheet, "MSO850");

            //validamos el ingreso al programa
            if (replySheet == null || replySheet.mapName != "MSM850A" || ValidateError(replySheet) || eFunctions.CheckReplyWarning(replySheet))
                throw new Exception("No se pudo establecer comunicación con el servicio");
            //Enviamos datos principales para activar los campos de labor
            arrayFields.Add("EMP_ID1I", labourEmployee.Employee);
            arrayFields.Add("TRAN_DATE1I", labourEmployee.TransactionDate);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";

            replySheet = proxySheet.submit(opSheet, requestSheet);

            //Continuamos en la pantalla pero con los campos de labor activos
            if (!ValidateError(replySheet) && replySheet.mapName == "MSM850A")
            {
                var labourFoundFlag = false;
                var rowMso = 1;
                //obtenemos los campos del reply
                var replyFields = new ArrayScreenNameValue(replySheet.screenFields);
                //son variables para determinar el cambio de screen real
                //reajustamos el valor de la tarea a un numérico ###
                if (string.IsNullOrWhiteSpace(labourEmployee.WorkOrderTask) && autoTaskAssigment)
                    if (!string.IsNullOrWhiteSpace(labourEmployee.WorkOrder))
                        labourEmployee.WorkOrderTask = labourEmployee.WorkOrderTask == "" ? "" : "001";

                //iniciamos el recorrido
                while (!labourFoundFlag)
                {
                    //comprobamos que: 1. Que no exista OT-Tarea cargada para que no se duplique, y 2. nos ubicamos en el último campo disponible
                    //si existe 1, entonces actualizamos la información. Si no, continuamos a ubicarnos en 2.
                    var isEmpty = replyFields.GetField("WO_PROJ1I" + rowMso).value.Equals("");
                    var sameWo = labourEmployee.WorkOrder.Equals(replyFields.GetField("WO_PROJ1I" + rowMso).value);
                    var screenTaskValue = replyFields.GetField("TASK1I" + rowMso).value;
                    var isWopInd = replyFields.GetField("WOP_IND1I" + rowMso).value.Equals("W");

                    if (string.IsNullOrWhiteSpace(screenTaskValue) && isWopInd && autoTaskAssigment)
                    {
                        screenTaskValue = "001";
                    }

                    var sameTask = labourEmployee.WorkOrderTask == screenTaskValue;
                    var sameEarnClass = string.IsNullOrWhiteSpace(labourEmployee.EarnCode) || labourEmployee.EarnCode.Equals(replyFields.GetField("EARN_CLASS1I" + rowMso).value);
                    //si se encuentra una posición para escribir la labor se activa el flag y se sale del while
                    if (isEmpty || (replaceExisting && sameWo && sameTask && sameEarnClass))
                    {
                        labourFoundFlag = true;
                        continue;
                    }
                    rowMso++;
                    //si es mayor que los registros del screen (5) entonces envíe el screen y prosiga con los del siguiente screen
                    if (rowMso <= 5) continue;

                    var previousReply = replyFields;

                    var errorLoopPos = 0;//para controlar que no se quede atrapado en loop infinito
                    //verifico la respuesta de este envío por errores y advertencias
                    while (replySheet.mapName == "MSM850A" && replyFields.GetField("TRAN_DATE1I").value == labourEmployee.TransactionDate && !ValidateError(replySheet))
                    {
                        //Creamos la nueva acción de envío reutilizando los elementos anteriores
                        requestSheet = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                        replyFields = new ArrayScreenNameValue(replySheet.screenFields);

                        //para asegurar que hubo un cambio real de pantalla. De lo contrario pasará hasta el final de los envíos
                        var currentReply = replyFields;
                        if (previousReply == currentReply)
                        {
                            errorLoopPos++;
                            if (errorLoopPos > 50)
                                throw new Exception("Error de ejecución. El proceso ha alcanzo el límite de intentos en posicionamiento");
                            continue;
                        }
                        break;
                    }
                    if (ValidateError(replySheet) || eFunctions.CheckReplyWarning(replySheet))
                        throw new Exception(replySheet.message);

                    rowMso = 1;
                }
                //ingresamos los elementos para los campos a enviar   
                arrayFields.Add("ORD_HRS1I" + rowMso, labourEmployee.LabourCostingHours);
                arrayFields.Add("OT_HRS1I" + rowMso, labourEmployee.OvertimeInd ? labourEmployee.LabourCostingHours : null);
                arrayFields.Add("ACCOUNT1I" + rowMso, labourEmployee.AccountCode);
                arrayFields.Add("WOP_IND1I" + rowMso, string.IsNullOrWhiteSpace(labourEmployee.WorkOrder) ? "P" : "W");
                arrayFields.Add("WO_PROJ1I" + rowMso, string.IsNullOrWhiteSpace(labourEmployee.WorkOrder) ? labourEmployee.Project : labourEmployee.WorkOrder);
                arrayFields.Add("TASK1I" + rowMso, labourEmployee.WorkOrderTask);
                arrayFields.Add("EQUIPMENT1I" + rowMso, labourEmployee.EquipmentNo);
                arrayFields.Add("UNITS_COMP1I" + rowMso, labourEmployee.UnitsComplete);
                arrayFields.Add("PC_COMP1I" + rowMso, labourEmployee.PercentComplete);
                arrayFields.Add("CODE_COMP1I" + rowMso, labourEmployee.CompletedCode);
                arrayFields.Add("EARN_CLASS1I" + rowMso, labourEmployee.EarnCode);
                arrayFields.Add("LABOUR_CLASS1I" + rowMso, labourEmployee.LabourClass);
                //enviamos la información
                requestSheet = new Screen.ScreenSubmitRequestDTO
                {
                    screenFields = arrayFields.ToArray(),
                    screenKey = "1"
                };
                replySheet = proxySheet.submit(opSheet, requestSheet);

                var errorLoopVal = 0;

                while (replySheet != null && !ValidateError(replySheet) &&//no existen errores
                       (eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys == "XMIT-Confirm" || replySheet.functionKeys == "XMIT-Validate"))//requiere confirmación
                {
                    errorLoopVal++;
                    if (errorLoopVal > 50)
                        throw new Exception("Error de ejecución. El proceso ha alcanzo el límite de intentos en confirmación");
                    requestSheet = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                    replySheet = proxySheet.submit(opSheet, requestSheet);
                }

                if (replySheet != null && replySheet.message.Length > 2 && replySheet.message.Substring(0, 2) == "X2")
                    throw new Exception(replySheet.message);

                return;
            }

            if (replySheet == null)
                throw new Exception("No se puede establecer conexión con el programa Mso850");

            throw new Exception(replySheet.message);
        }

        /// <summary>
        /// Verificar Error en Reply de Screen. Solo aplica para Mso850 por el comportamiento particular para algunos mensajes específicos
        /// </summary>
        /// <param name="reply"></param>
        /// <returns></returns>
        public static bool ValidateError(Screen.ScreenDTO reply)
        {
            //Si no existe un reply es error de ejecución. O si el reply tiene un error de datos
            if (reply == null)
            {
                Debugger.LogError("RibbonEllipse:ValidateError(Screen.ScreenDTO)", "null reply error");
                return true;
            }
            if (reply.message.Length < 2 || reply.message.Substring(0, 2) != "X2") return false;
            //0039:WO NOT ON FILE
            //7630 WO TASK NOT ON FILE
            //8438 HRS NORMALES INTRODUCIR PARA COD GANANCIAS TIPO OT
            //5331 WORK ORDER IS CLOSED TO COMMITMENT
            if (reply.message.Substring(3, 4).Equals("6008")) //6008:WARNING - VALUE IS ZERO (ORD_HRS); 
                return false;
            if (reply.message.Substring(3, 4).Equals("8852")) //8852 UNPROCESSED COSTING DETAILS DISPLAYED
                return false;
            if (reply.message.Substring(3, 4).Equals("4744")) //4744
                return false;
            if (reply.message.Substring(3, 4).Equals("8839")) //8839
                return false;

            Debugger.LogError("LabourCosting.ValidateError(ScreenDTO reply)", "Error: " + reply.message);
            return true;
        }
    }
}

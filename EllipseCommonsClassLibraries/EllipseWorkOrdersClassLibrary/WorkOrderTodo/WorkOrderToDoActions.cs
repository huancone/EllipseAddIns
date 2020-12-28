using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Utilities;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary.EquipmentReqmntsService;
using EllipseWorkOrdersClassLibrary.MaterialReqmntsService;
using EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using EllipseWorkOrdersClassLibrary.WorkOrderTaskService;

namespace EllipseWorkOrdersClassLibrary
{
    public static class WorkOrderToDoActions
    {
        public static List<WorkOrderToDoItem> FetchToDoItems(string urlService, string districtCode, string workOrder, string workOrderTask)
        {
            var opContext = new TodoListWorkOrderService.OperationContext();
            return FetchToDoItems(urlService, districtCode, workOrder, workOrderTask, opContext);
        }

        public static List<WorkOrderToDoItem> FetchToDoItems(string urlService, string districtCode, string workOrder, string workOrderTask, TodoListWorkOrderService.OperationContext opContext)
        {
            var toDoList = new List<WorkOrderToDoItem>();
            var service = new TodoListWorkOrderService.TodoListWorkOrderService();
            service.Url = urlService + "/TodoListWorkOrder";

            var itemDto = new TodoListWorkOrderService.TodoListWorkOrderSearchParam();
            itemDto.districtCode = districtCode;
            itemDto.workOrder = workOrder;
            itemDto.workOrderTask = workOrderTask;

            var reply = service.retrieveTodoInstances(opContext, itemDto, null);
            if (reply == null || reply.Length <= 0)
                throw new Exception("No se han encontrado elementos para el registro " + districtCode + " " + workOrder + " " + workOrderTask);

            string errors = "";
            foreach(var r in reply)
            {
                if(r.errors != null && r.errors.Length > 0)
                {
                    foreach(var e in r.errors)
                        errors = errors + e + ". ";
                }

                var toDoItem = new WorkOrderToDoItem(r.todoListItemInstanceDTO);
                toDoItem.DistrictCode = districtCode;
                toDoItem.WorkOrder = workOrder;
                toDoItem.WorkOrderTask = workOrderTask;
                toDoList.Add(toDoItem);
            }

            if (!string.IsNullOrWhiteSpace(errors))
                throw new Exception(errors);
            return toDoList;

        }

        public static WorkOrderToDoItem CreateToDoItems(string urlService, TodoListWorkOrderService.OperationContext opContext, WorkOrderToDoItem toDoItem)
        {
            var service = new TodoListWorkOrderService.TodoListWorkOrderService();
            service.Url = urlService + "/TodoListWorkOrder";

            var reply = service.createTodoInstance(opContext, toDoItem.ToWorkOrderTodoInstanceDto());
            

            string errors = "";
            if (reply.errors != null && reply.errors.Length > 0)
            {
                foreach (var e in reply.errors)
                    errors = errors + e + ". ";
            }

            if (!string.IsNullOrWhiteSpace(errors))
                throw new Exception(errors);

            var replyDto = reply.todoListItemInstanceDTO;
            if (string.IsNullOrWhiteSpace(replyDto.todoListItemInstanceId) || toDoItem.ItemName != replyDto.itemName)
                throw new Exception("No se ha podido crear ningún item");

            var replyItem = new WorkOrderToDoItem(replyDto);
            replyItem.DistrictCode = toDoItem.DistrictCode;
            replyItem.WorkOrder = toDoItem.WorkOrder;
            replyItem.WorkOrderTask = toDoItem.WorkOrderTask;

            return replyItem;
        }

        public static WorkOrderToDoItem DeleteToDoItems(string urlService, TodoListWorkOrderService.OperationContext opContext, WorkOrderToDoItem toDoItem)
        {
            var service = new TodoListWorkOrderService.TodoListWorkOrderService();
            service.Url = urlService + "/TodoListWorkOrder";

            var reply = service.retrieveTodoInstances(opContext, toDoItem.ToWorkOrderSearchParam(), null);
            
            TodoListWorkOrderService.TodoListWorkOrderServiceResult delReply = null;

            if (reply == null || reply.Length <= 0)
                throw new Exception("No se han encontrado elementos para eliminar " + toDoItem.DistrictCode + " " + toDoItem.WorkOrder + " " + toDoItem.WorkOrderTask);

            foreach (var r in reply)
            {
                var replyItem = new WorkOrderToDoItem(r.todoListItemInstanceDTO);
                
                if(replyItem.ToDoListItemInstanceId == toDoItem.ToDoListItemInstanceId || replyItem.Sequence == toDoItem.Sequence || replyItem.ItemName == toDoItem.ItemName)
                {
                    var item = replyItem.ToWorkOrderTodoInstanceDto();
                    item.districtCode = toDoItem.DistrictCode;
                    item.workOrder = toDoItem.WorkOrder;
                    item.workOrderTask = toDoItem.WorkOrderTask;

                    delReply = service.deleteTodoInstance(opContext, item);

                    string errors = "";
                    if (delReply != null && delReply.errors != null && delReply.errors.Length > 0)
                    {
                        foreach (var e in delReply.errors)
                            errors = errors + e.messageText + ". ";
                    }

                    if (!string.IsNullOrWhiteSpace(errors))
                        throw new Exception(errors);

                    replyItem.DistrictCode = toDoItem.DistrictCode;
                    replyItem.WorkOrder = toDoItem.WorkOrder;
                    replyItem.WorkOrderTask = toDoItem.WorkOrderTask;
                    return replyItem;
                }

            }

            throw new Exception("No se han encontrado elementos para eliminar " + toDoItem.DistrictCode + " " + toDoItem.WorkOrder + " " + toDoItem.WorkOrderTask);
        }

        public static WorkOrderToDoItem UpdateToDoItems(string urlService, TodoListWorkOrderService.OperationContext opContext, WorkOrderToDoItem toDoItem)
        {
            var service = new TodoListWorkOrderService.TodoListWorkOrderService();
            service.Url = urlService + "/TodoListWorkOrder";

            var reply = service.retrieveTodoInstances(opContext, toDoItem.ToWorkOrderSearchParam(), null);

            TodoListWorkOrderService.TodoListItemInstanceServiceResult updReply = null;

            if (reply == null || reply.Length <= 0)
                throw new Exception("No se han encontrado elementos para actualizar " + toDoItem.DistrictCode + " " + toDoItem.WorkOrder + " " + toDoItem.WorkOrderTask);

            foreach (var r in reply)
            {
                var replyItem = new WorkOrderToDoItem(r.todoListItemInstanceDTO);

                if (replyItem.ToDoListItemInstanceId == toDoItem.ToDoListItemInstanceId || replyItem.Sequence == toDoItem.Sequence || replyItem.ItemName == toDoItem.ItemName)
                {
                    toDoItem.ToDoListItemInstanceId = replyItem.ToDoListItemInstanceId;

                    updReply = service.updateTodoInstance(opContext, toDoItem.ToWorkOrderTodoInstanceDto());

                    string errors = "";
                    if (updReply != null && updReply.errors != null && updReply.errors.Length > 0)
                    {
                        foreach (var e in updReply.errors)
                            errors = errors + e.messageText + ". ";
                    }

                    if (!string.IsNullOrWhiteSpace(errors))
                        throw new Exception(errors);

                    replyItem = new WorkOrderToDoItem(updReply.todoListItemInstanceDTO);
                    replyItem.DistrictCode = toDoItem.DistrictCode;
                    replyItem.WorkOrder = toDoItem.WorkOrder;
                    replyItem.WorkOrderTask = toDoItem.WorkOrderTask;
                    return replyItem;
                }

            }

            throw new Exception("No se han encontrado elementos para actualizar " + toDoItem.DistrictCode + " " + toDoItem.WorkOrder + " " + toDoItem.WorkOrderTask);
        }

        public static TodoListWorkOrderService.OperationContext GetOperationContext(string districtCode, string userPosition)
        {
            var opContext = new TodoListWorkOrderService.OperationContext
            {
                district = districtCode,
                position = userPosition,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            return opContext;
        }
    }
}

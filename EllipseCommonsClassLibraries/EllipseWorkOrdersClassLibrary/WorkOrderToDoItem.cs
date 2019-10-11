using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseReferenceCodesClassLibrary;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;

namespace EllipseWorkOrdersClassLibrary
{
    public class WorkOrderToDoItem
    {
        public string DistrictCode;
        public string WorkOrder;
        public string WorkOrderTask;
        public decimal Sequence;
        public bool SequenceSpecified;
        public string ItemName;
        public DateTime RequiredByDate;
        public bool RequiredByDateSpecified;
        public DateTime ExpirationDate;
        public bool ExpirationDateSpecified;
        public bool NeededForRelease;
        public bool NeededForReleaseSpecified;

        public string StatusCode;
        public string ExternalReference;
        public string Notes;
        public string Owner;
        public string ToDoListItemInstanceId;
        public string TodoListWorkOrderId;
        public WorkOrderToDoItem()
        {

        }
        public WorkOrderToDoItem(TodoListWorkOrderService.TodoListItemInstanceDTO instanceDto)
        {
            ExpirationDate = instanceDto.expirationDate;
            ExpirationDateSpecified = instanceDto.expirationDateSpecified;
            ExternalReference = instanceDto.externalReference;
            ItemName = instanceDto.itemName;
            NeededForRelease = instanceDto.neededForRelease;
            NeededForReleaseSpecified = instanceDto.neededForReleaseSpecified;
            Notes = instanceDto.notes;
            Owner = instanceDto.owner;
            RequiredByDate = instanceDto.requiredByDate;
            RequiredByDateSpecified = instanceDto.requiredByDateSpecified;
            Sequence = instanceDto.sequence;
            SequenceSpecified = instanceDto.sequenceSpecified;
            StatusCode = instanceDto.statusCode;
            ToDoListItemInstanceId = instanceDto.todoListItemInstanceId;
            //TodoListWorkOrderId = todoListWorkOrderId;
        }

        public TodoListWorkOrderService.TodoListItemInstanceDTO ToListItemInstanceDto()
        {
            var dto = new TodoListWorkOrderService.TodoListItemInstanceDTO
            {
                expirationDate = ExpirationDate,
                expirationDateSpecified = ExpirationDateSpecified,
                externalReference = ExternalReference,
                itemName = ItemName,
                neededForRelease = NeededForRelease,
                neededForReleaseSpecified = NeededForReleaseSpecified,
                notes = Notes,
                owner = Owner,
                requiredByDate = RequiredByDate,
                requiredByDateSpecified = RequiredByDateSpecified,
                sequence = Sequence,
                sequenceSpecified = SequenceSpecified,
                statusCode = StatusCode,
                todoListItemInstanceId = ToDoListItemInstanceId,
                //todoListWorkOrderId = TodoListWorkOrderId
            };


            return dto;
        }

        public TodoListWorkOrderService.WorkOrderTodoInstanceDTO ToWorkOrderTodoInstanceDto()
        {
            var dto = new TodoListWorkOrderService.WorkOrderTodoInstanceDTO
            {
                districtCode = DistrictCode,
                workOrder = WorkOrder,
                workOrderTask = WorkOrderTask,
                expirationDate = ExpirationDate,
                expirationDateSpecified = ExpirationDateSpecified,
                externalReference = ExternalReference,
                itemName = ItemName,
                neededForRelease = NeededForRelease,
                neededForReleaseSpecified = NeededForReleaseSpecified,
                notes = Notes,
                owner = Owner,
                requiredByDate = RequiredByDate,
                requiredByDateSpecified = RequiredByDateSpecified,
                sequence = Sequence,
                sequenceSpecified = SequenceSpecified,
                statusCode = StatusCode,
                todoListItemInstanceId = ToDoListItemInstanceId,
                //todoListWorkOrderId = TodoListWorkOrderId
            };

            return dto;
        }
        public TodoListWorkOrderService.TodoListWorkOrderSearchParam ToWorkOrderSearchParam()
        {
            var dto = new TodoListWorkOrderService.TodoListWorkOrderSearchParam
            {
                districtCode = DistrictCode,
                workOrder = WorkOrder,
                workOrderTask = WorkOrderTask,
                //todoListWorkOrderId = TodoListWorkOrderId
            };

            return dto;
        }
        public TodoListWorkOrderService.TodoListWorkOrderDTO ToTodoListWorkOrderDto()
        {
            var dto = new TodoListWorkOrderService.TodoListWorkOrderDTO
            {
                districtCode = DistrictCode,
                workOrder = WorkOrder,
                workOrderTask = WorkOrderTask,
                todoListItemInstanceId = ToDoListItemInstanceId,
                todoListWorkOrderId = TodoListWorkOrderId
            };

            return dto;
        }
        
    }
}

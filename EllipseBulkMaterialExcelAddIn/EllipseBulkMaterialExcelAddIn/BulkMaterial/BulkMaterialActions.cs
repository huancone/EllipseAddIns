using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BMUService = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;
using BMUItemService = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;
using EllipseCommonsClassLibrary;

namespace EllipseBulkMaterialExcelAddIn
{
    public static class BulkMaterialActions
    {
        public static BMUService.BulkMaterialUsageSheetServiceResult ApplyHeader(BMUService.BulkMaterialUsageSheetService bmService, BMUService.OperationContext opContext, BMUService.BulkMaterialUsageSheetDTO requestSheet)
        {
            var reply = bmService.apply(opContext, requestSheet);
            var errorMessage = "";
            if (reply.errors.Length > 0)
            {
                foreach (var t in reply.errors)
                    errorMessage += " - " + t.messageText;

                if (!string.IsNullOrWhiteSpace(errorMessage))
                    throw new Exception(errorMessage);
            }

            return reply;
        }

        public static BMUService.BulkMaterialUsageSheetServiceResult UnApplyHeader(BMUService.BulkMaterialUsageSheetService bmService, BMUService.OperationContext opContext, BMUService.BulkMaterialUsageSheetDTO requestSheet, bool ignoreErrors = false)
        {
            var reply = bmService.unapply(opContext, requestSheet);
            var errorMessage = "";
            if (reply.errors.Length > 0)
            {
                foreach (var t in reply.errors)
                    errorMessage += " - " + t.messageText;

                if (!string.IsNullOrWhiteSpace(errorMessage) && !ignoreErrors)
                    throw new Exception(errorMessage);
            }

            return reply;
        }

        public static BMUService.BulkMaterialUsageSheetServiceResult DeleteHeader(BMUService.BulkMaterialUsageSheetService bmService, BMUService.OperationContext opContext, BMUService.BulkMaterialUsageSheetDTO requestSheet)
        {
            var reply = bmService.delete(opContext, requestSheet);
            var errorMessage = "";
            if (reply.errors.Length > 0)
            {
                foreach (var t in reply.errors)
                    errorMessage += " - " + t.messageText;

                if (!string.IsNullOrWhiteSpace(errorMessage))
                    throw new Exception(errorMessage);
            }

            return reply;
        }

        public static BMUService.BulkMaterialUsageSheetServiceResult CreateHeader(BMUService.BulkMaterialUsageSheetService bmService, BMUService.OperationContext opContext, BMUService.BulkMaterialUsageSheetDTO requestSheet)
        {
            var reply = bmService.create(opContext, requestSheet);

            var errorMessage = "";
            if (reply.errors.Length > 0)
            {
                foreach (var t in reply.errors)
                    errorMessage += " - " + t.messageText;

                if (!string.IsNullOrWhiteSpace(errorMessage))
                    throw new Exception(errorMessage);
            }

            return reply;
        }

        public static BMUItemService.BulkMaterialUsageSheetItemServiceResult AddItemToHeader(EllipseFunctions eFunctions, BMUItemService.BulkMaterialUsageSheetItemService bmItemService, BMUItemService.OperationContext opContext, BMUItemService.BulkMaterialUsageSheetItemDTO requestItem)
        {
            var profile = GetItemFuelCapacity(eFunctions, requestItem.equipmentReference, requestItem.bulkMaterialTypeId);

            if (!string.IsNullOrWhiteSpace(profile.Error))
                throw new Exception(profile.Error);
            if (requestItem.bulkMaterialTypeId == profile.FuelType && requestItem.quantity > profile.Capacity)
                throw new Exception("El valor ingresado supera la capacidad del Equipo!");

            var reply = bmItemService.create(opContext, requestItem);
            var errorMessage = "";

            if (reply.errors.Length > 0)
            {
                foreach (var t in reply.errors)
                    errorMessage += " - " + t.messageText;

                if (!string.IsNullOrWhiteSpace(errorMessage))
                    throw new Exception("ERROR" + errorMessage);
            }

            return reply;
        }


        public static Profile GetItemFuelCapacity(EllipseFunctions eFunctions, string equipNo, string fuelType)
        {
            var profile = new Profile();

            if (string.IsNullOrEmpty(equipNo))
                profile.Error = "Se requiere un número de equipo válido para determinar su capacidad";

            var sqlQuery = Queries.GetFuelCapacity(equipNo, eFunctions.dbReference, eFunctions.dbLink);
            var drEquipCapacity = eFunctions.GetQueryResult(sqlQuery);

            if (!drEquipCapacity.Read())
                profile.Error = "No se ha encontrado un perfil válido para el equipo proporcionado";

            if (!drEquipCapacity.IsClosed && drEquipCapacity.HasRows)
            {
                profile.Equipo = drEquipCapacity["EQUIP_NO"].ToString();
                profile.Egi = drEquipCapacity["EQUIP_GRP_ID"].ToString();
                profile.FuelType = drEquipCapacity["FUEL_OIL_TYPE"].ToString();
                profile.Capacity = Convert.ToDecimal(drEquipCapacity["FUEL_CAPACITY"].ToString());
                return profile;
            }
            else
            {
                profile.Error = "No existe un perfil estadístico de operación para el equipo";
                return profile;
            }
        }

        public static string GetBulkAccountCode(EllipseFunctions eFunctions, string equipNo)
        {
            try
            {
                if (string.IsNullOrEmpty(equipNo)) return "";

                var sqlQuery = Queries.GetBulkAccountCode(equipNo, eFunctions.dbReference, eFunctions.dbLink);
                var drQuery = eFunctions.GetQueryResult(sqlQuery);

                if (!drQuery.Read()) return "";

                if (!drQuery.IsClosed && drQuery.HasRows)
                    return drQuery["BULK_ACCOUNT"].ToString();
                else
                    return "";
            }
            catch(Exception ex)
            {
                Debugger.LogError("BulkMaterialActions.cs::GetBulkAccountCode", ex.Message);
                return "";
            }
        }
    }
}

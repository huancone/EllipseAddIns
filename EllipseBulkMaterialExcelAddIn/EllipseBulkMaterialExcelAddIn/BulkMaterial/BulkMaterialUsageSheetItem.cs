using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BMUSheetItem = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;

namespace EllipseBulkMaterialExcelAddIn.BulkMaterial
{
    public class BulkMaterialUsageSheetItem
    {
        public string AccountCode;
        public string BatchLotNumber;
        public string BinCode;
        public string BulkMaterialTypeId;
        public string BulkMaterialUsageSheetId;
        public string BulkMaterialUsageSheetItemId;
        public string ComponentCode;
        public string ConditionCode;
        public string ConditionMonitoringAction;
        public string EquipmentNumber;
        public string EquipmentReference;
        public string InventoryCategory;
        public string LastModifiedDate;
        public string LastModifiedTime;

        public string MeterReading;
        public string Modifier;
        public string OperationStatisticType;
        public string Quantity;
        public string SubLedger;
        public string SupplierReference;
        public string UnitPrice;
        public string UsageDate;
        public string UsageTime;
        public string UseByDate;

        public BMUSheetItem.Attribute[] customAttributes;

        public BMUSheetItem.BulkMaterialUsageSheetItemDTO ToDto()
        {
            var item = new BMUSheetItem.BulkMaterialUsageSheetItemDTO();
            item.accountCode = AccountCode;
            item.batchLotNumber = BatchLotNumber;
            item.binCode = BinCode;
            item.bulkMaterialTypeId = BulkMaterialTypeId;
            item.bulkMaterialUsageSheetId = BulkMaterialUsageSheetId;
            item.bulkMaterialUsageSheetItemId = BulkMaterialUsageSheetItemId;
            item.componentCode = ComponentCode;
            item.conditionCode = ConditionCode;
            item.conditionMonitoringAction = ConditionMonitoringAction;
            item.equipmentNumber = EquipmentNumber;
            item.equipmentReference = EquipmentReference;
            item.inventoryCategory = InventoryCategory;
            item.lastModifiedDate = LastModifiedDate;
            item.lastModifiedTime = LastModifiedTime;
            item.meterReading = string.IsNullOrEmpty(MeterReading) ? 0 : Convert.ToDecimal(MeterReading);
            item.meterReadingSpecified = string.IsNullOrEmpty(MeterReading) ? false : true;
            item.modifier = Modifier;
            item.operationStatisticType = OperationStatisticType;
            item.quantity = string.IsNullOrEmpty(Quantity) ? 0 : decimal.Round(Convert.ToDecimal(Quantity));
            item.quantitySpecified = string.IsNullOrEmpty(Quantity) ? false : true;
            item.subLedger = SubLedger;
            item.supplierReference = SupplierReference;
            item.unitPrice = string.IsNullOrEmpty(UnitPrice) ? 0 : Decimal.Parse(UnitPrice);
            item.unitPriceSpecified = string.IsNullOrEmpty(UnitPrice) ? false : true;
            item.usageDate = UsageDate;
            item.usageTime = UsageTime;
            item.useByDate = UseByDate;
            item.customAttributes = customAttributes;

            return item;
        }
    }
}

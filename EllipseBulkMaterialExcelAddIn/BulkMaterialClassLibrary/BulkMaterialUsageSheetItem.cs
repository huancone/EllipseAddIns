using System;
using SharedClassLibrary.Utilities;
using BMUSheetItem = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetItemService;

namespace BulkMaterialClassLibrary
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
        public string LastModifiedDate
        {
            get => MyUtilities.ToString(_lastModifiedDate);
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    _lastModifiedDate = null;
                else
                    _lastModifiedDate = MyUtilities.ToDate(value);
            }
        }
        private System.DateTime? _lastModifiedDate;
        public string LastModifiedTime;

        public string MeterReading;
        public string Modifier;
        public string OperationStatisticType;
        public string Quantity;
        public string SubLedger;
        public string SupplierReference;
        public string UnitPrice;
        public string UsageDate
        {
            get => MyUtilities.ToString(_usageDate);
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    _usageDate = null;
                else
                    _usageDate = MyUtilities.ToDate(value);
            }
            
        }
        private System.DateTime? _usageDate;
        public string UsageTime
        {
            get => MyUtilities.ToString(_usageTime, MyUtilities.DateTime.TimeDefaultFormat);
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    _usageTime = null;
                else
                    _usageTime = MyUtilities.ToDateTime(value, MyUtilities.DateTime.TimeDefaultFormat);
            }
        }
        private System.DateTime? _usageTime;
        public string UseByDate
        {
            get => MyUtilities.ToString(_usageByDate);
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    _usageByDate = null;
                else
                    _usageByDate = MyUtilities.ToDate(value);
            }
        }
        private System.DateTime? _usageByDate;

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
            if (_lastModifiedDate != null)
            {
                item.lastModifiedDate = (DateTime)_lastModifiedDate;
                item.lastModifiedDateSpecified = true;
            }
            item.lastModifiedTime = LastModifiedTime;
            item.meterReading = string.IsNullOrEmpty(MeterReading) ? 0 : Convert.ToDecimal(MeterReading);
            item.meterReadingSpecified = !string.IsNullOrEmpty(MeterReading);
            item.modifier = Modifier;
            item.operationStatisticType = OperationStatisticType;
            item.quantity = string.IsNullOrEmpty(Quantity) ? 0 : decimal.Round(Convert.ToDecimal(Quantity));
            item.quantitySpecified = !string.IsNullOrEmpty(Quantity);
            item.subLedger = SubLedger;
            item.supplierReference = SupplierReference;
            item.unitPrice = string.IsNullOrEmpty(UnitPrice) ? 0 : Decimal.Parse(UnitPrice);
            item.unitPriceSpecified = !string.IsNullOrEmpty(UnitPrice);
            if (_usageDate != null)
            {
                item.usageDate = (DateTime) _usageDate;
                item.usageDateSpecified = true;
            }
            if (_usageTime != null)
            {
                item.usageTime = (DateTime) _usageTime;
                item.usageTimeSpecified = true;
            }

            if (_usageByDate != null)
            {
                item.useByDate = (DateTime) _usageByDate;
                item.useByDateSpecified = true;
            }
            item.customAttributes = customAttributes;

            return item;
        }
    }
}

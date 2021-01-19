using System;
using SharedClassLibrary.Utilities;
using BMUSheet = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;

namespace BulkMaterialClassLibrary
{
    public class BulkMaterialUsageSheet
    {
        public BMUSheet.Attribute[] CustomAttributes;
        public string BulkMaterialUsageSheetId;
        public string DefaultAccountCode;
        public string DefaultBatchLotNumber;
        public string DefaultBulkMaterialTypeId;
        public string DefaultSubLedger;
        public string DefaultSupplierReference;
        public string DefaultUsageDate {
            get => MyUtilities.ToString(_defaultUsageDate);
            set => _defaultUsageDate = MyUtilities.ToDate(value);
        }
        private DateTime? _defaultUsageDate;
        public string DefaultUsageTime
        {
            get => MyUtilities.ToString(_defaultUsageTime);
            set => _defaultUsageTime = MyUtilities.ToDateTime(value, MyUtilities.DateTime.TimeDefaultFormat);
        }
        private DateTime? _defaultUsageTime;
        public string DefaultUseByDate
        {
            get => MyUtilities.ToString(_defaultUseByDate);
            set => _defaultUseByDate = MyUtilities.ToDate(value);
        }
        private DateTime? _defaultUseByDate;
        public string DistrictCode;
        public string EmployeeId;
        public string LastModifiedDate
        {
            get => MyUtilities.ToString(_lastModifiedDate);
            set => _lastModifiedDate = MyUtilities.ToDate(value);
        }
        private DateTime? _lastModifiedDate;
        public string LastModifiedTime { get; set; }
        
        public string RecordedBy;
        public string Status;
        public string SupplierNumber;
        public string SupplyCustomerAccountId;
        public string WarehouseId;

        public bool Equals(BulkMaterialUsageSheet bulkMaterialUsageSheet, bool ignoreNullSheetId = true)
        {
            if(!ignoreNullSheetId || (BulkMaterialUsageSheetId != null && bulkMaterialUsageSheet.BulkMaterialUsageSheetId != null))
                if (BulkMaterialUsageSheetId != bulkMaterialUsageSheet.BulkMaterialUsageSheetId)
                    return false;

            if (DefaultAccountCode != bulkMaterialUsageSheet.DefaultAccountCode)
                return false;
            if (DefaultBatchLotNumber != bulkMaterialUsageSheet.DefaultBatchLotNumber)
                return false;
            if (DefaultBulkMaterialTypeId != bulkMaterialUsageSheet.DefaultBulkMaterialTypeId)
                return false;
            if (DefaultSubLedger != bulkMaterialUsageSheet.DefaultSubLedger)
                return false;
            if (DefaultSupplierReference != bulkMaterialUsageSheet.DefaultSupplierReference)
                return false;
            if (DefaultUsageDate != bulkMaterialUsageSheet.DefaultUsageDate)
                return false;
            if (DefaultUsageTime != bulkMaterialUsageSheet.DefaultUsageTime)
                return false;
            if (DefaultUseByDate != bulkMaterialUsageSheet.DefaultUseByDate)
                return false;
            if (DistrictCode != bulkMaterialUsageSheet.DistrictCode)
                return false;
            if (EmployeeId != bulkMaterialUsageSheet.EmployeeId)
                return false;
            if (LastModifiedDate != bulkMaterialUsageSheet.LastModifiedDate)
                return false;
            if (LastModifiedTime != bulkMaterialUsageSheet.LastModifiedTime)
                return false;
            if (RecordedBy != bulkMaterialUsageSheet.RecordedBy)
                return false;
            if (Status != bulkMaterialUsageSheet.Status)
                return false;
            if (SupplierNumber != bulkMaterialUsageSheet.SupplierNumber)
                return false;
            if (SupplyCustomerAccountId != bulkMaterialUsageSheet.SupplyCustomerAccountId)
                return false;
            if (WarehouseId != bulkMaterialUsageSheet.WarehouseId)
                return false;

            return true;
        }

        public BMUSheet.BulkMaterialUsageSheetDTO ToDto()
        {
            var buSheet = new BMUSheet.BulkMaterialUsageSheetDTO();

            buSheet.bulkMaterialUsageSheetId = BulkMaterialUsageSheetId;
            buSheet.defaultAccountCode = DefaultAccountCode;
            buSheet.defaultBatchLotNumber = DefaultBatchLotNumber;
            buSheet.defaultBulkMaterialTypeId = DefaultBulkMaterialTypeId;
            buSheet.defaultSubLedger = DefaultSubLedger;
            buSheet.defaultSupplierReference = DefaultSupplierReference;
            if (_defaultUsageDate != null)
            {
                buSheet.defaultUsageDate = (DateTime) (_defaultUsageDate);
                buSheet.defaultUsageDateSpecified = true;
            }

            if (_defaultUsageTime != null)
            {
                buSheet.defaultUsageTime = (DateTime) _defaultUsageTime;
                buSheet.defaultUsageTimeSpecified = true;
            }

            if (_defaultUseByDate != null)
            {
                buSheet.defaultUseByDate = (DateTime) _defaultUseByDate;
                buSheet.defaultUseByDateSpecified = true;
            }

            buSheet.districtCode = DistrictCode;
            buSheet.employeeId = EmployeeId;
            if (_lastModifiedDate != null)
            {
                buSheet.lastModifiedDate = (DateTime) _lastModifiedDate;
                buSheet.lastModifiedDateSpecified = true;
            }

            buSheet.lastModifiedTime = LastModifiedTime;
            buSheet.recordedBy = RecordedBy;
            buSheet.status = Status;
            buSheet.supplierNumber = SupplierNumber;
            buSheet.supplyCustomerAccountId = SupplyCustomerAccountId;
            buSheet.warehouseId = WarehouseId;
            buSheet.customAttributes = CustomAttributes;

            return buSheet;
        }
    }
}

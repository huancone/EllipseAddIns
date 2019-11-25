using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BMUSheet = EllipseBulkMaterialExcelAddIn.BulkMaterialUsageSheetService;

namespace EllipseBulkMaterialExcelAddIn.BulkMaterial
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
        public string DefaultUsageDate;
        public string DefaultUsageTime;
        public string DefaultUseByDate;
        public string DistrictCode;
        public string EmployeeId;
        public string LastModifiedDate;
        public string LastModifiedTime;
        public string RecordedBy;
        public string Status;
        public string SupplierNumber;
        public string SupplyCustomerAccountId;
        public string WarehouseId;

        public BulkMaterialUsageSheet()
        {

        }

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
            buSheet.defaultUsageDate = DefaultUsageDate;
            buSheet.defaultUsageTime = DefaultUsageTime;
            buSheet.defaultUseByDate = DefaultUseByDate;
            buSheet.districtCode = DistrictCode;
            buSheet.employeeId = EmployeeId;
            buSheet.lastModifiedDate = LastModifiedDate;
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

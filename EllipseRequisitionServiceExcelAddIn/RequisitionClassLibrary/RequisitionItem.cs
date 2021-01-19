using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary
{
    public class RequisitionItem
    {
        public int Index;
        public string StockCode;
        public string UnitOfMeasure;
        public string ItemType;
        public decimal QuantityRequired;
        public decimal IssueRequisitionItem;
        public bool QuantityRequiredSpecified;
        public bool PartIssue;
        public bool RepairRequest;
        public bool RepairRequestProtect;
        public bool IssueDocoFlg;
        public bool PurchDocoFlg;
        public bool NarrativeExists;
        public string AlterStockCodeFlg;
        public bool IssueRequisitionItemSpecified;
        public bool RepairRequestSpecified;
        public bool RepairRequestProtectSpecified;
        public bool PartIssueSpecified;
        public bool IssueDocoFlgSpecified;
        public bool PurchDocoFlgSpecified;
        public bool NarrativeExistsSpecified;
        public bool AlterStockCodeFlgSpecified;
        public bool DirectOrderIndicator;


        public RequisitionItem()
        {
            IssueRequisitionItem = 0;
            RepairRequest = false;
            RepairRequestProtect = false;
            PartIssue = false;
            IssueDocoFlg = false;
            PurchDocoFlg = false;
            NarrativeExists = false;

            QuantityRequiredSpecified = true;

            IssueRequisitionItemSpecified = false;
            RepairRequestSpecified = false;
            RepairRequestProtectSpecified = false;
            PartIssueSpecified = false;
            IssueDocoFlgSpecified = false;
            PurchDocoFlgSpecified = false;
            NarrativeExistsSpecified = false;
            AlterStockCodeFlgSpecified = false;
        }

        public RequisitionService.RequisitionItemDTO GetRequisitionItemDto()
        {
            var item = new RequisitionService.RequisitionItemDTO
            {
                stockCode = StockCode,
                unitOfMeasure = UnitOfMeasure,
                itemType = ItemType,
                quantityRequired = QuantityRequired,
                issueRequisitionItem = IssueRequisitionItem,
                quantityRequiredSpecified = QuantityRequiredSpecified,
                partIssue = PartIssue,
                repairRequest = RepairRequest,
                repairRequestProtect = RepairRequestProtect,
                issueDocoFlg = IssueDocoFlg,
                purchDocoFlg = PurchDocoFlg,
                narrativeExists = NarrativeExists,
                alterStockCode = AlterStockCodeFlg,
                issueRequisitionItemSpecified = IssueRequisitionItemSpecified,
                repairRequestSpecified = RepairRequestSpecified,
                repairRequestProtectSpecified = RepairRequestProtectSpecified,
                partIssueSpecified = PartIssueSpecified,
                issueDocoFlgSpecified = IssueDocoFlgSpecified,
                purchDocoFlgSpecified = PurchDocoFlgSpecified,
                narrativeExistsSpecified = NarrativeExistsSpecified,
                alterStockCodeFlgSpecified = AlterStockCodeFlgSpecified
            };


            return item;
        }
    }

}

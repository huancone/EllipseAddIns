using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStockCodesExcelAddIn
{
    public class PurchaseOrder
    {
        public string PurchaseNumber;
        public string OrderType;
        public string NumberOfItems;
        public string OrderStatus;
        public string TotalEstimatedValue;
        public string AuthorizedStatus;
        public string SupplierNumber;
        public string SupplierName;
        public string DeliveryLocation;
        public string FreightCode;
        public string OrderDate;
        public string PurchaseOfficer;
        public string PurchaseTeam;
        public string Medium;
        public string OriginCode;
        public string Currency;

        public List<PurchaseOrderItem> Items;
    }
}

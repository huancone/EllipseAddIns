using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.SearchConstants
{
   

    public static class SearchFieldCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Requisition = new KeyValuePair<int, string>(1, "RequisitionNo");
        public static KeyValuePair<int, string> CreatedBy = new KeyValuePair<int, string>(2, "CreatedBy");
        public static KeyValuePair<int, string> RequestedBy = new KeyValuePair<int, string>(3, "RequestedBy");
        public static KeyValuePair<int, string> RequestedPos = new KeyValuePair<int, string>(4, "RequestedPos");
        public static KeyValuePair<int, string> AuthorizedBy = new KeyValuePair<int, string>(5, "AuthorizedBy");
        public static KeyValuePair<int, string> AuthorizedPos = new KeyValuePair<int, string>(6, "AuthorizedPos");
        public static KeyValuePair<int, string> AccountCode = new KeyValuePair<int, string>(7, "AccountCode");
        public static KeyValuePair<int, string> WorkOrder = new KeyValuePair<int, string>(8, "WorkOrder");
        //public static KeyValuePair<int, string> StockCode = new KeyValuePair<int, string>(9, "StockCode");
        //public static KeyValuePair<int, string> PartNumber = new KeyValuePair<int, string>(10, "PartNumber");
        public static KeyValuePair<int, string> WorkGroup = new KeyValuePair<int, string>(11, "WorkGroup");
        public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(12, "Equipment No");
        public static KeyValuePair<int, string> ProductiveUnit = new KeyValuePair<int, string>(13, "ProductiveUnit");
        public static KeyValuePair<int, string> ParentWorkOrder = new KeyValuePair<int, string>(14, "ParentWorkOrder");
        public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(15, "ListType");
        public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(16, "ListId");
        public static KeyValuePair<int, string> Egi = new KeyValuePair<int, string>(17, "EGI");
        public static KeyValuePair<int, string> Area = new KeyValuePair<int, string>(18, "Area");
        public static KeyValuePair<int, string> Quartermaster = new KeyValuePair<int, string>(19, "SuperIntendencia");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>>
            {
                None,
                WorkOrder,
                //StockCode,
                //PartNumber,
                WorkGroup,
                EquipmentReference,
                ProductiveUnit,
                CreatedBy,
                RequestedBy,
                AuthorizedBy,
                AuthorizedPos,
                AccountCode,
                ParentWorkOrder,
                ListId,
                ListType,
                Egi,
                Area,
                Quartermaster
            };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

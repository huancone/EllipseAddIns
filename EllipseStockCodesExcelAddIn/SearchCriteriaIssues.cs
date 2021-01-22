using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStockCodesExcelAddIn
{
    public static class SearchCriteriaIssues
    {
        public static KeyValuePair<int, string> Inventory = new KeyValuePair<int, string>(0, "Inventory");
        public static KeyValuePair<int, string> PurchaseOrder = new KeyValuePair<int, string>(1, "PurchaseOrder");
        public static KeyValuePair<int, string> Requisition = new KeyValuePair<int, string>(2, "Requisition");
        public static KeyValuePair<int, string> RequisitionDetailed = new KeyValuePair<int, string>(3, "Requisition Detailed");

        public static List<KeyValuePair<int, string>> GetSearchCriteriaIssues(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { Inventory, PurchaseOrder, Requisition, RequisitionDetailed };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

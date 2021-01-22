using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStockCodesExcelAddIn
{
    public static class SearchCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> StockCode = new KeyValuePair<int, string>(1, "StockCode");
        public static KeyValuePair<int, string> PartNumber = new KeyValuePair<int, string>(2, "PartNumber");
        public static KeyValuePair<int, string> ItemCode = new KeyValuePair<int, string>(3, "ItemCode");

        public static List<KeyValuePair<int, string>> GetSearchCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { None, StockCode, PartNumber, ItemCode };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

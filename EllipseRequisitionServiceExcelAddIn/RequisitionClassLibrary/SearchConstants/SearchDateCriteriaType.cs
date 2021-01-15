using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.SearchConstants
{
    public static class SearchDateCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Creation = new KeyValuePair<int, string>(1, "CreatedDate");
        //public static KeyValuePair<int, string> Closed = new KeyValuePair<int, string>(2, "ClosedDate");
        public static KeyValuePair<int, string> Required = new KeyValuePair<int, string>(3, "RequiredDate");
        public static KeyValuePair<int, string> WoRaisedDate = new KeyValuePair<int, string>(4, "WoRaisedDate");
        public static KeyValuePair<int, string> WoPlanStartDate = new KeyValuePair<int, string>(5, "WoPlanStarDate");
        public static KeyValuePair<int, string> IgnoreDate = new KeyValuePair<int, string>(6, "IgnoreDate");

        public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>>
            {
                None,
                Creation,
                //Closed,
                Required,
                WoRaisedDate,
                WoPlanStartDate,
                IgnoreDate
            };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

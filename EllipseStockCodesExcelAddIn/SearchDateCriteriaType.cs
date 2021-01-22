using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStockCodesExcelAddIn
{
    public static class SearchDateCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Raised = new KeyValuePair<int, string>(1, "Raised");
        //public static KeyValuePair<int, string> Closed = new KeyValuePair<int, string>(2, "Closed");
        //public static KeyValuePair<int, string> PlannedStart = new KeyValuePair<int, string>(3, "PlannedStart");
        //public static KeyValuePair<int, string> PlannedFinnish = new KeyValuePair<int, string>(4, "PlannedFinnish");
        //public static KeyValuePair<int, string> RequiredStart = new KeyValuePair<int, string>(5, "RequiredStart");
        //public static KeyValuePair<int, string> RequiredBy = new KeyValuePair<int, string>(6, "RequiredBy");
        //public static KeyValuePair<int, string> Modified = new KeyValuePair<int, string>(7, "Modified");
        //public static KeyValuePair<int, string> NotFinalized = new KeyValuePair<int, string>(8, "NotFinalized");
        //public static KeyValuePair<int, string> LastModified = new KeyValuePair<int, string>(9, "LastModified");
        //public static KeyValuePair<int, string> Finalized = new KeyValuePair<int, string>(10, "Finalized");

        public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
        {
            //var list = new List<KeyValuePair<int, string>> { None, Raised, Closed, PlannedStart, PlannedFinnish, RequiredStart, RequiredBy, Modified, NotFinalized };
            var list = new List<KeyValuePair<int, string>> { None, Raised };
            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

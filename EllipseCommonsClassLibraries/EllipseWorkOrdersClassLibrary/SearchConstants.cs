using System.Collections.Generic;
using System.Linq;

namespace EllipseWorkOrdersClassLibrary
{
    public static class SearchDateCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Raised = new KeyValuePair<int, string>(1, "Raised");
        public static KeyValuePair<int, string> Closed = new KeyValuePair<int, string>(2, "Closed");
        public static KeyValuePair<int, string> PlannedStart = new KeyValuePair<int, string>(3, "PlannedStart");
        public static KeyValuePair<int, string> PlannedFinnish = new KeyValuePair<int, string>(4, "PlannedFinnish");
        public static KeyValuePair<int, string> RequiredStart = new KeyValuePair<int, string>(5, "RequiredStart");
        public static KeyValuePair<int, string> RequiredBy = new KeyValuePair<int, string>(6, "RequiredBy");
        public static KeyValuePair<int, string> Modified = new KeyValuePair<int, string>(7, "Modified");
        public static KeyValuePair<int, string> NotFinalized = new KeyValuePair<int, string>(8, "NotFinalized");
        public static KeyValuePair<int, string> LastModified = new KeyValuePair<int, string>(9, "LastModified");
        //public static KeyValuePair<int, string> Finalized = new KeyValuePair<int, string>(10, "Finalized");

        public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> {None, Raised, Closed, PlannedStart, PlannedFinnish, RequiredStart, RequiredBy, Modified, NotFinalized};

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }

    public static class SearchFieldCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> WorkGroup = new KeyValuePair<int, string>(1, "WorkGroup");
        public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(2, "Equipment No");
        public static KeyValuePair<int, string> ProductiveUnit = new KeyValuePair<int, string>(3, "ProductiveUnit");
        public static KeyValuePair<int, string> Originator = new KeyValuePair<int, string>(4, "Originator");
        public static KeyValuePair<int, string> CompletedBy = new KeyValuePair<int, string>(5, "Originator");
        public static KeyValuePair<int, string> AccountCode = new KeyValuePair<int, string>(6, "AccountCode");
        public static KeyValuePair<int, string> WorkRequest = new KeyValuePair<int, string>(7, "WorkRequest");
        public static KeyValuePair<int, string> ParentWorkOrder = new KeyValuePair<int, string>(8, "ParentWorkOrder");
        public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(9, "ListType");
        public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(10, "ListId");
        public static KeyValuePair<int, string> Egi = new KeyValuePair<int, string>(11, "EGI");
        public static KeyValuePair<int, string> EquipmentClass = new KeyValuePair<int, string>(12, "Equipment Class");
        public static KeyValuePair<int, string> Area = new KeyValuePair<int, string>(13, "Area");
        public static KeyValuePair<int, string> Quartermaster = new KeyValuePair<int, string>(14, "SuperIntendencia");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> {None, WorkGroup, EquipmentReference, ProductiveUnit, Originator, CompletedBy, AccountCode, WorkRequest, ParentWorkOrder, ListId, ListType, Egi, EquipmentClass, Area, Quartermaster};

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

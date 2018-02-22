using System.Collections.Generic;
using System.Linq;

namespace EllipseJobsClassLibrary
{

    public static class SearchFieldCriteriaType
    {
        public static KeyValuePair<int, string> WorkGroup = new KeyValuePair<int, string>(1, "WorkGroup");
        public static KeyValuePair<int, string> Area = new KeyValuePair<int, string>(13, "Area");
        public static KeyValuePair<int, string> Quartermaster = new KeyValuePair<int, string>(14, "SuperIntendencia");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { WorkGroup, Area, Quartermaster };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }

    public static class SearchDateCriteriaType
    {
        public static KeyValuePair<int, string> PlannedStart = new KeyValuePair<int, string>(1, "PlannedStart");

        public static List<KeyValuePair<int, string>> GetSearchDateCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { PlannedStart };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }

}

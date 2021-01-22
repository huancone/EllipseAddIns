using System.Collections.Generic;
using System.Linq;

namespace EllipseMSO010ExcelAddIn
{

    public static class SearchFieldCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Type = new KeyValuePair<int, string>(1, "Table Type");
        public static KeyValuePair<int, string> Code = new KeyValuePair<int, string>(2, "Table Code");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { None, Type, Code };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

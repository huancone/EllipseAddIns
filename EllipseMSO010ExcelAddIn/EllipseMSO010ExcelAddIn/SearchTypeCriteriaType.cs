using System.Collections.Generic;
using System.Linq;

namespace EllipseMSO010ExcelAddIn
{
    public static class SearchTypeCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> EqualsTo = new KeyValuePair<int, string>(1, "Equal");
        public static KeyValuePair<int, string> StartsWith = new KeyValuePair<int, string>(2, "Starts With");
        public static KeyValuePair<int, string> EndsWith = new KeyValuePair<int, string>(3, "EndsWith");
        public static KeyValuePair<int, string> Contains = new KeyValuePair<int, string>(4, "Contains");

        public static List<KeyValuePair<int, string>> GetSearchTypeCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { None, EqualsTo, StartsWith, EndsWith, Contains };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

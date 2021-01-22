using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO010ExcelAddIn
{
    public static class StatusCode
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> Active = new KeyValuePair<int, string>(1, "Active");
        public static KeyValuePair<int, string> Inactive = new KeyValuePair<int, string>(2, "Inactive");

        public static List<KeyValuePair<int, string>> GetStatusList(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>> { None, Active, Inactive };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

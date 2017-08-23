using System.Collections.Generic;
using System.Linq;

namespace EllipseEquipmentClassLibrary
{
    public static class SearchFieldCriteria
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> EquipmentNo = new KeyValuePair<int, string>(1, "Equipment No");
        public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(2, "Equipment Ref");
        public static KeyValuePair<int, string> ProductiveUnit = new KeyValuePair<int, string>(3, "ProductiveUnit");
        public static KeyValuePair<int, string> EquipmentDescription = new KeyValuePair<int, string>(4, "Description");
        public static KeyValuePair<int, string> CreationUser = new KeyValuePair<int, string>(5, "Creation User");
        public static KeyValuePair<int, string> AccountCode = new KeyValuePair<int, string>(6, "AccountCode");
        public static KeyValuePair<int, string> Custodian = new KeyValuePair<int, string>(7, "Custodian");
        public static KeyValuePair<int, string> CustodianPosition = new KeyValuePair<int, string>(8, "Cust. Post.");
        public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(9, "ListType");
        public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(10, "ListId");
        public static KeyValuePair<int, string> Egi = new KeyValuePair<int, string>(11, "EGI");
        public static KeyValuePair<int, string> EquipmentClass = new KeyValuePair<int, string>(12, "Equipment Class");
        public static KeyValuePair<int, string> EquipmentType = new KeyValuePair<int, string>(13, "Equipment Type");
        public static KeyValuePair<int, string> EquipmentLocation = new KeyValuePair<int, string>(14, "Location");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>>
            {
                None, EquipmentNo, EquipmentReference, ProductiveUnit, EquipmentDescription, CreationUser, AccountCode, Custodian,
                CustodianPosition, ListId, ListType, Egi, EquipmentClass, EquipmentType, EquipmentLocation
            };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

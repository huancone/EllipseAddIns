using System.Collections.Generic;
using System.Linq;

namespace EllipseEquipmentClassLibrary
{
    public static class EquipListSearchFieldCriteria
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(1, "List Type");
        public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(2, "List Id");
        public static KeyValuePair<int, string> ListNumber = new KeyValuePair<int, string>(3, "List Number");
        public static KeyValuePair<int, string> ListDescription = new KeyValuePair<int, string>(4, "List Description");
        public static KeyValuePair<int, string> ListReference = new KeyValuePair<int, string>(5, "List Reference");
        public static KeyValuePair<int, string> EquipmentNo = new KeyValuePair<int, string>(6, "Equipment No");
        public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(7, "Equipment Reference");
        public static KeyValuePair<int, string> ListOwner = new KeyValuePair<int, string>(8, "List Owner");
        public static KeyValuePair<int, string> ListOwnerPosition = new KeyValuePair<int, string>(9, "List Owner Position");
        public static KeyValuePair<int, string> ListRaisedBy = new KeyValuePair<int, string>(10, "List Raised By");
        public static KeyValuePair<int, string> MemberEquipNo = new KeyValuePair<int, string>(11, "Member Equip No.");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>>
            {
                None, 
                ListType, 
                ListId, 
                ListNumber, 
                ListDescription, 
                ListReference, 
                EquipmentNo, 
                EquipmentReference, 
                ListOwner, 
                ListOwnerPosition, 
                ListRaisedBy, 
                MemberEquipNo
            };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}

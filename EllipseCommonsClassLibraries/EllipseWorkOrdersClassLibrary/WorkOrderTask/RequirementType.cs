using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseWorkOrdersClassLibrary
{
    public static class RequirementType
    {
        public static KeyValuePair<string, string> Material = new KeyValuePair<string, string>("MAT", "Material");
        public static KeyValuePair<string, string> Labour = new KeyValuePair<string, string>("LAB", "Labour");
        public static KeyValuePair<string, string> Equipment = new KeyValuePair<string, string>("EQP", "Equipment");
        public static KeyValuePair<string, string> All = new KeyValuePair<string, string>("ALL", "All Resources");

    }
}

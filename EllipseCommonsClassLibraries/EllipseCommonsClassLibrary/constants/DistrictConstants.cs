using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseCommonsClassLibrary
{
    public static class DistrictConstants
    {
        public static string DistrictIcor = "ICOR";
        public static string DistrictInstalations = "INST";
        public static string DefaultDistrict = "ICOR";

        public static List<string> GetDistrictList()
        {
            // ReSharper disable once UseObjectOrCollectionInitializer
            var districtList = new List<string>();
            districtList.Add(DistrictIcor);
            districtList.Add(DistrictInstalations);

            return districtList;
        }
    }
}

using System.Collections.Generic;

namespace SharedClassLibrary.Ellipse.Constants
{
    public static class Districts
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
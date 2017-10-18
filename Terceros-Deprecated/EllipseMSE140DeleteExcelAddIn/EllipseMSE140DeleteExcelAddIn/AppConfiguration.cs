using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSE140DeleteExcelAddIn
{
    public class AppConfiguration
    {
        public static string GetConfiguration(string Key)
        {
            string Value = "";
            try
            {
                Value = System.Configuration.ConfigurationManager.AppSettings[Key].ToString();
            }
            catch (Exception)
            {
                Value = "";
            }
            return Value;
        }
    }
}

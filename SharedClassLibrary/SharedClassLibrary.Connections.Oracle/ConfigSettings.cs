using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Configuration;

namespace SharedClassLibrary.Connections.Oracle
{
    public class ConfigSettings
    {
        public static string OracleTnsNode = "configuration/oracle.manageddataaccess.client/version/settings";

        /// <summary>
        ///     Actualiza la URL del archivo de TNS de Oracle
        /// </summary>
        /// <param name="newUrl"></param>
        public static void UpdateTnsUrlValue(string newUrl)
        {
            const string tnsItenName = "setting";
            var key = new KeyValuePair<string, string>("value", newUrl);

            RuntimeConfigSettings.EditNodeItemKeyValue(OracleTnsNode, tnsItenName, key);
        }

        public static string GetTnsUrlValue()
        {
            var rootNode = OracleTnsNode;
            const string node = "setting";
            const string keyName = "value";
            return RuntimeConfigSettings.GetNodeItemKeyValue(rootNode, node, keyName);
        }
    }
}

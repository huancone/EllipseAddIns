using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Xml;

// ReSharper disable LoopCanBePartlyConvertedToQuery
// ReSharper disable LoopCanBeConvertedToQuery

namespace CommonsClassLibrary
{
    public class RuntimeConfigSettings
    {
        //public static string OracleTnsNode = "configuration/oracle.manageddataaccess.client/version/settings/setting";
        public static string OracleTnsNode = "configuration/oracle.manageddataaccess.client/version/settings";

        /// <summary>
        ///     Adiciona un item especificado a la estructura raíz
        ///     <rootnode>
        ///         <node keyname1= keyvalue1 keyname2= keyvalue2 />
        ///     </rootnode>
        /// </summary>
        /// <param name="rootNode">La raíz nodo donde será adicionado el item nodo</param>
        /// <param name="node">Nombre del item nodo a adicionar</param>
        /// <param name="keyList">Lista de parámetros keyValuePair(string,string) a adicionar al item nodo</param>
        public static void AddNodeItem(string rootNode, string node, List<KeyValuePair<string, string>> keyList)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            // create new node <add key="Region" value="Canterbury" />
            var nodeRegion = xmlDoc.CreateElement(node);
            foreach (var key in keyList) nodeRegion.SetAttribute(key.Key, key.Value);
            xmlDoc.SelectSingleNode(rootNode).AppendChild(nodeRegion);

            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            ConfigurationManager.RefreshSection(rootNode);
        }

        /// <summary>
        ///     Edita la primera coincidencia del nodo item que tenga como llave el keyname especificado
        /// </summary>
        /// <param name="rootNode">La raíz nodo donde será editado el item nodo</param>
        /// <param name="node">Nombre del item nodo a editar</param>
        /// <param name="key">Parámetro KeyValuePair(name, value) con el nuevo valor</param>
        public static void EditNodeItemKeyValue(string rootNode, string node, KeyValuePair<string, string> key)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            var fullNode = "//" + rootNode + "/" + node + "[@" + key.Key + "]";
            xmlDoc.SelectSingleNode(fullNode).Attributes[key.Key].Value = key.Value;
            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            ConfigurationManager.RefreshSection(rootNode);
        }

        /// <summary>
        ///     Elimina la primera coincidencia del nodo item que tenga como llave el keyName, keyValue especificado
        /// </summary>
        /// <param name="rootNode">La raíz nodo de donde será eliminado el item nodo</param>
        /// <param name="node">Nombre del item nodo a eliminar</param>
        /// <param name="key">
        ///     Parámetro KeyValuePair(name, value) con el valor a eliminar. Si value es nulo solo tomará en cuenta
        ///     el keyName para la búsqueda
        /// </param>
        public static void DeleteNodeItem(string rootNode, string node, KeyValuePair<string, string> key)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            var keyValue = string.IsNullOrWhiteSpace(key.Value) ? "" : "='" + key.Value + "'";
            var fullNode = "//" + rootNode + "/" + node + "[@" + key.Key + keyValue + "]";
            var nodeItem = xmlDoc.SelectSingleNode(fullNode);
            nodeItem.ParentNode.RemoveChild(nodeItem);

            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            ConfigurationManager.RefreshSection(rootNode);
        }

        /// <summary>
        ///     Obtiene el valor del parámetro keyName de algún item nodo especificado
        /// </summary>
        /// <param name="rootNode">La raíz nodo de donde será consultado el item nodo</param>
        /// <param name="node">Nombre del item nodo a consultar</param>
        /// <param name="keyName">Nombre del atributo a consultar su valor</param>
        /// <returns></returns>
        public static string GetNodeItemKeyValue(string rootNode, string node, string keyName)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            var fullNode = "//" + rootNode + "/" + node + "[@" + keyName + "]";
            var nodeItem = xmlDoc.SelectSingleNode(fullNode);
            return nodeItem.Attributes[keyName].Value;
        }

        /// <summary>
        ///     Exporta el archivo de configuración a una cadena de texto
        /// </summary>
        /// <returns></returns>
        public static string PrintConfigFile()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            var stringWriter = new StringWriter();
            var xmlTextWriter = new XmlTextWriter(stringWriter);

            xmlDoc.WriteTo(xmlTextWriter);

            return stringWriter.ToString();
        }

        /// <summary>
        ///     Actualiza la URL del archivo de TNS de Oracle
        /// </summary>
        /// <param name="newUrl"></param>
        public static void UpdateTnsUrlValue(string newUrl)
        {
            const string tnsItenName = "setting";
            var key = new KeyValuePair<string, string>("value", newUrl);
            EditNodeItemKeyValue(OracleTnsNode, tnsItenName, key);
        }

        public static string GetTnsUrlValue()
        {
            var rootNode = OracleTnsNode;
            const string node = "setting";
            const string keyName = "value";
            return GetNodeItemKeyValue(rootNode, node, keyName);
        }
    }
}
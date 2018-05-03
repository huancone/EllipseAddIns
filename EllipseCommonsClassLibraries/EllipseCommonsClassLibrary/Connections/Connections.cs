using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
// ReSharper disable PossibleNullReferenceException
// ReSharper disable LoopCanBeConvertedToQuery

namespace EllipseCommonsClassLibrary.Connections
{
    public static class ServiceType
    {
        public static string PostService = "POST";
        public static string EwsService = "EWS";
    }
    public static class Environments
    {
        public static string EllipseProductivo = "Productivo";
        public static string EllipseContingencia = "Contingencia";
        public static string EllipseDesarrollo = "Desarrollo";
        public static string EllipseTest = "Test";

        public static string SigcorProductivo = "SIGCOPROD";
        public static string ScadaRdb = "SCADARDB";
        public static string CustomDatabase = "Personalizada";

        public const string DefaultDbReferenceName = "ELLIPSE";

        public static string GetServiceUrl(string enviroment, string serviceType = null)
        {
            var urlService = SelectServiceUrl(enviroment, serviceType);
            if (!string.IsNullOrWhiteSpace(urlService) && (urlService.EndsWith("/") || urlService.EndsWith("\\")))
                urlService = urlService.Substring(0, urlService.Length - 1);
            return urlService;
        }

        private static string SelectServiceUrl(string enviroment, string serviceType = null)
        {
            if (serviceType == null)
                serviceType = ServiceType.EwsService;

            var serviceFile = Configuration.ServiceFilePath + "\\" + Configuration.ConfigXmlFileName;
            if (!File.Exists(serviceFile))
                throw new Exception("No se puede leer el archivo de configuración de servicios de Ellipse. Asegúrese de que el archivo exista o cree un archivo local.");

            var xmlDoc = XDocument.Load(serviceFile);
            string urlServer;
            if (serviceType.Equals(ServiceType.EwsService))
            {

                if (enviroment == EllipseProductivo)
                    urlServer = WebService.Productivo;
                else if (enviroment == EllipseContingencia)
                    urlServer = WebService.Contingencia;
                else if (enviroment == EllipseDesarrollo)
                    urlServer = WebService.Desarrollo;
                else if (enviroment == EllipseTest)
                    urlServer = WebService.Test;
                else
                    urlServer = "/ellipse/webservice/" + enviroment;

                return xmlDoc.XPathSelectElement(urlServer + "[1]").Value;
            }
            if (serviceType.Equals(ServiceType.PostService))
            {
                if (enviroment == EllipseProductivo)
                    urlServer = UrlPost.Productivo;
                else if (enviroment == EllipseContingencia)
                    urlServer = UrlPost.Contingencia;
                else if (enviroment == EllipseDesarrollo)
                    urlServer = UrlPost.Desarrollo;
                else if (enviroment == EllipseTest)
                    urlServer = UrlPost.Test;
                else
                    urlServer = "/ellipse/url/" + enviroment;

                return xmlDoc.XPathSelectElement(urlServer + "[1]").Value;
            }
            throw new NullReferenceException("No se ha encontrado el servidor seleccionado");
        }
        public static List<string> GetEnviromentList()
        {
            var enviromentList = new List<string>();
            if (Configuration.IsServiceListForced)
            {
                var xmlDoc = new XmlDocument();
                var urlPath = Configuration.ServiceFilePath + "\\" + Configuration.ConfigXmlFileName;
                xmlDoc.Load(urlPath);

                const string fullNode = "//ellipse/url";
                var nodeItemList = xmlDoc.SelectSingleNode(fullNode).ChildNodes;

                foreach (XmlNode item in nodeItemList)
                    enviromentList.Add(item.Name);
            }
            else
            {
                enviromentList = new List<string>
                {
                    EllipseProductivo,
                    EllipseTest,
                    EllipseDesarrollo,
                    EllipseContingencia
                };
            }
            return enviromentList;
        }

        public static DatabaseItem GetDatabaseItem(string enviroment)
        {
            var dbItem = new DatabaseItem();

            try
            {
                if (Configuration.IsServiceListForced)
                {
                    var xmlDoc = new XmlDocument();
                    var urlPath = Configuration.LocalDataPath + Configuration.DatabaseXmlFileName;
                    xmlDoc.Load(urlPath);

                    const string fullNode = "//ellipse/connections";
                    var nodeItemList = xmlDoc.SelectSingleNode(fullNode).ChildNodes;

                    foreach (XmlNode item in nodeItemList)
                    {
                        if (!item.Name.Equals(enviroment)) continue;

                        var dbPassword = item.Attributes["dbpassword"] != null ? item.Attributes["dbpassword"].Value : null;
                        var dbEncodedPassword = item.Attributes["dbencodedpassword"] != null ? item.Attributes["dbencodedpassword"].Value : null;

                        dbItem.Name = item.Name;
                        dbItem.DbName = item.Attributes["dbname"] != null ? item.Attributes["dbname"].Value : null;
                        dbItem.DbUser = item.Attributes["dbuser"] != null ? item.Attributes["dbuser"].Value : null;
                        if (string.IsNullOrWhiteSpace(dbPassword) && !string.IsNullOrWhiteSpace(dbEncodedPassword))
                            dbItem.DbEncodedPassword = dbEncodedPassword;
                        else
                            dbItem.DbPassword = dbPassword;
                        dbItem.DbReference = item.Attributes["dbreference"] != null ? item.Attributes["dbreference"].Value : null;
                        dbItem.DbLink = item.Attributes["dblink"] != null ? item.Attributes["dblink"].Value : null;
                        dbItem.DbCatalog = item.Attributes["dbcatalog"] != null ? item.Attributes["dbcatalog"].Value : null;
                        return dbItem;
                    }
                }
                else
                {
                    if (enviroment == EllipseProductivo)
                    {
                        dbItem.Name = EllipseProductivo;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (enviroment == EllipseDesarrollo)
                    {
                        dbItem.Name = EllipseDesarrollo;
                        dbItem.DbName = "EL8DESA";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (enviroment == EllipseContingencia)
                    {
                        dbItem.Name = EllipseContingencia;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (enviroment == EllipseTest)
                    {
                        dbItem.Name = EllipseTest;
                        dbItem.DbName = "EL8TEST";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (enviroment == SigcorProductivo)
                    {
                        dbItem.Name = SigcorProductivo;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbPassword = "consulbo";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (enviroment == ScadaRdb)
                    {
                        dbItem.Name = ScadaRdb;
                        dbItem.DbName = "PBVFWL01";
                        dbItem.DbUser = "SCADARDBADMINGUI";
                        dbItem.DbPassword = "momia2011";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                        dbItem.DbCatalog = "SCADARDB.DBO";
                    }
                    return dbItem;
                }
                return dbItem;
            }
            catch (Exception ex)
            {
                Debugger.LogError("Connections:GetDatabaseItem(string) " + ex.Message, "No se ha encontrado el archivo xml de base de datos o el entorno seleccionado no existe en este archivo. Verifique la ruta del archivo y compruebe que la información del servidor existe");
                return null;
            }
        }

    }
}

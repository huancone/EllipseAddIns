using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using CommonsClassLibrary.Connections;
// ReSharper disable PossibleNullReferenceException
// ReSharper disable AccessToStaticMemberViaDerivedType

namespace EllipseCommonsClassLibrary.Connections
{
    public static class ServiceType
    {
        public static string PostService = "POST";
        public static string EwsService = "EWS";
    }
    public static class Environments
    {
        public const string EllipseProductivo = "Productivo";
        public const string EllipseContingencia = "Contingencia";
        public const string EllipseDesarrollo = "Desarrollo";
        public const string EllipseTest = "Test";

        public const string SigcorProductivo = "SIGCOR";
        public const string SigmanProductivo = "SIGMANPRD";
        public const string SigmanTest = "SIGMANTST";
        public const string EllipseSigmanProductivo = "Ellipse-Sigman-PRD";
        public const string EllipseSigmanTest = "Ellipse-Sigman-TST";
        public const string ScadaRdb = "SCADARDB";
        public const string CustomDatabase = "Personalizada";
        
        private const string EllProd = "ellprod";
        private const string EllCont = "ellcont";
        private const string EllDesa = "elldesa";
        private const string EllTest = "elltest";

        public const string DefaultDbReferenceName = "ELLIPSE";

        /// <summary>
        /// Obtiene la URL de conexión al servicio web de Ellipse
        /// </summary>
        /// <param name="environment">Nombre del ambiente al que se va a conectar (EnvironmentConstants.Ambiente)</param>
        /// <param name="serviceType">Tipo de conexión a realizar EWS/POST. Localizada en EnvironmentConstans.ServiceType</param>
        /// <returns>string: URL de la conexión</returns>
        public static string GetServiceUrl(string environment, string serviceType = null)
        {
            var urlService = SelectServiceUrl(environment, serviceType);
            if (!string.IsNullOrWhiteSpace(urlService) && (urlService.EndsWith("/") || urlService.EndsWith("\\")))
                urlService = urlService.Substring(0, urlService.Length - 1);
            return urlService;
        }

        private static string SelectServiceUrl(string environment, string serviceType = null)
        {
            if (serviceType == null)
                serviceType = ServiceType.EwsService;

            var serviceFile = Settings.CurrentSettings.ServiceFilePath + "\\" + Settings.CurrentSettings.ServicesConfigXmlFileName;
            var serviceFileBackUp = Settings.CurrentSettings.BackUpServiceFilePath + "\\" + Settings.CurrentSettings.ServicesConfigXmlFileName;
            var serviceFileLocal = Settings.CurrentSettings.DefaultLocalDataPath + "\\" + Settings.CurrentSettings.ServicesConfigXmlFileName;

            XDocument xmlDoc;
            if (File.Exists(serviceFile))
            {
                xmlDoc = XDocument.Load(serviceFile);
            }
            else if (File.Exists(serviceFileBackUp))
            {
                xmlDoc = XDocument.Load(serviceFileBackUp);
            }
            else if (File.Exists(serviceFileLocal))
            {
                xmlDoc = XDocument.Load(serviceFileLocal);
            }
            else
            {
                throw new Exception("No se puede leer el archivo de configuración de servicios de Ellipse. Asegúrese de que el archivo exista o cree un archivo local.");
            }

            string urlServer;
            if (serviceType.Equals(ServiceType.EwsService))
            {

                if (environment == EllipseProductivo)
                    urlServer = WebService.Productivo;
                else if (environment == EllipseContingencia)
                    urlServer = WebService.Contingencia;
                else if (environment == EllipseDesarrollo)
                    urlServer = WebService.Desarrollo;
                else if (environment == EllipseTest)
                    urlServer = WebService.Test;
                else
                    urlServer = "/ellipse/webservice/" + environment;

                return xmlDoc.XPathSelectElement(urlServer + "[1]").Value;
            }
            if (serviceType.Equals(ServiceType.PostService))
            {
                if (environment == EllipseProductivo)
                    urlServer = UrlPost.Productivo;
                else if (environment == EllipseContingencia)
                    urlServer = UrlPost.Contingencia;
                else if (environment == EllipseDesarrollo)
                    urlServer = UrlPost.Desarrollo;
                else if (environment == EllipseTest)
                    urlServer = UrlPost.Test;
                else
                    urlServer = "/ellipse/url/" + environment;

                return xmlDoc.XPathSelectElement(urlServer + "[1]").Value;
            }
            throw new NullReferenceException("No se ha encontrado el servidor seleccionado");
        }

        public static List<string> GetEnvironmentList()
        {
            var environmentList = new List<string>();
            if (Settings.CurrentSettings.IsServiceListForced)
            {
                var xmlDoc = new XmlDocument();
                var urlPath = Settings.CurrentSettings.ServiceFilePath + Settings.CurrentSettings.ServicesConfigXmlFileName;
                var urlPathBackUp = Settings.CurrentSettings.SecondaryServiceFilePath + Settings.CurrentSettings.ServicesConfigXmlFileName;
                var urlLocalPath = Settings.CurrentSettings.DefaultLocalDataPath + Settings.CurrentSettings.ServicesConfigXmlFileName;

                if (File.Exists(urlPath))
                {
                    xmlDoc.Load(urlPath);
                }
                else if (File.Exists(urlPathBackUp))
                {
                    xmlDoc.Load(urlPathBackUp);
                }
                else if (File.Exists(urlLocalPath))
                {
                    xmlDoc.Load(urlLocalPath);
                }
                else
                {
                    throw new Exception("No se puede leer el archivo de configuración de servicios de Ellipse. Asegúrese de que el archivo exista o cree un archivo local.");
                }
                
                const string fullNode = "//ellipse/url"; 
                var nodeItemList = xmlDoc.SelectSingleNode(fullNode).ChildNodes;

                foreach (XmlNode item in nodeItemList)
                    environmentList.Add(item.Name);
            }
            else
            {
                environmentList = new List<string>
                {
                    EllipseProductivo,
                    EllipseTest,
                    EllipseDesarrollo,
                    EllipseContingencia
                };
            }
            return environmentList;
        }

        public static DatabaseItem GetDatabaseItem(string environment)
        {
            var dbItem = new DatabaseItem();

            try
            {
                if (Settings.CurrentSettings.IsServiceListForced)
                {
                    var xmlDoc = new XmlDocument();
                    var urlPath = Path.Combine(Settings.CurrentSettings.LocalDataPath, Settings.CurrentSettings.DatabaseXmlFileName);
                    
                    xmlDoc.Load(urlPath);

                    const string fullNode = "//ellipse/connections";
                    var nodeItemList = xmlDoc.SelectSingleNode(fullNode).ChildNodes;

                    foreach (XmlNode item in nodeItemList)
                    {
                        if (!item.Name.Equals(environment)) continue;

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
                    if (environment.Equals(EllipseProductivo) || environment.Equals(EllProd))
                    {
                        dbItem.Name = EllipseProductivo;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseDesarrollo) || environment.Equals(EllDesa))
                    {
                        dbItem.Name = EllipseDesarrollo;
                        dbItem.DbName = "EL8DESA";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseContingencia) || environment.Equals(EllCont))
                    {
                        dbItem.Name = EllipseContingencia;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseTest) || environment.Equals(EllTest))
                    {
                        dbItem.Name = EllipseTest;
                        dbItem.DbName = "EL8TEST";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigcorProductivo)
                    {
                        dbItem.Name = SigcorProductivo;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbPassword = "consulbo";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigmanProductivo)
                    {
                        dbItem.Name = SigmanProductivo;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "SIGMAN";
                        dbItem.DbPassword = "sig0679";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigmanTest)
                    {
                        dbItem.Name = SigmanTest;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "SIGMAN";
                        dbItem.DbPassword = "sig0679";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == EllipseSigmanProductivo)
                    {
                        dbItem.Name = EllipseSigmanProductivo;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbPassword = "ventyx15";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                        dbItem.SecondaryDbLink = "@DBLSIGMAN";
                        dbItem.SecondaryDbReference = DefaultDbReferenceName;
                    }
                    else if (environment == EllipseSigmanTest)
                    {
                        dbItem.Name = EllipseSigmanTest;
                        dbItem.DbName = "EL8TEST";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbPassword = "ventyx";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                        dbItem.SecondaryDbLink = "@DBLSIGMAN"; //o @DBLSIG
                        dbItem.SecondaryDbReference = DefaultDbReferenceName;

                    }
                    else if (environment == ScadaRdb)
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
                Debugger.LogError("Connections:GetDatabaseItem(string) " + ex.Message, "Ha ocurrido un error al intentar conectarse al entorno del archivo xml. Asegúrese de que la base de datos o el entorno seleccionado sea válida. Verifique la ruta del archivo xml y compruebe que la información del servidor existe");
                return null;
            }
        }

    }
}

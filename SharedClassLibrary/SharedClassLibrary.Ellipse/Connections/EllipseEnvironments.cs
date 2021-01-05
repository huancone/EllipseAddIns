using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

// ReSharper disable PossibleNullReferenceException
// ReSharper disable AccessToStaticMemberViaDerivedType

namespace SharedClassLibrary.Ellipse.Connections
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
        
        internal const string EllProd = "ellprod";
        internal const string EllCont = "ellcont";
        internal const string EllDesa = "elldesa";
        internal const string EllTest = "elltest";

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

                        var dbPassword = item.Attributes["dbpassword"]?.Value;
                        var dbEncodedPassword = item.Attributes["dbencodedpassword"]?.Value;

                        dbItem.Name = item.Name;
                        dbItem.DbName = item.Attributes["dbname"]?.Value;
                        dbItem.DbUser = item.Attributes["dbuser"]?.Value;
                        if (string.IsNullOrWhiteSpace(dbPassword) && !string.IsNullOrWhiteSpace(dbEncodedPassword))
                            dbItem.DbEncodedPassword = dbEncodedPassword;
                        else
                            dbItem.DbPassword = dbPassword;
                        dbItem.DbReference = item.Attributes["dbreference"]?.Value;
                        dbItem.DbLink = item.Attributes["dblink"]?.Value;
                        dbItem.DbCatalog = item.Attributes["dbcatalog"]?.Value;
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
                        dbItem.DbEncodedPassword = @"EkJUiDBAY+nauIRmT33pNrLNoPapgUAw44M9aT0ZGcXdgIE/X4OLxD+22C2QMz2RqK+3SlBomkowWQcclWh94a+90BKkq+eL9KPaFJPcD9rEEc3VhEKoP2mrfR3OPWBL";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseDesarrollo) || environment.Equals(EllDesa))
                    {
                        dbItem.Name = EllipseDesarrollo;
                        dbItem.DbName = "EL8DESA";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbEncodedPassword = @"EkJUiDBAY+nauIRmT33pNrLNoPapgUAw44M9aT0ZGcXdgIE/X4OLxD+22C2QMz2RqK+3SlBomkowWQcclWh94a+90BKkq+eL9KPaFJPcD9rEEc3VhEKoP2mrfR3OPWBL";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseContingencia) || environment.Equals(EllCont))
                    {
                        dbItem.Name = EllipseContingencia;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbEncodedPassword = @"EkJUiDBAY+nauIRmT33pNrLNoPapgUAw44M9aT0ZGcXdgIE/X4OLxD+22C2QMz2RqK+3SlBomkowWQcclWh94a+90BKkq+eL9KPaFJPcD9rEEc3VhEKoP2mrfR3OPWBL";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment.Equals(EllipseTest) || environment.Equals(EllTest))
                    {
                        dbItem.Name = EllipseTest;
                        dbItem.DbName = "EL8TEST";
                        dbItem.DbUser = "SIGCON";
                        dbItem.DbEncodedPassword = @"EkJUiDBAY+nauIRmT33pNrLNoPapgUAw44M9aT0ZGcXdgIE/X4OLxD+22C2QMz2RqK+3SlBomkowWQcclWh94a+90BKkq+eL9KPaFJPcD9rEEc3VhEKoP2mrfR3OPWBL";
                        dbItem.DbLink = "";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigcorProductivo)
                    {
                        dbItem.Name = SigcorProductivo;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbEncodedPassword = @"rrm0HFcFN947tZwu5yAyaCvrALk9emYLn3SaNh2huucpBc6X6SoapF7jc1S1lnVzknUF6Z3DGrNABiwg2PSUnn5ERDzNlL34+EBG6jrSNv1P3NJxas5vy0C2fULYmy/G";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigmanProductivo)
                    {
                        dbItem.Name = SigmanProductivo;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "SIGMAN";
                        dbItem.DbEncodedPassword = @"2yqN2BVsTTW8mrK21olA5KEAEwRqMXds/CpySMMtN0uA5ZPsWWWZjsJcXTbCQxklGQLZCq6jYJOzmo4UNbEs503XWwI1KiX7+7WDgZ2Beems8lIsIBb++yKVlplNidFB";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == SigmanTest)
                    {
                        dbItem.Name = SigmanTest;
                        dbItem.DbName = "SIGCOPRD";
                        dbItem.DbUser = "SIGMAN";
                        dbItem.DbEncodedPassword = @"2yqN2BVsTTW8mrK21olA5KEAEwRqMXds/CpySMMtN0uA5ZPsWWWZjsJcXTbCQxklGQLZCq6jYJOzmo4UNbEs503XWwI1KiX7+7WDgZ2Beems8lIsIBb++yKVlplNidFB";
                        dbItem.DbLink = "@DBLELLIPSE8";
                        dbItem.DbReference = DefaultDbReferenceName;
                    }
                    else if (environment == EllipseSigmanProductivo)
                    {
                        dbItem.Name = EllipseSigmanProductivo;
                        dbItem.DbName = "EL8PROD";
                        dbItem.DbUser = "CONSULBO";
                        dbItem.DbEncodedPassword = @"5RgHgvloJ2S1Eaflx9oonNdkHnXiEhR71mv+hYO2mnkYSt1eH3rSUN3eWahzVMvrdRzH4p1+r6zi0KtVaST8OLlzlqJMQlMXpTSE/Zj4f0XHLa7zOpHTBzi+XE3N9y6f";
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
                        dbItem.DbEncodedPassword = @"cfnShslzZGN3WzBraEZOVjB/dvCylB2l8eQgEwyq0Q6oCEPZVcymeZh9qkAJnybgOkl71K8/C+iBW/duS8ED7Lj9CNMgMG7qzQr78uDG5RVRhkQ3pe4/tdjpDGzSijhd";
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
                        dbItem.DbEncodedPassword = @"z21xgEuA/HsQ8TKMHtQKGKcgTEV0/LaryE/KdhKfLnsGhzqRX7Paa1VwBUDFyJZ5qrYYUQMxTPH2zaHCmrvQmzKKggO9cWDYGbBu7Gs5tHjVurJIiGTBXoK3Lk+UZ+dQ";
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

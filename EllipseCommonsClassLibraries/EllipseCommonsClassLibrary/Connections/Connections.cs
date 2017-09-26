using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;

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

            if (serviceType.Equals(ServiceType.EwsService))
            {
                try
                {
                    //Primeramente intenta conseguir la URL del archivo local o de red
                    var localFile = Configuration.LocalDataPath + Configuration.ConfigXmlFileName;
                    var networkFile = Configuration.UrlServiceFileLocation + Configuration.ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EllipseProductivo)
                        urlServer = WebService.Productivo;
                    if (enviroment == EllipseContingencia)
                        urlServer = WebService.Contingencia;
                    if (enviroment == EllipseDesarrollo)
                        urlServer = WebService.Desarrollo;
                    if (enviroment == EllipseTest)
                        urlServer = WebService.Test;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    //Si no encuentra el archivo de red por alguna causa utiliza el valor predeterminado
                    if (enviroment == EllipseProductivo)
                        return WebService.UrlProductivo;
                    if (enviroment == EllipseContingencia)
                        return WebService.UrlContingencia;
                    if (enviroment == EllipseDesarrollo)
                        return WebService.UrlDesarrollo;
                    if (enviroment == EllipseTest)
                        return WebService.UrlTest;
                }

            }
            else if (serviceType.Equals(ServiceType.PostService))
            {
                try
                {
                    //Primeramente intenta conseguir la URL del archivo local o de red
                    var localFile = Configuration.LocalDataPath + Configuration.ConfigXmlFileName;
                    var networkFile = Configuration.UrlServiceFileLocation + Configuration.ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EllipseProductivo)
                        urlServer = UrlPost.Productivo;
                    if (enviroment == EllipseContingencia)
                        urlServer = UrlPost.Contingencia;
                    if (enviroment == EllipseDesarrollo)
                        urlServer = UrlPost.Desarrollo;
                    if (enviroment == EllipseTest)
                        urlServer = UrlPost.Test;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    if (enviroment == EllipseProductivo)
                        return UrlPost.UrlProductivo;
                    if (enviroment == EllipseContingencia)
                        return UrlPost.Contingencia;
                    if (enviroment == EllipseDesarrollo)
                        return UrlPost.Desarrollo;
                    if (enviroment == EllipseTest)
                        return UrlPost.Test;
                }
            }
            throw new NullReferenceException("No se ha encontrado el servidor seleccionado");
        }
        public static List<string> GetEnviromentList()
        {
            // ReSharper disable once UseObjectOrCollectionInitializer
            var enviromentList = new List<string>();
            enviromentList.Add(EllipseProductivo);
            enviromentList.Add(EllipseTest);
            enviromentList.Add(EllipseDesarrollo);
            enviromentList.Add(EllipseContingencia);

            return enviromentList;
        }

    }
}

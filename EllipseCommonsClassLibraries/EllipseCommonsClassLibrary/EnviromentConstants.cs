using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;

namespace EllipseCommonsClassLibrary
{

    public static class EnviromentConstants
    {
        public static class ServiceType
        {
            public static string PostService = "POST";
            public static string EwsService = "EWS";
        }

        public static string ConfigXmlFileName = @"EllipseConfiguration.xml";
        public static string EllipseProductivo = "Productivo";
        public static string EllipseContingencia = "Contingencia";
        public static string EllipseDesarrollo = "Desarrollo";
        public static string EllipseTest = "Test";
        public static string EllipseTest89 = "Test89";
        public static string SigcorProductivo = "SIGCOPROD";
        public static string ScadaRdb = "SCADARDB";
        public static string CustomDatabase = "Personalizada";
        public static string UrlServiceFileLocation = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
        //public static string SigcorTest = "SIGCOTEST";

        public static string EllipseVarNameServiceProductivo = @"/ellipse/webservice/ellprod";//XPath
        public static string EllipseVarNameServiceContingencia = @"/ellipse/webservice/ellcont";//XPath
        public static string EllipseVarNameServiceDesarrollo = @"/ellipse/webservice/elldesa";//XPath
        public static string EllipseVarNameServiceTest = @"/ellipse/webservice/elltest";//XPath
        public static string EllipseVarNameServiceTest89 = @"/ellipse/webservice/el89tst";//XPath

        public static string EllipseUrlServicesProductivo = "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services";
        public static string EllipseUrlServicesContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string EllipseUrlServicesDesarrollo = "http://ews-el8desa.lmnerp03.cerrejon.com/ews/services";
        public static string EllipseUrlServicesTest = "http://ews-el84test.bogdrp03.cerrejon.com/ews/services";
        public static string EllipseUrlServicesTest89 = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";

        public static string EllipseVarNamePostProductivo = @"/ellipse/url/ellprod";//XPath
        public static string EllipseVarNamePostContingencia = @"/ellipse/url/ellcont";//XPath
        public static string EllipseVarNamePostDesarrollo = @"/ellipse/url/elldesa";//XPath
        public static string EllipseVarNamePostTest = @"/ellipse/url/elltest";//XPath
        public static string EllipseVarNamePostTest89 = @"/ellipse/url/el89tst";//XPath

        public static string EllipseUrlPostServicesProductivo = "http://ellipse-el8prod.lmnerp01.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string EllipseUrlPostServicesContingencia = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string EllipseUrlPostServicesDesarrollo = "http://ellipse-el8desa.lmnerp03.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string EllipseUrlPostServicesTest = "http://ellipse-el84test.bogdrp03.cerrejon.com/ria-Ellipse-8.4.32_191/bind?app=";
        public static string EllipseUrlPostServicesTest89 = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.1_226/bind?app=";

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
                    var localFile = Debugger.LocalDataPath + ConfigXmlFileName;
                    var networkFile = UrlServiceFileLocation + ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EllipseProductivo)
                        urlServer = EllipseVarNameServiceProductivo;
                    if (enviroment == EllipseContingencia)
                        urlServer = EllipseVarNameServiceContingencia;
                    if (enviroment == EllipseDesarrollo)
                        urlServer = EllipseVarNameServiceDesarrollo;
                    if (enviroment == EllipseTest)
                        urlServer = EllipseVarNameServiceTest;
                    if (enviroment == EllipseTest89)
                        urlServer = EllipseVarNameServiceTest89;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    //Si no encuentra el archivo de red por alguna causa utiliza el valor predeterminado
                    if (enviroment == EllipseProductivo)
                        return EllipseUrlServicesProductivo;
                    if (enviroment == EllipseContingencia)
                        return EllipseUrlServicesContingencia;
                    if (enviroment == EllipseDesarrollo)
                        return EllipseUrlServicesDesarrollo;
                    if (enviroment == EllipseTest)
                        return EllipseUrlServicesTest;
                    if (enviroment == EllipseTest89)
                        return EllipseUrlServicesTest89;
                }

            }
            else if (serviceType.Equals(ServiceType.PostService))
            {
                try
                {
                    //Primeramente intenta conseguir la URL del archivo local o de red
                    var localFile = Debugger.LocalDataPath + ConfigXmlFileName;
                    var networkFile = UrlServiceFileLocation + ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EllipseProductivo)
                        urlServer = EllipseVarNamePostProductivo;
                    if (enviroment == EllipseContingencia)
                        urlServer = EllipseVarNamePostContingencia;
                    if (enviroment == EllipseDesarrollo)
                        urlServer = EllipseVarNamePostDesarrollo;
                    if (enviroment == EllipseTest)
                        urlServer = EllipseVarNamePostTest;
                    if (enviroment == EllipseTest89)
                        urlServer = EllipseVarNamePostTest89;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    if (enviroment == EllipseProductivo)
                        return EllipseUrlPostServicesProductivo;
                    if (enviroment == EllipseContingencia)
                        return EllipseUrlPostServicesContingencia;
                    if (enviroment == EllipseDesarrollo)
                        return EllipseUrlPostServicesDesarrollo;
                    if (enviroment == EllipseTest)
                        return EllipseUrlPostServicesTest;
                    if (enviroment == EllipseTest89)
                        return EllipseUrlPostServicesTest89;
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
            enviromentList.Add(EllipseTest89);

            return enviromentList;
        }

        public static void GenerateEllipseConfigurationXmlFile()
        {
            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<ellipse>";
            xmlFile += @"  <env>test</env>";
            xmlFile += @"  <url>";
            xmlFile += @"    <ellprod>" + EllipseUrlPostServicesProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + EllipseUrlPostServicesContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + EllipseUrlPostServicesDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + EllipseUrlPostServicesTest + "</elltest>";
            xmlFile += @"    <ell84test>" + EllipseUrlPostServicesTest89 + "</ell84test>";
            xmlFile += @"  </url>";
            xmlFile += @"  <webservice>";
            xmlFile += @"    <ellprod>" + EllipseUrlServicesProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + EllipseUrlServicesContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + EllipseUrlServicesDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + EllipseUrlServicesTest + "</elltest>";
            xmlFile += @"    <ell84test>" + EllipseUrlServicesTest89 + "</ell84test>";
            xmlFile += @"  </webservice>";
            xmlFile += @"</ellipse>";

            try
            {
                var configFilePath = Debugger.LocalDataPath;
                var configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void GenerateEllipseConfigurationXmlFile(string networkUrl)
        {
            try
            {
                var configFilePath = Debugger.LocalDataPath;
                var configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, networkUrl, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void DeleteEllipseConfigurationXmlFile()
        {
            try
            {
                FileWriter.DeleteFile(Debugger.LocalDataPath, ConfigXmlFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración", ex.Message);
                throw;
            }
        }
    }

    
}

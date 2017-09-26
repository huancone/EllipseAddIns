using System;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseCommonsClassLibrary.Connections
{
    public static class Configuration
    {
        public static string ConfigXmlFileName = @"EllipseConfiguration.xml";
        public static string UrlServiceFileLocation = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
        public static string UrlTnsnameFileLocation = @"%SystemDrive%\oracle\product\11.2.0\client\network\ADMIN\";
        private static string _localDataPath = @"%SystemDrive%\ellipse\";
        public static string EllipseAddInEnvironmentVariable = "EllipseAddinsHome";

        public static string LocalDataPath
        {
            get
            {
                var varHome = "" + Environment.GetEnvironmentVariable(EllipseAddInEnvironmentVariable, EnvironmentVariableTarget.User);
                var varHomeExpanded = Environment.ExpandEnvironmentVariables(varHome);
                return string.IsNullOrWhiteSpace(varHomeExpanded) ? _localDataPath : varHomeExpanded;
            }
            set
            {
                Environment.SetEnvironmentVariable(EllipseAddInEnvironmentVariable, value, EnvironmentVariableTarget.User);
                _localDataPath = value;
            }
        }

        public static void GenerateEllipseConfigurationXmlFile()
        {
            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<ellipse>";
            xmlFile += @"  <env>test</env>";
            xmlFile += @"  <url>";
            xmlFile += @"    <ellprod>" + UrlPost.UrlProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + UrlPost.UrlContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + UrlPost.UrlDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + UrlPost.UrlTest + "</elltest>";
            xmlFile += @"  </url>";
            xmlFile += @"  <webservice>";
            xmlFile += @"    <ellprod>" + WebService.UrlProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + WebService.UrlContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + WebService.UrlDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + WebService.UrlTest + "</elltest>";
            xmlFile += @"  </webservice>";
            xmlFile += @"</ellipse>";

            try
            {
                var configFilePath = LocalDataPath;
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
                var configFilePath = LocalDataPath;
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
                FileWriter.DeleteFile(LocalDataPath, ConfigXmlFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración", ex.Message);
                throw;
            }
        }
    
    }

    public static class WebService
    {
        public static string Productivo = @"/ellipse/webservice/ellprod";//XPath
        public static string Contingencia = @"/ellipse/webservice/ellcont";//XPath
        public static string Desarrollo = @"/ellipse/webservice/elldesa";//XPath
        public static string Test = @"/ellipse/webservice/elltest";//XPath
        public static string Test84 = @"/ellipse/webservice/ell84test";//XPath

        public static string UrlProductivo = "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services";
        public static string UrlContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlDesarrollo = "http://ews-el8desa.lmnerp03.cerrejon.com/ews/services";
        public static string UrlTest = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";
        public static string UrlTest84 = "http://ews-el84test.bogdrp03.cerrejon.com/ews/services";
    }

    public static class UrlPost
    {
        public static string Productivo = @"/ellipse/url/ellprod";//XPath
        public static string Contingencia = @"/ellipse/url/ellcont";//XPath
        public static string Desarrollo = @"/ellipse/url/elldesa";//XPath
        public static string Test = @"/ellipse/url/elltest";//XPath
        public static string Test84 = @"/ellipse/url/ell84test";//XPath

        public static string UrlProductivo = "http://ellipse-el8prod.lmnerp01.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string UrlContingencia = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string UrlDesarrollo = "http://ellipse-el8desa.lmnerp03.cerrejon.com/ria-Ellipse-8.4.29_31/bind?app=";
        public static string UrlTest = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.1_226/bind?app=";
        public static string UrlTest84 = "http://ellipse-el84test.bogdrp03.cerrejon.com/ria-Ellipse-8.4.32_191/bind?app=";
    }
}

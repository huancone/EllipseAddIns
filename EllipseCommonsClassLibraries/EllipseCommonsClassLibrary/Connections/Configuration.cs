using System;
using EllipseCommonsClassLibrary.Utilities;
using System.Reflection;
using System.IO;

namespace EllipseCommonsClassLibrary.Connections
{
    public static class Configuration
    {
        public static string DefaultServiceFilePath = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
        public static string DefaultTnsnamesFilePath = @"c:\oracle\product\11.2.0\client\network\ADMIN\";
        public static string DefaultLocalDataPath = @"c:\ellipse\";
        private const string EllipseHomeEnvironmentVariable = "EllipseAddInsHome";
        private const string EllipseServicesEnvironmentVariable = "EllipseServiceUrlFile";
       
        public const string ConfigXmlFileName = "EllipseConfiguration.xml";
        public const string TnsnamesFileName = "tnsnames.ora";

        public static string LocalDataPath
        {
            get
            {
                var varHome = "" + Environment.GetEnvironmentVariable(EllipseHomeEnvironmentVariable, EnvironmentVariableTarget.User);
                var varHomeExpanded = Environment.ExpandEnvironmentVariables(varHome);
                return string.IsNullOrWhiteSpace(varHomeExpanded) ? DefaultLocalDataPath : varHomeExpanded;
            }
            set
            {
                var currentVar = Environment.GetEnvironmentVariable(EllipseHomeEnvironmentVariable, EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(DefaultLocalDataPath))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(DefaultLocalDataPath))
                    Environment.SetEnvironmentVariable(EllipseHomeEnvironmentVariable, value, EnvironmentVariableTarget.User);
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //existe y es diferente a environment -> actualiza
                if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(EllipseHomeEnvironmentVariable, value, EnvironmentVariableTarget.User);
            }
        }
        public static string ServiceFilePath
        {
            get
            {
                var varService = "" + Environment.GetEnvironmentVariable(EllipseServicesEnvironmentVariable, EnvironmentVariableTarget.User);
                var varServiceExpanded = Environment.ExpandEnvironmentVariables(varService);
                return string.IsNullOrWhiteSpace(varServiceExpanded) ? DefaultServiceFilePath : varServiceExpanded;
            }
            set
            {
                var currentVar = Environment.GetEnvironmentVariable(EllipseServicesEnvironmentVariable, EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(DefaultServiceFilePath))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(DefaultServiceFilePath))
                    Environment.SetEnvironmentVariable(EllipseServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //existe y es diferente a environment -> actualiza
                if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(EllipseServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
            }
        }
        public static string TnsnamesFilePath
        {
            get { return DefaultTnsnamesFilePath; }
        }

        public static void GenerateEllipseConfigurationXmlFile(string targetUrl)
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
                if (FileWriter.NormalizePath(targetUrl, true).Equals(FileWriter.NormalizePath(DefaultServiceFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                const string configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void GenerateEllipseConfigurationXmlFile(string sourceUrl, string targetUrl)
        {
            try
            {
                if (FileWriter.NormalizePath(targetUrl, true).Equals(FileWriter.NormalizePath(DefaultServiceFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                const string configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
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
                if (FileWriter.NormalizePath(ServiceFilePath, true).Equals(FileWriter.NormalizePath(DefaultServiceFilePath, true)))
                    throw new Exception("No se puede eliminar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(ServiceFilePath, true);
                FileWriter.DeleteFile(configFilePath, ConfigXmlFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración", ex.Message);
                throw;
            }
        }

        public static void GenerateEllipseTnsnamesFile(string targetUrl)
        {
            try
            {
                if (File.Exists(TnsnamesFilePath + TnsnamesFileName))
                {
                    FileWriter.MoveFileToDirectory(TnsnamesFileName, TnsnamesFilePath, TnsnamesFileName + DateTime.Today.Year + DateTime.Today.Month + DateTime.Today.Day + ".BAK", TnsnamesFilePath);
                }
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                const string configFileName = TnsnamesFileName;

                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "EllipseCommonsClassLibrary.Resources.tnsnames.txt";
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                using (StreamReader reader = new StreamReader(stream))
                {
                    string tnsFileText = reader.ReadToEnd();
                    FileWriter.CreateDirectory(configFilePath);
                    FileWriter.WriteTextToFile(tnsFileText, configFileName, configFilePath);
                }

            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseTnsnamesFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void GenerateEllipseTnsnamesFile(string sourceUrl, string targetUrl)
        {
            try
            {
                if (FileWriter.NormalizePath(targetUrl, true).Equals(FileWriter.NormalizePath(DefaultTnsnamesFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo " + TnsnamesFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                const string configFileName = TnsnamesFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
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

        public static string UrlProductivo = "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services";
        public static string UrlContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlDesarrollo = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";
        public static string UrlTest = "http://ews-el84test.bogdrp03.cerrejon.com/ews/services";
    }

    public static class UrlPost
    {
        public static string Productivo = @"/ellipse/url/ellprod";//XPath
        public static string Contingencia = @"/ellipse/url/ellcont";//XPath
        public static string Desarrollo = @"/ellipse/url/elldesa";//XPath
        public static string Test = @"/ellipse/url/elltest";//XPath

        public static string UrlProductivo = "http://ellipse-el8prod.lmnerp01.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string UrlContingencia = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string UrlDesarrollo = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.1_226/bind?app=";
        public static string UrlTest = "http://ellipse-el84test.bogdrp03.cerrejon.com/ria-Ellipse-8.4.32_191/bind?app=";
    }
}

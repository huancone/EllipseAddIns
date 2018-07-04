using System;
using System.Collections.Generic;
using EllipseCommonsClassLibrary.Utilities;
using System.Reflection;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using EllipseCommonsClassLibrary.Utilities.RuntimeConfigSettings;

namespace EllipseCommonsClassLibrary.Connections
{
    public static class Configuration
    {
        public static string DefaultServiceFilePath = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
        public static string DefaultTnsnamesFilePath = @"c:\oracle\product\11.2.0\client\network\ADMIN\";
        public static string DefaultLocalDataPath = @"c:\ellipse\";
        private const string EllipseHomeEnvironmentVariable = "EllipseAddInsHome";
        private const string EllipseServicesEnvironmentVariable = "EllipseServiceUrlFile";
        private const string EllipseServicesForcedList = "EllipseServiceForcedList";

        public const string ConfigXmlFileName = "EllipseConfiguration.xml";
        public const string TnsnamesFileName = "tnsnames.ora";
        public const string DatabaseXmlFileName = "EllipseDatabases.xml";
        public const string EncryptPassPhrase = "hambingsdevel";

        public static bool IsServiceListForced
        {
            get
            {
                var varForced = "" + Environment.GetEnvironmentVariable(EllipseServicesForcedList, EnvironmentVariableTarget.User);
                var varForcedExpanded = Environment.ExpandEnvironmentVariables(varForced);
                return !string.IsNullOrWhiteSpace(varForcedExpanded) && varForcedExpanded.ToLower().Equals("true");
            }
            set
            {
                Environment.SetEnvironmentVariable(EllipseServicesForcedList, value.ToString(), EnvironmentVariableTarget.User);
            }
        }
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
            get
            {
                return RuntimeConfigSettings.GetTnsUrlValue(); ; 
            }
            set
            {
                if (value.Equals(RuntimeConfigSettings.GetTnsUrlValue()))
                    return;
                RuntimeConfigSettings.UpdateTnsUrlValue(value);
            }
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
                ServiceFilePath = DefaultServiceFilePath;
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
                if (FileWriter.NormalizePath(targetUrl, true).Equals(FileWriter.NormalizePath(DefaultTnsnamesFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo " + TnsnamesFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
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

        public static void GenerateEllipseDatabaseFile()
        {
            var databaseList = new List<DatabaseItem>();
            databaseList.Add(new DatabaseItem("Productivo", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("Contingencia", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("Desarrollo", "EL8DESA", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("Test", "EL8TEST", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("ellprod", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("ellcont", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("elldesa", "EL8DESA", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("elltest", "EL8TEST", "SIGCON", "ventyx", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("SCADARDB", "PBVFWL01", "SCADARDBADMINGUI", "momia2011", "SCADARDB.DBO", null, "SCADARDB"));
            databaseList.Add(new DatabaseItem("SIGCOR", "SIGCOPRD", "CONSULBO", "consulbo", "@DBLELLIPSE8", "ELLIPSE"));
            databaseList.Add(new DatabaseItem("SIGCOPRD", "SIGCOPRD", "CONSULBO", "consulbo", "@DBLELLIPSE8", "ELLIPSE"));

            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<ellipse>";
            xmlFile += @"  <connections>";
            foreach(var item in databaseList)
                xmlFile += @"    <" + item.Name + " dbname='" + item.DbName + "' dbuser='" + item.DbUser + "' dbpassword='' dbencodedpassword='" + item.DbEncodedPassword + "' dbreference='" + item.DbReference + "' dblink='" + item.DbLink + "' " + (string.IsNullOrWhiteSpace(item.DbCatalog) ? null : "dbcatalog='" + item.DbCatalog + "'")+ "/>";
            xmlFile += @"  </connections>";
            xmlFile += @"</ellipse>";

            try
            {
                var configFilePath = FileWriter.NormalizePath(LocalDataPath, true);
                const string configFileName = DatabaseXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseDatabaseFile", "No se puede crear el archivo de bases de datos\n" + ex.Message);
                throw;
            }
        }

        public static void DeleteEllipseDatabaseFile()
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(ServiceFilePath, true);
                FileWriter.DeleteFile(configFilePath, DatabaseXmlFileName);
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

        public static string UrlProductivo = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlDesarrollo = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";
        public static string UrlTest = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services/";
    }

    public static class UrlPost
    {
        public static string Productivo = @"/ellipse/url/ellprod";//XPath
        public static string Contingencia = @"/ellipse/url/ellcont";//XPath
        public static string Desarrollo = @"/ellipse/url/elldesa";//XPath
        public static string Test = @"/ellipse/url/elltest";//XPath

        public static string UrlProductivo = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.9.8_446/bind?app=";
        public static string UrlContingencia = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.9.8_446/bind?app=";
        public static string UrlDesarrollo = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.8_446/bind?app=";
        public static string UrlTest = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.8_446/bind?app=";
    }

    public class DatabaseItem
    {
        public string Name;
        public string DbName;
        public string DbUser;
        private string _dbEncodedPassword;
        public string DbLink;
        public string DbReference;
        public string DbCatalog;
        private string _dbPassword;

        public DatabaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference)
        {
            SetDataBaseItem(name, dbName, dbUser, dbPassword, dbReference, null, null);
        }
        public DatabaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink)
        {
            SetDataBaseItem(name, dbName, dbUser, dbPassword, dbLink, dbReference, null);
        }
        public DatabaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string dbCatalog)
        {
            SetDataBaseItem(name, dbName, dbUser, dbPassword, dbLink, dbReference, dbCatalog);
        }

        public DatabaseItem()
        {
            
        }

        private void SetDataBaseItem(string name, string dbName, string dbUser, string dbPassword,  string dbReference, string dbLink, string dbCatalog)
        {
            Name = name;
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPassword;
            DbReference = dbReference;
            DbLink = dbLink;
            DbCatalog = dbCatalog;

        }

        public string DbPassword
        {
            get { return string.IsNullOrWhiteSpace(_dbPassword) ? EncryptString.Decrypt(DbEncodedPassword, Configuration.EncryptPassPhrase) : _dbPassword; }
            set
            {
                _dbPassword = value;
                _dbEncodedPassword = EncryptString.Encrypt(value, Configuration.EncryptPassPhrase);
            }
        }
        public string DbEncodedPassword
        {
            get { return _dbEncodedPassword; }
            set
            {
                _dbEncodedPassword = value;
                _dbPassword = EncryptString.Decrypt(value, Configuration.EncryptPassPhrase);
            }
        }
    }
}

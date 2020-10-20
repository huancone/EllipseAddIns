using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using System.Reflection;
using System.Xml.Serialization;
using SharedClassLibrary.Connections;

// ReSharper disable AccessToStaticMemberViaDerivedType

namespace EllipseCommonsClassLibrary
{
    public class Settings : SharedClassLibrary.Configuration.Settings
    {
        public static Settings CurrentSettings;

        public Settings()
        {
            var settingOptions = new Options();
            Initialize(settingOptions);
        }
        public Settings(Options defaultProgramOptions)
        {
            Initialize(defaultProgramOptions);
        }

        public override void Initialize(Options defaultProgramOptions)
        {
            AssemblyProgram = new Settings.AssemblyItem(GetLastAssembly());
            //GeneralFolder
            DefaultLocalDataPath = @"c:\ellipse\";
            GeneralConfigFolder = @"addins\" + AssemblyProgram.AssemblyTitle;
            GeneralConfigFileName = "config.xml";
            DefaultRepositoryFilePath = @"\\lmnoas02\Shared\Sistemas\Mina\Proyecto Ellipse\Ellipse 8\ExcelAddIn_E8 (Loaders)\";

            //Windows Environment Variables
            ProgramEnvironmentHomeVariable = AssemblyProgram.AssemblyTitle + "Home";
            HomeEnvironmentVariable = "EllipseAddInsHome";
            ServicesEnvironmentVariable = "EllipseServiceUrlFile";
            SecondaryServicesEnvironmentVariable = "EllipseSecondaryServiceUrlFile";

            //Services & Databases Information
            ServicesForcedList = "EllipseServiceForcedList";
            ServicesConfigXmlFileName = "EllipseConfiguration.xml";
            TnsnamesFileName = "tnsnames.ora";
            DatabaseXmlFileName = "EllipseDatabases.xml";

            DefaultServiceFilePath = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
            SecondaryServiceFilePath = @"\\pbvshr01\SideLine\EllipsePopups\Ellipse8\";
            DefaultTnsnamesFilePath = @"c:\oracle\product\11.2.0\client\network\ADMIN\";

            //StaticReference Through All the Project
            EllipseCommonsClassLibrary.Debugger.LocalDataPath = LocalDataPath;
            CurrentSettings = this;
            //Options
            OptionsSettings = GetOptionsSettings(defaultProgramOptions);
        }

        #region -- Configuration Files Generation --
        public void GenerateEllipseConfigurationXmlFile(string targetUrl)
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
                //iniciamos las variables de directorio y archivos
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var defaultFilePath = FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true);
                var configFileName = CurrentSettings.ServicesConfigXmlFileName;

                //comprobamos que la ruta no corresponda a la ruta predeterminada
                if (configFilePath.Equals(defaultFilePath))
                    throw new Exception("No se puede reemplazar el archivo " + configFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");

                //creamos una copia de seguridad si el archivo existe
                if (FileWriter.CheckFileExist(Path.Combine(configFilePath, configFileName)))
                {
                    var backupFileName = configFileName + "_" + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK";
                    if (!FileWriter.CheckFileExist(Path.Combine(configFilePath, backupFileName)))
                        FileWriter.MoveFileToDirectory(configFileName, configFilePath, backupFileName, configFilePath);
                }


                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public void GenerateEllipseConfigurationXmlFile(string sourceUrl, string targetUrl)
        {
            try
            {
                //iniciamos las variables de directorio y archivos
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var defaultFilePath = FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                var configFileName = CurrentSettings.ServicesConfigXmlFileName;

                //comprobamos que la ruta no corresponda a la ruta predeterminada
                if (configFilePath.Equals(defaultFilePath))
                    throw new Exception("No se puede reemplazar el archivo " + configFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");

                //creamos una copia de seguridad si el archivo existe
                if (FileWriter.CheckFileExist(Path.Combine(configFilePath, configFileName)))
                {
                    var backupFileName = configFileName + "_" + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK";
                    if (!FileWriter.CheckFileExist(Path.Combine(configFilePath, backupFileName)))
                        FileWriter.MoveFileToDirectory(configFileName, configFilePath, backupFileName, configFilePath);
                }

                //realizamos la acción
                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public void DeleteEllipseConfigurationXmlFile()
        {
            try
            {
                //iniciamos las variables de directorio y archivos
                var configFilePath = FileWriter.NormalizePath(CurrentSettings.ServiceFilePath, true);
                var defaultFilePath = FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true);
                var configFileName = CurrentSettings.ServicesConfigXmlFileName;

                //comprobamos que la ruta no corresponda a la ruta predeterminada
                if (configFilePath.Equals(defaultFilePath))
                    throw new Exception("No se puede reemplazar el archivo " + configFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");

                //realizamos la acción
                FileWriter.DeleteFile(configFilePath, configFileName);
                //restablecemos al valor predeterminado
                CurrentSettings.ServiceFilePath = CurrentSettings.DefaultServiceFilePath;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DeleteEllipseConfigurationXmlFile()", "No se puede eliminar el archivo de configuración. " + ex.Message);
                throw;
            }
        }

        public void GenerateEllipseTnsnamesFile(string targetUrl)
        {
            try
            {
                //iniciamos las variables de directorio y archivos
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var defaultFilePath = FileWriter.NormalizePath(CurrentSettings.DefaultTnsnamesFilePath, true);
                var configFileName = CurrentSettings.TnsnamesFileName;

                //comprobamos que la ruta no corresponda a la ruta predeterminada
                if (configFilePath.Equals(defaultFilePath))
                    throw new Exception("No se puede reemplazar el archivo " + configFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");

                //creamos una copia de seguridad si el archivo existe
                if (FileWriter.CheckFileExist(Path.Combine(configFilePath, configFileName)))
                {
                    var backupFileName = configFileName + "_" + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK";
                    if (!FileWriter.CheckFileExist(Path.Combine(configFilePath, backupFileName)))
                        FileWriter.MoveFileToDirectory(configFileName, configFilePath, backupFileName, configFilePath);
                }

                //realizamos la acción
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "EllipseCommonsClassLibrary.Resources.tnsnames.txt";
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                using (var reader = new StreamReader(stream))
                {
                    var tnsFileText = reader.ReadToEnd();
                    FileWriter.CreateDirectory(configFilePath);
                    FileWriter.WriteTextToFile(tnsFileText, configFileName, configFilePath);
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseTnsnamesFile(string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public void GenerateEllipseTnsnamesFile(string sourceUrl, string targetUrl)
        {
            try
            {
                //iniciamos las variables de directorio y archivo
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var defaultFilePath = FileWriter.NormalizePath(CurrentSettings.DefaultTnsnamesFilePath, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                var configFileName = CurrentSettings.TnsnamesFileName;

                //comprobamos que la ruta no corresponda a la ruta predeterminada
                if (configFilePath.Equals(defaultFilePath))
                    throw new Exception("No se puede reemplazar el archivo " + configFileName + " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");

                //creamos una copia de seguridad si el archivo existe
                if (FileWriter.CheckFileExist(Path.Combine(configFilePath, configFileName)))
                {
                    var backupFileName = configFileName + "_" + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK";
                    if (!FileWriter.CheckFileExist(Path.Combine(configFilePath, backupFileName)))
                        FileWriter.MoveFileToDirectory(configFileName, configFilePath,backupFileName, configFilePath);
                }

                //realizamos la acción
                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseTnsnamesFile(string, string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public void GenerateEllipseDatabaseFile(string targetUrl = null)
        {
            var databaseList = new List<DatabaseItem>();
            databaseList.Add(new DatabaseItem("Productivo", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("Contingencia", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("Desarrollo", "EL8DESA", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("Test", "EL8TEST", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("ellprod", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("ellcont", "EL8PROD", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("elldesa", "EL8DESA", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("elltest", "EL8TEST", "SIGCON", "ventyx", "ELLIPSE", null, null));
            databaseList.Add(new DatabaseItem("SCADARDB", "PBVFWL01", "SCADARDBADMINGUI", "momia2011", "SCADARDB.DBO",
                null, "SCADARDB"));
            databaseList.Add(new DatabaseItem("SIGCOR", "SIGCOPRD", "CONSULBO", "consulbo", "@DBLELLIPSE8", "ELLIPSE", null));
            databaseList.Add(
                new DatabaseItem("SIGCOPRD", "SIGCOPRD", "CONSULBO", "consulbo", "@DBLELLIPSE8", "ELLIPSE", null));

            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<ellipse>";
            xmlFile += @"  <connections>";
            foreach (var item in databaseList)
                xmlFile += @"    <" + item.Name + " dbname='" + item.DbName + "' dbuser='" + item.DbUser +
                           "' dbpassword='' dbencodedpassword='" + item.DbEncodedPassword + "' dbreference='" +
                           item.DbReference + "' dblink='" + item.DbLink + "' " +
                           (string.IsNullOrWhiteSpace(item.DbCatalog) ? null : "dbcatalog='" + item.DbCatalog + "'") +
                           "/>";
            xmlFile += @"  </connections>";
            xmlFile += @"</ellipse>";

            try
            {
                //iniciamos las variables de directorio y archivo
                if (targetUrl == null)
                    targetUrl = CurrentSettings.LocalDataPath;
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var configFileName = CurrentSettings.DatabaseXmlFileName;
                
                //creamos una copia de seguridad si el archivo existe
                if (FileWriter.CheckFileExist(Path.Combine(configFilePath, configFileName)))
                {
                    var backupFileName = configFileName + "_" + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK";
                    if (!FileWriter.CheckFileExist(Path.Combine(configFilePath, backupFileName)))
                        FileWriter.MoveFileToDirectory(configFileName, configFilePath, backupFileName, configFilePath);
                }

                //realizamos la acción
                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseDatabaseFile",
                    "No se puede crear el archivo de bases de datos\n" + ex.Message);
                throw;
            }
        }

        public void DeleteEllipseDatabaseFile(string targetUrl = null)
        {
            try
            {
                //iniciamos las variables de directorio y archivo
                if (targetUrl == null)
                    targetUrl = CurrentSettings.LocalDataPath;
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var configFileName = CurrentSettings.DatabaseXmlFileName;
                
                FileWriter.DeleteFile(configFilePath, configFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración", ex.Message);
                throw;
            }
        }

        #endregion
    }

    public class RuntimeConfigSettings : SharedClassLibrary.Configuration.RuntimeConfigSettings
    {

    }
}

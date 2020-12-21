using System;
using System.Collections.Generic;
using System.IO;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using System.Reflection;
using System.Windows.Forms;
using SharedClassLibrary.Configuration;
using SharedClassLibrary.Connections;

// ReSharper disable AccessToStaticMemberViaDerivedType

namespace EllipseCommonsClassLibrary
{
    public class Options : SharedClassLibrary.Configuration.Options
    {

    }
    public class Settings : SharedClassLibrary.Configuration.ISettings
    {
        public static Settings CurrentSettings;
        public Settings()
        {
            Initialize();
        }

        public void Initialize()
        {
            var lastAssembly = SharedClassLibrary.Configuration.AssemblyItem.GetLastAssembly();
            ProgramTitle = new SharedClassLibrary.Configuration.AssemblyItem(lastAssembly).AssemblyTitle;
            //GeneralFolder
            DefaultLocalDataPath = @"c:\ellipse\";
            GeneralConfigFolder = @"addins\" + ProgramTitle;
            GeneralConfigFileName = "config.xml";
            DefaultRepositoryFilePath = @"\\lmnoas02\Shared\Sistemas\Mina\Proyecto Ellipse\Ellipse 8\ExcelAddIn_E8 (Loaders)\";

            //Windows Environment Variables
            ProgramEnvironmentHomeVariable = ProgramTitle + "Home";
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
            //Option Settings
            OptionsSettings = GetOptionsSettings();
            if (_optionsSettings == null)
                _optionsSettings = new EllipseCommonsClassLibrary.Options();
        }

        public string DefaultRepositoryFilePath { get; set; }
        public string HomeEnvironmentVariable { get; set; }
        public string ServicesEnvironmentVariable { get; set; }
        public string SecondaryServicesEnvironmentVariable { get; set; }
        public string ServicesForcedList { get; set; }
        public string ServicesConfigXmlFileName { get; set; }
        public string TnsnamesFileName { get; set; }
        public string DatabaseXmlFileName { get; set; }
        public string DefaultServiceFilePath { get; set; }
        public string SecondaryServiceFilePath { get; set; }
        public string DefaultTnsnamesFilePath { get; set; }
        public string DefaultLocalDataPath { get; set; }
        public string ProgramEnvironmentHomeVariable { get; set; }
        public string ProgramTitle { get; set; }
        public string GeneralConfigFileName { get; set; }
        public string GeneralConfigFolder { get; set; }

        private IOptions _optionsSettings;

        #region -- SettingOptions Methods --
        public IOptions OptionsSettings
        {
            get
            {
                try
                {
                    if (_optionsSettings != null) return _optionsSettings;

                    var path = LocalDataPath;
                    var option = (Options)Utilities.MyUtilities.Xml.DeserializeXmlToObject(Path.Combine(LocalDataPath, GeneralConfigFolder, GeneralConfigFileName), typeof(Options));
                    _optionsSettings = option;
                    return _optionsSettings;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Se ha producido un error al intentar cargar la configuración de " + ProgramTitle + ". Se continuará con la configuración predeterminada si esta existe. " + ex.Message, "Error a cargar Opciones de Configuración");
                    return null;
                }
            }
            set => UpdateOptionsSettings(value);
        }
        private IOptions GetOptionsSettings()
        {
            try
            {
                if (_optionsSettings != null) return _optionsSettings;

                var path = LocalDataPath;
                var option = (Options)Utilities.MyUtilities.Xml.DeserializeXmlToObject(Path.Combine(LocalDataPath, GeneralConfigFolder, GeneralConfigFileName), typeof(Options));

                return option;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se ha producido un error al intentar cargar la configuración de " + ProgramTitle + ". Se continuará con la configuración predeterminada. " + ex.Message, "Error a cargar Opciones de Configuración");
                return OptionsSettings?.DefaultOptions != null ? OptionsSettings.DefaultOptions : null;
            }
        }
        public void SetDefaultOptionsSettings(IOptions defaultProgramOptions)
        {
            if (OptionsSettings == null)
                OptionsSettings = defaultProgramOptions;
        }
        public IOptions CreateOptionsSettingFile(IOptions optionsSettings = null)
        {
            // Serialize the configuration object to a file
            return UpdateOptionsSettings(optionsSettings);
        }

        public void DeleteConfigurationXmlFile()
        {
            throw new NotImplementedException();
        }

        public void DeleteDatabaseFile()
        {
            throw new NotImplementedException();
        }

        public IOptions UpdateOptionsSettings(IOptions optionsSettings = null)
        {
            try
            {
                if (optionsSettings != null)
                    _optionsSettings = optionsSettings;

                FileWriter.CreateDirectory(Path.Combine(LocalDataPath, GeneralConfigFolder));

                if (_optionsSettings?.OptionsList != null)
                    MyUtilities.Xml.SerializeObjectToXml(Path.Combine(LocalDataPath, GeneralConfigFolder, GeneralConfigFileName), _optionsSettings);
                else
                    throw new Exception("No hay opciones disponibles");

                return _optionsSettings;
            }
            catch (Exception ex)
            {
                Debugger.LogError("UpdateOptionsSettings(IOptions)",
                    "Error al intentar actualizar las opciones. \n" + ex.Message);
                throw;
            }
        }

        #endregion

        #region -- Variable Accessors --
        public bool IsServiceListForced
        {
            get
            {
                var varForced =
                    "" + Environment.GetEnvironmentVariable(ServicesForcedList, EnvironmentVariableTarget.User);
                var varForcedExpanded = Environment.ExpandEnvironmentVariables(varForced);
                return !string.IsNullOrWhiteSpace(varForcedExpanded) && varForcedExpanded.ToLower().Equals("true");
            }
            set
            {
                Environment.SetEnvironmentVariable(ServicesForcedList, value.ToString(),
                    EnvironmentVariableTarget.User);
            }
        }

        public string LocalDataPath
        {
            get
            {
                var varHome = "" + Environment.GetEnvironmentVariable(HomeEnvironmentVariable,
                                  EnvironmentVariableTarget.User);
                var varHomeExpanded = Environment.ExpandEnvironmentVariables(varHome);
                return string.IsNullOrWhiteSpace(varHomeExpanded) ? DefaultLocalDataPath : varHomeExpanded;
            }
            set
            {
                var currentVar = Environment.GetEnvironmentVariable(HomeEnvironmentVariable, EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(DefaultLocalDataPath))
                    return;

                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;

                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(DefaultLocalDataPath))
                    Environment.SetEnvironmentVariable(HomeEnvironmentVariable, value, EnvironmentVariableTarget.User);

                //existe y es diferente a environment -> actualiza
                else if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(HomeEnvironmentVariable, value, EnvironmentVariableTarget.User);
            }
        }

        public string BackUpServiceFilePath
        {
            get
            {
                var varService = "" + Environment.GetEnvironmentVariable(SecondaryServicesEnvironmentVariable,
                                     EnvironmentVariableTarget.User);
                var varServiceExpanded = Environment.ExpandEnvironmentVariables(varService);
                return string.IsNullOrWhiteSpace(varServiceExpanded) ? SecondaryServiceFilePath : varServiceExpanded;
            }
            set
            {
                var currentVar = Environment.GetEnvironmentVariable(SecondaryServicesEnvironmentVariable,
                    EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(SecondaryServiceFilePath))
                    return;
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(SecondaryServiceFilePath))
                    Environment.SetEnvironmentVariable(SecondaryServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
                //existe y es diferente a environment -> actualiza
                else if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(SecondaryServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
            }
        }

        public string ServiceFilePath
        {
            get
            {
                var varService = "" + Environment.GetEnvironmentVariable(ServicesEnvironmentVariable,
                                     EnvironmentVariableTarget.User);
                var varServiceExpanded = Environment.ExpandEnvironmentVariables(varService);
                return string.IsNullOrWhiteSpace(varServiceExpanded) ? DefaultServiceFilePath : varServiceExpanded;
            }
            set
            {
                var currentVar = Environment.GetEnvironmentVariable(ServicesEnvironmentVariable,
                    EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(DefaultServiceFilePath))
                    return;
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(DefaultServiceFilePath))
                    Environment.SetEnvironmentVariable(ServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
                //existe y es diferente a environment -> actualiza
                else if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(ServicesEnvironmentVariable, value, EnvironmentVariableTarget.User);
            }
        }

        public string TnsnamesFilePath
        {
            get { return RuntimeConfigSettings.GetTnsUrlValue(); }
            set
            {
                if (value.Equals(RuntimeConfigSettings.GetTnsUrlValue()))
                    return;
                RuntimeConfigSettings.UpdateTnsUrlValue(value);
            }
        }
        #endregion
        #region -- Configuration Files Generation --
        public void GenerateConfigurationXmlFile(string targetUrl)
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

        public void GenerateConfigurationXmlFile(string sourceUrl, string targetUrl)
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

        public void GenerateDatabaseFile()
        {
            throw new NotImplementedException();
        }

        public void GenerateTnsnamesFile(string targetUrl)
        {
            throw new NotImplementedException();
        }

        public void DeletConfigurationXmlFile()
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

        public void GenerateTnsnamesFile(string sourceUrl, string targetUrl)
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

        public void GenerateDatabaseFile(string targetUrl = null)
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

        public void DeleteDatabaseFile(string targetUrl = null)
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

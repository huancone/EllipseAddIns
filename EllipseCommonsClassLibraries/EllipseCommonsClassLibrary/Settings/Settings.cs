using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using EllipseCommonsClassLibrary.Utilities;

namespace CommonsClassLibrary
{
    public partial class Settings
    {
        public static Settings CurrentSettings;
        public string DefaultRepositoryFilePath;

        public string HomeEnvironmentVariable;
        public string ServicesEnvironmentVariable;
        public string SecondaryServicesEnvironmentVariable;
        public string ServicesForcedList;

        public string ServicesConfigXmlFileName;
        public string TnsnamesFileName;
        public string DatabaseXmlFileName;

        public string DefaultServiceFilePath;
        public string SecondaryServiceFilePath;
        public string DefaultTnsnamesFilePath;
        public string DefaultLocalDataPath;

        public string ProgramEnvironmentHomeVariable;
        public AssemblyItem AssemblyProgram;
        public string GeneralConfigFileName;
        public string GeneralConfigFolder;
        public Configuration Config;
        public Options OptionsSettings;

        public Settings()
        {
        }
        public Settings(Options defaultProgramOptions)
        {
            Initialize(defaultProgramOptions);
        }

        public void Initialize(Options defaultProgramOptions)
        {
            AssemblyProgram = new Settings.AssemblyItem(GetLastAssembly());
            //GeneralFolder
            DefaultLocalDataPath = @"c:\project\";
            GeneralConfigFolder = @"apps\" + AssemblyProgram.AssemblyTitle;
            GeneralConfigFileName = "settings.xml";
            DefaultRepositoryFilePath = @"\\project\repository\";

            //Windows Environment Variables
            ProgramEnvironmentHomeVariable = AssemblyProgram.AssemblyTitle + "Home";
            HomeEnvironmentVariable = "ProjectHome";
            ServicesEnvironmentVariable = "ProjectServiceUrlFile";
            SecondaryServicesEnvironmentVariable = "ProjectSecondaryServiceUrlFile";

            //Services & Databases Information
            ServicesForcedList = "ProjectServiceForcedList";
            ServicesConfigXmlFileName = "ProjectServices.xml";
            TnsnamesFileName = "tnsnames.ora";
            DatabaseXmlFileName = "ProjectDatabases.xml";

            DefaultServiceFilePath = @"\\project\";
            SecondaryServiceFilePath = @"\\project\";
            DefaultTnsnamesFilePath = @"c:\project\network\ADMIN\";

            //StaticReference Through All the Project
            Debugger.LocalDataPath = LocalDataPath;
            CurrentSettings = this;
            //Options
            OptionsSettings = GetOptionsSettings(defaultProgramOptions);
        }



        #region -- SettingOptions Methods & Class --
        public Options CreateOptionsSettingFile(Options optionsSettings = null)
        {
            // Serialize the configuration object to a file
            return UpdateOptionsSettings(optionsSettings);
        }
        
        public Options UpdateOptionsSettings(Options optionsSettings = null)
        {
            if (optionsSettings == null)
                optionsSettings = OptionsSettings;

            CommonsClassLibrary.Utilities.FileWriter.CreateDirectory(Path.Combine(LocalDataPath, GeneralConfigFolder));

            if (optionsSettings != null && optionsSettings.OptionsList != null)
                Utilities.MyUtilities.Xml.SerializeObjectToXml(Path.Combine(LocalDataPath, GeneralConfigFolder, GeneralConfigFileName), optionsSettings);
            return optionsSettings;
        }
        public Options GetOptionsSettings(Options defaultOptionsSettings)
        {
            try
            {
                var path = LocalDataPath;
                var option = (Options)Utilities.MyUtilities.Xml.DeserializeXmlToObject(Path.Combine(LocalDataPath, GeneralConfigFolder, GeneralConfigFileName), typeof(Options));
                option.SetDefaultOptions(defaultOptionsSettings);
                return option;
            }
            catch (DirectoryNotFoundException)
            {
                return UpdateOptionsSettings(defaultOptionsSettings);
            }
            catch (FileNotFoundException)
            {
                return UpdateOptionsSettings(defaultOptionsSettings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se ha producido un error al intentar cargar la configuración de " + AssemblyProgram.AssemblyTitle + ". Se continuará con la configuración predeterminada. " + ex.Message, "Error a cargar Opciones de Configuración");
                return defaultOptionsSettings;
            }
        }

        /// <summary>
        /// This Configuration class is basically just a set of 
        /// properties with a couple of static methods to manage
        /// the serialization to and deserialization from a
        /// simple XML file.
        /// </summary>
        [Serializable]
        public class Options
        {
            public List<ConfigValuePair<string, string>> OptionsList;
            private Options _defaultSettingOptions;
            public void SetOption(string key, string value)
            {
                if (OptionsList == null)
                    OptionsList = new List<ConfigValuePair<string, string>>();


                //Valido primeramente si existe y lo reemplazo si existe
                foreach (var item in OptionsList)
                {
                    if (item.Key != null && item.Key.Equals(key))
                    {
                        var newItem = new ConfigValuePair<string, string>(item.Key, value);
                        OptionsList[OptionsList.IndexOf(item)] = newItem;
                        return;
                    }
                }

                //Lo agrego si no existe
                OptionsList.Add(new ConfigValuePair<string, string>(key, value));
            }

            public void SetDefaultOptions(Options defaultSettingOptions)
            {
                _defaultSettingOptions = defaultSettingOptions;
            }
            public string GetOptionValue(ConfigValuePair<string, string> defaultOptionItem)
            {
                var currentItemValue = GetOptionValue(defaultOptionItem.Key);
                if (!string.IsNullOrWhiteSpace(currentItemValue))
                    return currentItemValue;

                SetOption(defaultOptionItem.Key, defaultOptionItem.Value);
                return defaultOptionItem.Value;
            }

            public string  GetOptionValue(string key)
            {
                foreach (var item in OptionsList)
                    if (item.Key.Equals(key))
                        return item.Value;

                if (_defaultSettingOptions != null)
                {
                    var defaultItem = _defaultSettingOptions.GetOption(key);
                    if (defaultItem.Key != null)
                    {
                        SetOption(defaultItem.Key, defaultItem.Value);
                        return defaultItem.Value;
                    }
                }

                return null;
            }

            public ConfigValuePair<string, string> GetOption(string key)
            {
                foreach (var item in OptionsList)
                    if (item.Key.Equals(key))
                        return item;
                if (_defaultSettingOptions != null)
                {
                    var defaultItem = _defaultSettingOptions.GetOption(key);
                    if (defaultItem.Key != null)
                    {
                        SetOption(defaultItem.Key, defaultItem.Value);
                        return defaultItem;
                    }
                }

                return new ConfigValuePair<string, string>();
            }
        }

        [Serializable]
        [XmlType(TypeName = "optionItem")]
        public struct ConfigValuePair<TK, TV>
        {
            public ConfigValuePair(TK key, TV value)
            {
                Key = key;
                Value = value;
            }
            public TK Key
            { get; set; }

            public TV Value
            { get; set; }
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

        #region -- Configuration Files Generator --
        public void GenerateConfigurationXmlFile(string targetUrl)
        {
            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<services>";
            xmlFile += @"  <url>";
            xmlFile += @"    <servername> http://url/ </servername>";
            xmlFile += @"  </url>";
            xmlFile += @"  <webservice>";
            xmlFile += @"    <servername> http://url/ </servername>";
            xmlFile += @"  </webservice>";
            xmlFile += @"</ellipse>";

            try
            {
                if (FileWriter.NormalizePath(targetUrl, true)
                    .Equals(FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true)))
                    throw new Exception(
                        "No se puede reemplazar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var configFileName = CurrentSettings.ServicesConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateConfigurationXmlFile(string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public void GenerateConfigurationXmlFile(string sourceUrl, string targetUrl)
        {
            try
            {
                if (FileWriter.NormalizePath(targetUrl, true)
                    .Equals(FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true)))
                    throw new Exception(
                        "No se puede reemplazar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                var configFileName = CurrentSettings.ServicesConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateConfigurationXmlFile(string, string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }
        public void DeleteConfigurationXmlFile()
        {
            try
            {
                if (FileWriter.NormalizePath(CurrentSettings.ServiceFilePath, true)
                    .Equals(FileWriter.NormalizePath(CurrentSettings.DefaultServiceFilePath, true)))
                    throw new Exception(
                        "No se puede eliminar el archivo de configuración original del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(CurrentSettings.ServiceFilePath, true);
                FileWriter.DeleteFile(configFilePath, CurrentSettings.ServicesConfigXmlFileName);
                CurrentSettings.ServiceFilePath = CurrentSettings.DefaultServiceFilePath;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DeleteConfigurationXmlFile()", "No se puede eliminar el archivo de configuración. " + ex.Message);
                throw;
            }
        }
        public void GenerateTnsnamesFile(string targetUrl)
        {
            try
            {
                if (FileWriter.NormalizePath(targetUrl, true)
                    .Equals(FileWriter.NormalizePath(CurrentSettings.DefaultTnsnamesFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo " + CurrentSettings.TnsnamesFileName +
                                        " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                if (File.Exists(CurrentSettings.TnsnamesFilePath + CurrentSettings.TnsnamesFileName))
                    FileWriter.MoveFileToDirectory(CurrentSettings.TnsnamesFileName, CurrentSettings.TnsnamesFilePath,
                        CurrentSettings.TnsnamesFileName + System.DateTime.Today.Year + System.DateTime.Today.Month + System.DateTime.Today.Day + ".BAK",
                        CurrentSettings.TnsnamesFilePath);
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var configFileName = CurrentSettings.TnsnamesFileName;

                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "CommonsClassLibrary.Resources.tnsnames.txt";
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
                Debugger.LogError("GenerateTnsnamesFile(string, string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }
        public void GenerateTnsnamesFile(string sourceUrl, string targetUrl)
        {
            try
            {
                if (FileWriter.NormalizePath(targetUrl, true)
                    .Equals(FileWriter.NormalizePath(CurrentSettings.DefaultTnsnamesFilePath, true)))
                    throw new Exception("No se puede reemplazar el archivo " + CurrentSettings.TnsnamesFileName +
                                        " del sistema. Si desea modificarlo, comuníquese con el administrador del sistema");
                var configFilePath = FileWriter.NormalizePath(targetUrl, true);
                var sourceFilePath = FileWriter.NormalizePath(sourceUrl, true);
                var configFileName = CurrentSettings.TnsnamesFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, sourceFilePath, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateTnsnamesFile(string, string)",
                    "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }
        public void GenerateDatabaseFile()
        {
            var databaseList = new List<Connections.DatabaseItem>();
            databaseList.Add(new Connections.DatabaseItem("Name", "DbName", "DbUser", "DbPassword", "DbReference", "DbLink", "DbCatalog"));

            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<databases>";
            xmlFile += @"  <connections>";
            foreach (var item in databaseList)
                xmlFile += @"    <" + item.Name + " dbname='" + item.DbName + "' dbuser='" + item.DbUser +
                           "' dbpassword='' dbencodedpassword='" + item.DbEncodedPassword + "' dbreference='" +
                           item.DbReference + "' dblink='" + item.DbLink + "' " +
                           (string.IsNullOrWhiteSpace(item.DbCatalog) ? null : "dbcatalog='" + item.DbCatalog + "'") +
                           "/>";
            xmlFile += @"  </connections>";
            xmlFile += @"</databases>";

            try
            {
                var configFilePath = FileWriter.NormalizePath(CurrentSettings.LocalDataPath, true);
                var configFileName = CurrentSettings.DatabaseXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateDatabaseFile",
                    "No se puede crear el archivo de bases de datos\n" + ex.Message);
                throw;
            }
        }
        public void DeleteDatabaseFile()
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(CurrentSettings.ServiceFilePath, true);
                FileWriter.DeleteFile(configFilePath, CurrentSettings.DatabaseXmlFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración", ex.Message);
                throw;
            }
        }
        #endregion
    }

   


}


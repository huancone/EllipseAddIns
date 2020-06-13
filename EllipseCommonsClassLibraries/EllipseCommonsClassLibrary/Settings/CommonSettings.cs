using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using EllipseCommonsClassLibrary;

namespace EllipseCommonsClassLibrary.Settings
{
    public class CommonSettings
    {

        private string _defaultLocalDataPath;
        private string _addinEnvironmentHomeVariable;
        private AssemblyItem _assembly;
        public string FileName = "config.xml";
        public Options Configuration;
        public Options DefaultConfiguration;

        public CommonSettings()
        {
            DefaultConfiguration = null;
            _assembly = new AssemblyItem(Assembly.GetCallingAssembly());
            _defaultLocalDataPath = EllipseCommonsClassLibrary.Connections.Configuration.DefaultLocalDataPath + @"addins\" + _assembly.AssemblyTitle;
            _addinEnvironmentHomeVariable = _assembly.AssemblyTitle + "Home";
            Configuration = GetSettings(DefaultConfiguration);
        }

        public CommonSettings(Options defaultConfiguration)
        {
            DefaultConfiguration = defaultConfiguration;
            _assembly = new AssemblyItem(Assembly.GetCallingAssembly());
            _defaultLocalDataPath = EllipseCommonsClassLibrary.Connections.Configuration.DefaultLocalDataPath + @"addins\" + _assembly.AssemblyTitle;
            _addinEnvironmentHomeVariable = _assembly.AssemblyTitle + "Home";
            Configuration = GetSettings(DefaultConfiguration);
        }

        public Options CreateSettingsFile()
        {
            // Serialize the configuration object to a file
            return UpdateSettings(Configuration);
        }
        public Options CreateSettingsFile(Options configOptions)
        {
            // Serialize the configuration object to a file
            return UpdateSettings(configOptions);
        }

        public Options UpdateSettings()
        {
            string path = LocalDataPath;
            EllipseCommonsClassLibrary.Utilities.FileWriter.CreateDirectory(path);
            Serialize(path + @"\config.xml", Configuration);
            return Configuration;
        }
        public Options UpdateSettings(Options options)
        {
            string path = LocalDataPath;
            EllipseCommonsClassLibrary.Utilities.FileWriter.CreateDirectory(path);
            Serialize(path + @"\config.xml", options);
            return options;
        }
        public Options GetSettings(Options defaultOptions)
        {
            try
            {
                string path = LocalDataPath;
                var c = Deserialize(path + @"\config.xml");
                return c;
            }
            catch (DirectoryNotFoundException)
            {
                return UpdateSettings(defaultOptions);
            }
            catch (FileNotFoundException)
            {
                return UpdateSettings(defaultOptions);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se ha producido un error al intentar cargar la configuración de " + _assembly.AssemblyTitle + ". Se continuará con la configuración predeterminada. " + ex.Message, "Error a cargar Opciones de Configuración");
                return defaultOptions;
            }
        }
        static void Serialize(string file, Options options)
        {
            System.Xml.Serialization.XmlSerializer xs
               = new System.Xml.Serialization.XmlSerializer(options.GetType());
            StreamWriter writer = File.CreateText(file);
            xs.Serialize(writer, options);
            writer.Flush();
            writer.Close();
        }
        static Options Deserialize(string file)
        {
            System.Xml.Serialization.XmlSerializer xs
               = new System.Xml.Serialization.XmlSerializer(
                  typeof(Options));
            StreamReader reader = File.OpenText(file);
            Options options = (Options)xs.Deserialize(reader);
            reader.Close();
            return options;
        }

        #region -- Configuration Class --
        /// <summary>
        /// This Configuration class is basically just a set of 
        /// properties with a couple of static methods to manage
        /// the serialization to and deserialization from a
        /// simple XML file.
        /// </summary>
        [Serializable]
        public class Options
        {
            public List<ConfigValuePair<string, string>> Settings;
            private Options _defaultOptions;
            public Options(string fileName)
            {
                Serialize(fileName, this);
            }

            public Options()
            {

            }
            public void SetOption(string key, string value)
            {
                if (Settings == null)
                    Settings = new List<ConfigValuePair<string, string>>();


                //Valido primeramente si existe y lo reemplazo si existe
                foreach (var item in Settings)
                {
                    if (item.Key != null && item.Key.Equals(key))
                    {
                        var newItem = new ConfigValuePair<string, string>(item.Key, value);
                        Settings[Settings.IndexOf(item)] = newItem;
                        return;
                    }
                }

                //Lo agrego si no existe
                Settings.Add(new ConfigValuePair<string, string>(key, value));
            }

            public void SetDefaultOptions(Options defaultOptions)
            {
                _defaultOptions = defaultOptions;
            }
            public string GetOptionValue(ConfigValuePair<string, string> defaultItem)
            {
                var currentItemValue = GetOptionValue(defaultItem.Key);
                if (!string.IsNullOrWhiteSpace(currentItemValue))
                    return currentItemValue;
                SetOption(defaultItem.Key, defaultItem.Value);
                return defaultItem.Value;
            }

            public string  GetOptionValue(string key)
            {
                foreach (var item in Settings)
                    if (item.Key.Equals(key))
                        return item.Value;

                if (_defaultOptions != null)
                {

                    var defaultItem = _defaultOptions.GetOption(key);
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
                foreach (var item in Settings)
                    if (item.Key.Equals(key))
                        return item;
                if (_defaultOptions != null)
                {
                    var defaultItem = _defaultOptions.GetOption(key);
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

        #region -- EnviromentVariables --
        public string LocalDataPath
        {
            get
            {
                var varHome = "" + Environment.GetEnvironmentVariable(_addinEnvironmentHomeVariable,
                                  EnvironmentVariableTarget.User);
                var varHomeExpanded = Environment.ExpandEnvironmentVariables(varHome);
                return string.IsNullOrWhiteSpace(varHomeExpanded) ? _defaultLocalDataPath : varHomeExpanded;
            }
            set
            {
                var currentVar =
                    Environment.GetEnvironmentVariable(_addinEnvironmentHomeVariable, EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(_defaultLocalDataPath))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(_defaultLocalDataPath))
                    Environment.SetEnvironmentVariable(_addinEnvironmentHomeVariable, value,
                        EnvironmentVariableTarget.User);
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //existe y es diferente a environment -> actualiza
                if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(_addinEnvironmentHomeVariable, value,
                        EnvironmentVariableTarget.User);
            }
        }
        #endregion
    }
}

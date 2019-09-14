using System;
using System.Configuration;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.Assurance
{
    public static class AssuranceSettings
    {

        public static string DefaultLocalDataPath = EllipseCommonsClassLibrary.Connections.Configuration.DefaultLocalDataPath + @"addins\EllipseRequisitionServiceExcelAddin";
        public static string AddinEnvironmentHomeVariable = "EllipseRequisitionServiceExcelAddInHome";

        public static Configuration CreateSettingsFile()
        {
            // Create a new configuration object
            // and initialize some variables
            Configuration c = new Configuration
            {
                costCenterMatch = true,
                minItemValue = 100
            };

            // Serialize the configuration object to a file
            return UpdateSettings(c);
        }
        public static Configuration UpdateSettings(Configuration c)
        {
            string path = LocalDataPath;
            EllipseCommonsClassLibrary.Utilities.FileWriter.CreateDirectory(path);
            Serialize(path + @"\config.xml", c);
            return c;
        }
        public static Configuration GetSettings()
        {
            string path = LocalDataPath;
            var c = Deserialize(path + @"\config.xml");
            return c;
        }
        static void Serialize(string file, Configuration c)
        {
            System.Xml.Serialization.XmlSerializer xs
               = new System.Xml.Serialization.XmlSerializer(c.GetType());
            StreamWriter writer = File.CreateText(file);
            xs.Serialize(writer, c);
            writer.Flush();
            writer.Close();
        }
        static Configuration Deserialize(string file)
        {
            System.Xml.Serialization.XmlSerializer xs
               = new System.Xml.Serialization.XmlSerializer(
                  typeof(Configuration));
            StreamReader reader = File.OpenText(file);
            Configuration c = (Configuration)xs.Deserialize(reader);
            reader.Close();
            return c;
        }

        #region -- Configuration Class --
        /// <summary>
        /// This Configuration class is basically just a set of 
        /// properties with a couple of static methods to manage
        /// the serialization to and deserialization from a
        /// simple XML file.
        /// </summary>
        [Serializable]
        public class Configuration
        {
            bool _costCenterMatch;
            int _minItemValue;

            public Configuration()
            {
                Serialize("config.xml", this);
            }


            public int minItemValue
            {
                get { return _minItemValue; }
                set { _minItemValue = value; }
            }
            public bool costCenterMatch
            {
                get { return _costCenterMatch; }
                set { _costCenterMatch = value; }
            }
        }
        #endregion

        #region -- EnviromentVariables --
        public static string LocalDataPath
        {
            get
            {
                var varHome = "" + Environment.GetEnvironmentVariable(AddinEnvironmentHomeVariable,
                                  EnvironmentVariableTarget.User);
                var varHomeExpanded = Environment.ExpandEnvironmentVariables(varHome);
                return string.IsNullOrWhiteSpace(varHomeExpanded) ? DefaultLocalDataPath : varHomeExpanded;
            }
            set
            {
                var currentVar =
                    Environment.GetEnvironmentVariable(AddinEnvironmentHomeVariable, EnvironmentVariableTarget.User);
                //no existe y es igual a _origen -> no hace nada
                if (string.IsNullOrWhiteSpace(currentVar) && value.Equals(DefaultLocalDataPath))
                    return;
                //no existe y es diferente a _origen -> actualiza
                if (string.IsNullOrWhiteSpace(currentVar) && !value.Equals(DefaultLocalDataPath))
                    Environment.SetEnvironmentVariable(AddinEnvironmentHomeVariable, value,
                        EnvironmentVariableTarget.User);
                //existe y es igual a environment -> no hace nada
                if (!string.IsNullOrWhiteSpace(currentVar) && value.Equals(currentVar))
                    return;
                //existe y es diferente a environment -> actualiza
                if (!string.IsNullOrWhiteSpace(currentVar) && !value.Equals(currentVar))
                    Environment.SetEnvironmentVariable(AddinEnvironmentHomeVariable, value,
                        EnvironmentVariableTarget.User);
            }
        }
        #endregion
    }
}

namespace SharedClassLibrary.Configuration
{
    public interface ISettings
    {
        void Initialize();
        string DefaultRepositoryFilePath { get; set; }
        string HomeEnvironmentVariable { get; set; }
        string ServicesEnvironmentVariable { get; set; }
        string SecondaryServicesEnvironmentVariable { get; set; }
        string ServicesForcedList { get; set; }

        string ServicesConfigXmlFileName { get; set; }
        string TnsnamesFileName { get; set; }
        string DatabaseXmlFileName { get; set; }

        string DefaultServiceFilePath { get; set; }
        string SecondaryServiceFilePath { get; set; }
        string DefaultTnsnamesFilePath { get; set; }
        string DefaultLocalDataPath { get; set; }

        string ProgramEnvironmentHomeVariable { get; set; }
        string ProgramTitle { get; set; }
        string GeneralConfigFileName { get; set; }
        string GeneralConfigFolder { get; set; }
        string BackUpServiceFilePath { get; set; }
        bool IsServiceListForced { get; set; }
        string LocalDataPath { get; set; }
        string ServiceFilePath { get; set; }
        string TnsnamesFilePath { get; set; }
        void DeleteConfigurationXmlFile();
        void DeleteDatabaseFile();
        void GenerateConfigurationXmlFile(string targetUrl);
        void GenerateConfigurationXmlFile(string sourceUrl, string targetUrl);
        void GenerateDatabaseFile();
        void GenerateTnsnamesFile(string targetUrl);
        void GenerateTnsnamesFile(string sourceUrl, string targetUrl);
        void SetDefaultCustomSettingValue(string key, string value);
        string GetDefaultCustomSettingValue(string key);
        void SetCustomSettingValue(string key, string value);
        string GetCustomSettingValue(string key);

        void SaveCustomSettings();
        void LoadCustomSettings();

    }
}
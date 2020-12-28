using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Configuration;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Configuration
{
    public class Settings : SharedClassLibrary.Configuration.Settings
    {
        public Settings()
        {
            Initialize(null);
        }
        public Settings(Options defaultProgramOptions)
        {
            Initialize(defaultProgramOptions);
        }
        public override void Initialize(Options defaultProgramOptions)
        {
            AssemblyProgram = new Settings.AssemblyItem(GetLastAssembly());
            //GeneralFolder
            DefaultLocalDataPath = @"c:\lsdatamodel\";
            GeneralConfigFolder = @"app\" + AssemblyProgram.AssemblyTitle;
            GeneralConfigFileName = "settings.xml";
            DefaultRepositoryFilePath = @"c:\lsdatamodel\";

            //Windows Environment Variables
            ProgramEnvironmentHomeVariable = AssemblyProgram.AssemblyTitle + "Home";
            HomeEnvironmentVariable = "LogsheetDatamodelHome";
            ServicesEnvironmentVariable = "LogsheetDatamodelServiceUrlFile";
            SecondaryServicesEnvironmentVariable = "LogsheetDatamodelSecondaryServiceUrlFile";

            //Services & Databases Information
            ServicesForcedList = "LogsheetDatamodelServiceForcedList";
            ServicesConfigXmlFileName = "LogsheetDatamodelServices.xml";
            TnsnamesFileName = "tnsnames.ora";
            DatabaseXmlFileName = "LogsheetDatamodelDatabases.xml";

            DefaultServiceFilePath = @"\\LogsheetDatamodel\";
            SecondaryServiceFilePath = @"\\LogsheetDatamodel\";
            DefaultTnsnamesFilePath = @"c:\LogsheetDatamodel\network\ADMIN\";

            //StaticReference Through All the Project
            Debugger.LocalDataPath = LocalDataPath;
            //Options
            if(defaultProgramOptions != null)
                OptionsSettings = GetOptionsSettings(defaultProgramOptions);
        }
    }
}

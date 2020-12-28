using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LogsheetDatamodelLibrary.Configuration
{
    public static class LsdmConfig
    {
        public static Settings Settings;
        public static DataSource DataSource;
        public static Login Login = new Login();

        public static bool AutoPullValues = true;

        public static class DatabaseInformation
        {
            public static int ModelIdLengthLimit = 20;
            public static int AttributeDescLengthLimit = 60;
            public static int VarcharLengthLimit = 20;
            public static int ModelDescLengthLimit = 120;
            public static int UsernameLengthLimit = 12;

            public static string TableValidationSources = "validation_sources";
            public static string TableMeasureTypes = "measure_types";
            public static string TableValidationItems = "validation_items";
            public static string TableDatamodels = "datamodels";
            public static string TableMeasures = "measures";
            public static string TableModelAttributes = "model_attributes";
            public static string TableDatasheets = "datasheets";
            public static string TableValueDatetimes = "value_datetimes";
            public static string TableValueNumerics = "value_numerics";
            public static string TableValueVarchars = "value_varchars";
            public static string TableValueTexts = "value_texts";
        }
    }
}

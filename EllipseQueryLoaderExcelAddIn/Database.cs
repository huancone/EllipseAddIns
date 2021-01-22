using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseQueryLoaderExcelAddIn
{
    public static class Database
    {
        public static string SqlDatabase = "SQL";
        public static string OracleDatabase = "ORACLE";

        public static class ParamType
        {
            public const string None = "";
            public const string Equal = "=";
            public const string InList = "IN";
            public const string GreatherThan = ">";
            public const string LessThan = "<";
            public const string GreatherEqualThan = ">=";
            public const string LessEqualThan = "<=";
            public const string DifferentThan = "<>";

            public static List<string> GetParamList()
            {
                var paramList = new List<string>
                {
                    None,
                    Equal,
                    InList,
                    GreatherThan,
                    LessThan,
                    GreatherEqualThan,
                    LessEqualThan,
                    DifferentThan
                };
                return paramList;
            }
        }
    }
}

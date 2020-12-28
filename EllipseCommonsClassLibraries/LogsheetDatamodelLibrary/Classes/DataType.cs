using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LogsheetDatamodelLibrary
{
    public static class DataTypes
    {
        public static string Date = "DATE";
        public static string DateTime = "DATETIME";
        public static string Numeric = "NUMERIC";
        public static string Varchar = "VARCHAR";
        public static string Text = "TEXT";

        public static List<string> GetList()
        {
            var list = new List<string>
            {
                Numeric,
                Varchar,
                Date,
                DateTime,
                Text
            };

            return list;
        }
    }
}

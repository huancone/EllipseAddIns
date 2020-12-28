using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Models
{
    public class ValidationSource : ISimpleObjectModelSql
    {
        public string Name;
        public string DbName;
        public string DbUser;
        public string DbPassword;
        public string DbReference;
        public string DbLink;
        public string PasswordEncodedType;

        public void SetFromDataRecord(IDataRecord dr)
        {
            Name = dr["source_name"].ToString().Trim();
            DbName = dr["db_name"].ToString().Trim();
            DbUser = dr["db_user"].ToString().Trim();
            DbPassword = dr["db_password"].ToString().Trim();
            DbLink = dr["db_link"].ToString().Trim();
            DbReference = dr["db_reference"].ToString().Trim();
            PasswordEncodedType = dr["password_encoded_type"].ToString().Trim();
        }

        public ValidationSource()
        {

        }

        public ValidationSource(IDataRecord dr)
        {
            SetFromDataRecord(dr);
        }
        public ValidationSource(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string passwordEncodedType)
        {
            Name = name;
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPassword;
            DbReference = dbReference;
            DbLink = dbLink;
            PasswordEncodedType = passwordEncodedType;
        }
        public bool Equals(ValidationSource validationSource)
        {
            if (!MyUtilities.EqualUpper(Name, validationSource.Name))
                return false;
            if (!MyUtilities.EqualUpper(DbName, validationSource.DbName))
                return false;
            if (!MyUtilities.EqualUpper(DbUser, validationSource.DbUser))
                return false;
            if (!MyUtilities.EqualUpper(DbPassword, validationSource.DbPassword))
                return false;
            if (!MyUtilities.EqualUpper(DbLink, validationSource.DbLink))
                return false;
            if (!MyUtilities.EqualUpper(DbReference, validationSource.DbReference))
                return false;
            if (!MyUtilities.EqualUpper(PasswordEncodedType, validationSource.PasswordEncodedType))
                return false;
            return true;
        }

        public static class EncryptionTypeValues
        {
            public static string Default = "DEFAULT";

            public static List<string> GetList()
            {
                var list = new List<string>
                {
                    Default
                };

                return list;
            }
        }
    }
}

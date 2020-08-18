using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseCommonsClassLibrary.Connections
{
    public class DatabaseItem
    {
        private string _dbEncodedPassword;
        private string _dbPassword;
        public string DbCatalog;
        public string DbLink;
        public string DbName;
        public string DbReference;
        public string DbUser;
        public string Name;
        public string SecondaryDbLink;
        public string SecondaryDbReference;

        public DatabaseItem(string dbName, string dbUser, string dbPassword)
        {
            SetDataBaseItem(dbName, dbUser, dbPassword, null, null, null);
        }

        public DatabaseItem(string dbName, string dbUser, string dbPassword, string dbReference,
            string dbLink)
        {
            SetDataBaseItem(dbName, dbUser, dbPassword, dbReference, dbLink, null);
        }

        public DatabaseItem(string dbName, string dbUser, string dbPassword, string dbReference,
            string dbLink, string dbCatalog)
        {
            SetDataBaseItem(dbName, dbUser, dbPassword, dbReference, dbLink, dbCatalog);
        }
        public DatabaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference,
            string dbLink, string dbCatalog)
        {
            SetDataBaseItem(name, dbName, dbUser, dbPassword, dbReference, dbLink, dbCatalog);
        }
        public DatabaseItem()
        {
        }

        public string DbPassword
        {
            get
            {
                return string.IsNullOrWhiteSpace(_dbPassword)
                    ? EncryptString.Decrypt(DbEncodedPassword, Configuration.EncryptPassPhrase)
                    : _dbPassword;
            }
            set
            {
                _dbPassword = value;
                if(_dbPassword != null)
                    _dbEncodedPassword = EncryptString.Encrypt(value, Configuration.EncryptPassPhrase);
            }
        }

        public string DbEncodedPassword
        {
            get { return _dbEncodedPassword; }
            set
            {
                _dbEncodedPassword = value;
                if (_dbEncodedPassword != null)
                    _dbPassword = EncryptString.Decrypt(value, Configuration.EncryptPassPhrase);
            }
        }
        private void SetDataBaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference,
            string dbLink, string dbCatalog)
        {
            Name = name;
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPassword;
            DbReference = dbReference;
            DbLink = dbLink;
            DbCatalog = dbCatalog;
        }
        private void SetDataBaseItem(string dbName, string dbUser, string dbPassword, string dbReference,
            string dbLink, string dbCatalog)
        {
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPassword;
            DbReference = dbReference;
            DbLink = dbLink;
            DbCatalog = dbCatalog;
        }
    }
}

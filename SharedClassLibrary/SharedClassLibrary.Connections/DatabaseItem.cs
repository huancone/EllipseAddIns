using SharedClassLibrary.Utilities;

//Shared Class Library - ExcelStyleCells
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

namespace SharedClassLibrary.Connections
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
                    ? Utilities.Encryption.Decrypt(DbEncodedPassword, Utilities.Encryption.EncryptPassPhrase)
                    : _dbPassword;
            }
            set
            {
                _dbPassword = value;
                if(_dbPassword != null)
                    _dbEncodedPassword = Utilities.Encryption.Encrypt(value, Utilities.Encryption.EncryptPassPhrase);
            }
        }

        public string DbEncodedPassword
        {
            get { return _dbEncodedPassword; }
            set
            {
                _dbEncodedPassword = value;
                if (_dbEncodedPassword != null)
                    _dbPassword = Encryption.Decrypt(value, Utilities.Encryption.EncryptPassPhrase);
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

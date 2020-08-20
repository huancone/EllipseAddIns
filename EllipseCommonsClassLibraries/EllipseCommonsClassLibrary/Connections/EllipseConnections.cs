using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseCommonsClassLibrary.Connections
{
    public class DatabaseItem : CommonsClassLibrary.Connections.DatabaseItem
    {
        #region constructors
        public DatabaseItem(string dbName, string dbUser, string dbPassword) : 
            base(dbName, dbUser, dbPassword)
        {
        }

        public DatabaseItem(string dbName, string dbUser, string dbPassword, string dbReference, string dbLink) : 
            base(dbName, dbUser, dbPassword, dbReference, dbLink)
        {
        }

        public DatabaseItem(string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string dbCatalog) : 
            base(dbName, dbUser, dbPassword, dbReference, dbLink, dbCatalog)
        {
        }
        public DatabaseItem(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string dbCatalog) :
            base(name, dbName, dbUser, dbPassword, dbReference, dbLink, dbCatalog)
        {
            
        }
        public DatabaseItem()
        {
        }
        #endregion
    }

    public class OracleConnector : CommonsClassLibrary.Connections.OracleConnector
    {
        public OracleConnector(string dbName, string dbUser, string dbPass) :base (dbName, dbUser, dbPass)
        {}
        public OracleConnector(string connectionString) :base(connectionString)
        {}

        public OracleConnector(DatabaseItem dbItem)
        {
            DbName = dbItem.DbName;
            DbUser = dbItem.DbUser;
            DbPassword = dbItem.DbPassword;
            DbLink = dbItem.DbLink;
            DbReference = dbItem.DbReference;
            DbCatalog = dbItem.DbCatalog;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection();
        }
    }
}

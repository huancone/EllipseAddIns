using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class ValidationSourceSqlAccess
    {
        public static int CreateValidationSource(OracleConnector connector, string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string passwordEncodedType)
        {
            return UpdateValidationSource(connector, name, dbName, dbUser, dbPassword, dbReference, dbLink, passwordEncodedType);
        }

        public static int UpdateValidationSource(OracleConnector connector, string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string passwordEncodedType)
        {
            var sqlQpc = GetUpdateValidationSourceQuery(connector.DbReference, connector.DbLink, name, dbName, dbUser, dbPassword, dbReference, dbLink, passwordEncodedType);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadValidationSource(OracleConnector connector, string name)
        {
            var sqlQpc = GetReadValidationSourceQuery(connector.DbReference, connector.DbLink, name);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadValidationSource(OracleConnector connector, string name, string keyword)
        {
            var sqlQpc = GetReadValidationSourceQuery(connector.DbReference, connector.DbLink, name, keyword);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadValidationSource(OracleConnector connector)
        {
            var sqlQpc = GetReadValidationSourceQuery(connector.DbReference, connector.DbLink);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int DeleteValidationSource(OracleConnector connector, string name)
        {
            var sqlQpc = GetDeleteValidationSourceQuery(connector.DbReference, connector.DbLink, name);

            return connector.ExecuteQuery(sqlQpc);
        }
        
        #region Queries
        private static IQueryParamCollection GetUpdateValidationSourceQuery(string dbReference, string dbLink, string name, string dbName, string dbUser, string dbPassword, string dbViReference, string dbViLink, string passwordEncodedType)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationSources;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                name = ("" + name).ToUpper();
                dbName = ("" + dbName).ToUpper();
                dbUser = ("" + dbUser).ToUpper();
                dbViReference = ("" + dbViReference).ToUpper();
                dbViLink = ("" + dbViLink).ToUpper();
                passwordEncodedType = ("" + passwordEncodedType).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " VS USING " +
                        " (SELECT " +
                        "  :" + nameof(name) + " source_name, " +
                        "  :" + nameof(dbName) + " db_name, " +
                        "  :" + nameof(dbUser) + " db_user, " +
                        "  :" + nameof(dbPassword) + " db_password, " +
                        "  :" + nameof(dbViReference) + " db_reference, " +
                        "  :" + nameof(dbViLink) + " db_link, " +
                        "  :" + nameof(passwordEncodedType) + " password_encoded_type " +
                        "  FROM DUAL) IVS ON ( " +
                        "  VS.source_name = IVS.source_name " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   VS.db_name = IVS.db_name," +
                        "   VS.db_user = IVS.db_user," +
                        "   VS.db_password = IVS.db_password," +
                        "   VS.db_reference = IVS.db_reference," +
                        "   VS.db_link = IVS.db_link," +
                        "   VS.password_encoded_type = IVS.password_encoded_type" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   source_name, " +
                        "   db_name, " +
                        "   db_user, " +
                        "   db_password, " +
                        "   db_reference, " +
                        "   db_link, " +
                        "   password_encoded_type " +
                        " ) " +
                        " VALUES(" +
                        "   IVS.source_name, " +
                        "   IVS.db_name, " +
                        "   IVS.db_user, " +
                        "   IVS.db_password, " +
                        "   IVS.db_reference, " +
                        "   IVS.db_link, " +
                        "   IVS.password_encoded_type " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(name), name));
            qpCollection.AddParam(new OracleParameter(nameof(dbName), dbName));
            qpCollection.AddParam(new OracleParameter(nameof(dbUser), dbUser));
            qpCollection.AddParam(new OracleParameter(nameof(dbPassword), dbPassword));
            qpCollection.AddParam(new OracleParameter(nameof(dbViReference), dbViReference));
            qpCollection.AddParam(new OracleParameter(nameof(dbViLink), dbViLink));
            qpCollection.AddParam(new OracleParameter(nameof(passwordEncodedType), passwordEncodedType));

            return qpCollection;
        }


        private static IQueryParamCollection GetReadValidationSourceQuery(string dbReference, string dbLink)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationSources;

            var query = " SELECT VS.source_name, VS.db_name, VS.db_user, VS.db_password, VS.db_link, VS.db_reference, VS.password_encoded_type" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VS" +
                        " ORDER BY source_name, db_name, db_user";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);

            return qpCollection;
        }
        private static IQueryParamCollection GetReadValidationSourceQuery(string dbReference, string dbLink, string name)
        {
            if (name == null)
                return GetReadValidationSourceQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationSources;

            if (!LsdmConfig.DataSource.CaseSensitive)
                name = ("" + name).ToUpper();

            var query = " SELECT VS.source_name, VS.db_name, VS.db_user, VS.db_password, VS.db_link, VS.db_reference, VS.password_encoded_type" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VS" +
                        " WHERE VS.source_name = :" + nameof(name) + "" +
                        " ORDER BY source_name, db_name, db_user";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(name), name));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValidationSourceQuery(string dbReference, string dbLink, string name, string keyword)
        {
            if(string.IsNullOrWhiteSpace(keyword))
                return GetReadValidationSourceQuery(dbReference, dbLink, name);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationSources;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                name = ("" + name).ToUpper();
                keyword = ("" + keyword).ToUpper();
            }

            keyword = "%" + keyword + "%";
            var query = " SELECT VS.source_name, VS.db_name, VS.db_user, VS.db_password, VS.db_link, VS.db_reference, VS.password_encoded_type" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VS" +
                        " WHERE VS.source_name = :" + nameof(name) + " OR VS.source_name LIKE :" + nameof(keyword) + " OR VS.db_name LIKE :" + nameof(keyword) + " OR VS.db_user LIKE :" + nameof(keyword) + " " +
                        " ORDER BY source_name, db_name, db_user";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(name), name));
            qpCollection.AddParam(new OracleParameter(nameof(keyword), keyword));

            return qpCollection;
        }

        private static IQueryParamCollection GetDeleteValidationSourceQuery(string dbReference, string dbLink, string name)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationSources;

            if (!LsdmConfig.DataSource.CaseSensitive)
                name = ("" + name).ToUpper();

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " source_name = :" + nameof(name) +"";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(name), name));

            return qpCollection;
        }
        #endregion
    }
}

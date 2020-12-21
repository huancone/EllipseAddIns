using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class DatamodelSqlAccess
    {
        public static int CreateModel(OracleConnector connector, string id, string description, string creationUser, bool activeStatus)
        {

            return UpdateModel(connector, id, description, creationUser, activeStatus);
        }

        public static int UpdateModel(OracleConnector connector, string id, string description, string user, bool activeStatus)
        {
            var sqlQpc = GetUpdateModelQuery(connector.DbReference, connector.DbLink, id, description, user, activeStatus);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadModel(OracleConnector connector, string id)
        {
            var sqlQpc = GetReadModelQuery(connector.DbReference, connector.DbLink, id);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadModel(OracleConnector connector, string code, string keyword)
        {
            var sqlQpc = GetReadModelQuery(connector.DbReference, connector.DbLink, code, keyword);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadModel(OracleConnector connector)
        {
            var sqlQpc = GetReadModelQuery(connector.DbReference, connector.DbLink);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int DeleteModel(OracleConnector connector, string id)
        {
            var sqlQpc = GetDeleteModelQuery(connector.DbReference, connector.DbLink, id);

            return connector.ExecuteQuery(sqlQpc);
        }

        public static int UpdateModelLastModification(OracleConnector connector, string modelId, string user)
        {
            var sqlQpc = GetUpdateLastModificationModelQuery(connector.DbReference, connector.DbLink, modelId, user);

            return connector.ExecuteQuery(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateModelQuery(string dbReference, string dbLink, string id, string description, string userName, bool activeStatus)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            var activeStatusInt = MyUtilities.ToInteger(activeStatus);
            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                id = ("" + id).ToUpper();
                description = ("" + description).ToUpper();
                userName = ("" + userName).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " DM USING " +
                        " (SELECT " +
                        "  :" + nameof(id) + " model_id, " +
                        "  :" + nameof(description) + " model_desc, " +
                        "  CURRENT_TIMESTAMP creation_date, " +
                        "  :" + nameof(userName) + " creation_user, " +
                        "  CURRENT_TIMESTAMP last_mod_date, " +
                        "  :" + nameof(userName) + " last_mod_user, " +
                        "  :" + nameof(activeStatusInt) + " active_status " +
                        "  FROM DUAL) IDM ON ( " +
                        "  DM.model_id = IDM.model_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   DM.model_desc = IDM.model_desc," +
                        "   DM.last_mod_date = IDM.last_mod_date," +
                        "   DM.last_mod_user = IDM.last_mod_user," +
                        "   DM.active_status = IDM.active_status" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   model_id, " +
                        "   model_desc, " +
                        "   creation_date, " +
                        "   creation_user, " +
                        "   last_mod_date, " +
                        "   last_mod_user, " +
                        "   active_status " +
                        " ) " +
                        " VALUES(" +
                        "   IDM.model_id, " +
                        "   IDM.model_desc, " +
                        "   IDM.creation_date, " +
                        "   IDM.creation_user, " +
                        "   IDM.last_mod_date, " +
                        "   IDM.last_mod_user, " +
                        "   IDM.active_status " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            qpCollection.AddParam(new OracleParameter(nameof(description), description));
            qpCollection.AddParam(new OracleParameter(nameof(userName), userName));
            qpCollection.AddParam(new OracleParameter(nameof(activeStatusInt), activeStatusInt));
            
            return qpCollection;
        }


        private static IQueryParamCollection GetReadModelQuery(string dbReference, string dbLink)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            var query = " SELECT DM.model_id, DM.model_desc, DM.creation_date, DM.creation_user, DM.last_mod_date, DM.last_mod_user, DM.active_status" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DM" +
                        " ORDER BY model_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            
            return qpCollection;
        }

        private static IQueryParamCollection GetReadModelQuery(string dbReference, string dbLink, string id)
        {
            if (id == null)
                return GetReadModelQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            if (!LsdmConfig.DataSource.CaseSensitive)
                id = ("" + id).ToUpper();

            var query = " SELECT DM.model_id, DM.model_desc, DM.creation_date, DM.creation_user, DM.last_mod_date, DM.last_mod_user, DM.active_status" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DM" +
                        " WHERE DM.model_id = :" + nameof(id) + "" +
                        " ORDER BY model_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadModelQuery(string dbReference, string dbLink, string id, string keyword)
        {
            if(string.IsNullOrWhiteSpace(keyword))
                return GetReadModelQuery(dbReference, dbLink, id);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                id = ("" + id).ToUpper();
                keyword = ("" + keyword).ToUpper();
            }

            keyword = "%" + keyword + "%";
            var query = " SELECT DM.model_id, DM.model_desc, DM.creation_date, DM.creation_user, DM.last_mod_date, DM.last_mod_user, DM.active_status" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DM" +
                        " WHERE DM.model_id = :" + nameof(id) + " OR DM.model_desc LIKE :" + keyword + " "+
                        " ORDER BY model_id, model_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            qpCollection.AddParam(new OracleParameter(nameof(keyword), keyword));

            return qpCollection;
        }

        private static IQueryParamCollection GetDeleteModelQuery(string dbReference, string dbLink, string id)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            if (!LsdmConfig.DataSource.CaseSensitive)
                id = ("" + id).ToUpper();

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " model_id = :" + nameof(id) +"";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }

        private static IQueryParamCollection GetUpdateLastModificationModelQuery(string dbReference, string dbLink, string modelId, string userName)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatamodels;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                userName = ("" + userName).ToUpper();
            }

            var query = " UPDATE " +
                        " " + dbReference + tableName + dbLink +
                        " SET " +
                        "   last_mod_user = :" + nameof(userName) + "," +
                        "   last_mod_date = CURRENT_TIMESTAMP" +
                        " WHERE " +
                        "   model_id = :" + nameof(modelId) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(userName), userName));
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));

            return qpCollection;
        }

        #endregion
    }
}

using System;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class DatasheetSqlAccess
    {
        public static int CreateDatasheetHeader(OracleConnector connector, string modelId, int? sheedId, DateTime date, string shift, string sequenceId, string userName)
        {
            return UpdateDatasheetHeader(connector, modelId, sheedId, date, shift, sequenceId, userName);
        }

        public static int UpdateDatasheetHeader(OracleConnector connector, string modelId, int? sheedId, DateTime date, string shift, string sequenceId, string userName)
        {
            var sqlQpc = GetUpdateDatasheetHeaderQuery(connector.DbReference, connector.DbLink, modelId, sheedId, date, shift, sequenceId, userName);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadDatasheetHeader(OracleConnector connector, int? sheetId)
        {
            var sqlQpc = GetReadDatasheetHeaderQuery(connector.DbReference, connector.DbLink, sheetId);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadDatasheetHeader(OracleConnector connector, string modelId, DateTime date)
        {
            return ReadDatasheetHeader(connector, modelId, date, null, null);
        }

        public static IDataReader ReadDatasheetHeader(OracleConnector connector, string modelId, DateTime startDate, DateTime finishDate)
        {
            var sqlQpc = GetReadDatasheetHeaderQuery(connector.DbReference, connector.DbLink, modelId, startDate, finishDate);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadDatasheetHeader(OracleConnector connector, string modelId, DateTime date, string shift, string sequenceId)
        {
            var sqlQpc = GetReadDatasheetHeaderQuery(connector.DbReference, connector.DbLink, modelId, date, shift, sequenceId);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int DeleteDatasheetHeader(OracleConnector connector, int? sheetId)
        {
            var sqlQpc = GetDeleteDatasheetHeaderQuery(connector.DbReference, connector.DbLink, sheetId);

            return connector.ExecuteQuery(sqlQpc);
        }

        public static int DeleteDatasheetHeader(OracleConnector connector, string modelId, DateTime date, string shift, string sequenceId)
        {
            var sqlQpc = GetDeleteDatasheetHeaderQuery(connector.DbReference, connector.DbLink, modelId, date, shift, sequenceId);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static int UpdateDatasheetHeaderLastModification(OracleConnector connector, int? id, string user)
        {
            var sqlQpc = GetUpdateLastModificationDatasheetQuery(connector.DbReference, connector.DbLink, id, user);

            return connector.ExecuteQuery(sqlQpc);
        }
        #region Queries
        private static IQueryParamCollection GetUpdateDatasheetHeaderQuery(string dbReference, string dbLink, string modelId, int? sheetId, DateTime sheetDate, string shift, string sequenceId, string userName)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                shift = ("" + shift).ToUpper();
                sequenceId = ("" + sequenceId).ToUpper();
                userName = ("" + userName).ToUpper();
            }
            sheetDate = sheetDate.Date;//Only Date Parte (No Time)

            if (string.IsNullOrWhiteSpace(shift))
                shift = " ";
            if (string.IsNullOrWhiteSpace(sequenceId))
                sequenceId = " ";

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " DS USING " +
                        " (SELECT " +
                        "  :" + nameof(modelId) + " model_id, " +
                        (sheetId != null ? "  :" + nameof(sheetId) + " sheet_id, " : null) +
                        "  :" + nameof(sheetDate) + " sheet_date, " +
                        "  :" + nameof(shift) + " shift, " +
                        "  :" + nameof(sequenceId) + " sequence_id, " +
                        "  :" + nameof(userName) + " creation_user, " +
                        "  CURRENT_TIMESTAMP creation_date, " +
                        "  :" + nameof(userName) + " last_mod_user, " +
                        "  CURRENT_TIMESTAMP last_mod_date " +
                        "  FROM DUAL) IDS ON ( ";

            if (sheetId != null)
                query += "  DS.sheet_id = IDS.sheet_id ";
            else
                query += "  DS.model_id = IDS.model_id " +
                        "  AND DS.sheet_date = IDS.sheet_date " +
                        "  AND DS.shift = IDS.shift " +
                        "  AND DS.sequence_id = IDS.sequence_id";

            query += " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        //Not allowed because primary/unique keys shouldn't be updated
                        //Not allowed to be changed due to historic and compatibility with attribute types
                        //"   DS.model_id = IDS.model_id," + 
                        //"   DS.sheet_date = IDS.sheet_date," +
                        //"   DS.shift = IDS.shift," +
                        //"   DS.sequence_id = IDS.sequence_id," +
                        "   DS.last_mod_date = IDS.last_mod_date," +
                        "   DS.last_mod_user = IDS.last_mod_user" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   model_id, " +
                        (sheetId != null ? "   sheet_id, " : null) +
                        "   sheet_date, " +
                        "   shift, " +
                        "   sequence_id, " +
                        "   creation_user, " +
                        "   creation_date, " +
                        "   last_mod_user, " +
                        "   last_mod_date " +
                        " ) " +
                        " VALUES(" +
                        "   IDS.model_id, " +
                        (sheetId != null ? "   IDS.sheet_id, " : null) +
                        "   IDS.sheet_date, " +
                        "   IDS.shift, " +
                        "   IDS.sequence_id, " +
                        "   IDS.creation_user, " +
                        "   IDS.creation_date, " +
                        "   IDS.last_mod_user, " +
                        "   IDS.last_mod_date " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            if(sheetId != null)
                qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));
            qpCollection.AddParam(new OracleParameter(nameof(sheetDate), sheetDate));
            qpCollection.AddParam(new OracleParameter(nameof(shift), shift));
            qpCollection.AddParam(new OracleParameter(nameof(sequenceId), sequenceId));
            qpCollection.AddParam(new OracleParameter(nameof(userName), userName));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadDatasheetHeaderQuery(string dbReference, string dbLink, int? sheetId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            var query = " SELECT DS.model_id, DS.sheet_id, DS.sheet_date, DS.shift, DS.sequence_id, DS.creation_user, DS.creation_date, DS.last_mod_user, DS.last_mod_date" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DS" +
                        " DS DS.sheet_id = :" + nameof(sheetId) + "" +
                        " ORDER BY model_id, sheet_date, shift, sequence_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadDatasheetHeaderQuery(string dbReference, string dbLink, string modelId, DateTime sheetDate, string shift, string sequenceId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                shift = ("" + shift).ToUpper();
                sequenceId = ("" + sequenceId).ToUpper();
            }
            sheetDate = sheetDate.Date;//Only Date Parte (No Time)

            if (string.IsNullOrWhiteSpace(shift))
                shift = " ";
            if (string.IsNullOrWhiteSpace(sequenceId))
                sequenceId = " ";

            var query = " SELECT DS.model_id, DS.sheet_id, DS.sheet_date, DS.shift, DS.sequence_id, DS.creation_user, DS.creation_date, DS.last_mod_user, DS.last_mod_date" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DS" +
                        " WHERE DS.model_id = :" + nameof(modelId) + "" +
                        "   AND DS.sheet_date = :" + nameof(sheetDate) + "" +
                        "   AND DS.shift = :" + nameof(shift) + "" +
                        "   AND DS.sequence_id = :" + nameof(sequenceId) + "" +
                        " ORDER BY model_id, sheet_date, shift, sequence_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(sheetDate), sheetDate));
            qpCollection.AddParam(new OracleParameter(nameof(shift), shift));
            qpCollection.AddParam(new OracleParameter(nameof(sequenceId), sequenceId));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadDatasheetHeaderQuery(string dbReference, string dbLink, string modelId, DateTime startDate, DateTime finishDate)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            if (!LsdmConfig.DataSource.CaseSensitive)
                modelId = ("" + modelId).ToUpper();
            startDate = startDate.Date;//Only Date Parte (No Time)
            finishDate = finishDate.Date;//Only Date Parte (No Time) //Warning: Due to inserts of Datasheet with only date this should not have any negative effect in the outer boundary

            var query = " SELECT DS.model_id, DS.sheet_id, DS.sheet_date, DS.shift, DS.sequence_id, DS.creation_user, DS.creation_date, DS.last_mod_user, DS.last_mod_date" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " DS" +
                        " WHERE DS.model_id = :" + nameof(modelId) + "" +
                        "   AND DS.sheet_date >= :" + nameof(startDate) + "" +
                        "   AND DS.sheet_date <= :" + nameof(finishDate) + "" +
                        " ORDER BY model_id, sheet_date, shift, sequence_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(startDate), startDate));
            qpCollection.AddParam(new OracleParameter(nameof(finishDate), finishDate));

            return qpCollection;
        }
        private static IQueryParamCollection GetDeleteDatasheetHeaderQuery(string dbReference, string dbLink, int? sheetId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " sheet_id = :" + nameof(sheetId) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));

            return qpCollection;
        }

        private static IQueryParamCollection GetDeleteDatasheetHeaderQuery(string dbReference, string dbLink, string modelId, DateTime date, string shift, string sequenceId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                shift = ("" + shift).ToUpper();
                sequenceId = ("" + sequenceId).ToUpper();
            }
            date = date.Date;//Only Date Parte (No Time) //Warning: Due to inserts of Datasheet with only date this should not have any negative effect

            if (string.IsNullOrWhiteSpace(shift))
                shift = " ";
            if (string.IsNullOrWhiteSpace(sequenceId))
                sequenceId = " ";

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        "   DS.model_id = :" + nameof(modelId) + "" +
                        "   AND DS.sheet_date = :" + nameof(date) + "" +
                        "   AND DS.shift = :" + nameof(shift) + "" +
                        "   AND DS.sequence_id = :" + nameof(sequenceId) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(date), date));
            qpCollection.AddParam(new OracleParameter(nameof(shift), shift));
            qpCollection.AddParam(new OracleParameter(nameof(sequenceId), sequenceId));

            return qpCollection;
        }

        private static IQueryParamCollection GetUpdateLastModificationDatasheetQuery(string dbReference, string dbLink, int? id, string userName)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableDatasheets;

            if (!LsdmConfig.DataSource.CaseSensitive)
                userName = ("" + userName).ToUpper();

            var query = " UPDATE " +
                        " " + dbReference + tableName + dbLink +
                        " SET " +
                        "   last_mod_user = :" + nameof(userName) + "," +
                        "   last_mod_date = CURRENT_TIMESTAMP" +
                        " WHERE " +
                        "   sheet_id = :" + nameof(id) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(userName), userName));
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }
        #endregion
    }
}

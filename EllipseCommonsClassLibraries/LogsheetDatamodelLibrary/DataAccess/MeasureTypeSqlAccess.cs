using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public static class MeasureTypeSqlAccess
    {
        public static int CreateMeasureType(OracleConnector connector, int? id, string description)
        {
            return UpdateMeasureType(connector, id, description);
        }

        public static IDataReader ReadMeasureType(OracleConnector connector, int? id)
        {
            var sqlQpc = GetReadMeasureTypeQuery(connector.DbReference, connector.DbLink, id);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadMeasureType(OracleConnector connector, string description, bool exactMatch = false)
        {
            var sqlQpc = GetReadMeasureTypeQuery(connector.DbReference, connector.DbLink, description, exactMatch);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadMeasureType(OracleConnector connector)
        {
            var sqlQpc = GetReadMeasureTypeQuery(connector.DbReference, connector.DbLink);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int UpdateMeasureType(OracleConnector connector, int? id, string description)
        {
            var sqlQpc = GetUpdateMeasureTypeQuery(connector.DbReference, connector.DbLink, id, description);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static int DeleteMeasureType(OracleConnector connector, int? id)
        {
            var sqlQpc = GetDeleteMeasureTypeQuery(connector.DbReference, connector.DbLink, id);

            return connector.ExecuteQuery(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateMeasureTypeQuery(string dbReference, string dbLink, int? id, string description)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                description = ("" + description).ToUpper();

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " MT USING " +
                        " (SELECT " +
                        "  :" + nameof(id) + " measure_type_id, " +
                        "  :" + nameof(description) + " measure_type_desc" +
                        "  FROM DUAL) IMT ON ( " +
                        "  MT.measure_type_id = IMT.measure_type_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   MT.measure_type_desc = IMT.measure_type_desc" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   measure_type_id, " +
                        "   measure_type_desc " +
                        " ) " +
                        " VALUES(" +
                        "   IMT.measure_type_id, " +
                        "   IMT.measure_type_desc " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            qpCollection.AddParam(new OracleParameter(nameof(description), description));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadMeasureTypeQuery(string dbReference, string dbLink)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            var query = " SELECT measure_type_id, measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " ORDER BY measure_type_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var qpCollection = new OracleQueryParamCollection(query);

            return qpCollection;
        }
        private static IQueryParamCollection GetReadMeasureTypeQuery(string dbReference, string dbLink, int? id)
        {
            if(id == null)
                return GetReadMeasureTypeQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            var query = " SELECT measure_type_id, measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " measure_type_id = :" + nameof(id) + "" +
                        " ORDER BY measure_type_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadMeasureTypeQuery(string dbReference, string dbLink, string description, bool exactMatch = false)
        {
            if (string.IsNullOrWhiteSpace(description))
                return GetReadMeasureTypeQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                description = ("" + description).ToUpper();

            string descCondition;
            if (exactMatch)
                descCondition = " measure_type_desc = :" + nameof(description) + "";
            else
            {
                description = "%" + description + "%";
                descCondition = " measure_type_desc LIKE :" + nameof(description) + "";
            }

            var query = " SELECT measure_type_id, measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + 
                        " WHERE" +
                        descCondition +
                        " ORDER BY measure_type_desc, measure_type_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(description), description));

            return qpCollection;
        }

        private static IQueryParamCollection GetDeleteMeasureTypeQuery(string dbReference, string dbLink, int? id)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " measure_type_id = :" + nameof(id);

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }
        #endregion
    }
}

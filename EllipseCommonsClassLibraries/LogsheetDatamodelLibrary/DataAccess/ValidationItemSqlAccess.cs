using System;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;
using SharedClassLibrary;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class ValidationItemSqlAccess
    {
        public static int CreateValidationItem(OracleConnector connector, string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {

            return UpdateValidationItem(connector, sourceName, id, description, sourceTable, sourceColumn, sortable, distinctFilter);
        }

        public static int UpdateValidationItem(OracleConnector connector, string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            var sqlQpc = GetUpdateValidationItemQuery(connector.DbReference, connector.DbLink, sourceName, id, description, sourceTable, sourceColumn, sortable, distinctFilter);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadValidationItem(OracleConnector connector)
        {
            var sqlQpc = GetReadValidationItemQuery(connector.DbReference, connector.DbLink);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadValidationItem(OracleConnector connector, int? id)
        {
            var sqlQpc = GetReadValidationItemQuery(connector.DbReference, connector.DbLink, id);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadValidationItem(OracleConnector connector, string sourceName)
        {
            var sqlQpc = GetReadValidationItemQuery(connector.DbReference, connector.DbLink, sourceName);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadValidationItem(OracleConnector connector, string sourceName, string keyword)
        {
            var sqlQpc = GetReadValidationItemQuery(connector.DbReference, connector.DbLink, sourceName, keyword);
            return connector.GetQueryResult(sqlQpc);
        }
        public static int DeleteValidationItem(OracleConnector connector, int? id)
        {
            var sqlQpc = GetDeleteValidationItemQuery(connector.DbReference, connector.DbLink, id);

            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader GetValidationListValues(OracleConnector connector, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            var sqlQpc = GetValidationListValuesQuery(connector.DbReference, connector.DbLink, sourceTable, sourceColumn, sortable, distinctFilter);
            return connector.GetQueryResult(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateValidationItemQuery(string dbReference, string dbLink, string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;
            
            //Sanitization and Sql Injection Control for NonParameters Input
            char[] allowChars = { '_', ' ' };
            var sanitSourceTable = MyUtilities.RemoveSpecialCharacters(sourceTable, allowChars);
            var sanitSourceColumn = MyUtilities.RemoveSpecialCharacters(sourceColumn, allowChars);
            if (!sourceTable.Equals(sanitSourceTable, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(Resources.Error_ValidateInput + " " + LsdmResource.ValidationItem_SourceTable, nameof(sourceTable));
            if (!sourceColumn.Equals(sanitSourceColumn, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(Resources.Error_ValidateInput + " " + LsdmResource.ValidationItem_SourceColumn, nameof(sourceColumn));
            //

            var sortableInt = MyUtilities.ToInteger(sortable);
            var distinctFilterInt = MyUtilities.ToInteger(distinctFilter);
            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                sourceName = ("" + sourceName).ToUpper();
                description = ("" + description).ToUpper();
                sourceTable = ("" + sourceTable).ToUpper();
                sourceColumn = ("" + sourceColumn).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " VI USING " +
                        " (SELECT " +
                        "  :" + nameof(sourceName) + " source_name, " +
                        "  :" + nameof(id) + " valid_item_id, " +
                        "  :" + nameof(description) + " valid_item_desc, " +
                        "  :" + nameof(sourceTable) + " valid_item_source_table, " +
                        "  :" + nameof(sourceColumn) + " valid_item_source_column, " +
                        "  :" + nameof(sortableInt) + " sortable, " +
                        "  :" + nameof(distinctFilterInt) + " distinct_filter " +
                        "  FROM DUAL) IVI ON ( " +
                        "  VI.valid_item_id = IVI.valid_item_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   VI.source_name = IVI.source_name," +
                        "   VI.valid_item_desc = IVI.valid_item_desc," +
                        "   VI.valid_item_source_table = IVI.valid_item_source_table," +
                        "   VI.valid_item_source_column = IVI.valid_item_source_column," +
                        "   VI.sortable = IVI.sortable," +
                        "   VI.distinct_filter = IVI.distinct_filter" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   source_name, " +
                        "   valid_item_id, " +
                        "   valid_item_desc, " +
                        "   valid_item_source_table, " +
                        "   valid_item_source_column, " +
                        "   sortable, " +
                        "   distinct_filter " +
                        " ) " +
                        " VALUES(" +
                        "   IVI.source_name, " +
                        "   IVI.valid_item_id, " +
                        "   IVI.valid_item_desc, " +
                        "   IVI.valid_item_source_table, " +
                        "   IVI.valid_item_source_column, " +
                        "   IVI.sortable, " +
                        "   IVI.distinct_filter " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sourceName), sourceName));
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            qpCollection.AddParam(new OracleParameter(nameof(description), description));
            qpCollection.AddParam(new OracleParameter(nameof(sourceTable), sourceTable));
            qpCollection.AddParam(new OracleParameter(nameof(sourceColumn), sourceColumn));
            qpCollection.AddParam(new OracleParameter(nameof(sortableInt), sortableInt));
            qpCollection.AddParam(new OracleParameter(nameof(distinctFilterInt), distinctFilterInt));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValidationItemQuery(string dbReference, string dbLink)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;

            var query = " SELECT VI.source_name, VI.valid_item_id, VI.valid_item_desc, VI.valid_item_source_table, VI.valid_item_source_column, VI.sortable, VI.distinct_filter" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VI" +
                        " ORDER BY source_name, valid_item_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var qpCollection = new OracleQueryParamCollection(query);

            return qpCollection;
        }
        private static IQueryParamCollection GetReadValidationItemQuery(string dbReference, string dbLink, int? id)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;

            var query = " SELECT VI.source_name, VI.valid_item_id, VI.valid_item_desc, VI.valid_item_source_table, VI.valid_item_source_column, VI.sortable, VI.distinct_filter" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VI" +
                        " WHERE VI.valid_item_id = :" + nameof(id) + "" +
                        " ORDER BY source_name, valid_item_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValidationItemQuery(string dbReference, string dbLink, string sourceName)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;

            if (!LsdmConfig.DataSource.CaseSensitive)
                sourceName = ("" + sourceName).ToUpper();

            var query = " SELECT VI.source_name, VI.valid_item_id, VI.valid_item_desc, VI.valid_item_source_table, VI.valid_item_source_column, VI.sortable, VI.distinct_filter" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VI" +
                        " WHERE VI.source_name = :" + nameof(sourceName) + "" +
                        " ORDER BY source_name, valid_item_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sourceName), sourceName));

            return qpCollection;
        }
        private static IQueryParamCollection GetReadValidationItemQuery(string dbReference, string dbLink, string sourceName, string keyword)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                sourceName = ("" + sourceName).ToUpper();
                keyword = ("" + keyword).ToUpper();
            }

            keyword = "%" + keyword + "%";
            var query = " SELECT VI.source_name, VI.valid_item_id, VI.valid_item_desc, VI.valid_item_source_table, VI.valid_item_source_column, VI.sortable, VI.distinct_filter" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " VI" +
                        " WHERE VI.source_name = :" + nameof(sourceName) + " OR VI.source_name LIKE :" + nameof(keyword) + " OR VI.valid_item_desc LIKE :" + nameof(keyword) + "" +
                        " ORDER BY source_name, valid_item_desc";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sourceName), sourceName));
            qpCollection.AddParam(new OracleParameter(nameof(keyword), keyword));

            return qpCollection;
        }
        private static IQueryParamCollection GetDeleteValidationItemQuery(string dbReference, string dbLink, int? id)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValidationItems;

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " valid_item_id = :" + nameof(id) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }

        private static IQueryParamCollection GetValidationListValuesQuery(string dbReference, string dbLink, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = sourceTable;

            //Sanitization and Sql Injection Control for NonParameters Input
            char[] allowChars = {'_', ' '};
            var sanitSourceTable = MyUtilities.RemoveSpecialCharacters(sourceTable, allowChars);
            var sanitSourceColumn = MyUtilities.RemoveSpecialCharacters(sourceColumn, allowChars);
            if (!sourceTable.Equals(sanitSourceTable, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(Resources.Error_ValidateInput + " " + LsdmResource.ValidationItem_SourceTable, nameof(sourceTable));
            if (!sourceColumn.Equals(sanitSourceColumn, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(Resources.Error_ValidateInput + " " + LsdmResource.ValidationItem_SourceColumn, nameof(sourceColumn));
            //

            var distinctParam = distinctFilter ? " DISTINCT " : "";
            var sortableParam = sortable ? " ORDER BY " + sourceColumn : "";

            var query = " SELECT " + distinctParam + sourceColumn +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + "" +
                        sortableParam;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);

            return qpCollection;
        }
        #endregion
    }
}

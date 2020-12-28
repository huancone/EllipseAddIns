using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class MeasureSqlAccess
    {
        public static int CreateMeasure(OracleConnector connector, int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId)
        {
            return UpdateMeasure(connector, id, code, name, description, units, activeStatus, measureTypeId);
        }

        public static int UpdateMeasure(OracleConnector connector, int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId)
        {
            var sqlQpc = GetUpdateMeasureQuery(connector.DbReference, connector.DbLink, id, code, name, description, units, activeStatus, measureTypeId);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadMeasure(OracleConnector connector, int? id)
        {
            var sqlQpc = GetReadMeasureQuery(connector.DbReference, connector.DbLink, id);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadMeasure(OracleConnector connector, string code)
        {
            var sqlQpc = GetReadMeasureQuery(connector.DbReference, connector.DbLink, code);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadMeasure(OracleConnector connector, string code, string keyword)
        {
            var sqlQpc = GetReadMeasureQuery(connector.DbReference, connector.DbLink, code, keyword);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadMeasure(OracleConnector connector, int? measureTypeId, string measureCode)
        {
            var sqlQpc = GetReadMeasureQuery(connector.DbReference, connector.DbLink, measureTypeId, measureCode);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadMeasure(OracleConnector connector)
        {
            var sqlQpc = GetReadMeasureQuery(connector.DbReference, connector.DbLink);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int DeleteMeasure(OracleConnector connector, int? id)
        {
            var sqlQpc = GetDeleteMeasureQuery(connector.DbReference, connector.DbLink, id);

            return connector.ExecuteQuery(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateMeasureQuery(string dbReference, string dbLink, int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;

            var activeStatusInt = MyUtilities.ToInteger(activeStatus);
            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                code = ("" + code).ToUpper();
                name = ("" + name).ToUpper();
                description = ("" + description).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " MS USING " +
                        " (SELECT " +
                        "  :" + nameof(id) + " measure_id, " +
                        "  :" + nameof(code) + " measure_code, " +
                        "  :" + nameof(name) + " measure_name, " +
                        "  :" + nameof(description) + " measure_desc, " +
                        "  :" + nameof(units) + " measure_units, " +
                        "  :" + nameof(activeStatusInt) + " measure_status, " +
                        "  :" + nameof(measureTypeId) + " measure_type_id" +
                        "  FROM DUAL) IMS ON ( " +
                        "  MS.measure_id = IMS.measure_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   MS.measure_code = IMS.measure_code," +
                        "   MS.measure_name = IMS.measure_name," +
                        "   MS.measure_desc = IMS.measure_desc," +
                        "   MS.measure_units = IMS.measure_units," +
                        "   MS.measure_status = IMS.measure_status," +
                        "   MS.measure_type_id = IMS.measure_type_id" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   measure_id, " +
                        "   measure_code, " +
                        "   measure_name, " +
                        "   measure_desc, " +
                        "   measure_units, " +
                        "   measure_status, " +
                        "   measure_type_id " +
                        " ) " +
                        " VALUES(" +
                        "   IMS.measure_id, " +
                        "   IMS.measure_code, " +
                        "   IMS.measure_name, " +
                        "   IMS.measure_desc, " +
                        "   IMS.measure_units, " +
                        "   IMS.measure_status, " +
                        "   IMS.measure_type_id " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            qpCollection.AddParam(new OracleParameter(nameof(code), code));
            qpCollection.AddParam(new OracleParameter(nameof(name), name));
            qpCollection.AddParam(new OracleParameter(nameof(description), description));
            qpCollection.AddParam(new OracleParameter(nameof(units), units));
            qpCollection.AddParam(new OracleParameter(nameof(activeStatusInt), activeStatusInt));
            qpCollection.AddParam(new OracleParameter(nameof(measureTypeId), measureTypeId));

            return qpCollection;
        }


        private static IQueryParamCollection GetReadMeasureQuery(string dbReference, string dbLink)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;
            var tableNameMt = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            var query = " SELECT MS.measure_id, MS.measure_code, MS.measure_name, MS.measure_desc, MS.measure_units, MS.measure_status, MS.measure_type_id, MST.measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MS LEFT JOIN " + dbReference + tableNameMt + dbLink + " MST ON" +
                        " MS.measure_type_id = MST.measure_type_id " +
                        " ORDER BY measure_code";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);

            return qpCollection;
        }
        private static IQueryParamCollection GetReadMeasureQuery(string dbReference, string dbLink, int? id)
        {
            if (id == null)
                return GetReadMeasureQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;
            var tableNameMt = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            var query = " SELECT MS.measure_id, MS.measure_code, MS.measure_name, MS.measure_desc, MS.measure_units, MS.measure_status, MS.measure_type_id, MST.measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MS LEFT JOIN " + dbReference + tableNameMt + dbLink + " MST ON" +
                        " MS.measure_type_id = MST.measure_type_id " +
                        " WHERE MS.measure_id = :" + nameof(id) + "" +
                        " ORDER BY measure_code";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));
            
            return qpCollection;
        }

        private static IQueryParamCollection GetReadMeasureQuery(string dbReference, string dbLink, string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return GetReadMeasureQuery(dbReference, dbLink);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;
            var tableNameMt = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                code = ("" + code).ToUpper();

            var query = " SELECT MS.measure_id, MS.measure_code, MS.measure_name, MS.measure_desc, MS.measure_units, MS.measure_status, MS.measure_type_id, MST.measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MS LEFT JOIN " + dbReference + tableNameMt + dbLink + " MST ON" +
                        " MS.measure_type_id = MST.measure_type_id " +
                        " WHERE MS.measure_code = :" + nameof(code) + "" +
                        " ORDER BY measure_code, measure_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(code), code));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadMeasureQuery(string dbReference, string dbLink, string code, string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
                return GetReadMeasureQuery(dbReference, dbLink, code);

            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;
            var tableNameMt = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                code = ("" + code).ToUpper();
                keyword = ("" + keyword).ToUpper();
            }

            keyword = "%" + keyword + "%";
            var query = " SELECT MS.measure_id, MS.measure_code, MS.measure_name, MS.measure_desc, MS.measure_units, MS.measure_status, MS.measure_type_id, MST.measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MS LEFT JOIN " + dbReference + tableNameMt + dbLink + " MST ON" +
                        " MS.measure_type_id = MST.measure_type_id " +
                        " WHERE MS.measure_code = '" + nameof(code) + "' OR MS.measure_name LIKE :" + nameof(keyword) + " OR MS.measure_desc LIKE :" + nameof(keyword) + " " +
                        " ORDER BY measure_code, measure_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(code), code));
            qpCollection.AddParam(new OracleParameter(nameof(keyword), keyword));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadMeasureQuery(string dbReference, string dbLink, int? measureTypeId, string code)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;
            var tableNameMt = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasureTypes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                code = ("" + code).ToUpper();

            var query = " SELECT MS.measure_id, MS.measure_code, MS.measure_name, MS.measure_desc, MS.measure_units, MS.measure_status, MS.measure_type_id, MST.measure_type_desc" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MS LEFT JOIN " + dbReference + tableNameMt + dbLink + " MST ON" +
                        " MS.measure_type_id = MST.measure_type_id " +
                        " WHERE MS.measure_code = :" + nameof(code) + " AND MS.measure_type_id = :" + nameof(measureTypeId) + " " +
                        " ORDER BY measure_code, measure_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(code), code));
            qpCollection.AddParam(new OracleParameter(nameof(measureTypeId), measureTypeId));

            return qpCollection;
        }
        private static IQueryParamCollection GetDeleteMeasureQuery(string dbReference, string dbLink, int? id)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableMeasures;

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " measure_id = :" + nameof(id);

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(id), id));

            return qpCollection;
        }
        #endregion
    }
}

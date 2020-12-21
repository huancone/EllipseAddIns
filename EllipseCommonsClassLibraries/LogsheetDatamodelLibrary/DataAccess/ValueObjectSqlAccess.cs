using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public static class ValueObjectSqlAccess
    {
        public static int CreateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, decimal value)
        {
            return UpdateValueObject(connector, modelId, sheetId, attributeId, value);
        }
        public static int CreateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, string value)
        {
            return UpdateValueObject(connector, modelId, sheetId, attributeId, value);
        }
        public static int CreateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, DateTime value)
        {
            return UpdateValueObject(connector, modelId, sheetId, attributeId, value);
        }
        public static int CreateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, string datatype, object value)
        {

            return UpdateValueObject(connector, modelId, sheetId, attributeId, datatype, value);
        }
        public static int UpdateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, decimal value)
        {
            var sqlQpc = GetUpdateValueObjectQuery(connector.DbReference, connector.DbLink, modelId, sheetId, attributeId, value);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static int UpdateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, string value)
        {
            var sqlQpc = GetUpdateValueObjectQuery(connector.DbReference, connector.DbLink, modelId, sheetId, attributeId, value);
            return connector.ExecuteQuery(sqlQpc);
        }
        public static int UpdateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, DateTime value)
        {
            var sqlQpc = GetUpdateValueObjectQuery(connector.DbReference, connector.DbLink, modelId, sheetId, attributeId, value);
            return connector.ExecuteQuery(sqlQpc);
        }
        public static int UpdateValueObject(OracleConnector connector, string modelId, int? sheetId, string attributeId, string datatype, object value)
        {
            var sqlQpc = GetUpdateValueObjectQuery(connector.DbReference, connector.DbLink, modelId, sheetId, attributeId, datatype, value);
            return connector.ExecuteQuery(sqlQpc);
        }
        private static IQueryParamCollection GetUpdateValueObjectQuery(string dbReference, string dbLink, string modelId, int? sheetId, string attributeId, decimal value)
        {
            return GetUpdateValueObjectQuery(dbReference, dbLink, modelId, sheetId, attributeId, DataTypes.Numeric, value);
        }
        private static IQueryParamCollection GetUpdateValueObjectQuery(string dbReference, string dbLink, string modelId, int? sheetId, string attributeId, DateTime value)
        {
            return GetUpdateValueObjectQuery(dbReference, dbLink, modelId, sheetId, attributeId, DataTypes.DateTime, value);
        }
        private static IQueryParamCollection GetUpdateValueObjectQuery(string dbReference, string dbLink, string modelId, int? sheetId, string attributeId, string value)
        {
            return GetUpdateValueObjectQuery(dbReference, dbLink, modelId, sheetId, attributeId, DataTypes.Varchar, value);
        }


        public static IDataReader ReadValueObject(OracleConnector connector, int? sheetId, string attributeId)
        {
            var sqlQpc = GetReadValueObjectQuery(connector.DbReference, connector.DbLink, sheetId, attributeId);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadValueObject(OracleConnector connector, string modelId, DateTime date, string shift, string sequenceId, string attributeId)
        {
            var sqlQpc = GetReadValueObjectQuery(connector.DbReference, connector.DbLink, modelId, date, shift, sequenceId, attributeId);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadValueObject(OracleConnector connector, int? sheetId)
        {
            var sqlQpc = GetReadValueObjectQuery(connector.DbReference, connector.DbLink, sheetId);
            return connector.GetQueryResult(sqlQpc);
        }
        public static IDataReader ReadValueObject(OracleConnector connector, string modelId, DateTime date, string shift, string sequenceId)
        {
            var sqlQpc = GetReadValueObjectQuery(connector.DbReference, connector.DbLink, modelId, date, shift, sequenceId);
            return connector.GetQueryResult(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateValueObjectQuery(string dbReference, string dbLink, string modelId, int? sheetId, string attributeId, string datatype, object value)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;

            var tableName = LsdmConfig.DataSource.DataBasePrefix;
            if(datatype.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
                tableName += LsdmConfig.DatabaseInformation.TableValueNumerics;
            else if (datatype.Equals(DataTypes.Varchar, StringComparison.OrdinalIgnoreCase))
                tableName += LsdmConfig.DatabaseInformation.TableValueVarchars;
            else if (datatype.Equals(DataTypes.Text, StringComparison.OrdinalIgnoreCase))
                tableName += LsdmConfig.DatabaseInformation.TableValueTexts;
            else if (datatype.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase) || datatype.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                tableName += LsdmConfig.DatabaseInformation.TableValueDatetimes;
            else 
                throw new ArgumentException(LsdmResource.Error_DataType_Invalid, nameof(datatype));

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                attributeId = ("" + attributeId).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " VAL USING " +
                        " (SELECT " +
                        "  :" + nameof(modelId) + " model_id, " +
                        "  :" + nameof(sheetId) + " sheet_id, " +
                        "  :" + nameof(attributeId) + " attribute_id, " +
                        "  :" + nameof(value) + " value " +
                        "  FROM DUAL) IVAL ON ( " +
                        "  VAL.model_id = IVAL.model_id " +
                        "  AND VAL.sheet_id = IVAL.sheet_id " +
                        "  AND VAL.attribute_id = IVAL.attribute_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   VAL.value = IVAL.value" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   model_id, " +
                        "   sheet_id, " +
                        "   attribute_id, " +
                        "   value " +
                        " ) " +
                        " VALUES(" +
                        "   IVAL.model_id, " +
                        "   IVAL.sheet_id, " +
                        "   IVAL.attribute_id, " +
                        "   IVAL.value " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));
            if (datatype.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
                qpCollection.AddParam(new OracleParameter(nameof(value), (decimal)value));
            else if (datatype.Equals(DataTypes.Varchar, StringComparison.OrdinalIgnoreCase))
                qpCollection.AddParam(new OracleParameter(nameof(value), (string)value));
            else if (datatype.Equals(DataTypes.Text, StringComparison.OrdinalIgnoreCase))
                qpCollection.AddParam(new OracleParameter(nameof(value), (string)value));
            else if (datatype.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase) || datatype.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                qpCollection.AddParam(new OracleParameter(nameof(value), (DateTime)value));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValueObjectQuery(string dbReference, string dbLink, string modelId, DateTime date, string shift, string sequenceId, string attributeId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableNameNumerics = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueNumerics;
            var tableNameVarchars = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueVarchars;
            var tableNameDatetimes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueDatetimes;
            var tableNameTexts = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueTexts;
            var tableNameAttributes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                attributeId = ("" + attributeId).ToUpper();
                shift = ("" + shift).ToUpper();
                sequenceId = ("" + sequenceId).ToUpper();
            }
            date = date.Date;//Only Date Parte (No Time)

            var query = " WITH OVALS AS (" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, value datetime_value, NULL numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameDatetimes + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " AND VAL.attribute_id = :" + nameof(attributeId) +" UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, value numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameNumerics + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " AND VAL.attribute_id = :" + nameof(attributeId) +" UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, NULL numeric_value, value varchar_value, NULL text_value FROM " + dbReference + tableNameVarchars + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " AND VAL.attribute_id = :" + nameof(attributeId) +" UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, NULL numeric_value, NULL varchar_value, value text_value FROM " + dbReference + tableNameTexts + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " AND VAL.attribute_id = :" + nameof(attributeId) +"" +
                        " )" +
                        " SELECT OVALS.model_id, OVALS.sheet_id, OVALS.attribute_id, OVALS.datetime_value, OVALS.numeric_value, OVALS.varchar_value, OVALS.text_value, MA.datatype" +
                        " FROM OVALS LEFT JOIN " + dbReference + tableNameAttributes + dbLink + " MA ON OVALS.model_id = MA.model_id AND OVALS.attribute_id = MA.attribute_id" + 
                        " ORDER BY OVALS.model_id, OVALS.sheet_id, OVALS.attribute_Id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(date), date));
            qpCollection.AddParam(new OracleParameter(nameof(shift), shift));
            qpCollection.AddParam(new OracleParameter(nameof(sequenceId), sequenceId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));
            

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValueObjectQuery(string dbReference, string dbLink, string modelId, DateTime date, string shift, string sequenceId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableNameNumerics = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueNumerics;
            var tableNameVarchars = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueVarchars;
            var tableNameDatetimes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueDatetimes;
            var tableNameTexts = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueTexts;
            var tableNameAttributes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                shift = ("" + shift).ToUpper();
                sequenceId = ("" + sequenceId).ToUpper();
            }
            date = date.Date;//Only Date Parte (No Time)

            var query = " WITH OVALS AS (" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, value datetime_value, NULL numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameDatetimes + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, value numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameNumerics + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, NULL numeric_value, value varchar_value, NULL text_value FROM " + dbReference + tableNameVarchars + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + " UNION ALL" +
                        "   SELECT DS.model_id, DS.sheet_id, VAL.attribute_id, NULL datetime_value, NULL numeric_value, NULL varchar_value, value text_value FROM " + dbReference + tableNameTexts + dbLink + " VAL" +
                        "     JOIN lsdm_datasheets DS on VAL.model_id = DS.model_id AND VAL.sheet_id = DS.sheet_id" +
                        "     WHERE DS.model_id = :" + nameof(modelId) + " AND DS.sheet_date = :" + nameof(date) + " AND DS.shift = :" + nameof(shift) + " AND DS.sequence_id = :" + nameof(sequenceId) + "" +
                        " )" +
                        " SELECT OVALS.model_id, OVALS.sheet_id, OVALS.attribute_id, OVALS.datetime_value, OVALS.numeric_value, OVALS.varchar_value, OVALS.text_value, MA.datatype" +
                        " FROM OVALS LEFT JOIN " + dbReference + tableNameAttributes + dbLink + " MA ON OVALS.model_id = MA.model_id AND OVALS.attribute_id = MA.attribute_id" +
                        " ORDER BY OVALS.model_id, OVALS.sheet_id, OVALS.attribute_Id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(date), date));
            qpCollection.AddParam(new OracleParameter(nameof(shift), shift));
            qpCollection.AddParam(new OracleParameter(nameof(sequenceId), sequenceId));

            return qpCollection;
        }
        private static IQueryParamCollection GetReadValueObjectQuery(string dbReference, string dbLink, int? sheetId, string attributeId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableNameNumerics = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueNumerics;
            var tableNameVarchars = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueVarchars;
            var tableNameDatetimes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueDatetimes;
            var tableNameTexts = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueTexts;
            var tableNameAttributes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                attributeId = ("" + attributeId).ToUpper();
            

            var query = " WITH OVALS AS (" +
                        "   SELECT model_id, sheet_id, attribute_id, value datetime_value, NULL numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameDatetimes + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " AND attribute_id = :" + nameof(attributeId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, value numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameNumerics + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " AND attribute_id = :" + nameof(attributeId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, NULL numeric_value, value varchar_value, NULL text_value FROM " + dbReference + tableNameVarchars + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " AND attribute_id = :" + nameof(attributeId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, NULL numeric_value, NULL varchar_value, value text_value FROM " + dbReference + tableNameTexts + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " AND attribute_id = :" + nameof(attributeId) + "" +
                        " )" +
                        " SELECT OVALS.model_id, OVALS.sheet_id, OVALS.attribute_id, OVALS.datetime_value, OVALS.numeric_value, OVALS.varchar_value, OVALS.text_value, MA.datatype" +
                        " FROM OVALS LEFT JOIN " + dbReference + tableNameAttributes + dbLink + " MA ON OVALS.model_id = MA.model_id AND OVALS.attribute_id = MA.attribute_id" +
                        " ORDER BY OVALS.model_id, OVALS.sheet_id, OVALS.attribute_Id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadValueObjectQuery(string dbReference, string dbLink, int? sheetId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableNameNumerics = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueNumerics;
            var tableNameVarchars = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueVarchars;
            var tableNameDatetimes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueDatetimes;
            var tableNameTexts = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueTexts;
            var tableNameAttributes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            var query = " WITH OVALS AS (" +
                        "   SELECT model_id, sheet_id, attribute_id, value datetime_value, NULL numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameDatetimes + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, value numeric_value, NULL varchar_value, NULL text_value FROM " + dbReference + tableNameNumerics + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, NULL numeric_value, value varchar_value, NULL text_value FROM " + dbReference + tableNameVarchars + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + " UNION ALL" +
                        "   SELECT model_id, sheet_id, attribute_id, NULL datetime_value, NULL numeric_value, NULL varchar_value, value text_value FROM " + dbReference + tableNameTexts + dbLink +
                        "     WHERE sheet_id = :" + nameof(sheetId) + "" +
                        " )" +
                        " SELECT OVALS.model_id, OVALS.sheet_id, OVALS.attribute_id, OVALS.datetime_value, OVALS.numeric_value, OVALS.varchar_value, OVALS.text_value, MA.datatype" +
                        " FROM OVALS LEFT JOIN " + dbReference + tableNameAttributes + dbLink + " MA ON OVALS.model_id = MA.model_id AND OVALS.attribute_id = MA.attribute_id" +
                        " ORDER BY OVALS.model_id, OVALS.sheet_id, OVALS.attribute_Id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));

            return qpCollection;
        }
        private static IQueryParamCollection GetDeleteValueObjectQuery(string dbReference, string dbLink, int? sheetId, string attributeId, string dataType)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableNameNumerics = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueNumerics;
            var tableNameVarchars = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueVarchars;
            var tableNameDatetimes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueDatetimes;
            var tableNameTexts = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableValueTexts;
            var tableNameAttributes = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            var tableName = "";
            if(dataType.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
                tableName = tableNameNumerics;
            else if (dataType.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase) || dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                tableName = tableNameDatetimes;
            else if (dataType.Equals(DataTypes.Varchar, StringComparison.OrdinalIgnoreCase))
                tableName = tableNameVarchars;
            else if (dataType.Equals(DataTypes.Text, StringComparison.OrdinalIgnoreCase))
                tableName = tableNameTexts;
            else
                throw new ArgumentException(LsdmResource.Error_DataType_Invalid, nameof(dataType));

            if (!LsdmConfig.DataSource.CaseSensitive)
                attributeId = ("" + attributeId).ToUpper();

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " sheet_id = :" + nameof(sheetId) + "" +
                        " AND attribute_id = :" + nameof(attributeId) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(sheetId), sheetId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));

            return qpCollection;
        }
        #endregion
    }
}

using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Oracle.ManagedDataAccess.Client;

namespace LogsheetDatamodelLibrary.DataAccess
{
    public class ModelAttributeSqlAccess
    {
        public static int CreateModelAttribute(OracleConnector connector, string modelId, string attributeId, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId)
        {

            return UpdateModelAttribute(connector, modelId, attributeId, description, dataType, sheetIndex, maxLength, maxPrecision, maxScale, allowNull, defaultValue, activeStatus, measureId, validItemId);
        }

        public static int UpdateModelAttribute(OracleConnector connector, string modelId, string attributeId, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId)
        {
            var sqlQpc = GetUpdateModelAttributeQuery(connector.DbReference, connector.DbLink, modelId, attributeId, description, dataType, sheetIndex, maxLength, maxPrecision, maxScale, allowNull, defaultValue, activeStatus, measureId, validItemId);
            return connector.ExecuteQuery(sqlQpc);
        }

        public static IDataReader ReadModelAttribute(OracleConnector connector, string modelId, bool activeOnly)
        {
            var sqlQpc = GetReadModelAttributeQuery(connector.DbReference, connector.DbLink, modelId, activeOnly);
            return connector.GetQueryResult(sqlQpc);
        }

        public static IDataReader ReadModelAttribute(OracleConnector connector, string modelId, string attributeId)
        {
            var sqlQpc = GetReadModelAttributeQuery(connector.DbReference, connector.DbLink, modelId, attributeId);
            return connector.GetQueryResult(sqlQpc);
        }

        public static int DeleteModelAttribute(OracleConnector connector, string modelId, string attributeId)
        {
            var sqlQpc = GetDeleteModelAttributeQuery(connector.DbReference, connector.DbLink, modelId, attributeId);

            return connector.ExecuteQuery(sqlQpc);
        }

        #region Queries
        private static IQueryParamCollection GetUpdateModelAttributeQuery(string dbReference, string dbLink, string modelId, string attributeId, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            var allowNullInt = MyUtilities.ToInteger(allowNull);
            var activeStatusInt = MyUtilities.ToInteger(activeStatus);
            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                attributeId = ("" + attributeId).ToUpper();
                description = ("" + description).ToUpper();
                dataType = ("" + dataType).ToUpper();
                defaultValue = ("" + defaultValue).ToUpper();
            }

            var query = "MERGE INTO " + dbReference + tableName + dbLink + " MA USING " +
                        " (SELECT " +
                        "  :" + nameof(modelId) + " model_id, " +
                        "  :" + nameof(attributeId) + " attribute_id, " +
                        "  :" + nameof(description) + " attribute_desc, " +
                        "  :" + nameof(dataType) + " datatype, " +
                        "  :" + nameof(sheetIndex) + " sheet_index, " +
                        "  :" + nameof(maxLength) + " max_length, " +
                        "  :" + nameof(maxPrecision) + " max_precision, " +
                        "  :" + nameof(maxScale) + " max_scale, " +
                        "  :" + nameof(allowNullInt) + " allow_null, " +
                        "  :" + nameof(defaultValue) + " default_value, " +
                        "  :" + nameof(activeStatusInt) + " active_status, " +
                        "  :" + nameof(measureId) + " measure_id, " +
                        "  :" + nameof(validItemId) + " valid_item_id " +
                        "  FROM DUAL) IMA ON ( " +
                        "  MA.model_id = IMA.model_id " +
                        "  AND MA.attribute_id = IMA.attribute_id " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   MA.attribute_desc = IMA.attribute_desc," +
                        //Not allowed to be changed due to historic reference types
                        //"   MA.datatype = IMA.datatype," + 
                        "   MA.sheet_index = IMA.sheet_index," +
                        "   MA.max_length = IMA.max_length," +
                        "   MA.max_precision = IMA.max_precision," +
                        "   MA.max_scale = IMA.max_scale," +
                        "   MA.allow_null = IMA.allow_null," +
                        "   MA.default_value = IMA.default_value," +
                        "   MA.active_status = IMA.active_status," +
                        "   MA.measure_id = IMA.measure_id," +
                        "   MA.valid_item_id = IMA.valid_item_id" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   model_id, " +
                        "   attribute_id, " +
                        "   attribute_desc, " +
                        "   datatype, " +
                        "   sheet_index, " +
                        "   max_length, " +
                        "   max_precision, " +
                        "   max_scale, " +
                        "   allow_null, " +
                        "   default_value, " +
                        "   active_status, " +
                        "   measure_id, " +
                        "   valid_item_id " +
                        " ) " +
                        " VALUES(" +
                        "   IMA.model_id, " +
                        "   IMA.attribute_id, " +
                        "   IMA.attribute_desc, " +
                        "   IMA.datatype, " +
                        "   IMA.sheet_index, " +
                        "   IMA.max_length, " +
                        "   IMA.max_precision, " +
                        "   IMA.max_scale, " +
                        "   IMA.allow_null, " +
                        "   IMA.default_value, " +
                        "   IMA.active_status, " +
                        "   IMA.measure_id, " +
                        "   IMA.valid_item_id " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));
            qpCollection.AddParam(new OracleParameter(nameof(description), description));
            qpCollection.AddParam(new OracleParameter(nameof(dataType), dataType));
            qpCollection.AddParam(new OracleParameter(nameof(sheetIndex), sheetIndex));
            qpCollection.AddParam(new OracleParameter(nameof(maxLength), maxLength));
            qpCollection.AddParam(new OracleParameter(nameof(maxPrecision), maxPrecision));
            qpCollection.AddParam(new OracleParameter(nameof(maxScale), maxScale));
            qpCollection.AddParam(new OracleParameter(nameof(allowNullInt), allowNullInt));
            qpCollection.AddParam(new OracleParameter(nameof(defaultValue), defaultValue));
            qpCollection.AddParam(new OracleParameter(nameof(activeStatusInt), activeStatusInt));
            qpCollection.AddParam(new OracleParameter(nameof(measureId), measureId));
            qpCollection.AddParam(new OracleParameter(nameof(validItemId), validItemId));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadModelAttributeQuery(string dbReference, string dbLink, string modelId, bool activeOnly = true)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
                modelId = ("" + modelId).ToUpper();

            var activeParam = activeOnly ? "   AND MA.active_status = '1'" : null;

            var query = " SELECT MA.model_id, MA.attribute_id, MA.attribute_desc, MA.datatype, MA.sheet_index, MA.max_length, MA.max_precision, MA.max_scale, MA.allow_null, MA.default_value, MA.active_status, MA.measure_id, MA.valid_item_id" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MA" +
                        " WHERE MA.model_id = :" + nameof(modelId) + "" +
                        activeParam +
                        " ORDER BY model_id, sheet_index, attribute_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));

            return qpCollection;
        }

        private static IQueryParamCollection GetReadModelAttributeQuery(string dbReference, string dbLink, string modelId, string attributeId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                attributeId = ("" + attributeId).ToUpper();
            }

            var query = " SELECT MA.model_id, MA.attribute_id, MA.attribute_desc, MA.datatype, MA.sheet_index, MA.max_length, MA.max_precision, MA.max_scale, MA.allow_null, MA.default_value, MA.active_status, MA.measure_id, MA.valid_item_id" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink + " MA" +
                        " WHERE MA.model_id = :" + nameof(modelId) + "" +
                        "   AND MA.attribute_id = :" + nameof(attributeId) + "" +
                        " ORDER BY model_id, sheet_index, attribute_id";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));

            return qpCollection;
        }

        private static IQueryParamCollection GetDeleteModelAttributeQuery(string dbReference, string dbLink, string modelId, string attributeId)
        {
            if (!string.IsNullOrWhiteSpace(dbReference))
                dbReference = dbReference + ".";
            if (!string.IsNullOrWhiteSpace(dbLink))
                dbLink = "@" + dbLink;
            var tableName = LsdmConfig.DataSource.DataBasePrefix + LsdmConfig.DatabaseInformation.TableModelAttributes;

            if (!LsdmConfig.DataSource.CaseSensitive)
            {
                modelId = ("" + modelId).ToUpper();
                attributeId = ("" + attributeId).ToUpper();
            }

            var query = " DELETE" +
                        " FROM" +
                        " " + dbReference + tableName + dbLink +
                        " WHERE" +
                        " model_id = :" + nameof(modelId) + "" +
                        " AND attribute_id = :" + nameof(attributeId) + "";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(modelId), modelId));
            qpCollection.AddParam(new OracleParameter(nameof(attributeId), attributeId));

            return qpCollection;
        }
        #endregion
    }
}

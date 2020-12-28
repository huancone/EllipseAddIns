using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using LogsheetDatamodelLibrary.DataAccess;
using LogsheetDatamodelLibrary.Models;
using SharedClassLibrary;
using SharedClassLibrary.Classes;

namespace LogsheetDatamodelLibrary.Controllers
{
    public class ValueObjectController
    {
        
        private static ValueObject GetFromDataRecord(IDataRecord dr)
        {
            return new ValueObject(dr);
        }
        public static ReplyMessage Create(ValueObject item, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            connector.BeginTransaction();

            try
            {
                var replyValidation = ValidateValueObjectRestrictions(item, connector);
                if (replyValidation.Message.StartsWith(LsdmResource.Results_Failed, StringComparison.OrdinalIgnoreCase))
                    return replyValidation;

                var user = LsdmConfig.Login.User;

                int result;
                if (string.IsNullOrWhiteSpace(item.DataType))
                    result = ValueObjectSqlAccess.CreateValueObject(connector, item.ModelId, item.SheetId, item.AttributeId, item.Value);
                else
                    result = ValueObjectSqlAccess.CreateValueObject(connector, item.ModelId, item.SheetId, item.AttributeId, item.DataType, item.Value);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    connector.Rollback();
                    return reply;
                }

                var lastModResult = DatasheetController.UpdateHeaderLastModification(item.SheetId, user);
                if (!lastModResult.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                    return lastModResult;

                reply.Message = LsdmResource.Results_Success;
                connector.Commit();
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ValueObjectController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                connector.Rollback();
                return reply;
            }
            finally
            {

            }
        }

        public static ReplyMessage Create(List<ValueObject> itemList, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            connector.BeginTransaction();
            reply.Message = LsdmResource.Results_Success;

            try
            {
                foreach (var item in itemList)
                {
                    try
                    {
                        var replyValidation = ValidateValueObjectRestrictions(item, connector);
                        if (replyValidation.Message.StartsWith(LsdmResource.Results_Failed, StringComparison.OrdinalIgnoreCase))
                        {
                            foreach(var error in replyValidation.Errors)
                                reply.AddError(item?.AttributeId + " " + error);
                            continue;
                        }

                        var user = LsdmConfig.Login.User;

                        int result;
                        if (string.IsNullOrWhiteSpace(item.DataType))
                            result = ValueObjectSqlAccess.CreateValueObject(connector, item.ModelId, item.SheetId, item.AttributeId, item.Value);
                        else
                            result = ValueObjectSqlAccess.CreateValueObject(connector, item.ModelId, item.SheetId, item.AttributeId, item.DataType, item.Value);

                        if (result == 0)
                        {
                            reply.Message = LsdmResource.Results_Warning;
                            reply.AddWarning(item?.AttributeId + " " + LsdmResource.Results_NoRecordsAffected);
                            continue;
                        }

                        var lastModReply = DatasheetController.UpdateHeaderLastModification(item.SheetId, user);
                        if (!lastModReply.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                        {
                            reply.AddWarning(lastModReply.Warnings);
                            foreach(var error in lastModReply.Errors)
                                reply.AddError(item?.AttributeId + " " + error);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("ValueObjectController:Create(List<ValueObject>)", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                        reply.AddError(item?.AttributeId + " " + ex.Message);
                    }
                }

                if (reply.Warnings.Length > 0)
                    reply.Message = LsdmResource.Results_Warning;
                if (reply.Errors.Length > 0)
                    reply.Message = LsdmResource.Results_Failed;

                if(reply.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                    connector.Commit();
                else
                    connector.Rollback();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ValueObjectController:Create(List<ValueObject>)", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.AddError(ex.Message);
                connector.Rollback();
            }
            return reply;
        }

        public static ReplyMessage Create(string modelId, int? sheetId, string attributeId, decimal value, OracleConnector connector = null)
        {
            return Create(modelId, sheetId, attributeId, DataTypes.Numeric, value, connector);
        }
        public static ReplyMessage Create(string modelId, int? sheetId, string attributeId, DateTime value, OracleConnector connector = null)
        {
            return Create(modelId, sheetId, attributeId, DataTypes.DateTime, value, connector);
        }
        public static ReplyMessage Create(string modelId, int? sheetId, string attributeId, string value, OracleConnector connector = null)
        {
            
            return Create(modelId, sheetId, attributeId, DataTypes.Varchar, value, connector);
        }

        public static ReplyMessage Create(string modelId, int? sheetId, string attributeId, string dataType, object value, OracleConnector connector = null)
        {
            var item = new ValueObject {ModelId = modelId, SheetId = sheetId, AttributeId = attributeId};
            item.SetDataType(dataType);

            if(dataType.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
                item.SetValue((decimal)value);
            else if (dataType.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase) || dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                item.SetValue((DateTime)value);
            else if (dataType.Equals(DataTypes.Varchar, StringComparison.OrdinalIgnoreCase))
                item.SetValue((string)value);
            else if (dataType.Equals(DataTypes.Text, StringComparison.OrdinalIgnoreCase))
                item.SetValue((string)value);

            return Create(item, connector);
        }


        public static ValueObject Read(int? sheetId, string attributeId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = ValueObjectSqlAccess.ReadValueObject(connector, sheetId, attributeId);

            if (dr == null || dr.IsClosed) return null;

            dr.Read();
            
            var item = GetFromDataRecord(dr);

            return item;
        }
        public static ValueObject Read(string modelId, DateTime date, string shift, string sequenceId, string attributeId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = ValueObjectSqlAccess.ReadValueObject(connector, modelId, date, shift, sequenceId, attributeId);

            if (dr == null || dr.IsClosed) return null;

            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static List<ValueObject> Read(int? sheetId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<ValueObject>();

            var dr = ValueObjectSqlAccess.ReadValueObject(connector, sheetId);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<ValueObject> Read(string modelId, DateTime date, string shift, string sequenceId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<ValueObject>();

            var dr = ValueObjectSqlAccess.ReadValueObject(connector, modelId, date, shift, sequenceId);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        private static ReplyMessage ValidateValueObjectRestrictions(ValueObject item, OracleConnector connector = null)
        {
            return ValidateValueObjectRestrictions(item.ModelId, item.AttributeId, item.DataType, item.Value, connector);
        }

        private static ReplyMessage ValidateValueObjectRestrictions(string modelId, string attributeId, string dataType, object value, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var replyMessage = new ReplyMessage();

            var attribute = ModelAttributeController.Read(modelId, attributeId, connector);
            
            if(!MyUtilities.IsTrue(attribute.AllowNull) && value == null)
                replyMessage.AddError(LsdmResource.Error_ValueObject_NullValue);

            if (dataType.Equals(DataTypes.Numeric))
            {
                if (value == null)
                {
                    replyMessage.AddError("[" + attributeId + "] " + LsdmResource.Error_ValueObject_NumericValidation_Scale + " (" + attribute.MaxScale + ")");
                }
                else
                {
                    var numValue = (decimal) value;
                    var sqlDecimal = new SqlDecimal(numValue);
                    var numPrecision = (int) sqlDecimal.Precision;
                    var numScale = (int) sqlDecimal.Scale;
                    if (attribute.MaxPrecision != null && numPrecision > attribute.MaxPrecision)
                        replyMessage.AddError("[" + attributeId + "] " + LsdmResource.Error_ValueObject_NumericValidation_Precision + " (" + attribute.MaxPrecision + ")");
                    if (attribute.MaxScale != null && numScale > attribute.MaxScale)
                        replyMessage.AddError("[" + attributeId + "] " + LsdmResource.Error_ValueObject_NumericValidation_Scale + " (" + attribute.MaxScale + ")");
                }
            }
            else if (dataType.Equals(DataTypes.Varchar))
            {
                var strValue = (string) value;
                var strLength = ("" + strValue).Length;

                if (attribute.MaxLength != null && strLength > attribute.MaxLength)
                    replyMessage.AddError("[" + attributeId + "] " + LsdmResource.Error_Attribute_VarcharValidation_LengthExceeded + " (" + attribute.MaxScale + ")");
                else if (strLength > LsdmConfig.DatabaseInformation.VarcharLengthLimit)
                    replyMessage.AddError("[" + attributeId + "] " + LsdmResource.Error_Attribute_VarcharValidation_LengthExceeded + " (" + LsdmConfig.DatabaseInformation.VarcharLengthLimit + ")");

            }
            else if (dataType.Equals(DataTypes.DateTime) || dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
            {
                //no validations required
            }
            else if (dataType.Equals(DataTypes.Text))
            {
                //no validations required
            }
            else
                throw new ArgumentException(LsdmResource.Error_DataType_Invalid, nameof(dataType));

            replyMessage.Message = LsdmResource.Results_Success;

            if (replyMessage.Errors.Length > 0)
                replyMessage.Message = LsdmResource.Results_Failed;

            return replyMessage;
        }
    }
}

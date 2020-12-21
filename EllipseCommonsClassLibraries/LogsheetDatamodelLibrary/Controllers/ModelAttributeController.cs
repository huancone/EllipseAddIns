using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using LogsheetDatamodelLibrary.DataAccess;
using LogsheetDatamodelLibrary.Models;
using SharedClassLibrary;
using SharedClassLibrary.Classes;

namespace LogsheetDatamodelLibrary.Controllers
{
    public class ModelAttributeController
    {
        
        private static ModelAttribute GetFromDataRecord(IDataRecord dr)
        {
            return new ModelAttribute(dr);
        }
        public static ReplyMessage Create(ModelAttribute attribute, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();

                var replyValidation = ValidateModelAttributeRestrictions(attribute);
                if (replyValidation.Message.StartsWith(LsdmResource.Results_Failed, StringComparison.OrdinalIgnoreCase))
                    return replyValidation;

                connector.BeginTransaction();
                
                var user = LsdmConfig.Login.User;

                var result = ModelAttributeSqlAccess.CreateModelAttribute(connector, attribute.ModelId, attribute.Id, attribute.Description, attribute.DataType, attribute.SheetIndex, attribute.MaxLength, attribute.MaxPrecision, attribute.MaxScale, attribute.AllowNull, attribute.DefaultValue, attribute.ActiveStatus, attribute.MeasureId, attribute.ValidationItemId);
                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    connector.Rollback();
                    return reply;
                }

                var lastModResult = DatamodelController.UpdateLastModification(attribute.ModelId, user);
                if (!lastModResult.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                    return lastModResult;

                reply.Message = LsdmResource.Results_Success;
                connector.Commit();
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ModelAttributeController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);

                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                connector.Rollback();
                return reply;
            }
        }


        public static ReplyMessage Create(string modelId, string attributeId, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId, OracleConnector connector = null)
        {
            var item = new ModelAttribute(modelId, attributeId, description, dataType, sheetIndex, maxLength, maxPrecision, maxScale, allowNull, defaultValue, activeStatus, measureId, validItemId);

            return Create(item, connector);
        }

        public static List<ModelAttribute> Read(string modelId, bool activeOnly = true, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<ModelAttribute>();

            var dr = ModelAttributeSqlAccess.ReadModelAttribute(connector, modelId, activeOnly);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }


        public static ModelAttribute Read(string modelId, string attributeId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = ModelAttributeSqlAccess.ReadModelAttribute(connector, modelId, attributeId);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }
        
        public static ReplyMessage Update(ModelAttribute attribute, OracleConnector connector = null)
        {
            return Update(attribute.ModelId, attribute.Id, attribute.Description, attribute.DataType, attribute.SheetIndex, attribute.MaxLength, attribute.MaxPrecision, attribute.MaxScale, attribute.AllowNull, attribute.DefaultValue, attribute.ActiveStatus, attribute.MeasureId, attribute.ValidationItemId, connector);
        }

        public static ReplyMessage Update(string modelId, string attributeId, string description, string dataType, int? sheetIndex, int? maxLength, int? maxPrecision, int? maxScale, bool allowNull, string defaultValue, bool activeStatus, int? measureId, int? validItemId, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();

                var replyValidation = ValidateModelAttributeRestrictions(dataType, maxLength, maxPrecision, maxScale, defaultValue);
                if (replyValidation.Message.StartsWith(LsdmResource.Results_Failed, StringComparison.OrdinalIgnoreCase))
                    return replyValidation;
                connector.BeginTransaction();

                var user = LsdmConfig.Login.User;

                var result = ModelAttributeSqlAccess.UpdateModelAttribute(connector, modelId, attributeId, description, dataType, sheetIndex, maxLength, maxPrecision, maxScale, allowNull, defaultValue, activeStatus, measureId, validItemId);
                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    connector.Rollback();
                    return reply;
                }

                var lastModResult = DatamodelController.UpdateLastModification(modelId, user);
                if (!lastModResult.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                    return lastModResult;

                reply.Message = LsdmResource.Results_Success;
                connector.Commit();
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ModelAttributeController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                connector.Rollback();
                return reply;
            }
        }


        public static ReplyMessage Delete(ModelAttribute attribute, OracleConnector connector = null)
        {
            return Delete(attribute.ModelId, attribute.Id, connector);
        }

        public static ReplyMessage Delete(string modelId, string attributeId, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            connector.BeginTransaction();
            try
            {
                var user = LsdmConfig.Login.User;

                var result = ModelAttributeSqlAccess.DeleteModelAttribute(connector, modelId, attributeId);
                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    connector.Rollback();
                    return reply;
                }

                var lastModResult = DatamodelController.UpdateLastModification(modelId, user);
                if (!lastModResult.Message.StartsWith(LsdmResource.Results_Success, StringComparison.OrdinalIgnoreCase))
                    return lastModResult;

                reply.Message = LsdmResource.Results_Success;
                connector.Commit();
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ModelAttributeController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                connector.Rollback();
                return reply;
            }
        }

        private static ReplyMessage ValidateModelAttributeRestrictions(ModelAttribute item)
        {
            return ValidateModelAttributeRestrictions(item.DataType, item.MaxLength, item.MaxPrecision, item.MaxScale, item.DefaultValue);
        }
        private static ReplyMessage ValidateModelAttributeRestrictions(string dataType, int? maxLength, int? maxPrecision, int? maxScale, string defaultValue)
        {
            var replyMessage = new ReplyMessage();
            replyMessage.Message = LsdmResource.Results_Success;

            if (dataType.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
            {
                if (maxLength != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_NumericValidation_LenghtNotRequired);
                if(maxPrecision != null && maxScale != null && maxScale > maxPrecision) 
                    replyMessage.AddError(LsdmResource.Error_Attribute_NumericValidation_ScaleGreaterThanPrecision);
            }
            else if (dataType.Equals(DataTypes.Varchar, StringComparison.OrdinalIgnoreCase))
            {
                if (maxPrecision != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_VarcharValidation_PrecisionNotRequired);
                if (maxScale != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_VarcharValidation_ScaleNotRequired);
                if(maxLength != null && maxLength > LsdmConfig.DatabaseInformation.VarcharLengthLimit)
                    replyMessage.AddError(LsdmResource.Error_Attribute_VarcharValidation_LengthExceeded + " ("+ LsdmConfig.DatabaseInformation.VarcharLengthLimit + ")");
            }
            else if (dataType.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase) || dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
            {
                if (maxPrecision != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_DatetimeValidation_PrecisionNotRequired);
                if (maxScale != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_DatetimeValidation_ScaleNotRequired);
                if (maxLength != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_DatetimeValidation_LengthNotRequired);
            }
            else if (dataType.Equals(DataTypes.Text, StringComparison.OrdinalIgnoreCase))
            {
                if (maxPrecision != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_VarcharValidation_PrecisionNotRequired);
                if (maxScale != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_VarcharValidation_ScaleNotRequired);
                if (maxLength != null)
                    replyMessage.AddError(LsdmResource.Error_Attribute_TextValidation_LengthNotRequired);
                if (!string.IsNullOrWhiteSpace(defaultValue) && defaultValue.Length > LsdmConfig.DatabaseInformation.VarcharLengthLimit)
                    replyMessage.AddWarning(LsdmResource.Error_Attribute_TextValidation_DefaultValueWarning);
            }
            else
                throw new ArgumentException(LsdmResource.Error_DataType_Invalid, nameof(dataType));

            if (replyMessage.Warnings.Length > 0)
                replyMessage.Message = LsdmResource.Results_Warning;

            if (replyMessage.Errors.Length > 0)
                replyMessage.Message = LsdmResource.Results_Failed;

            return replyMessage;
        }

    }
}

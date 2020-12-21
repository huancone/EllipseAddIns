using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using LogsheetDatamodelLibrary.DataAccess;
using LogsheetDatamodelLibrary.Models;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Controllers
{
    public class ValidationItemController
    {
        private static ValidationItem GetFromDataRecord(IDataRecord dr)
        {
            return new ValidationItem(dr);
        }
        public static ReplyMessage Create(ValidationItem validItem, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = ValidationItemSqlAccess.CreateValidationItem(connector, validItem.SourceName, validItem.Id, validItem.Description, validItem.SourceTable, validItem.SourceColumn, validItem.Sortable, validItem.DistinctFilter);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = LsdmResource.Results_Success;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ValidationItemController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Create(string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter, OracleConnector connector = null)
        {
            var item = new ValidationItem(sourceName, id, description, sourceTable, sourceColumn, sortable, distinctFilter);

            return Create(item, connector);
        }

        public static List<ValidationItem> Read(OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<ValidationItem>();

            var dr = ValidationItemSqlAccess.ReadValidationItem(connector);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static ValidationItem Read(int? validItemId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = ValidationItemSqlAccess.ReadValidationItem(connector, validItemId);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static List<ValidationItem> Read(string sourceName, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<ValidationItem>();

            var dr = ValidationItemSqlAccess.ReadValidationItem(connector, sourceName);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<ValidationItem> Read(string sourceName, string keyword, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<ValidationItem>();

            var dr = ValidationItemSqlAccess.ReadValidationItem(connector, sourceName, keyword);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static ReplyMessage Update(ValidationItem validItem, OracleConnector connector = null)
        {
            return Update(validItem.SourceName, validItem.Id, validItem.Description, validItem.SourceTable, validItem.SourceColumn, validItem.Sortable, validItem.DistinctFilter, connector);
        }

        public static ReplyMessage Update(string sourceName, int? id, string description, string sourceTable, string sourceColumn, bool sortable, bool distinctFilter, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = ValidationItemSqlAccess.UpdateValidationItem(connector, sourceName, id, description, sourceTable, sourceColumn, sortable, distinctFilter);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = LsdmResource.Results_Success;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ValidationItemController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Delete(ValidationItem validItem, OracleConnector connector = null)
        {
            return Delete(validItem.Id, connector);
        }

        public static ReplyMessage Delete(int? id, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = ValidationItemSqlAccess.DeleteValidationItem(connector, id);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = LsdmResource.Results_Success;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ValidationItemController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }

        public static List<string> GetValidationValuesList(int? id)
        {
            var list = new List<string>();
            var validItem = ValidationItemController.Read(id);
            var source = ValidationSourceController.ReadFirst(validItem.SourceName);
            var dbItem = new DatabaseItem();
            dbItem.Name = source.Name;
            dbItem.DbName = source.DbName;
            dbItem.DbUser = source.DbUser;
            dbItem.DbPassword = source.DbPassword;
            dbItem.DbReference = source.DbReference;
            dbItem.DbLink = source.DbLink;

            if (source.PasswordEncodedType.Equals(ValidationSource.EncryptionTypeValues.Default))
                dbItem.DbPassword = Encryption.Decrypt(source.DbPassword, Encryption.EncryptPassPhrase);

            var connector = new OracleConnector(dbItem);
            try
            {
                
                var dr = ValidationItemSqlAccess.GetValidationListValues(connector, validItem.SourceTable, validItem.SourceColumn, validItem.Sortable, validItem.DistinctFilter);
                if (dr == null || dr.IsClosed) return list;
                while (dr.Read())
                {
                    var item = dr["sourcename"].ToString().Trim();

                    list.Add(item);
                }
            }
            finally
            {
                connector.CloseConnection();
            }
            return list;
        }
    }
}

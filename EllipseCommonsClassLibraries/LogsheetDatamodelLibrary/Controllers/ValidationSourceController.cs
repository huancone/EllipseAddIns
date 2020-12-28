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
    public class ValidationSourceController
    {
        private static ValidationSource GetFromDataRecord(IDataRecord dr)
        {
            return new ValidationSource(dr);
        }
        public static ReplyMessage Create(ValidationSource source, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                string encryptedPass;
                if (source.PasswordEncodedType.Equals(ValidationSource.EncryptionTypeValues.Default))
                    encryptedPass = Encryption.Encrypt(source.DbPassword, Encryption.EncryptPassPhrase);
                else
                    encryptedPass = source.DbPassword;

                var result = ValidationSourceSqlAccess.CreateValidationSource(connector, source.Name, source.DbName, source.DbUser, encryptedPass, source.DbReference, source.DbLink, source.PasswordEncodedType);

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
                Debugger.LogError("ValidationSourceController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Create(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string passwordEncodedType, OracleConnector connector = null)
        {
            var item = new ValidationSource();
            item.Name = name;
            item.DbName = dbName;
            item.DbUser = dbUser;
            item.DbLink = dbLink;
            item.DbReference = dbReference;
            item.DbPassword = dbPassword;
            item.PasswordEncodedType = passwordEncodedType;


            return Create(item, connector);
        }

        public static List<ValidationSource> Read(OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<ValidationSource>();

            var dr = ValidationSourceSqlAccess.ReadValidationSource(connector);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }
        public static List<ValidationSource> Read(string name, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<ValidationSource>();

            var dr = ValidationSourceSqlAccess.ReadValidationSource(connector, name);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }


        public static List<ValidationSource> Read(string name, string keyword, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<ValidationSource>();

            var dr = ValidationSourceSqlAccess.ReadValidationSource(connector, name, keyword);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static ValidationSource ReadFirst(string name, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = ValidationSourceSqlAccess.ReadValidationSource(connector, name);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static ReplyMessage Update(ValidationSource source, OracleConnector connector = null)
        {
            return Update(source.Name, source.DbName, source.DbUser, source.DbPassword, source.DbReference, source.DbLink, source.PasswordEncodedType, connector);
        }

        public static ReplyMessage Update(string name, string dbName, string dbUser, string dbPassword, string dbReference, string dbLink, string passwordEncodedType, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                string encryptedPass;
                if (passwordEncodedType.Equals(ValidationSource.EncryptionTypeValues.Default))
                    encryptedPass = Encryption.Encrypt(dbPassword, Encryption.EncryptPassPhrase);
                else
                    encryptedPass = dbPassword;

                var result = ValidationSourceSqlAccess.UpdateValidationSource(connector, name, dbName, dbUser, encryptedPass, dbReference, dbLink, passwordEncodedType);

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
                Debugger.LogError("ValidationSourceController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Delete(ValidationSource source, OracleConnector connector = null)
        {
            return Delete(source.Name, connector);
        }

        public static ReplyMessage Delete(string name, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = ValidationSourceSqlAccess.DeleteValidationSource(connector, name);

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
                Debugger.LogError("ValidationSourceController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);

                return reply;
            }
        }
    }
}

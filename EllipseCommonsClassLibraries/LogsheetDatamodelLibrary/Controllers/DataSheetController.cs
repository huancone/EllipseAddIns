using System;
using System.Collections.Generic;
using System.Data;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary.Connections;
using LogsheetDatamodelLibrary.DataAccess;
using LogsheetDatamodelLibrary.Models;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Controllers
{
    public class DatasheetController
    {
        
        private static Datasheet GetFromDataRecord(IDataRecord dr)
        {
            return new Datasheet(dr);
        }

        public static List<ModelAttribute> GetHeaderAttributes(string modelId, bool activeOnly = true, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            if (string.IsNullOrWhiteSpace(modelId))
                throw new ArgumentNullException(nameof(modelId), LsdmResource.Error_Model_IdRequired);

            return ModelAttributeController.Read(modelId, activeOnly, connector);
        }

        public static ReplyMessage CreateHeader(Datasheet datasheet, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = DatasheetSqlAccess.CreateDatasheetHeader(connector, datasheet.ModelId, datasheet.Id, datasheet.Date, datasheet.Shift, datasheet.SequenceId, datasheet.CreationUser);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = Resources.Results_Created;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DatasheetController:CreateHeader()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage CreateHeader(string modelId, int? id, DateTime date, string shift, string sequenceId, string creationUser, OracleConnector connector = null)
        {
            var item = new Datasheet();
            item.ModelId = modelId;
            item.Id = id;
            item.Date = date;
            item.Shift = shift;
            item.SequenceId = sequenceId;
            item.CreationUser = creationUser;

            return CreateHeader(item, connector);
        }

        public static List<Datasheet> ReadHeader(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<Datasheet>();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, id);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<Datasheet> ReadHeader(string modelId, DateTime date, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Datasheet>();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, modelId, date, date);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<Datasheet> ReadHeader(string modelId, DateTime startDate, DateTime finishDate, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Datasheet>();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, modelId, startDate, finishDate);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<Datasheet> ReadHeader(string modelId, DateTime startDate, string shift, string sequenceId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Datasheet>();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, modelId, startDate, shift, sequenceId);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }
        
        public static Datasheet ReadFirstHeader(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, id);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static Datasheet ReadFirstHeader(string modelId, DateTime startDate, string shift, string sequenceId, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = DatasheetSqlAccess.ReadDatasheetHeader(connector, modelId, startDate, shift, sequenceId);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static ReplyMessage UpdateHeader(Datasheet datasheet, OracleConnector connector = null)
        {
            return UpdateHeader(datasheet.ModelId, datasheet.Id, datasheet.Date, datasheet.Shift, datasheet.SequenceId, datasheet.LastModUser, connector);
        }

        public static ReplyMessage UpdateHeader(string modelId, int? id, DateTime date, string shift, string sequenceId, string user, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = DatasheetSqlAccess.UpdateDatasheetHeader(connector, modelId, id, date, shift, sequenceId, user);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = Resources.Results_Updated;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DatasheetController:UpdateHeader()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }

        public static ReplyMessage Delete(Datasheet datasheet, OracleConnector connector = null)
        {
            return Delete(datasheet.Id, connector);
        }

        public static ReplyMessage Delete(int? id, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = DatasheetSqlAccess.DeleteDatasheetHeader(connector, id);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = Resources.Results_Deleted;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DatasheetController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }

        public static ReplyMessage UpdateHeaderLastModification(int? id, string user, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = DatasheetSqlAccess.UpdateDatasheetHeaderLastModification(connector, id, user);

                if (result == 0)
                {
                    reply.Message = LsdmResource.Results_Warning;
                    reply.AddWarning(LsdmResource.Results_NoRecordsAffected);
                    return reply;
                }

                reply.Message = Resources.Results_Updated;
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DatasheetController:UpdateHeaderLastModification()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }

        public static List<ValueObject> GetDataSheetValueObjects(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            if (id == null)
                throw new ArgumentNullException(nameof(id), LsdmResource.Error_Datasheet_IdRequired);

            return ValueObjectController.Read(id, connector);
        }


    }
}

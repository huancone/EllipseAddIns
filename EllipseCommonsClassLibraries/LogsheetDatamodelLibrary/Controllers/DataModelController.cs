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
    public class DatamodelController
    {
        private static Datamodel GetFromDataRecord(IDataRecord dr)
        {
            return new Datamodel(dr);
        }

        public static List<ModelAttribute> GetModelAttributes(string modelId, bool activeOnly = true, IDbConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            if (string.IsNullOrWhiteSpace(modelId))
                throw new ArgumentNullException(nameof(modelId), LsdmResource.Error_Model_IdRequired);

            return ModelAttributeController.Read(modelId, activeOnly, connector);
        }


        public static ReplyMessage Create(Datamodel model, IDbConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = DatamodelSqlAccess.CreateModel(connector, model.Id, model.Description, model.CreationUser, model.ActiveStatus);

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
                Debugger.LogError("DatamodelController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);

                return reply;
            }
        }


        public static ReplyMessage Create(string id, string description, string creationUser, bool activeStatus, IDbConnector connector = null)
        {
            var model = new Datamodel();
            model.Id = id;
            model.Description = description;
            model.CreationUser = creationUser;
            model.ActiveStatus = activeStatus;

            return Create(model, connector);
        }

        public static List<Datamodel> Read(IDbConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Datamodel>();

            var dr = DatamodelSqlAccess.ReadModel(connector);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }
        public static List<Datamodel> Read(string id, IDbConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<Datamodel>();

            var dr = DatamodelSqlAccess.ReadModel(connector, id);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }


        public static List<Datamodel> Read(string id, string keyword, IDbConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Datamodel>();

            var dr = DatamodelSqlAccess.ReadModel(connector, id, keyword);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static Datamodel ReadFirst(string id, IDbConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = DatamodelSqlAccess.ReadModel(connector, id);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }
        

        public static ReplyMessage Update(Datamodel model, IDbConnector connector = null)
        {
            return Update(model.Id, model.Description, model.CreationUser, model.ActiveStatus, connector);
        }

        public static ReplyMessage Update(string id, string description, string user, bool activeStatus, IDbConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();

                var result = DatamodelSqlAccess.UpdateModel(connector, id, description, user, activeStatus);
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
                Debugger.LogError("DatamodelController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Delete(Datamodel model, IDbConnector connector = null)
        {
            return Delete(model.Id, connector);
        }

        public static ReplyMessage Delete(string id, IDbConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();

                var result = DatamodelSqlAccess.DeleteModel(connector, id);
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
                Debugger.LogError("DatamodelController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);

                return reply;
            }
        }

        public static ReplyMessage UpdateLastModification(string id, string user, IDbConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();

                var result = DatamodelSqlAccess.UpdateModelLastModification(connector, id, user);
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
                Debugger.LogError("DatamodelController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }
    }
}

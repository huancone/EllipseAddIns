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
    public class MeasureTypeController
    {
        private static MeasureType GetFromDataRecord(IDataRecord dr)
        {
            return new MeasureType(dr);
        }
        public static ReplyMessage Create(MeasureType measureType, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureTypeSqlAccess.CreateMeasureType(connector, measureType.Id, measureType.Description);

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
                Debugger.LogError("MeasureTypeController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }

        public static ReplyMessage Create(int? id, string description, OracleConnector connector = null)
        {
            var measureType = new MeasureType(id, description);

            return Create(measureType, connector);
        }

        public static List<MeasureType> Read(OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<MeasureType>();

            var dr = MeasureTypeSqlAccess.ReadMeasureType(connector);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }
        public static List<MeasureType> Read(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<MeasureType>();

            var dr = MeasureTypeSqlAccess.ReadMeasureType(connector, id);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<MeasureType> Read(string description, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<MeasureType>();

            var dr = MeasureTypeSqlAccess.ReadMeasureType(connector, description);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static MeasureType ReadFirst(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = MeasureTypeSqlAccess.ReadMeasureType(connector, id);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static MeasureType ReadFirst(string description, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = MeasureTypeSqlAccess.ReadMeasureType(connector, description, true);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }
        public static ReplyMessage Update(MeasureType measureType, OracleConnector connector = null)
        {
            return Update(measureType.Id, measureType.Description, connector);
        }

        public static ReplyMessage Update(int? id, string newDescription, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureTypeSqlAccess.UpdateMeasureType(connector, id, newDescription);

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
                Debugger.LogError("MeasureTypeController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }
        
        public static ReplyMessage Delete(MeasureType measureType, OracleConnector connector = null)
        {
            return Delete(measureType.Id, connector);
        }

        public static ReplyMessage Delete(int? id, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureTypeSqlAccess.DeleteMeasureType(connector, id);

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
                Debugger.LogError("MeasureTypeController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }
    }
}

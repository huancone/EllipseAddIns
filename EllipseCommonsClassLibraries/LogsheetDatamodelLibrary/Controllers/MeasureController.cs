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
    public class MeasureController
    {
        private static Measure GetFromDataRecord(IDataRecord dr)
        {
            return new Measure(dr);
        }
        public static ReplyMessage Create(Measure measure, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureSqlAccess.CreateMeasure(connector, measure.Id, measure.Code, measure.Name, measure.Description, measure.Units, measure.ActiveStatus, measure.MeasureTypeId);

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
                Debugger.LogError("MeasureController:Create()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Create(int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId, OracleConnector connector = null)
        {
            var measure = new Measure(id, code, name, description, units, activeStatus, measureTypeId);

            return Create(measure, connector);
        }

        public static List<Measure> Read(OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Measure>();

            var dr = MeasureSqlAccess.ReadMeasure(connector);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }
        public static List<Measure> Read(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();
            var list = new List<Measure>();

            var dr = MeasureSqlAccess.ReadMeasure(connector, id);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<Measure> Read(string code, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Measure>();

            var dr = MeasureSqlAccess.ReadMeasure(connector, code);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static List<Measure> Read(string code, string keyword, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var list = new List<Measure>();

            var dr = MeasureSqlAccess.ReadMeasure(connector, code, keyword);

            if (dr == null || dr.IsClosed) return list;
            while (dr.Read())
            {
                var item = GetFromDataRecord(dr);

                list.Add(item);
            }

            return list;
        }

        public static Measure ReadFirst(int? id, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = MeasureSqlAccess.ReadMeasure(connector, id);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static Measure ReadFirst(string code, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = MeasureSqlAccess.ReadMeasure(connector, code);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }

        public static Measure ReadFirst(int? measureTypeId, string measureCode, OracleConnector connector = null)
        {
            if (connector == null)
                connector = LsdmConfig.DataSource.GetOracleConnector();

            var dr = MeasureSqlAccess.ReadMeasure(connector, measureTypeId, measureCode);

            if (dr == null || dr.IsClosed) return null;
            dr.Read();

            var item = GetFromDataRecord(dr);

            return item;
        }


        public static ReplyMessage Update(Measure measure, OracleConnector connector = null)
        {
            return Update(measure.Id, measure.Code, measure.Name, measure.Description, measure.Units, measure.ActiveStatus, measure.MeasureTypeId, connector);
        }

        public static ReplyMessage Update(int? id, string code, string name, string description, string units, bool activeStatus, int? measureTypeId, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureSqlAccess.UpdateMeasure(connector, id, code, name, description, units, activeStatus, measureTypeId);

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
                Debugger.LogError("MeasureController:Update()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }


        public static ReplyMessage Delete(Measure measure, OracleConnector connector = null)
        {
            return Delete(measure.Id, connector);
        }

        public static ReplyMessage Delete(int? id, OracleConnector connector = null)
        {
            var reply = new ReplyMessage();
            try
            {
                if (connector == null)
                    connector = LsdmConfig.DataSource.GetOracleConnector();
                var result = MeasureSqlAccess.DeleteMeasure(connector, id);

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
                Debugger.LogError("MeasureController:Delete()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                reply.Message = LsdmResource.Results_Failed;
                reply.AddError(ex.Message);
                return reply;
            }
        }
    }
}

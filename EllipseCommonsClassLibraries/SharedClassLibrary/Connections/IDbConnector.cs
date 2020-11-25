using System.Collections.Generic;
using System.Data;

namespace SharedClassLibrary.Connections
{
    public interface IDbConnector
    {
        int ConnectionTimeOut { get; set; }
        string DbCatalog { get; set; }
        IDbCommand DbCommand { get; }
        IDbConnection DbConnection { get; }
        string DbLink { get; set; }
        string DbName { get; set; }
        string DbPassword { get; set; }
        string DbReference { get; set; }
        IDbTransaction DbTransaction { get; }
        string DbUser { get; set; }
        bool PoolingDataBase { get; set; }
        void BeginTransaction();
        void CancelConnection();
        void CloseConnection(bool dispose = false);
        void Commit();
        int ExecuteQuery(IQueryParamCollection queryParamCollection);
        int ExecuteQuery(string sqlQuery);
        int ExecuteQuery(string sqlQuery, List<IDbDataParameter> parameters);
        int ExecuteQuery(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar);
        DataSet GetDataSetQueryResult(IQueryParamCollection queryParamCollection);
        DataSet GetDataSetQueryResult(string sqlQuery);
        DataSet GetDataSetQueryResult(string sqlQuery, List<IDbDataParameter> parameters);
        DataSet GetDataSetQueryResult(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar);
        //long GetFetchSize();
        IDataReader GetQueryResult(IQueryParamCollection queryParamCollection);
        IDataReader GetQueryResult(string sqlQuery);
        IDataReader GetQueryResult(string sqlQuery, List<IDbDataParameter> parameters);
        IDataReader GetQueryResult(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar);
        List<T> GetQueryResult<T>(IQueryParamCollection queryParamCollection) where T : ISimpleObjectModelSql, new();
        List<T> GetQueryResult<T>(string sqlQuery) where T : ISimpleObjectModelSql, new();
        List<T> GetQueryResult<T>(string sqlQuery, List<IDbDataParameter> parameters) where T : ISimpleObjectModelSql, new();
        List<T> GetQueryResult<T>(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar) where T : ISimpleObjectModelSql, new();
        void RestartConnection();
        void Rollback();
        //void SetFetchSize(long size);
        void StartConnection();
        void StartConnection(string connectionString);
        void StartConnection(string dbName, string dbUser, string dbPass);
    }
}
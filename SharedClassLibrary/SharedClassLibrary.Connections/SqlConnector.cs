using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using SharedClassLibrary.Utilities;

//Shared Class Library - SqlConnector
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

namespace SharedClassLibrary.Connections
{
    public class SqlConnector : IDbConnector
    {
        private string _currentConnectionString;
        private int _queryAttempt;
        private SqlTransaction _transaction;
        public int ConnectionTimeOut { get; set; } //default ODP 15
        public string DbCatalog { get; set; } //para algunas bases de datos
        public string DbLink { get; set; }
        public string DbName { get; set; }
        public string DbPassword { get; set; }
        public string DbReference { get; set; }
        public string DbUser { get; set; } //Ej. SIGCON, CONSULBO
        public int MaxQueryAttempts;

        public IDbCommand DbCommand => _sqlComm;
        public IDbConnection DbConnection => _sqlConn;
        public IDbTransaction DbTransaction => _transaction;

        private SqlCommand _sqlComm;
        private SqlConnection _sqlConn;
        public bool PoolingDataBase { get; set; }

        public SqlConnector()
        {
        }

        public SqlConnector(string dbName, string dbUser, string dbPass)
        {
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPass;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection();
        }

        public SqlConnector(string connectionString)
        {
            _currentConnectionString = connectionString;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection(_currentConnectionString);
        }

        public SqlConnector(DatabaseItem dbItem)
        {
            DbName = dbItem.DbName;
            DbUser = dbItem.DbUser;
            DbPassword = dbItem.DbPassword;
            DbLink = dbItem.DbLink;
            DbReference = dbItem.DbReference;
            DbCatalog = dbItem.DbCatalog;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection();
        }

        public void StartConnection()
        {
            if(!string.IsNullOrWhiteSpace(_currentConnectionString))
                StartConnection(_currentConnectionString);
            else
            {
                var connectionString = "Data Source=" + DbName + ";User ID=" + DbUser + ";Password=" + DbPassword + "; Connection Timeout=" + ConnectionTimeOut + "; Pooling=" + PoolingDataBase.ToString().ToLower();
                StartConnection(connectionString);
            }
        }

        public void StartConnection(string dbName, string dbUser, string dbPass)
        {
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPass;
            StartConnection();
        }

        /// <summary>
        /// </summary>
        /// <param name="connectionString">
        ///     string: anula la configuración predeterminada por la especificada en la cadena de
        ///     conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")
        /// </param>
        public void StartConnection(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                StartConnection();
                return;
            }

            if (string.IsNullOrWhiteSpace(_currentConnectionString) || _currentConnectionString != connectionString)
                _currentConnectionString = connectionString;

            if (_sqlConn != null)
            {
                Rollback();

                _sqlConn.Close();
                _sqlConn.Dispose();
            }

            _sqlConn = new SqlConnection(connectionString);

            if (_sqlComm == null)
                _sqlComm = new SqlCommand();
        }

        public void RestartConnection()
        {
            StartConnection(_currentConnectionString);
        }

        public void BeginTransaction()
        {
            if (_transaction != null)
                try
                {
                    _transaction.Rollback();
                }
                catch
                {
                    //ignored
                }

            if (_sqlConn.State != ConnectionState.Open)
                _sqlConn.Open();
            _transaction = _sqlConn.BeginTransaction();
        }

        public void Commit()
        {
            if (_transaction == null)
                return;
            _transaction.Commit();
            _transaction = null;
            if (_sqlComm != null)
                _sqlComm.Transaction = null;
        }

        public void Rollback()
        {
            if (_transaction == null)
                return;
            _transaction.Rollback();
            _transaction = null;
            if (_sqlComm != null)
                _sqlComm.Transaction = null;
        }


        public void CancelConnection()
        {
            if (_sqlConn != null && _sqlComm != null)
                _sqlComm.Cancel();
        }

        /// <summary>
        ///     Cierra la conexión realizada para la consulta
        /// </summary>
        /// <param name="dispose">Libera los recursos del ejecutar de comandos y de la conexión</param>
        public void CloseConnection(bool dispose = false)
        {
            //This will avoid the default autocommit behaviour when connection closes
            if (_transaction != null && _sqlConn != null)
            {
                Rollback();
                _transaction.Dispose();
            }

            if (_sqlConn != null)
            {
                if (_sqlConn.State != ConnectionState.Closed)
                    _sqlConn.Close();
                if (dispose)
                {
                    _sqlConn.Dispose();
                    _sqlConn = null;
                }
            }

            if (_sqlComm != null && dispose)
            {
                _sqlComm.Dispose();
                _sqlComm = null;
            }
        }



        public long GetFetchSize()
        {
            throw new NotImplementedException();
        }
        public void SetFetchSize(long size)
        {
            throw new NotImplementedException();
        }

        #region ExecuteQuery - Execute Implementations
        public int ExecuteQuery(IQueryParamCollection queryParamCollection)
        {
            Debugger.LogQuery(queryParamCollection.GetGeneratedSql());
            _queryAttempt++;

            try
            {
                if (_sqlComm == null || _sqlConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (_sqlConn.State != ConnectionState.Open && _transaction == null)
                    _sqlConn.Open();
                _sqlComm.Connection = _sqlConn;
                _sqlComm.CommandText = queryParamCollection.CommandText;
                _sqlComm.Parameters.Clear();

                if (queryParamCollection.Parameters != null)
                    foreach (var p in queryParamCollection.Parameters)
                        _sqlComm.Parameters.AddWithValue(p.ParameterName, p.Value);
                if (_transaction != null)
                    _sqlComm.Transaction = _transaction;
                _queryAttempt = 0;
                return _sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _queryAttempt++;

                //TO DO
                /*
                //Timeout error handlers
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut);
                    GetDataSetQueryResult(queryParamCollection);
                }
                */
                Debugger.LogError("SqlConnector:ExecuteQuery(queryParamCollection)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }

        public int ExecuteQuery(string sqlQuery)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery);
            return ExecuteQuery(queryParamCollection);
        }

        public int ExecuteQuery(string sqlQuery, List<IDbDataParameter> parameters)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters);
            return ExecuteQuery(queryParamCollection);
        }

        public int ExecuteQuery(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters, escapeChar);
            return ExecuteQuery(queryParamCollection);
        }
        #endregion

        #region List<T> GetQueryResult - Generic Implementations
        public List<T> GetQueryResult<T>(IQueryParamCollection queryParamCollection) where T : ISimpleObjectModelSql, new()
        {
            var dr = GetQueryResult(queryParamCollection);

            var list = new List<T>();
            if (dr == null || dr.IsClosed)
                return list;

            while (dr.Read())
            {
                var item = new T();
                item.SetFromDataRecord(dr);
                list.Add(item);
            }

            return list;
        }
        public List<T> GetQueryResult<T>(string sqlQuery) where T : ISimpleObjectModelSql, new()
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery);

            return GetQueryResult<T>(queryParamCollection);
        }

        public List<T> GetQueryResult<T>(string sqlQuery, List<IDbDataParameter> parameters) where T : ISimpleObjectModelSql, new()
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters);

            return GetQueryResult<T>(queryParamCollection);
        }
        public List<T> GetQueryResult<T>(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar) where T : ISimpleObjectModelSql, new()
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters, escapeChar);
            return GetQueryResult<T>(queryParamCollection);
        }
        #endregion

        #region DataSet GetQueryResults - DataSet Implementations

        /// <summary>
        ///     Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(string sqlQuery)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery);
            return GetDataSetQueryResult(queryParamCollection);
        }

        /// <summary>
        ///     Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="queryParamCollection">Objeto de colección de query y parámetros de oracle</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(IQueryParamCollection queryParamCollection)
        {
            Debugger.LogQuery(queryParamCollection.GetGeneratedSql());
            _queryAttempt++;
            try
            {
                if (_sqlComm == null || _sqlConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (_sqlConn.State != ConnectionState.Open && _transaction == null)
                    _sqlConn.Open();
                _sqlComm.Connection = _sqlConn;
                _sqlComm.CommandText = queryParamCollection.CommandText;
                _sqlComm.Parameters.Clear();

                if (queryParamCollection.Parameters != null)
                    foreach (var p in queryParamCollection.Parameters)
                        _sqlComm.Parameters.AddWithValue(p.ParameterName, p.Value);
                if (_transaction != null)
                    _sqlComm.Transaction = _transaction;
                _queryAttempt = 0;
                var ds = new DataSet();
                var adapter = new SqlDataAdapter(_sqlComm);
                adapter.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut);
                    GetDataSetQueryResult(queryParamCollection);
                }

                Debugger.LogError("SqlConnector:GetDataSetQueryResult(queryParamCollection)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }

        /// <summary>
        ///     Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">List of parameters</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(string sqlQuery, List<IDbDataParameter> parameters)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters);
            return GetDataSetQueryResult(queryParamCollection);
        }

        /// <summary>
        ///     Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">List of parameters</param>
        /// <param name="escapeChar">Escape char for parameteres. Sql Default '@'</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters, escapeChar);
            return GetDataSetQueryResult(queryParamCollection);
        }
        #endregion

        #region IDataReader GetQueryResult - IDataReader Implementations
        /// <summary>
        ///     Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <returns>SqlDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery);
            return GetQueryResult(queryParamCollection);
        }

        /// <summary>
        ///     Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">List of parameters</param>
        /// <returns>SqlDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, List<IDbDataParameter> parameters)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters);
            return GetQueryResult(queryParamCollection);
        }

        /// <summary>
        ///     Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">List of parameters</param>
        /// <param name="escapeChar">Escape char for parameteres. Sql Default '@'</param>
        /// <returns>SqlDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, List<IDbDataParameter> parameters, char escapeChar)
        {
            var queryParamCollection = new SqlQueryParamCollection(sqlQuery, parameters, escapeChar);
            return GetQueryResult(queryParamCollection);
        }

        /// <summary>
        ///     Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="queryParamCollection">Objeto de colección de query y parámetros de oracle</param>
        /// <returns>SqlDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(IQueryParamCollection queryParamCollection)
        {
            Debugger.LogQuery(queryParamCollection.GetGeneratedSql());
            _queryAttempt++;

            try
            {
                if (_sqlComm == null || _sqlConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (_sqlConn.State != ConnectionState.Open && _transaction == null)
                    _sqlConn.Open();
                _sqlComm.Connection = _sqlConn;
                _sqlComm.CommandText = queryParamCollection.CommandText;
                _sqlComm.Parameters.Clear();

                if (queryParamCollection.Parameters != null)
                    foreach (var p in queryParamCollection.Parameters)
                        _sqlComm.Parameters.AddWithValue(p.ParameterName, p.Value);

                if (_transaction != null)
                    _sqlComm.Transaction = _transaction;
                _queryAttempt = 0;
                return _sqlComm.ExecuteReader();
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut * 10);
                    GetQueryResult(queryParamCollection);
                }

                Debugger.LogError("SqlConnector:GetQueryResult(QueryParamCollection)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }
        #endregion

        public void Dispose()
        {
            _transaction?.Dispose();
            _sqlComm?.Dispose();
            _sqlConn?.Dispose();
        }
    }
}
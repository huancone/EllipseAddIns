using System;
using System.Data;
using System.Threading;
using Oracle.ManagedDataAccess.Client;

namespace CommonsClassLibrary.Connections
{
    public class OracleConnector
    {
        private OracleConnection OracleConn;
        private OracleCommand OracleComm;
        public string DbName;
        public string DbUser; //Ej. SIGCON, CONSULBO
        public string DbPassword;
        public string DbCatalog; //para algunas bases de datos
        public string DbLink;
        public string DbReference;
        public int ConnectionTimeOut;//default ODP 15
        public bool PoolingDataBase;//default ODP true
        private string _currentConnectionString;
        public int MaxQueryAttempts;
        private int _queryAttempt;
        private OracleTransaction _transaction;

        public OracleConnector()
        {}
        public OracleConnector(string dbName, string dbUser, string dbPass)
        {
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPass;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection();
        }
        public OracleConnector(string connectionString)
        {
            _currentConnectionString = connectionString;
            ConnectionTimeOut = 15;
            PoolingDataBase = true;
            MaxQueryAttempts = 3;
            StartConnection(_currentConnectionString);
        }

        public OracleConnector(DatabaseItem dbItem)
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
            var connectionString = "Data Source=" + DbName + ";User ID=" + DbUser + ";Password=" + DbPassword + "; Connection Timeout=" + ConnectionTimeOut + "; Pooling=" + PoolingDataBase.ToString().ToLower();
            StartConnection(connectionString);
        }
        public void StartConnection(string dbName, string dbUser, string dbPass)
        {
            DbName = dbName;
            DbUser = dbUser;
            DbPassword = dbPass;
            StartConnection();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="connectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        public void StartConnection(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                StartConnection();
                return;
            }
            if(string.IsNullOrWhiteSpace(_currentConnectionString) || _currentConnectionString != connectionString)
                _currentConnectionString = connectionString;

            if (OracleConn != null)
            {
                Rollback();

                OracleConn.Close();
                OracleConn.Dispose();

            }

            OracleConn = new OracleConnection(connectionString);

            if (OracleComm == null)
                OracleComm = new OracleCommand();
        }

        public void RestartConnection()
        {
            StartConnection(_currentConnectionString);
        }
        public void BeginTransaction()
        {
            if (_transaction != null)
            {
                try
                {
                    _transaction.Rollback();
                }
                catch
                {
                    //ignored
                }
            }
            if (OracleConn.State != ConnectionState.Open)
                OracleConn.Open();
            _transaction = OracleConn.BeginTransaction();
        }

        public void Commit()
        {
            if (_transaction == null)
                return;
            _transaction.Commit();
            _transaction = null;
            if (OracleComm != null)
                OracleComm.Transaction = null;
        }

        public void Rollback()
        {
            if (_transaction == null)
                return;
            _transaction.Rollback();
            _transaction = null;
            if (OracleComm != null)
                OracleComm.Transaction = null;
        }
        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <returns>OracleDataReader: Conjunto de resultados de la consulta</returns>
        public OracleDataReader GetQueryResult(string sqlQuery)
        {
            Debugger.LogQuery(sqlQuery);
            _queryAttempt++;

            try
            {
                if(OracleComm == null || OracleConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (OracleConn.State != ConnectionState.Open && _transaction == null)
                    OracleConn.Open();
                OracleComm.Connection = OracleConn;
                OracleComm.CommandText = sqlQuery;
                if (_transaction != null)
                    OracleComm.Transaction = _transaction;
                _queryAttempt = 0;
                return OracleComm.ExecuteReader();
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut * 10);
                    GetQueryResult(sqlQuery);
                }

                Debugger.LogError("OracleConnector:GetQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }
        /// <summary>
        /// Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(string sqlQuery)
        {
            Debugger.LogQuery(sqlQuery);
            _queryAttempt++;
            try
            {
                if (OracleComm == null || OracleConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (OracleConn.State != ConnectionState.Open && _transaction == null)
                    OracleConn.Open();
                OracleComm.Connection = OracleConn;
                OracleComm.CommandText = sqlQuery;
                if (_transaction != null)
                    OracleComm.Transaction = _transaction;
                _queryAttempt = 0;
                var ds = new DataSet();
                var adapter = new OracleDataAdapter(OracleComm);
                adapter.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut);
                    GetDataSetQueryResult(sqlQuery);
                }

                Debugger.LogError("OracleConnector:GetDataSetQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }


        public int ExecuteQuery(string sqlQuery, string customConnectionString = null)
        {
            Debugger.LogQuery(sqlQuery);
            _queryAttempt++;

            try
            {
                if (OracleComm == null || OracleConn == null)
                    throw new ArgumentException("Database connection error: Make sure the Database connector is set and the connection is not disposed");
                if (OracleConn.State != ConnectionState.Open && _transaction == null)
                    OracleConn.Open();
                OracleComm.Connection = OracleConn;
                OracleComm.CommandText = sqlQuery;
                if (_transaction != null)
                    OracleComm.Transaction = _transaction;
                _queryAttempt = 0;
                return OracleComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < MaxQueryAttempts)
                {
                    Thread.Sleep(ConnectionTimeOut);
                    GetDataSetQueryResult(sqlQuery);
                }

                Debugger.LogError("OracleConnector:ExecuteQuery(string, string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
        }

        public void CancelConnection()
        {
            if(OracleConn != null && OracleComm != null)
                OracleComm.Cancel();
        }
        /// <summary>
        /// Cierra la conexión realizada para la consulta
        /// </summary>
        /// <param name="dispose">Libera los recursos del ejecutar de comandos y de la conexión</param>
        public void CloseConnection(bool dispose = false)
        {
            //This will avoid the default autocommit behaviour when connection closes
            if (_transaction != null && OracleConn != null)
            {
                Rollback();
                _transaction.Dispose();
            }

            if (OracleConn != null)
            {
                if (OracleConn.State != ConnectionState.Closed)
                    OracleConn.Close();
                if (dispose)
                {
                    OracleConn.Dispose();
                    OracleConn = null;
                }
            }
            if (OracleComm != null && dispose)
            {
                OracleComm.Dispose();
                OracleComm = null;
            }
        }
    }
}

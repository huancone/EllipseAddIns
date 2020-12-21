using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary.Configuration
{
    public class DataSource
    {
        public string DataBasePrefix = "LSDM_";
        public bool CaseSensitive = false;
        private DatabaseItem _dbItem;

        private SqlConnection _sqlConn;
        private SqlCommand _sqlComm;

        private string _currentConnectionString;
        private IDbConnector _databaseConnector;

        private int _connectionTimeOut = 30;//default ODP 15
        private bool _poolingDataBase = true;//default ODP true

        private int _queryAttempt;

        public string DbLink
        {
            get { return _dbItem.DbLink; }
        }
        public string DbReference
        {
            get { return _dbItem.DbReference; }
        }

        public IDbConnector GetOracleConnector()
        {
            return _databaseConnector;
        }
        
        /// <summary>
        /// Limpia las variables de referencia a bases de datos
        /// </summary>
        private void CleanDbSettings()
        {
            if (_dbItem != null)
            {
                _dbItem.Name = null;
                _dbItem.DbName = null;
                _dbItem.DbUser = null;
                _dbItem.DbCatalog = null;
                _dbItem.DbPassword = null;
                _dbItem.DbLink = null;
                _dbItem.DbReference = null;
            }

            if (_databaseConnector != null)
                _databaseConnector.CloseConnection(true);
        }
        /// <summary>
        /// Establece la base de datos según la información ingresada
        /// </summary>
        /// <param name="dbname">Nombre de base de datos (Ej. EL8PROD, EL8TEST)</param>
        /// <param name="dbuser">Usuario de conexión a la base de datos</param>
        /// <param name="dbpass">Contraseña de conexión del usuario a la base de datos</param>
        /// <param name="dblink">Enlace de consulta (Ej. @DBLMIMS)</param>
        /// <param name="dbreference">Referencia de la base de datos (Ej. ELLIPSE, MIMS)</param>
        /// <param name="dbcatalog"></param>
        /// <returns>True</returns>
        // ReSharper disable once InconsistentNaming
        public bool SetDBSettings(string dbname, string dbuser, string dbpass, string dblink, string dbreference, string dbcatalog = null)
        {
            CleanDbSettings();
            _dbItem = new DatabaseItem(dbname, dbuser, dbpass, dblink, dbreference, dbcatalog);
            _databaseConnector = new SharedClassLibrary.Connections.Oracle.OracleConnector(_dbItem);
            return true;
        }

        /// <summary>
        /// Establece la base de datos según la información ingresada
        /// </summary>
        /// <param name="dbname">Nombre de base de datos (Ej. EL8PROD, EL8TEST)</param>
        /// <param name="dbuser">Usuario de conexión a la base de datos</param>
        /// <param name="dbpass">Contraseña de conexión del usuario a la base de datos</param>
        /// <param name="dbcatalog"></param>
        /// <returns>True</returns>
        // ReSharper disable once InconsistentNaming
        public bool SetDBSettings(string dbname, string dbuser, string dbpass, string dbcatalog = null)
        {
            CleanDbSettings();
            _dbItem = new DatabaseItem(dbname, dbuser, dbpass, "", "", dbcatalog);
            _databaseConnector = new DbConnector(_dbItem);

            return true;
        }
        public void SetConnectionTimeOut(int timeout)
        {
            _connectionTimeOut = timeout;
            if (_databaseConnector != null)
                _databaseConnector.ConnectionTimeOut = _connectionTimeOut;
        }

        public int GetConnectionTimeOut()
        {
            return _connectionTimeOut;
        }
        public void SetConnectionPoolingType(bool pooling)
        {
            _poolingDataBase = pooling;
            if (_databaseConnector != null)
                _databaseConnector.PoolingDataBase = pooling;
        }

        public bool GetConnectionPoolingType()
        {
            return _poolingDataBase;
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>OracleDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _databaseConnector.StartConnection(customConnectionString);
            return _databaseConnector.GetQueryResult(sqlQuery);
        }
        /// <summary>
        /// Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public DataSet GetDataSetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _databaseConnector.StartConnection(customConnectionString);
            return _databaseConnector.GetDataSetQueryResult(sqlQuery);
        }
        public int ExecuteQuery(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _databaseConnector.StartConnection(customConnectionString);
            return _databaseConnector.ExecuteQuery(sqlQuery);
        }

        public void BeginTransaction()
        {
            _databaseConnector.BeginTransaction();
        }
        public void Commit()
        {
            _databaseConnector.Commit();
        }
        public void RollBack()
        {
            _databaseConnector.Rollback();
        }
        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>OracleDataReader: Conjunto de resultados de la consulta</returns>
        public SqlDataReader GetSqlQueryResult(string sqlQuery, string customConnectionString = null)
        {
            Debugger.LogQuery(sqlQuery);
            var dbcatalog = "";
            if (_dbItem.DbCatalog != null && !string.IsNullOrWhiteSpace(dbcatalog))
                dbcatalog = "Initial Catalog=" + _dbItem.DbCatalog + "; ";
            var defaultConnectionString = "Data Source=" + _dbItem.DbName + "; " + dbcatalog + "User Id=" + _dbItem.DbUser + "; Password=" + _dbItem.DbPassword + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnectionString;

            if (_sqlConn == null || _currentConnectionString != connectionString)
                _sqlConn = new SqlConnection(connectionString);
            _currentConnectionString = connectionString;

            if (_sqlComm == null)
                _sqlComm = new SqlCommand();
            _queryAttempt++;
            try
            {
                _sqlConn.Open();
                _sqlComm.Connection = _sqlConn;
                _sqlComm.CommandText = sqlQuery;

                _queryAttempt = 0;
                return _sqlComm.ExecuteReader();
            }
            catch (Exception ex)
            {
                Debugger.LogError("EllipseFunctions:GetSqlQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                _queryAttempt = 0;
                throw;
            }
        }
        public DataSet GetDataSetSqlQueryResult(string sqlQuery, string customConnectionString = null)
        {
            Debugger.LogQuery(sqlQuery);
            var dbcatalog = "";
            if (_dbItem.DbCatalog != null && !string.IsNullOrWhiteSpace(dbcatalog))
                dbcatalog = "Initial Catalog=" + _dbItem.DbCatalog + "; ";
            var defaultConnectionString = "Data Source=" + _dbItem.DbName + "; " + dbcatalog + "User Id=" + _dbItem.DbUser + "; Password=" + _dbItem.DbPassword + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnectionString;

            if (_sqlConn == null || _currentConnectionString != connectionString)
                _sqlConn = new SqlConnection(connectionString);
            _currentConnectionString = connectionString;

            if (_sqlComm == null)
                _sqlComm = new SqlCommand();
            _queryAttempt++;
            try
            {
                _sqlConn.Open();
                _sqlComm.Connection = _sqlConn;
                _sqlComm.CommandText = sqlQuery;

                _queryAttempt = 0;
                var ds = new DataSet();
                var adapter = new SqlDataAdapter(_sqlComm);
                adapter.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                Debugger.LogError("EllipseFunctions:GetSqlQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                _queryAttempt = 0;
                throw;
            }
        }
        /// <summary>
        /// Cierra la conexión realizada para la consulta
        /// </summary>
        public void CloseConnection(bool dispose = true)
        {
            if (_databaseConnector != null)
                _databaseConnector.CloseConnection(dispose);

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
            // ReSharper disable once InvertIf
            if (_sqlComm != null && dispose)
            {
                _sqlComm.Dispose();
                _sqlComm = null;
            }
        }
        /// <summary>
        /// Cancela la acción que esté realizando la conexión, pero no cierra la conexión
        /// </summary>
        public void CancelConnection()
        {
            if (_databaseConnector != null)
                _databaseConnector.CancelConnection();

            if (_sqlConn != null && _sqlComm != null)
            {
                _sqlComm.Cancel();
            }
        }

        public static class ProgramAccessType
        {
            public static int Full = 2;
            public static int ReviewOnly = 1;
            public static int AnyAccess = 99;
        }

    }
}

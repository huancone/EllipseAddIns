using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Utilities;

namespace SharedClassLibrary.Ellipse
{
    public class EllipseFunctions : IDisposable
    {
        private DatabaseItem _dbItem;
        /*
        private SqlConnection _sqlConn;
        private SqlCommand _sqlComm;
        */
        private string _currentEnvironment;
        private OracleConnector _oracleConnector;
        private SqlConnector _sqlConnector;
        private int _connectionTimeOut = 30;//default ODP 15
        private bool _poolingDataBase = true;//default ODP true
        //public PostService PostServiceProxy;

        public string DbLink => _dbItem.DbLink;
        public string DbReference => _dbItem.DbReference;

        /// <summary>
        /// Constructor de la clase. Inicia la clase con el nombre de ambientes disponibles (Ej. Productivo, Test, etc) y sus respectivas direcciones web de conexión a los web services
        /// </summary>
        public EllipseFunctions()
        {
            if(!Settings.CurrentSettings.IsServiceListForced)
              SetDBSettings(Environments.EllipseProductivo);
        }

        public EllipseFunctions(EllipseFunctions ellipseFunctions)
        {
            SetDBSettings(ellipseFunctions.GetCurrentEnvironment());
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
                _dbItem.SetDatabaseType(IxDataBaseType.Undefined);
            }

            CloseConnection(true);
            SetCurrentEnvironment(null);
        }
        /// <summary>
        /// Establece un ambiente de producción con el que van a realizarse las consultas/conexiones
        /// </summary>
        /// <param name="environment">Especifica el ambiente con el que va a conectar</param>
        /// <returns></returns>
        // ReSharper disable once InconsistentNaming
        public bool SetDBSettings(string environment)
        {
            CleanDbSettings();
            var dbItem = Environments.GetDatabaseItem(environment);
            if(dbItem == null || dbItem.Name.Equals(null))
                throw new NullReferenceException("No se puede encontrar la base de datos seleccionada. Verifique que eligió un servidor válido y que la base de datos relacionada existe");

            _dbItem = dbItem;

            _initiateConnector();
            SetCurrentEnvironment(environment);
            return true;
        }

        private void _initiateConnector()
        {
            if (_dbItem.DbType.Equals(IxDataBaseType.SqlServer))
                _sqlConnector = new SqlConnector(_dbItem);
            else
                _oracleConnector = new OracleConnector(_dbItem);
        }

        private void _startConnection(string customConnectionString)
        {
            if(_dbItem.DbType == IxDataBaseType.SqlServer)
                _sqlConnector.StartConnection(customConnectionString);
            else
                _oracleConnector.StartConnection(customConnectionString);
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
            _initiateConnector();
            SetCurrentEnvironment(Environments.CustomDatabase);
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
            _dbItem = new DatabaseItem(dbname, dbuser, dbpass, Environments.DefaultDbReferenceName, "", dbcatalog);
            _initiateConnector();
            SetCurrentEnvironment(Environments.CustomDatabase);
            return true;
        }

        public void SetConnectionTimeOut(int timeout)
        {
            _connectionTimeOut = timeout;
            if (_oracleConnector != null)
                _oracleConnector.ConnectionTimeOut = _connectionTimeOut;
            if (_sqlConnector != null)
                _sqlConnector.ConnectionTimeOut = _connectionTimeOut;
        }

        public int GetConnectionTimeOut()
        {
            return _connectionTimeOut;
        }

        public void SetConnectionPoolingType(bool pooling)
        {
            _poolingDataBase = pooling;
            if (_oracleConnector != null)
                _oracleConnector.PoolingDataBase = _poolingDataBase;
            if (_sqlConnector != null)
                _sqlConnector.PoolingDataBase = _poolingDataBase;
        }

        public bool GetConnectionPoolingType()
        {
            return _poolingDataBase;
        }

        public string GetCurrentEnvironment()
        {
            return _currentEnvironment;
        }

        public void SetCurrentEnvironment(string environment)
        {
            _currentEnvironment = environment;
        }

        /// <summary>
        /// Obtiene la URL de conexión al servicio web de Ellipse
        /// </summary>
        /// <param name="environment">Nombre del ambiente al que se va a conectar (EnvironmentConstants.Ambiente)</param>
        /// <param name="serviceType">Tipo de conexión a realizar EWS/POST. Localizada en EnvironmentConstans.ServiceType</param>
        /// <returns>string: URL de la conexión</returns>
        [Obsolete("Function is deprecated. Please use EllipseCommonsClassLibrary.Connections.Environments.GetServiceUrl")]
        public string GetServicesUrl(string environment, string serviceType = null)
        {
            return Environments.GetServiceUrl(environment, serviceType);
        }


        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="queryParamCollection">Objeto de colección de query y parámetros de consulta</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(IQueryParamCollection queryParamCollection, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResult(queryParamCollection);
            
            return _oracleConnector.GetQueryResult(queryParamCollection);
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="queryParamCollection">Objeto de colección de query y parámetros de consulta</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public Task<IDataReader> GetQueryResultAsync(IQueryParamCollection queryParamCollection, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResultAsync(queryParamCollection);

            return _oracleConnector.GetQueryResultAsync(queryParamCollection);
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            if(!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResult(sqlQuery);

            return _oracleConnector.GetQueryResult(sqlQuery);
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public Task<IDataReader> GetQueryResultAsync(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResultAsync(sqlQuery);

            return _oracleConnector.GetQueryResultAsync(sqlQuery);
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">Lista de parámetros de la consulta</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, List<IDbDataParameter> parameters, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResult(sqlQuery, parameters);

            return _oracleConnector.GetQueryResult(sqlQuery, parameters);
        }


        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="parameters">Lista de parámetros de la consulta</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public Task<IDataReader> GetQueryResultAsync(string sqlQuery, List<IDbDataParameter> parameters, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetQueryResultAsync(sqlQuery, parameters);

            return _oracleConnector.GetQueryResultAsync(sqlQuery, parameters);
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
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetDataSetQueryResult(sqlQuery);
            return _oracleConnector.GetDataSetQueryResult(sqlQuery);
        }

        /// <summary>
        /// Obtiene el data set con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>DataSet: Conjunto de resultados de la consulta</returns>
        public Task<DataSet> GetDataSetQueryResultAsync(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.GetDataSetQueryResultAsync(sqlQuery);
            return _oracleConnector.GetDataSetQueryResultAsync(sqlQuery);
        }

        public int ExecuteQuery(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);
            
            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.ExecuteQuery(sqlQuery);

            return _oracleConnector.ExecuteQuery(sqlQuery);
        }

        public Task<int> ExecuteQueryAsync(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _startConnection(customConnectionString);

            if (_dbItem.DbType == IxDataBaseType.SqlServer)
                return _sqlConnector.ExecuteQueryAsync(sqlQuery);

            return _oracleConnector.ExecuteQueryAsync(sqlQuery);
        }

        public void BeginTransaction()
        { 
            _sqlConnector?.BeginTransaction();
            _oracleConnector?.BeginTransaction();
        }
        public void Commit()
        {
            _sqlConnector?.Commit();
            _oracleConnector?.Commit();
        }
        public void RollBack()
        {
            _sqlConnector?.Rollback();
            _oracleConnector?.Rollback();
        }
        
        /// <summary>
        /// Cierra la conexión realizada para la consulta
        /// </summary>
        public void CloseConnection(bool dispose = true)
        {
            _oracleConnector?.CloseConnection(dispose);
            _sqlConnector?.CloseConnection(dispose);
        }
        /// <summary>
        /// Cancela la acción que esté realizando la conexión, pero no cierra la conexión
        /// </summary>
        public void CancelConnection()
        {
            _oracleConnector?.CancelConnection();
            _sqlConnector?.CancelConnection();
        }
        /// <summary>
        /// Revertir Operación. Solo aplica para ScreenService (MSO)
        /// </summary>
        /// <param name="opScreen"></param>
        /// <param name="proxyScreen"></param>
        /// <returns></returns>
        public bool RevertOperation(Screen.OperationContext opScreen, Screen.ScreenService proxyScreen)
        {
            //forzar inicio de pantalla
            var requestScreen = new Screen.ScreenSubmitRequestDTO();
            var prevProgram = "0";
            var actualProgram = "1";

            while (!actualProgram.Equals(prevProgram))
            {
                try
                {
                    requestScreen.screenFields = null;
                    requestScreen.screenKey = "3";
                    var replyScreen = proxyScreen.submit(opScreen, requestScreen);
                    prevProgram = actualProgram;
                    actualProgram = replyScreen.mapName;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("RibbonEllipse:revertOperation(Screen.OperationContext, Screen.ScreenService)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                    prevProgram = actualProgram;

                }
            }
            return false;
        }

        /// <summary>
        /// Verificar Error en Reply de Screen. Solo aplica para ScreenService (MSO)
        /// </summary>
        /// <param name="reply"></param>
        /// <returns></returns>
        public bool CheckReplyError(Screen.ScreenDTO reply)
        {
            //Si no existe un reply es error de ejecución. O si el reply tiene un error de datos
            if (reply == null)
            {
                Debugger.LogError("RibbonEllipse:checkReplyError(Screen.ScreenDTO)", "Se ha producido un error en tiempo de ejecución: null reply error");
                return true;
            }
            // ReSharper disable once InvertIf
            if (reply.message.Length >= 2 && reply.message.Substring(0, 2) == "X2")
            {
                Debugger.LogError("RibbonEllipse: checkReplyError(Screen.ScreenDTO)", reply.message);
                return true;
            }
            return false;
        }
        /// <summary>
        /// Verificar Warning en Reply de Screen. Solo aplica para ScreenService (MSO)
        /// </summary>
        /// <param name="reply"></param>
        /// <returns></returns>
        public bool CheckReplyWarning(Screen.ScreenDTO reply)
        {
            //Si no existe un reply es error de ejecución. O si el reply tiene un warning de datos
            if (reply == null)
            {
                Debugger.LogError("RibbonEllipse:checkReplyWarning(Screen.ScreenDTO)", "Se ha producido un error en tiempo de ejecución: null reply error");
                return true;
            }
            if (reply.message != null && reply.message.Length >= 2 && reply.message.Substring(0, 2) == "W2")
            {
                Debugger.LogWarning("Warning", reply.message);

                return true;
            }
            if (reply.message == null || reply.functionKeys == null || !reply.functionKeys.StartsWith("XMIT-WARNING"))
                return false;

            Debugger.LogWarning("Warning", reply.functionKeys);
            return true;
        }
        
        /// <summary>
        /// Verifica si un usuario tiene acceso a una aplicación especificada de Ellipse
        /// </summary>
        /// <param name="environment">Ambiente a verificar</param>
        /// <param name="districtCode">Distrito</param>
        /// <param name="userName">Nombre de usuario</param>
        /// <param name="codeProgram">Código del Programa (Ej. MSEWOT, MSO720)</param>
        /// <param name="accessType">Tipo de acceso a verificar (ProgramAccessType.Full, ProgramAccessType.ReviewObly, etc)</param>
        /// <returns></returns>
        public bool CheckUserProgramAccess(string environment, string districtCode, string userName, string codeProgram, int accessType)
        {
            SetDBSettings(environment);
            var query = "" +
                           " WITH EPROFILES AS(" +
                           " SELECT" +
                           "     EMPOS.EMPLOYEE_ID," +
                           "     EMPOS.POSITION_ID," +
                           "     DECODE(TRIM(POSITION_PROFILE.DSTRCT_CODE), NULL, EMPLOYEE_PROFILE.DSTRCT_CODE) DSTRCT_CODE," +
                           "     DECODE (TRIM ( EMPOS.GLOBAL_PROFILE ), NULL, DECODE ( TRIM ( POS.GLOBAL_PROFILE ), NULL, DECODE ( TRIM ( POSITION_PROFILE.GLOBAL_PROFILE ), NULL, EMPLOYEE_PROFILE.GLOBAL_PROFILE, POSITION_PROFILE.GLOBAL_PROFILE ), POS.GLOBAL_PROFILE ), EMPOS.GLOBAL_PROFILE ) PROFILE" +
                           "   FROM" +
                           "     ELLIPSE.MSF878 EMPOS" +
                           "     LEFT JOIN ELLIPSE.MSF020 POSITION_PROFILE" +
                           "       ON POSITION_PROFILE.ENTITY = EMPOS.POSITION_ID AND POSITION_PROFILE.ENTRY_TYPE = 'G'" +
                           "     LEFT JOIN ELLIPSE.MSF020 EMPLOYEE_PROFILE" +
                           "       ON EMPLOYEE_PROFILE.ENTITY = EMPOS.EMPLOYEE_ID AND EMPLOYEE_PROFILE.ENTRY_TYPE = 'S'" +
                           "     INNER JOIN ELLIPSE.MSF870 POS" +
                           "       ON EMPOS.POSITION_ID = POS.POSITION_ID" +
                           "     INNER JOIN ELLIPSE.MSF810 EMP" +
                           "       ON EMPOS.EMPLOYEE_ID = EMP.EMPLOYEE_ID" +
                           "   WHERE" +
                           "   TO_DATE((99999999 - EMPOS.INV_STR_DATE), 'YYYYMMDD') <= SYSDATE" +
                           "   AND TO_DATE(DECODE(EMPOS.POS_STOP_DATE, NULL, '99991231', '00000000', '99991231', EMPOS.POS_STOP_DATE), 'YYYYMMDD')             >= SYSDATE" +
                           "   AND EMP.EMPLOYEE_ID = '" + userName + "'" +
                           "   AND (POSITION_PROFILE.DSTRCT_CODE = '" + districtCode + "' OR EMPLOYEE_PROFILE.DSTRCT_CODE = '" + districtCode + "'))" +
                           " SELECT *" +
                           " FROM EPROFILES JOIN ELLIPSE.MSF02A PACCESS ON EPROFILES.PROFILE = PACCESS.ENTITY" +
                           " WHERE PACCESS.APPLICATION_NAME = '" + codeProgram + "' AND ACCESS_LEVEL = '" + accessType + "'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var dReader = GetQueryResult(query);

            var result = !(dReader == null || dReader.IsClosed || !dReader.Read());
            
            return result;
        }

        public static class ProgramAccessType
        {
            public static int Full = 2;
            public static int ReviewOnly = 1;
            public static int AnyAccess = 99;
        }

        public List<KeyValuePair<string, string>> GetItemCodesKeyValuePairs(string tableType, bool activeOnly = true)
        {
            return GetItemCodesKeyValuePairs(tableType, activeOnly, null);
        }
        public List<KeyValuePair<string, string>> GetItemCodesKeyValuePairs(string tableType, bool activeOnly, string additionalQueryParameters)
        {
            var listItems = new List<KeyValuePair<string, string>>();
            var paramActiveOnly = activeOnly ? " AND ACTIVE_FLAG = 'Y'" : "";

            var query = "SELECT * FROM " + _dbItem.DbReference + ".MSF010" + _dbItem.DbLink + " WHERE TABLE_TYPE = '" + tableType + "'" +
                        paramActiveOnly + " " + additionalQueryParameters;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drItemCodes = GetQueryResult(query);

            if (drItemCodes == null || drItemCodes.IsClosed) return listItems;
            while (drItemCodes.Read())
            {
                var item = new KeyValuePair<string, string>(
                    drItemCodes["TABLE_CODE"].ToString().Trim(),
                    drItemCodes["TABLE_DESC"].ToString().Trim());
                listItems.Add(item);
            }

            return listItems;
        }

        public List<string> GetItemCodesString(string tableType, bool activeOnly = true)
        {
            return GetItemCodesString(tableType, activeOnly, null);
        }
        public List<string> GetItemCodesString(string tableType, bool activeOnly, string additionalQueryParameters)
        {
            var listItems = new List<string>();
            var paramActiveOnly = activeOnly ? " AND ACTIVE_FLAG = 'Y'" : "";

            var query = "SELECT * FROM " + _dbItem.DbReference + ".MSF010" + _dbItem.DbLink + " WHERE TABLE_TYPE = '" + tableType + "'" +
                        paramActiveOnly + " " + additionalQueryParameters;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drItemCodes = GetQueryResult(query);

            if (drItemCodes == null || drItemCodes.IsClosed) return listItems;
            while (drItemCodes.Read())
            {
                var item = "" + drItemCodes["TABLE_CODE"].ToString().Trim() + " - "  + drItemCodes["TABLE_DESC"].ToString().Trim();
                listItems.Add(item);
            }

            return listItems;
        }
        public List<EllipseCodeItem> GetItemCodes(string tableType, bool activeOnly, string additionalQueryParameters)
        {
            var listItems = new List<EllipseCodeItem>();
            var paramActiveOnly = activeOnly ? " AND ACTIVE_FLAG = 'Y'" : "";

            var query = "SELECT * FROM " + _dbItem.DbReference + ".MSF010" + _dbItem.DbLink + " WHERE TABLE_TYPE = '" + tableType + "'" +
                        paramActiveOnly + " " + additionalQueryParameters;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var drItemCodes = GetQueryResult(query);

            if (drItemCodes == null || drItemCodes.IsClosed) return listItems;
            while (drItemCodes.Read())
            {
                var item = new EllipseCodeItem(
                    drItemCodes["TABLE_CODE"].ToString().Trim(), 
                    drItemCodes["TABLE_DESC"].ToString().Trim(), 
                    drItemCodes["TABLE_TYPE"].ToString().Trim(), 
                    drItemCodes["ASSOC_REC"].ToString().Trim(), 
                    drItemCodes["ACTIVE_FLAG"].ToString().Trim());
                listItems.Add(item);
            }

            return listItems;
        }

        public List<EllipseCodeItem> GetItemCodes(string tableType, bool activeOnly = true)
        {
            return GetItemCodes(tableType, activeOnly, null);
        }

        public Dictionary<string, string> GetItemCodesDictionary(string tableType)
        {
            var itemList = GetItemCodes(tableType);
            
            return itemList.ToDictionary(item => item.Code, item => item.Description);
        }

        /*
        public static void DebugScreen(Screen.ScreenSubmitRequestDTO request, Screen.ScreenDTO reply, string filename)
        {
            var requestJson = new JavaScriptSerializer().Serialize(request.screenFields);
            var replyJson = new JavaScriptSerializer().Serialize(reply.screenFields);
            var filePath = Settings.CurrentSettings.LocalDataPath + @"debugger\";
            FileWriter.AppendTextToFile(requestJson, "ScreenRequest.txt", filePath);
            FileWriter.AppendTextToFile(replyJson, "ScreenReply.txt", filePath);
        }*/

        //Post methods deprecated
        /*
        public PostService SetPostService(string ellipseUser, string ellipsePswd, string ellipsePost, string EllipseDstrct, string urlService)
        {
            PostServiceProxy = new PostService(ellipseUser, ellipsePswd, ellipsePost, EllipseDstrct, urlService);
            return PostServiceProxy;
        }

        public ResponseDto InitiatePostConnection()
        {
            if(PostServiceProxy == null)
                throw new Exception("No se puede iniciar un servicio post no establecido");
            return PostServiceProxy.InitConexion();
        }

        public ResponseDto ExecutePostRequest(string xmlRequest)
        {
            return PostServiceProxy.ExecutePostRequest(xmlRequest);
        }
        */

        public void Dispose()
        {
            _oracleConnector?.Dispose();
            _sqlConnector?.Dispose();
        }
    }
    

    
    
}

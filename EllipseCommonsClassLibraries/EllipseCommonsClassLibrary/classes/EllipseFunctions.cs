using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary.Classes;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using EllipseCommonsClassLibrary.Connections;
using Oracle.ManagedDataAccess.Client;
using MyUtilities = SharedClassLibrary.Utilities.MyUtilities;
// ReSharper disable AccessToStaticMemberViaDerivedType

namespace EllipseCommonsClassLibrary
{
    public class EllipseFunctions
    {
        private DatabaseItem _dbItem;

        private SqlConnection _sqlConn;
        private SqlCommand _sqlComm;

        private string _currentConnectionString;
        private string _currentEnvironment;
        private OracleConnector _oracleConnector;
        private int _connectionTimeOut = 30;//default ODP 15
        private bool _poolingDataBase = true;//default ODP true
        public PostService PostServiceProxy;
        private int _queryAttempt;

        public string DbLink
        {
            get { return _dbItem.DbLink; }
        }
        public string DbReference
        {
            get { return _dbItem.DbReference; }
        }
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
            }

            if(_oracleConnector != null)
                _oracleConnector.CloseConnection(true);
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
                throw new NullReferenceException("No se puede encontrar la base de datos seleccionada. Verifique que eligió un servidor de ellipse válido y que la base de datos relacionada existe");

            _dbItem = dbItem;
            _oracleConnector = new OracleConnector(_dbItem);

            SetCurrentEnvironment(environment);
            return true;
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
            _oracleConnector = new OracleConnector(_dbItem);
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
            _oracleConnector = new OracleConnector(_dbItem);
            SetCurrentEnvironment(Environments.CustomDatabase);
            return true;
        }
        public void SetConnectionTimeOut(int timeout)
        {
            _connectionTimeOut = timeout;
            if (_oracleConnector != null)
                _oracleConnector.ConnectionTimeOut = _connectionTimeOut;
        }

        public int GetConnectionTimeOut()
        {
            return _connectionTimeOut;
        }
        public void SetConnectionPoolingType(bool pooling)
        {
            _poolingDataBase = pooling;
            if (_oracleConnector != null)
                _oracleConnector.PoolingDataBase = pooling;
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
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>IDataReader: Conjunto de resultados de la consulta</returns>
        public IDataReader GetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            if(!string.IsNullOrWhiteSpace(customConnectionString))
                _oracleConnector.StartConnection(customConnectionString);
            return _oracleConnector.GetQueryResult(sqlQuery);
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
                _oracleConnector.StartConnection(customConnectionString);
            return _oracleConnector.GetDataSetQueryResult(sqlQuery);
        }
        public int ExecuteQuery(string sqlQuery, string customConnectionString = null)
        {
            if (!string.IsNullOrWhiteSpace(customConnectionString))
                _oracleConnector.StartConnection(customConnectionString);
            return _oracleConnector.ExecuteQuery(sqlQuery);
        }

        public void BeginTransaction()
        {
            _oracleConnector.BeginTransaction();
        }
        public void Commit()
        {
            _oracleConnector.Commit();
        }
        public void RollBack()
        {
            _oracleConnector.Rollback();
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

            if(_sqlComm == null)
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

            if(_sqlComm == null)
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
            if(_oracleConnector != null)
                _oracleConnector.CloseConnection(dispose);

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
            if (_oracleConnector != null)
                _oracleConnector.CancelConnection();

            if (_sqlConn != null && _sqlComm != null)
            {
                _sqlComm.Cancel();
            }
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

        [Obsolete("CheckReplyError(ResponseDto) is deprecated, please use System.Web.Services.Ellipse.Post.ResponseDto GotErrorMessages() or GetStringErrorMessages()")]
        public bool CheckReplyError(System.Web.Services.Ellipse.Post.ResponseDto reply)
        {
            
            if (!reply.GotErrorMessages()) return true;
            var errorMessage = "";
            foreach (var msg in reply.Errors)
                errorMessage += msg.Field + " " + msg.Text;
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
            return false;
        }

        [Obsolete("CheckReplyWarning(ResponseDto) is deprecated, please use System.Web.Services.Ellipse.Post.ResponseDto GotWarningMessages() or GetStringWarningMessages()")]
        public bool CheckReplyWarning(System.Web.Services.Ellipse.Post.ResponseDto reply)
        {
            
            if (!reply.GotWarningMessages()) return true;
            var warningMessage = "";
            foreach (var msg in reply.Warnings)
                warningMessage += msg.Field + " " + msg.Text;
            if (!warningMessage.Equals(""))
                throw new Exception(warningMessage);
            return false;
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

        public List<EllipseCodeItem> GetItemCodes(string tableType, bool activeOnly, string additionalQueryParameters)
        {
            var listItems = new List<EllipseCodeItem>();
            var paramActiveOnly = activeOnly ? " AND ACTIVE_FLAG = 'Y'" : "";

            var query = "SELECT * FROM " + _dbItem.DbReference + ".MSF010" + _dbItem.DbLink + " WHERE TABLE_TYPE = '" + tableType + "'" +
                        paramActiveOnly + " " + additionalQueryParameters;
            query = EllipseCommonsClassLibrary.Utilities.MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
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

        public Dictionary<string, string> GetDictionaryItemCodes(string tableType)
        {
            var itemList = GetItemCodes(tableType);
            return itemList.ToDictionary(item => item.code, item => item.description);
        }

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
    }
    

    
    
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using Oracle.ManagedDataAccess.Client;
using Screen = EllipseCommonsClassLibrary.ScreenService;

using System.Threading;

namespace EllipseCommonsClassLibrary
{
    public class EllipseFunctions
    {
        private string _dbname;
        private string _dbuser; //Ej. SIGCON, CONSULBO
        private string _dbpass;
        private string _dbcatalog; //para algunas bases de datos
        // ReSharper disable once InconsistentNaming
        public string dbLink; //Ej. @DBLMIMS, @DBLELLIPSE8
        // ReSharper disable once InconsistentNaming
        public string dbReference; //Ej. MIMSPROD, ELLIPSE

        private SqlConnection _sqlConn;
        private SqlCommand _sqlComm;
        private OracleConnection _sqlOracleConn;
        private OracleCommand _sqlOracleComm;
        private string _currentConnectionString;

        private string _currentEnviroment;
        public string CurrentEnviroment { get; set; }
    
        private int _connectionTimeOut = 30;//default ODP 15
        private bool _poolingDataBase = true;//default ODP true
        public PostService PostServiceProxy;
        private int _queryAttempt;
        private string _defaultDbReferenceName = "ELLIPSE";
        /// <summary>
        /// Constructor de la clase. Inicia la clase con el nombre de ambientes disponibles (Ej. Productivo, Test, etc) y sus respectivas direcciones web de conexión a los web services
        /// </summary>
        public EllipseFunctions()
        {
            SetDBSettings(EnviromentConstants.EllipseProductivo);
        }

        public EllipseFunctions(EllipseFunctions ellipseFunctions)
        {
            SetDBSettings(ellipseFunctions.GetCurrentEnviroment());
        }
        /// <summary>
        /// Limpia las variables de referencia a bases de datos
        /// </summary>
        private void CleanDbSettings()
        {
            _dbname = null;
            _dbuser = null;
            _dbcatalog = null;
            _dbpass = null;
            dbLink = null;
            dbReference = null;
            SetCurrentEnviroment(null);
        }
        /// <summary>
        /// Establece un ambiente de producción con el que van a realizarse las consultas/conexiones
        /// </summary>
        /// <param name="enviroment">Especifica el ambiente con el que va a conectar</param>
        /// <returns></returns>
        // ReSharper disable once InconsistentNaming
        public bool SetDBSettings(string enviroment)
        {
            CleanDbSettings();
            if(enviroment == EnviromentConstants.EllipseProductivo)
            {
                _dbname = "EL8PROD";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = _defaultDbReferenceName;
            }
            else if(enviroment == EnviromentConstants.EllipseTest)
            {
                _dbname = "EL8TEST";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = _defaultDbReferenceName;
            }
            else if (enviroment == EnviromentConstants.EllipseDesarrollo)
            {
                _dbname = "EL8DESA";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = _defaultDbReferenceName;
            }
            else if(enviroment == EnviromentConstants.EllipseContingencia)
            {
                _dbname = "EL8PROD";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = _defaultDbReferenceName;
            }
            else if (enviroment == EnviromentConstants.SigcorProductivo)
            {
                _dbname = "SIGCOPRD";
                _dbuser = "CONSULBO";
                _dbpass = "consulbo";
                dbLink = "@DBLELLIPSE8";
                dbReference = _defaultDbReferenceName;
            }
            else if (enviroment == EnviromentConstants.ScadaRdb)
            {
                _dbname = "PBVFWL01";
                _dbcatalog = "SCADARDB";
                _dbuser = "SCADARDBADMINGUI";
                _dbpass = "momia2011";
                dbLink = "";
                dbReference = "SCADARDB.DBO";
            }
            else
            {
                throw new NullReferenceException("NO SE PUEDE ENCONTRAR LA BASE DE DATOS SELECCIONADA");
            }
            SetCurrentEnviroment(enviroment);
            return true;
        }

        public void SetConnectionTimeOut(int timeout)
        {
            _connectionTimeOut = timeout;
        }

        public int GetConnectionTimeOut()
        {
            return _connectionTimeOut;
        }
        public void SetConnectionPoolingType(bool pooling)
        {
            _poolingDataBase = pooling;
        }

        public bool GetConnectionPoolingType()
        {
            return _poolingDataBase;
        }
        public string GetCurrentEnviroment()
        {
            return _currentEnviroment;
        }
        public void SetCurrentEnviroment(string enviroment)
        {
            _currentEnviroment = enviroment;
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
            _dbname = dbname;
            _dbuser = dbuser;
            _dbcatalog = dbcatalog;
            _dbpass = dbpass;
            dbLink = dblink;
            dbReference = dbreference;
            SetCurrentEnviroment(EnviromentConstants.CustomDatabase);
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
            _dbname = dbname;
            _dbuser = dbuser;
            _dbcatalog = dbcatalog;
            _dbpass = dbpass;
            dbLink = "";
            dbReference = _defaultDbReferenceName;
            SetCurrentEnviroment(EnviromentConstants.CustomDatabase);
            return true;
        }
        /// <summary>
        /// Obtiene la URL de conexión al servicio web de Ellipse
        /// </summary>
        /// <param name="enviroment">Nombre del ambiente al que se va a conectar (EnviromentConstants.Ambiente)</param>
        /// <param name="serviceType">Tipo de conexión a realizar EWS/POST. Localizada en EnviromentConstans.ServiceType</param>
        /// <returns>string: URL de la conexión</returns>
        public string GetServicesUrl(string enviroment, string serviceType = null)
        {
            return EnviromentConstants.GetServiceUrl(enviroment, serviceType);
        }

        /// <summary>
        /// Obtiene el data reader con los resultados de una consulta
        /// </summary>
        /// <param name="sqlQuery">Query a consultar</param>
        /// <param name="customConnectionString">string: anula la configuración predeterminada por la especificada en la cadena de conexión (Ej. "Data Source=DBNAME; User ID=USERID; Passwork=PASSWORD")</param>
        /// <returns>OracleDataReader: Conjunto de resultados de la consulta</returns>
        public OracleDataReader GetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            Debugger.LogQuery(sqlQuery);
            var defaultConnString = "Data Source=" + _dbname + ";User ID=" + _dbuser + ";Password=" + _dbpass + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnString;

            if (_sqlOracleConn == null || _currentConnectionString != connectionString)
                _sqlOracleConn = new OracleConnection(connectionString);
            _currentConnectionString = connectionString;
            
            _sqlOracleComm = new OracleCommand();

            _queryAttempt++;
            try
            {
                if (_sqlOracleConn.State != ConnectionState.Open)
                    _sqlOracleConn.Open();
                _sqlOracleComm.Connection = _sqlOracleConn;
                _sqlOracleComm.CommandText = sqlQuery;

                _queryAttempt = 0;
                return _sqlOracleComm.ExecuteReader();
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < 3)
                {
                    Thread.Sleep(_connectionTimeOut * 10);
                    GetQueryResult(sqlQuery, customConnectionString);
                }
                
                Debugger.LogError("EllipseFunctions:GetQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

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
        public DataSet GetDataSetQueryResult(string sqlQuery, string customConnectionString = null)
        {
            Debugger.LogQuery(sqlQuery);
            var defaultConnString = "Data Source=" + _dbname + ";User ID=" + _dbuser + ";Password=" + _dbpass + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnString;

            if (_sqlOracleConn == null || _currentConnectionString != connectionString)
                _sqlOracleConn = new OracleConnection(connectionString);
            _currentConnectionString = connectionString;

            _sqlOracleComm = new OracleCommand();

            _queryAttempt++;
            try
            {
                if (_sqlOracleConn.State != ConnectionState.Open)
                    _sqlOracleConn.Open();
                _sqlOracleComm.Connection = _sqlOracleConn;
                _sqlOracleComm.CommandText = sqlQuery;

                _queryAttempt = 0;
                var ds = new DataSet();
                var adapter = new OracleDataAdapter(_sqlOracleComm);
                adapter.Fill(ds);
                CloseConnection(false);
                return ds;
            }
            catch (Exception ex)
            {
                _queryAttempt++;
                if (ex.Message.Contains("ORA-12516") && _queryAttempt < 3)
                {
                    Thread.Sleep(_connectionTimeOut);
                    GetQueryResult(sqlQuery, customConnectionString);
                }

                Debugger.LogError("EllipseFunctions:GetQueryResult(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                _queryAttempt = 0;
                throw;
            }
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
            if (_dbcatalog != null && !string.IsNullOrWhiteSpace(dbcatalog))
                dbcatalog = "Initial Catalog=" + _dbcatalog + "; ";
            var defaultConnectionString = "Data Source=" + _dbname + "; " + dbcatalog + "User Id=" + _dbuser + "; Password=" + _dbpass + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnectionString;

            if (_sqlConn == null || _currentConnectionString != connectionString)
                _sqlConn = new SqlConnection(connectionString);
            _currentConnectionString = connectionString;

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
            if (_dbcatalog != null && !string.IsNullOrWhiteSpace(dbcatalog))
                dbcatalog = "Initial Catalog=" + _dbcatalog + "; ";
            var defaultConnectionString = "Data Source=" + _dbname + "; " + dbcatalog + "User Id=" + _dbuser + "; Password=" + _dbpass + "; Connection Timeout=" + _connectionTimeOut + "; Pooling=" + _poolingDataBase.ToString().ToLower();

            var connectionString = customConnectionString ?? defaultConnectionString;

            if (_sqlConn == null || _currentConnectionString != connectionString)
                _sqlConn = new SqlConnection(connectionString);
            _currentConnectionString = connectionString;

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
                CloseConnection(false);
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
            if (_sqlOracleConn != null && _sqlOracleConn.State != ConnectionState.Closed)
            {
                _sqlOracleConn.Close();
                if (dispose)
                {
                    _sqlOracleComm.Dispose();
                    _sqlOracleConn.Dispose();
                    _sqlOracleComm = null;
                    _sqlOracleConn = null;
                }
            }
            // ReSharper disable once InvertIf
            if (_sqlConn != null && _sqlConn.State != ConnectionState.Closed)
            {
                _sqlConn.Close();
                // ReSharper disable once InvertIf
                if (dispose)
                {
                    _sqlConn.Dispose();
                    _sqlComm.Dispose();
                    _sqlConn = null;
                    _sqlComm = null;
                }
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

        public bool CheckReplyError(ResponseDTO reply)
        {
            
            if (!reply.GotErrorMessages()) return true;
            var errorMessage = "";
            foreach (var msg in reply.Errors)
                errorMessage += msg.Field + " " + msg.Text;
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
            return false;
        }
        public bool CheckReplyWarning(ResponseDTO reply)
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
        /// <param name="enviroment">Ambiente a verificar</param>
        /// <param name="districtCode">Distrito</param>
        /// <param name="userName">Nombre de usuario</param>
        /// <param name="codeProgram">Código del Programa (Ej. MSEWOT, MSO720)</param>
        /// <param name="accessType">Tipo de acceso a verificar (ProgramAccessType.Full, ProgramAccessType.ReviewObly, etc)</param>
        /// <returns></returns>
        public bool CheckUserProgramAccess(string enviroment, string districtCode, string userName, string codeProgram, int accessType)
        {
            SetDBSettings(enviroment);
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

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var dReader = GetQueryResult(query);

            var result = !(dReader == null || dReader.IsClosed || !dReader.HasRows || !dReader.Read());

            CloseConnection();
            return result;
        }

        public static class ProgramAccessType
        {
            public static int Full = 2;
            public static int ReviewOnly = 1;
            public static int AnyAccess = 99;
        }

        public List<EllipseCodeItem> GetItemCodes(string tableType)
        {
            var listItems = new List<EllipseCodeItem>();
            var query = "SELECT * FROM " + dbReference + ".MSF010" + dbLink + " WHERE TABLE_TYPE = '" + tableType + "' AND ACTIVE_FLAG = 'Y'";
            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            var drItemCodes = GetQueryResult(query);

            if (drItemCodes == null || drItemCodes.IsClosed || !drItemCodes.HasRows) return listItems;
            while (drItemCodes.Read())
            {
                var item = new EllipseCodeItem(drItemCodes["TABLE_CODE"].ToString().Trim(), drItemCodes["TABLE_DESC"].ToString().Trim(), drItemCodes["TABLE_TYPE"].ToString().Trim(), drItemCodes["ASSOC_REC"].ToString().Trim());
                listItems.Add(item);
            }

            return listItems;
        }

        public Dictionary<string, string> GetDictionaryItemCodes(string tableType)
        {
            var itemList = GetItemCodes(tableType);
            return itemList.ToDictionary(item => item.code, item => item.description);
        }

        public void SetPostService(string ellipseUser, string ellipsePswd, string ellipsePost, string ellipseDsct, string urlService)
        {
            PostServiceProxy = new PostService(ellipseUser, ellipsePswd, ellipsePost, ellipseDsct, urlService);
        }

        public ResponseDTO InitiatePostConnection()
        {
            if(PostServiceProxy == null)
                throw new Exception("No se puede iniciar un servicio post no establecido");
            return PostServiceProxy.InitConexion();
        }

        public ResponseDTO ExecutePostRequest(string xmlRequest)
        {
            return PostServiceProxy.ExecutePostRequest(xmlRequest);
        }
    }
    

    
    
}

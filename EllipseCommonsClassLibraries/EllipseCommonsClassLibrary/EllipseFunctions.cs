using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web.Services.Ellipse.Post;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using System.Xml.Linq;
using System.Xml.XPath;
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
                dbReference = "ELLIPSE";
            }
            else if(enviroment == EnviromentConstants.EllipseTest)
            {
                _dbname = "EL8TEST";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = "ELLIPSE";
            }
            else if (enviroment == EnviromentConstants.EllipseDesarrollo)
            {
                _dbname = "EL8DESA";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = "ELLIPSE";
            }
            else if(enviroment == EnviromentConstants.EllipseContingencia)
            {
                _dbname = "EL8PROD";
                _dbuser = "SIGCON";
                _dbpass = "ventyx";
                dbLink = "";
                dbReference = "ELLIPSE";
            }
            else if (enviroment == EnviromentConstants.SigcorProductivo)
            {
                _dbname = "SIGCOPRD";
                _dbuser = "CONSULBO";
                _dbpass = "consulbo";
                dbLink = "@DBLELLIPSE8";
                dbReference = "ELLIPSE";
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
            dbReference = "ELLIPSE";
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
            if (serviceType == null)
                serviceType = EnviromentConstants.ServiceType.EwsService;

            if (serviceType.Equals(EnviromentConstants.ServiceType.EwsService))
            {
                try
                {
                    //Primeramente intenta conseguir la URL del archivo local o de red
                    var localFile = Debugger.LocalDataPath + EnviromentConstants.ConfigXmlFileName;
                    var networkFile = EnviromentConstants.UrlServiceFileLocation + EnviromentConstants.ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EnviromentConstants.EllipseProductivo)
                        urlServer = EnviromentConstants.EllipseVarNameServiceProductivo;
                    if (enviroment == EnviromentConstants.EllipseContingencia)
                        urlServer = EnviromentConstants.EllipseVarNameServiceContingencia;
                    if (enviroment == EnviromentConstants.EllipseTest)
                        urlServer = EnviromentConstants.EllipseVarNameServiceTest;
                    if (enviroment == EnviromentConstants.EllipseDesarrollo)
                        urlServer = EnviromentConstants.EllipseVarNameServiceDesarrollo;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    //Si no encuentra el archivo de red por alguna causa utiliza el valor predeterminado
                    if (enviroment == EnviromentConstants.EllipseProductivo)
                        return EnviromentConstants.EllipseUrlServicesProductivo;
                    if (enviroment == EnviromentConstants.EllipseContingencia)
                        return EnviromentConstants.EllipseUrlServicesContingencia;
                    if (enviroment == EnviromentConstants.EllipseTest)
                        return EnviromentConstants.EllipseUrlServicesTest;
                    if (enviroment == EnviromentConstants.EllipseDesarrollo)
                        return EnviromentConstants.EllipseUrlServicesDesarrollo;
                }
                
            }
            else if (serviceType.Equals(EnviromentConstants.ServiceType.PostService))
            {
                try
                {
                    //Primeramente intenta conseguir la URL del archivo local o de red
                    var localFile = Debugger.LocalDataPath + EnviromentConstants.ConfigXmlFileName;
                    var networkFile = EnviromentConstants.UrlServiceFileLocation + EnviromentConstants.ConfigXmlFileName;
                    var doc = File.Exists(localFile) ? XDocument.Load(localFile) : XDocument.Load(networkFile);

                    var urlServer = "";

                    if (enviroment == EnviromentConstants.EllipseProductivo)
                        urlServer = EnviromentConstants.EllipseVarNamePostProductivo;
                    if (enviroment == EnviromentConstants.EllipseContingencia)
                        urlServer = EnviromentConstants.EllipseVarNamePostContingencia;
                    if (enviroment == EnviromentConstants.EllipseTest)
                        urlServer = EnviromentConstants.EllipseVarNamePostTest;
                    if (enviroment == EnviromentConstants.EllipseDesarrollo)
                        urlServer = EnviromentConstants.EllipseVarNamePostDesarrollo;

                    return doc.XPathSelectElement(urlServer + "[1]").Value;
                }
                catch (Exception)
                {
                    if (enviroment == EnviromentConstants.EllipseProductivo)
                        return EnviromentConstants.EllipseUrlPostServicesProductivo;
                    if (enviroment == EnviromentConstants.EllipseContingencia)
                        return EnviromentConstants.EllipseUrlPostServicesContingencia;
                    if (enviroment == EnviromentConstants.EllipseTest)
                        return EnviromentConstants.EllipseUrlPostServicesTest;
                    if (enviroment == EnviromentConstants.EllipseDesarrollo)
                        return EnviromentConstants.EllipseUrlPostServicesDesarrollo;
                }
            }
            throw new NullReferenceException("No se ha encontrado el servidor seleccionado");
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
    public static class EnviromentConstants
    {
        public static class ServiceType
        {
            public static string PostService = "POST";
            public static string EwsService = "EWS";
        }

        public static string ConfigXmlFileName = @"EllipseConfiguration.xml";
        public static string EllipseProductivo = "Productivo";
        public static string EllipseTest = "Test";
        public static string EllipseContingencia = "Contingencia";
        public static string EllipseDesarrollo = "Desarrollo";
        public static string SigcorProductivo = "SIGCOPROD";
        public static string ScadaRdb = "SCADARDB";
        public static string CustomDatabase = "Personalizada";
        public static string UrlServiceFileLocation = @"\\lmnoas02\SideLine\EllipsePopups\Ellipse8\";
        //public static string SigcorTest = "SIGCOTEST";

        public static string EllipseVarNameServiceProductivo = @"/ellipse/webservice/ellprod";//XPath
        public static string EllipseVarNameServiceContingencia = @"/ellipse/webservice/ellcont";//XPath
        public static string EllipseVarNameServiceDesarrollo = @"/ellipse/webservice/elldesa";//XPath
        public static string EllipseVarNameServiceTest = @"/ellipse/webservice/elltest";//XPath

        public static string EllipseUrlServicesProductivo = "http://ews-el8prod.lmnerp01.cerrejon.com/ews/services";
        public static string EllipseUrlServicesContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string EllipseUrlServicesDesarrollo = "http://ews-el8desa.lmnerp03.cerrejon.com/ews/services";
        public static string EllipseUrlServicesTest = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";

        public static string EllipseVarNamePostProductivo = @"/ellipse/url/ellprod";//XPath
        public static string EllipseVarNamePostContingencia = @"/ellipse/url/ellcont";//XPath
        public static string EllipseVarNamePostDesarrollo = @"/ellipse/url/elldesa";//XPath
        public static string EllipseVarNamePostTest = @"/ellipse/url/elltest";//XPath

        public static string EllipseUrlPostServicesProductivo = "http://ellipse-el8prod.lmnerp01.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string EllipseUrlPostServicesContingencia = "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";
        public static string EllipseUrlPostServicesDesarrollo = "http://ellipse-el8desa.lmnerp03.cerrejon.com/ria-Ellipse-8.4.29_31/bind?app=";
        public static string EllipseUrlPostServicesTest = "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.4.31_112/bind?app=";


        public static List<string> GetEnviromentList() {
            // ReSharper disable once UseObjectOrCollectionInitializer
            var enviromentList = new List<string>();
            enviromentList.Add(EllipseProductivo);
            enviromentList.Add(EllipseTest);
            enviromentList.Add(EllipseDesarrollo);
            enviromentList.Add(EllipseContingencia);

            return enviromentList;
        }

        public static void GenerateEllipseConfigurationXmlFile()
        {
            var xmlFile = "";

            xmlFile += @"<?xml version=""1.0"" encoding=""UTF-8""?>";
            xmlFile += @"<ellipse>";
            xmlFile += @"  <env>test</env>";
            xmlFile += @"  <url>";
            xmlFile += @"    <ellprod>" + EllipseUrlPostServicesProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + EllipseUrlPostServicesContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + EllipseUrlPostServicesDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + EllipseUrlPostServicesTest + "</elltest>";
            xmlFile += @"  </url>";
            xmlFile += @"  <webservice>";
            xmlFile += @"    <ellprod>" + EllipseUrlServicesProductivo + "</ellprod>";
            xmlFile += @"    <ellcont>" + EllipseUrlServicesContingencia + "</ellcont>";
            xmlFile += @"    <elldesa>" + EllipseUrlServicesDesarrollo + "</elldesa>";
            xmlFile += @"    <elltest>" + EllipseUrlServicesTest + "</elltest>";
            xmlFile += @"  </webservice>";
            xmlFile += @"</ellipse>";

            try
            {
                var configFilePath = Debugger.LocalDataPath;
                var configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.WriteTextToFile(xmlFile, configFileName, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void GenerateEllipseConfigurationXmlFile(string networkUrl)
        {
            try
            {
                var configFilePath = Debugger.LocalDataPath;
                var configFileName = ConfigXmlFileName;

                FileWriter.CreateDirectory(configFilePath);
                FileWriter.CopyFileToDirectory(configFileName, networkUrl, configFilePath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("GenerateEllipseConfigurationXmlFile", "No se puede crear el archivo de configuración\n" + ex.Message);
                throw;
            }
        }

        public static void DeleteEllipseConfigurationXmlFile()
        {
            try
            {
                FileWriter.DeleteFile(Debugger.LocalDataPath, ConfigXmlFileName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("No se puede eliminar el archivo de configuración",  ex.Message);
                throw;
            }
        }
    }

    public static class DistrictConstants
    {
        public static string DistrictIcor = "ICOR";
        public static string DistrictInstalations = "INST";
        public static string DefaultDistrict = "ICOR";

        public static List<string> GetDistrictList()
        {
            // ReSharper disable once UseObjectOrCollectionInitializer
            var districtList = new List<string>();
            districtList.Add(DistrictIcor);
            districtList.Add(DistrictInstalations);

            return districtList;
        }
    }
    public static class WoTypeMtType
    {
        /// <summary>
        /// Obtiene listado de objeto de Tipo de Orden vs Tipo de mantenimiento (MT Type, MT Desc, OT Type, OT Desc)
        /// </summary>
        /// <returns></returns>
        public static List<WoTypeMtTypeCode> GetWoTypeMtTypeList()
        {
            var typeList = new List<WoTypeMtTypeCode>
                {
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "CA", "CALIBRACION"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "CO", "CAMBIO DE COMPONENTE MAYOR"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "IP", "SERVICIOS E INSPECCIONES (SEIS)"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "IS", "INSPECCIONES"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "LA", "LAVADO"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "LU", "LUBRICACION"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "OH", "OVERHAUL"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "PB", "PRECIO BASE"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "RM", "REPARACION/CAMBIO DE COMPONENTE MENOR"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "RP", "REPARACIONES PROGRAMADAS"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "SN", "SERVICIO NO CONFORME"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AA", "ANALISIS DE ACEITES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AC", "ANALISIS DE COMBUSTIBLE"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AV", "ANALISIS DE VIBRACIONES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "BC", "BASADA EN CONDICION"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "BK", "PRUEBA BAKER"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "DU", "DETECCION ULTRASONICA"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "ET", "CORRIENTES DE EDDY"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "IE", "INSPECCION ESTRUCTURAL"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "IR", "INSPECCION TERMOGRAFICA"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "ME", "MEDICIONES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "MT", "PARTICULAS MAGNETICAS"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "P0", "ANÁLISIS REFRIGERANTE"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "PT", "TINTAS PENETRANTES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "UT", "ULTRASONIDO"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "VI", "VIDEO"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "VT", "INSPECCION VISUAL"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "WR", "WINDROCK"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "RE", "REPARACION"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "A", "ACCIDENTE"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "AT", "ATENTADO"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "CO", "CAMBIO DE COMPONENTE MAYOR"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "DO", "DAÑO OPERACIONAL"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "G", "GARANTIA"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "SE", "SERVICIO IMIS"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "ET", "ESTUDIO TECNICO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "FA", "FABRICACION"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "MC", "CAMBIO DE EQUIPO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "MN", "MONTAJE NUEVO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "RD", "REDISEÑO O MODIFICACIONES"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AF", "ANALISIS DE FALLAS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AR", "ANALISIS DE RESULTADOS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AS", "ACTIVIDADES  DE SIO & MEDIO AMBIENTE"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "FA", "FABRICACION"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "MP", "MOVILIZACIÓN DE COMPONENTES"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "OP", "OPERACION DE EQUIPOS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "SL", "SIN LABOR"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "SM", "SOPORTE AL MANTENIMIENTO")

                };
            return typeList;
        }

        public static Dictionary<string, string> GetPriorityCodeList()
        {
            var dictionaryList = new Dictionary<string, string>
            {
                {"P0", "CRÍTICA - Detener equipo inmediatamente"},
                {"P1", "URGENTE - Programar en ventana en curso"},
                {"P2", "PRIORITARIA - Programar más tardar en ventana siguiente"},
                {"P3", "RUTINA - Programar en el próximo PM"},
                {"P4", "RUTINA - Programar según oportunidad"},
                {"BE", "INST/IMIS - EMERGENCIA - Atención 1h Cierre 7 días"},
                {"B1", "INST/IMIS - ALTA - Atención 48h Cierre 7 días"},
                {"B2", "INST/IMIS - NORMAL - Atención 6 días Cierre 15 días"},
                {"B3", "INST/IMIS - BAJA - Atención 9 días cierre 30 días"},
            };

            return dictionaryList;
        }
        /// <summary>
        /// Obtiene arreglo Dictionary{key, value} con listado de los códigos de Tipo de Orden admitidos {codigo, descripcion}
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetWoTypeList()
        {
            var listType = GetWoTypeMtTypeList();

            var woTypeList = new Dictionary<string, string>();
            foreach (var type in listType.Where(type => !woTypeList.ContainsKey(type.WoTypeCode)))
            {
                woTypeList.Add(type.WoTypeCode, type.WoTypeDesc);
            }
            return woTypeList;
        }


        /// <summary>
        /// Obtiene arreglo Dictionary{key, value} con listado de los códigos de Tipo de Mantenimiento admitidos {codigo, descripcion}
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetMtTypeList()
        {
            var listType = GetWoTypeMtTypeList();

            var mtTypeList = new Dictionary<string, string>();
            foreach (var type in listType.Where(type => !mtTypeList.ContainsKey(type.MtTypeCode)))
            {
                mtTypeList.Add(type.MtTypeCode, type.MtTypeDesc);
            }
            return mtTypeList;
        }
        /// <summary>
        /// Valida la prioridad de una orden/std establecida para MDC
        /// </summary>
        /// <param name="priority">string: código de prioridad</param>
        /// <param name="district">string: distrito al que pertenece la orden-std</param>
        /// <param name="workGroup">string: grupo de trabajo</param>
        /// <returns>true si la prioridad es válida, false si no es válida</returns>
        public static bool ValidatePriority(string priority, string district = null, string workGroup = null)
        {
            if (priority == null)
                return false;

            priority = priority.Trim();

            if (district == null || district.Trim().Equals("ICOR"))
            {
                if (priority.Equals("P0") || priority.Equals("P1") || priority.Equals("P2") || priority.Equals("P3") || priority.Equals("P4"))
                    return true;
            }
            else if (district.Trim().Equals("INST"))
            {
                if (workGroup != null && (workGroup.Trim().Equals("AAPREV") && (priority.Equals("P0") || priority.Equals("P1") || priority.Equals("P2") || priority.Equals("P3") || priority.Equals("P4"))))
                    return true;
                if (priority.Equals("B1") || priority.Equals("B2") || priority.Equals("B3"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Valida la relación de Tipo de Orden vs Tipo de Mantenimiento de una orden/std establecida para MDC
        /// </summary>
        /// <param name="woType">string: Tipo de Orden</param>
        /// <param name="mtType">string: Tipo de Mantenimiento</param>
        /// <returns>true si la relación es válida, false si no es válida</returns>
        public static bool ValidateWoMtTypeCode(string woType, string mtType)
        {
            if (woType == null || mtType == null)
                return false;

            woType = woType.Trim();
            mtType = mtType.Trim();

            var typeList = GetWoTypeMtTypeList();

            return typeList.Any(type => woType == type.WoTypeCode && mtType == type.MtTypeCode);
        }

        public class WoTypeMtTypeCode
        {
            public string MtTypeCode;
            public string MtTypeDesc;
            public string WoTypeCode;
            public string WoTypeDesc;

            public WoTypeMtTypeCode(string mtTypeCode, string mtTypeDesc, string woTypeCode, string woTypeDesc)
            {
                MtTypeCode = mtTypeCode;
                MtTypeDesc = mtTypeDesc;
                WoTypeCode = woTypeCode;
                WoTypeDesc = woTypeDesc;
            }
        }


    }
    public static class GroupConstants
    {
        public static List<WorkGroup> GetWorkGroupList()
        {
            var groupList = new List<WorkGroup>
            {
                new WorkGroup("AAPREV", "Mantenimiento Aire Acondicionado", "INST", "MINA", "INST"),
                new WorkGroup("BASE9", "OPERACION Y ATENCION BASE 9", "SOP", "ENERGIA", "ICOR"),
                new WorkGroup("CALLCEN", "Call Center", "INST", "IMIS", "INST"),
                new WorkGroup("CARGUE2", "Cargue 2", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("CAT2401", "U.A.S. CAMIONES CAT240", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("CAT789C", "Camion 190 ton cat789C", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("CTC", "INSPECCIONES VIAS Y MANTTO. DEL CTC", "MDC", "FFCC", "ICOR"),
                new WorkGroup("EH320", "U.A.S. CAMIONES DE 320 MINA NORTE", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("ELIVIA1", "GRUPO DE TRABAJO DE LIVIANOS", "SOP", "LIVIANOS", "ICOR"),
                new WorkGroup("EMEDIA1", "GRUPO DE TRABAJO MEDIANOS", "SOP", "MEDIANOS", "ICOR"),
                new WorkGroup("EQAUXV", "MTTO.EQUIPO VIAS FFCC", "MDC", "FFCC", "ICOR"),
                new WorkGroup("GI&T", "Grupo de Inspección y Tecnología", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("GRUAS", "UAS GRUAS Y MANEJADORES DE LLANTA", "SOP", "GRUAS", "ICOR"),
                new WorkGroup("IALIAL1", "GRUPO DE SEIS Transformadores y Distrib.", "SOP", "ENERGIA", "ICOR"),
                new WorkGroup("IAPTAL1", "GRUPO SEIS Taller & soporte SER", "SOP", "ENERGIA", "ICOR"),
                new WorkGroup("IBOMBA1", "SEIS DE BOMBAS Super. de Servicio", "SOP", "ENERGIA", "ICOR"),
                new WorkGroup("ICARROS", "MTTO.VAGONES", "MDC", "FFCC", "ICOR"),
                new WorkGroup("LLANTAS", "TALLER DE LLANTAS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("LUBRICA", "LABORES TALLER DE LUBRICACION", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("L1350", "CARGADORES LETORNEAU L1350", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("MCARGA", "UAS MANEJO DE CARGA - OPERADORES", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("MTIL17", "MANTENIMIENTO INDUSTRIAL&PLANTA-AGUA", "INST", "MINA", "ICOR"),
                new WorkGroup("MTOLOC", "MTTO. LOCOMOTORAS", "MDC", "FFCC", "ICOR"),
                new WorkGroup("MTTOSOP", "UAS EQUIPO DE SOPORTE", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("ORUGAS", "TRACTORES DE ORUGAS D9L Y D11N", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("PCSERVI", "PLANTA DE CARBON", "MDC", "PTAS", "ICOR"),
                new WorkGroup("PHIDCAS", "PALAS HIDRAULICAS MINA", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("PHS", "UAS PALAS ELECTRICAS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("PPELCOP", "TALLER ELECTRICO/ELECTRONICO", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOAA", "MANTENIMIENTO DE AIRES PBV", "MDC", "PBV", "INST"),
                new WorkGroup("PTOBAND", "MANTTO.BANDAS TRANSPORTADORAS CARBON", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOCAR", "MANTTO MECANICO EQUIPOS DE MANCARB", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOCE", "CARGA Y ESTIBA PUERTO BOLIVAR", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOCP8", "GRUPO SEIS CONTRATO REDES Y MONTAJES", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOINS", "GRUPO DE INSPECCIONES ESTRUCTURALES PBV", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOMET", "CONTRATISTA METALISTERIA Y PINTURA PBV", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOMIN", "MANTENIMIENTO INSTALACIONES PBV", "MDC", "PBV", "INST"),
                new WorkGroup("PTOOM1", "GRUPO DE MANTENIMIENTO MOTORES DIESEL", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOPRED", "GRUPO PREDICTIVOS PUERTO BOLIVAR", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOSEG", "GRUPO CONTROLES CRITICOS", "MDC", "PBV", "ICOR"),
                new WorkGroup("PTOTM", "TALLER MECANICO/PLANTA AGUA -PBV", "MDC", "PBV", "ICOR"),
                new WorkGroup("RDCAMPO", "GRUPO DE MANTENIMIENTO MOTORES EN CAMPO", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RDCOMPO", "REPARACIÓN DE COMPONENTES MENORES DE MOT", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RDIESEL", "RECONSTRUCCION DE MOTORES", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RECHID", "GRUPO REC.HIDRAULICA DE PRONOSTICOS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RHMENOR", "RECONSTRUIR COMP.MENORES HIDRAULICOS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RELECII", "GRUPO PARA REPARACION DE COMPO. PROGRAMA", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("REMAQ", "MAQUINAS HERRAMIENTAS RECONSTRUCCION", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RERODA", "RECONSTRUCCION TREN DE RODAJE", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("RESOLD", "RECONSTRUCCION SOLDADURA", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("T&A", "TALLER MANTTO.ROLDAN -TRAFICO & ADUANA", "MDC", "PBV", "ICOR"),
                new WorkGroup("TANQ777", "UAS DE TANQUEROS Y TRAILLAS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("TRACLLA", "UAS DE TRACTORES DE LLANTAS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("VIAS", "UAS DE MOTONIVELADORAS", "MNTTO", "MINA", "ICOR"),
                new WorkGroup("VIASM", "MANTENIMIENTO VIAS MINA", "MDC", "FFCC", "ICOR"),
                new WorkGroup("VIASP", "MANTENIMIENTO DE VIAS PUERTO", "MDC", "FFCC", "ICOR")
            };
            return groupList;
        }

       
        public class WorkGroup
        {
            public string Name;
            public string Description;
            public string Area;
            public string Details;
            public string DistrictCode;

            public WorkGroup(string name, string description, string area, string details, string districtCode)
            {
                Name = name;
                Description = description;
                Area = area;
                Details = details;
                DistrictCode = districtCode;
            }
        }
        
    }

    public class ReplyMessage
    {
        public string[] Errors;
        public string[] Warnings;
        public string Message;
    }
}

using EllipseScreenLibrary.ScreenService;
using System;
using System.Collections.Generic;
using System.Text;
using EllipseScreenLibrary.Properties;



namespace EllipseScreenLibrary
{
    public class Ellipse
    {
        /// <summary>
        /// Objeto MSO que permite interactuar con las pantallas de Ellipse.
        /// </summary>
        public ScreenDTO MSO { get; set; }
        /// <summary>
        /// Listado de Campos que seran enviados a una pantalla de Ellipse despues de la ejecucion del metodo ExecuteMSO.
        /// </summary>
        /// <seealso cref="ExecuteMSO(String, bool)">
        public List<ScreenNameValueDTO> ScreenFields { get; set; }
        
        private ScreenSubmitRequestDTO ScreenSubmit { get; set; }
        private ScreenService.ScreenService Proxy { get; set; }
        private OperationContext Context { get; set; }

        public static String OK = "1";
        public static String CANCEL = "3";

        public static String TRANSMIT = "1";
        public static String F1_KEY = "2";
        public static String F2_KEY = "3";
        public static String F3_KEY = "4";
        public static String F4_KEY = "5";
        public static String F5_KEY = "6";
        public static String F6_KEY = "7";
        public static String F7_KEY = "8";
        public static String F8_KEY = "9";
        public static String F9_KEY = "10";
        public static String F10_KEY = "11";

        /// <summary>
        /// Constructor Basico.
        /// </summary>
        public Ellipse() { }

        /// <summary>
        /// Constructor que inicializa el contexto de los servicio de Ellipse.
        /// </summary>
        /// <param name="District">El distrito del usuario que se va a conectar a Ellipse.</param>
        /// <param name="Position">La posición del usuario que se va a conectar a Ellipse.</param>
        /// <param name="Instances">Numero de instancias concurrentes que puede tener el objeto. Máximo 100.</param>
        /// <param name="ShowWarnings">Permite excluir los mensajes de advertencia en las ventanas de MSO. Posibles valores true o false.</param>
        public Ellipse(String District, String Position, int Instances, bool ShowWarnings)
        {
            Context = new OperationContext();
            Context.district = District;
            Context.position = Position;
            Context.maxInstances = Instances;
            Context.returnWarnings = ShowWarnings;
        }

        public static void DebugScreen(ScreenService.ScreenSubmitRequestDTO request, ScreenService.ScreenDTO reply, string filename)
        {
            //var requestJson = new JsonSerializer.Serialize(request.screenFields);
            //var replyJson = new JavaScriptSerializer().Serialize(reply.screenFields);
            //
            // = "c:\\ellipse\\" + @"debugger\";
            //ppendTextToFile(requestJson, "ScreenRequest.txt", filePath);
            //ppendTextToFile(replyJson, "ScreenReply.txt", filePath);
        }
        /// <summary>
        /// Crea la conexión con el servicio ScreenService de Ellipse
        /// </summary>
        /// <param name="URL">La URL del servicio dependiendo del ambiente al cual se requiere conectar.</param>
        public void InitMSOInstance(String URL)
        {
            Proxy = new ScreenService.ScreenService();
            Proxy.Url = URL+ "ScreenService";

        }

        /// <summary>
        /// Realiza la inicialización de los campos a llenar en una pantalla. Esto se debe hace siempre y cuando se requiera enviar nuevos campos al método ExecuteMSO.
        /// </summary>
        /// <seealso cref="ExecuteMSO(String, bool)">
        public void InitScreenFields()
        {
            ScreenFields = new List<ScreenNameValueDTO>();
        }

        /// <summary>
        /// Verifica si el objeto MSO se encuentra correctamente.
        /// </summary>
        /// <returns>Retorna true si el objeto se encuentra apto para procesar o verificar su contenido, retorna false en el caso contrario.</returns>
        public bool IsMSOCorrect()
        {
            if (MSO != null && MSO.mapName != null && !MSO.mapName.Equals(""))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Verifica si el objeto MSO luego de su ejecución encontró algún error.
        /// </summary>
        /// <returns>Retorna true si existe algun error, retorna false en el caso contrario.</returns>
        public bool IsMSOError()
        {
            if (IsMSOCorrect())
            {
                if (!MSO.message.Trim().Equals(""))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Verifica si el objeto MSO luego de su ejecución encontró algún error diferente a una advertencia.
        /// </summary>
        /// <returns>Retorna true si existe algun error sin adevretencia, retorna false en el caso contrario.</returns>
        public bool IsMSOErrorNotWarning()
        {
            if (IsMSOCorrect())
            {
                if (!MSO.message.Trim().Equals("") && !MSO.message.ToUpper().StartsWith("W"))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Verifica si el objeto MSO se encuentra en la pantalla seleccionada.
        /// </summary>
        /// <param name="ScreenApp">Es la pantalla que se requiere validar.</param>
        /// <returns>Retorna true si el objeto MSO se encuentra en la pantalla refrenciada, retorna false en el caso contrario.</returns>
        public bool IsScreenNameCorrect(String ScreenApp)
        {
            if (IsMSOCorrect())
            {
                if (MSO.mapName.Equals(ScreenApp))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Metodo que permite enviar una petición a los servicio de Ellipse.
        /// </summary>
        /// <param name="ScreenKey">Es el comando que se requiere ejecutar.</param>
        /// <param name="IsNeedFields">Si el MSO no requiere parametro este campo se debe enviar en false, en el caso contrario se debe enviar en true.</param>
        public void ExecuteMSO(String ScreenKey, bool IsNeedFields)
        {
            if (IsMSOCorrect())
            {
                if (IsNeedFields)
                {
                    ScreenSubmit.screenFields = ScreenFields.ToArray();
                }
                else
                {
                    ScreenSubmit.screenFields = null;
                }
                ScreenSubmit.screenKey = ScreenKey;
                MSO = Proxy.submit(Context, ScreenSubmit);
            }
        }

        public bool RevertOperation()
        {
            //forzar inicio de pantalla
            var requestScreen = new ScreenService.ScreenSubmitRequestDTO();
            var prevProgram = "0";
            var actualProgram = "1";

            if (Proxy == null || Context == null)
                return false;
            while (!actualProgram.Equals(prevProgram))
            {
                try
                {
                    requestScreen.screenFields = null;
                    requestScreen.screenKey = "3";
                    var replyScreen = Proxy.submit(Context, requestScreen);
                    prevProgram = actualProgram;
                    actualProgram = replyScreen.mapName;
                }
                catch (Exception ex)
                {
                    //Debugger.LogError("RibbonEllipse:revertOperation(Screen.OperationContext, Screen.ScreenService)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                    prevProgram = actualProgram;

                }
            }
            return false;
        }

        /// <summary>
        /// Metodo que permite cerrar las conexiones.
        public void ClosedConnection()
        {
            try
            {
                Proxy.Abort();
                Proxy.Dispose();
            }
            catch (Exception ex)
            {

            }
        }




        /// <summary>
        /// Metodo que permite ejecutar un programa en Ellipse.
        /// </summary>
        /// <param name="Screen">Programa que se va a Ejecutar.</param>
        /// <param name="ScreenApp">Pantalla a la cual debe ingresar de forma inmediata la ejecucion del Programa de Ellipse.</param>
        public void ExecuteScreen(String Screen, String ScreenApp)
        {
            MSO = Proxy.executeScreen(Context, Screen);
            ScreenSubmit = new ScreenSubmitRequestDTO();
            if (!IsScreenNameCorrect(ScreenApp))
            {
                ExecuteMSO(Ellipse.F2_KEY, false);
                MSO = Proxy.executeScreen(Context, Screen);
            }
        }

        /// <summary>
        /// Metodo que permite ejecutar un programa en Ellipse creando un Nuevo Contexto de Operacion.
        /// </summary>
        /// <param name="Screen">Programa que se va a Ejecutar.</param>
        /// <param name="ScreenApp">Pantalla a la cual debe ingresar de forma inmediata la ejecucion del Programa de Ellipse.</param>
        public void ExecuteScreenNewContext(String Screen, String ScreenApp)
        {
            String District = Context.district;
            String Position = Context.position;
            int Instances = Context.maxInstances;
            bool ShowWarnings = Context.returnWarnings;

            Context = new OperationContext();
            Context.district = District;
            Context.position = Position;
            Context.maxInstances = Instances;
            Context.returnWarnings = ShowWarnings;

            MSO = Proxy.executeScreen(Context, Screen);
            ScreenSubmit = new ScreenSubmitRequestDTO();
            if (!IsScreenNameCorrect(ScreenApp))
            {
                ExecuteMSO(Ellipse.F2_KEY, false);
                MSO = Proxy.executeScreen(Context, Screen);
            }
        }

        /// <summary>
        /// Estable el valor de un parametro en una pantalla de Ellipse.
        /// </summary>
        /// <param name="FieldName">Nombre del parametro que se desea cambiar.</param>
        /// <param name="Value">Valor que se desea enviar al parametro de la pantalla de Ellipse.</param>
        public void SetMSOFieldValue(String FieldName, String Value)
        {
            ScreenNameValueDTO Field = new ScreenNameValueDTO();
            Field.fieldName = FieldName;
            Field.value = Value;
            ScreenFields.Add(Field);
        }

        /// <summary>
        /// Estable el valor de un parametro en una pantalla de Ellipse, validando que el campo Value no sea Nulo.
        /// </summary>
        /// <param name="FieldName">Nombre del parametro que se desea cambiar.</param>
        /// <param name="Value">Valor que se desea enviar al parametro de la pantalla de Ellipse.</param>
        public void SetMSOFieldValueValidateNull(String FieldName, String Value)
        {
            if (Value != null)
            {
                ScreenNameValueDTO Field = new ScreenNameValueDTO();
                Field.fieldName = FieldName;
                Field.value = Value;
                ScreenFields.Add(Field);
            }
        }

        /// <summary>
        /// Estable el valor de un parametro en una pantalla de Ellipse, validando que el campo Value no sea Nulo ni Vacio.
        /// </summary>
        /// <param name="FieldName">Nombre del parametro que se desea cambiar.</param>
        /// <param name="Value">Valor que se desea enviar al parametro de la pantalla de Ellipse.</param>
        public void SetMSOFieldValueValidateEmpty(String FieldName, String Value)
        {
            if (Value != null && !Value.Trim().Equals(""))
            {
                ScreenNameValueDTO Field = new ScreenNameValueDTO();
                Field.fieldName = FieldName;
                Field.value = Value;
                ScreenFields.Add(Field);
            }
        }

        /// <summary>
        /// Obtiene el valor de un parametro en una pantalla de Ellipse.
        /// </summary>
        /// <param name="FieldName">Nombre del parametro que se desea buscar su valor.</param>
        /// <returns>Retorna el valor que se encuentre establecido en el parametro.</returns>
        public String GetMSOFieldValue(String FieldName)
        {
            String Value = "";
            if (IsMSOCorrect())
            {
                ScreenFieldDTO[] ListFields = MSO.screenFields;
                foreach (ScreenFieldDTO Field in ListFields)
                {
                    if (Field.fieldName.Trim().Equals(FieldName.Trim()))
                    {
                        Value = Field.value.Trim();
                        break;
                    }
                }

            }
            return Value;
        }

        /// <summary>
        /// Metodo que permite la recuperacion de errores de un objeto MSO.
        /// </summary>
        /// <returns>Retorna el mensaje de Error arrojado por el objeto MSO.</returns>
        public String GetMSOError()
        {
            String Error = "";
            if (IsMSOCorrect())
            {
                if (IsMSOError())
                {
                    Error = "<" + MSO.mapName + "><" + MSO.currentCursorFieldName + "> " + MSO.message;
                }
            }
            return Error;
        }

        /// <summary>
        /// Metodo que retorna el valor de cadena de un objeto cualquiera.
        /// </summary>
        /// <param name="obj">Cualquier tipo de Objeto</param>
        /// <returns>Retorna su valor de cadena siempre que el objeto no sea nulo.</returns>
        public static String GetStringValueFromObject(object obj)
        {
            if (obj != null)
            {
                return obj.ToString();
            }
            return "";
        }

        /// <summary>
        /// Metodo que retorna el valor de cadena de un objeto cualquiera sin espacios.
        /// </summary>
        /// <param name="obj">Cualquier tipo de Objeto</param>
        /// <returns>Retorna su valor de cadena siempre que el objeto no sea nulo.</returns>
        public static String GetTrimStringValueFromObject(object obj)
        {
            if (obj != null)
            {
                return obj.ToString().Trim();
            }
            return "";
        }
    }
}

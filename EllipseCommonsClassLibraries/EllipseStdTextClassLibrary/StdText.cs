using System;
using System.Linq;
using SharedClassLibrary.Utilities;

namespace EllipseStdTextClassLibrary
{
    public static class StdText
    {
        public static bool DebugErrors = false; //Muestra las alertas de debugger. Default: false
        private static int _lineLength = 60;

        /// <summary>
        /// Establece la longitud máxima por línea para los StdText. Ellipse gestiona los 60 de forma predeterminada.
        /// </summary>
        /// <param name="lineLenght"></param>
        public static void SetLineLength(int lineLenght)
        {
            _lineLength = lineLenght;
        }
        /// <summary>
        /// Obtiene el texto de un elemento. En la obtención se reagrupa el texto obviando la división de caracteres por línea para casi todos los casos
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextService.OperationContext: Contexto del servicio stdText. Puede crear uno mediante el uso de StdText.getStdTextOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <returns>string: Texto del elemento ingresado. Retorna vacío si el Id no existe</returns>
        public static string GetText(string urlService, StdTextService.OperationContext opContext, string stdTextId)
        {
            using (var stdTextService = new StdTextService.StdTextService())
            {
                try
                {


                    var requestSt = new StdTextService.StdTextServiceGetTextRequestDTO();
                    var requiredAtributes = new StdTextService.StdTextServiceGetTextRequiredAttributesDTO();
                    //se cargan los parámetros de la orden
                    stdTextService.Url = urlService + "/StdText";

                    //se cargan los parámetros de la solicitud
                    requestSt.stdTextId = stdTextId;
                    //se envía la acción
                    var replySt = stdTextService.getText(opContext, requestSt, requiredAtributes, "");

                    var fullText = "";

                    foreach (var line in replySt.replyElements.SelectMany(block => block.textLine))
                    {
                        fullText = fullText + line;
                        if (line.Length < _lineLength)
                            fullText = fullText + "\n";
                    }

                    return fullText;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("StdText:getText(String, StdTextService.OperationContext, string)", ex.Message);
                    throw;
                }
            }
        }

        /// <summary>
        /// Obtiene el texto de un elemento. Se mantiene la división de caracteres por línea
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextCustomService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <returns>string: Texto del elemento ingresado. Retorna vacío si el Id no existe</returns>
        public static string GetText(string urlService, StdTextCustomService.OperationContext opContext, string stdTextId)
        {
            try
            {
                // ejecuta las acciones del servicio
                using (var proxySt = new StdTextCustomService.StdTextCustomService
                {
                    Url = urlService + "/StdTextCustom"
                })
                {
                    // se envía la acción
                    var replySt = proxySt.getExtendedText(opContext, stdTextId);

                    return replySt;
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("StdText:getText(String, StdTextService.OperationContext, string)", ex.Message);
                throw;
            }
        }


        /// <summary>
        /// Establece el texto para un stdTextID. Actualmente arroja una excepción por el LINE COUNT. Usar el setCustomText
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <param name="text">string: Texto a ingresar para el stdTextId</param>
        /// <returns>bool: True si se culmina sin problemas</returns>
        [Obsolete("setText está obsoleto. Genera errores y no es utilizado. Utilizar setText(StdTextCustomService.OperationContext) en su lugar",
            true)]
        public static bool SetText(string urlService, StdTextService.OperationContext opContext, string stdTextId, string text)
        {
            using (var stdTextService = new StdTextService.StdTextService())
            {
                try
                {


                    var requestSt = new StdTextService.StdTextServiceSetTextRequestDTO();

                    //se cargan los parámetros de la orden
                    stdTextService.Url = urlService + "/StdText";

                    //se cargan los parámetros de la solicitud
                    requestSt.stdTextId = stdTextId;
                    var splittedText = MyUtilities.SplitText(text, _lineLength);
                    requestSt.textLine = splittedText;
                    //se envía la acción
                    stdTextService.setText(opContext, requestSt);
                    return true;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("StdText:setText(String, StdTextService.OperationContext, string, string)", ex.Message);
                    throw;
                }
            }
        }

        /// <summary>
        /// Establece el texto para un control de texto de tipo stdText
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextCustomService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <param name="text">string: Texto a ingresar para el stdTextId</param>
        /// <returns>bool: True si se culmina sin problemas</returns>
        public static bool SetText(string urlService, StdTextCustomService.OperationContext opContext, string stdTextId, string text)
        {
            try
            {
                if (text == null)
                    text = "";
                using (var stdTextCustomService = new StdTextCustomService.StdTextCustomService {Url = urlService + "/StdTextCustom"})
                {

                    //text = SpliceText(text, _lineLength);
                    var arrayText = MyUtilities.SplitText(text, _lineLength);
                    //se envía la acción

                    //proxySt.setExtendedText(opContext, stdTextId, text)
                    stdTextCustomService.setExtendedTextWithArray(opContext, stdTextId, arrayText);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError(
                    "StdText:setText(String, StdTextCustomService.OperationContext, string, string)", ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Establece el texto para el encabezado de un elemento stdText de id stdTextId
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextCustomService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <param name="text">string: Texto a ingresar para el stdTextId</param>
        /// <returns>bool: True si se culmina sin problemas</returns>
        public static bool SetHeader(string urlService, StdTextService.OperationContext opContext, string stdTextId, string text)
        {
            using (var stdTextService = new StdTextService.StdTextService())
            {
                try
                {

                    var requestSt = new StdTextService.StdTextServiceSetHeadingRequestDTO();

                    //se cargan los parámetros de la orden
                    stdTextService.Url = urlService + "/StdText";

                    //se cargan los parámetros de la solicitud
                    requestSt.stdTextId = stdTextId;
                    requestSt.headingLine = text;

                    //
                    //se envía la acción
                    stdTextService.setHeading(opContext, requestSt);

                    return true;
                }
                catch (Exception ex)
                {
                    Debugger.LogError("StdText:setHeading(String, StdTextService.OperationContext, string, string)", ex.Message);
                    throw;
                }
            }
        }
        /// <summary>
        /// Establece el texto para el encabezado de un elemento stdText de id stdTextId
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextCustomService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <param name="text">string: Texto a ingresar para el stdTextId</param>
        /// <returns>bool: True si se culmina sin problemas</returns>
        public static bool SetHeader(string urlService, StdTextCustomService.OperationContext opContext, string stdTextId, string text)
        {
            try
            {
                using (var stdTextCustomService = new StdTextCustomService.StdTextCustomService {Url = urlService + "/StdTextCustom"})
                {

                    //se envía la acción
                    stdTextCustomService.setExtendedTextHeading(opContext, stdTextId, text);

                    return true;
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("StdText:setHeading(String, StdTextService.OperationContext, string, string)", ex.Message);
                throw;
            }
        }
        /// <summary>
        /// Obtiene el encabezado de un elemento a partir de su stdTextId
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextService.OperationContext: Contexto del servicio stdText. Puede crear uno mediante el uso de StdText.getNewOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <returns>string: Encabezado del elemento ingresado</returns>
        public static string GetHeader(string urlService, StdTextService.OperationContext opContext, string stdTextId)
        {
            using (var stdTextService = new StdTextService.StdTextService())
            {

                var requestParameters = new StdTextService.StdTextServiceGetHeadingRequestDTO();

                //se cargan los parámetros de la orden
                stdTextService.Url = urlService + "/StdText";

                //se cargan los parámetros de la solicitud
                requestParameters.stdTextId = stdTextId;
                //
                //se envía la acción

                var replySt = stdTextService.getHeading(opContext, requestParameters);

                return replySt.headingLine;
            }

        }
        /// <summary>
        /// Obtiene el encabezado de un elemento a partir de su stdTextId
        /// </summary>
        /// <param name="urlService">string: URL del servicio (EFunctions.getServicesURL(drpEnvironment.SelectedItem.Label))</param>
        /// <param name="opContext">StdTextCustomService.OperationContext: Contexto del servicio stdTextCustom. Puede crear uno mediante el uso de StdText.getCustomOpContext()</param>
        /// <param name="stdTextId">string: Tipo[2], Distrito[4], Id[8] (Ej. WOICORIF039909)</param>
        /// <returns>string: Encabezado del elemento ingresado</returns>
        public static string GetHeader(string urlService, StdTextCustomService.OperationContext opContext, string stdTextId)
        {
            using (var stdTextCustomService = new StdTextCustomService.StdTextCustomService {Url = urlService + "/StdTextCustom"})
            {
                //se envía la acción
                return stdTextCustomService.getExtendedTextHeading(opContext, stdTextId);
            }
        }
        /// <summary>
        /// Crea un nuevo operador de contexto para los métodos de la clase
        /// </summary>
        /// <param name="district">string: Distrito donde se va a crear el contexto</param>
        /// <param name="position">string: Posición donde se va a crear el contexto</param>
        /// <param name="maxInstances">int: Número máximo de instancias</param>
        /// <param name="returnWarnings">bool: True no ignora las advertencias</param>
        /// <returns></returns>
        public static StdTextService.OperationContext GetStdTextOpContext(string district, string position,
            int maxInstances, bool returnWarnings)
        {
            var opContext = new StdTextService.OperationContext
            {
                district = district,
                position = position,
                maxInstances = maxInstances,
                maxInstancesSpecified = true,
                returnWarnings = returnWarnings,
                returnWarningsSpecified = true
            };

            return opContext;
        }

        /// <summary>
        /// Crea un nuevo operador de contexto para los métodos de la clase
        /// </summary>
        /// <param name="district">string: Distrito donde se va a crear el contexto</param>
        /// <param name="position">string: Posición donde se va a crear el contexto</param>
        /// <param name="maxInstances">int: Número máximo de instancias</param>
        /// <param name="returnWarnings">bool: True no ignora las advertencias</param>
        /// <returns></returns>
        public static StdTextCustomService.OperationContext GetCustomOpContext(string district, string position,
            int maxInstances, bool returnWarnings)
        {
            var opContext = new StdTextCustomService.OperationContext
            {
                district = district,
                position = position,
                maxInstances = maxInstances,
                maxInstancesSpecified = true,
                returnWarnings = returnWarnings,
                returnWarningsSpecified = true
            };

            return opContext;
        }
        
    }
}
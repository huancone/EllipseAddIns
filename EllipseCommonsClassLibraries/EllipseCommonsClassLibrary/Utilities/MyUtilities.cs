using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using EllipseCommonsClassLibrary.Classes;
using Screen = EllipseCommonsClassLibrary.ScreenService;


namespace EllipseCommonsClassLibrary.Utilities
{
    public class MyUtilities
    {
        /// <summary>
        ///     Obtiene una cadena con el nombre de una variable dada
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item">Variable a obtener el nombre</param>
        /// <returns>string: nombre de una variable (Ej. int numeroEntero = 3; //output: numeroEntero)</returns>
        public static string GetVarName<T>(T item) where T : class
        {
            return typeof(T).GetProperties()[0].Name;
        }

        /// <summary>
        ///     Divide el text ingresado en un arreglo string[] teniendo en cuenta los saltos de línea y la longitud de línea
        ///     máxima deseada
        /// </summary>
        /// <param name="text">string: Texto a segmentar</param>
        /// <param name="chunkSize">string: Tamaño del segmento</param>
        /// <returns>string[]: arreglo con la segmentación del texto ingresado</returns>
        public static string[] SplitText(string text, int chunkSize)
        {
            var textArray = new List<string>();
            if (text == null)
                return null;

            if (!text.Contains("\n") && text.Length <= chunkSize)
            {
                textArray.Add(text);
            }
            else
            {
                var charArray = text.ToCharArray();
                var iChunk = 0;
                var newLine = "";
                for (var i = 0; i < charArray.Length; i++)
                {
                    if (iChunk >= chunkSize || charArray[i] == '\n')
                    {
                        textArray.Add(newLine);
                        newLine = "";
                        iChunk = 0;
                        if (charArray[i] == '\n')
                            i++;
                    }

                    newLine = newLine + charArray[i];
                    iChunk++;
                }

                if (newLine.Length > 0)
                    textArray.Add(newLine);
            }

            return textArray.ToArray();
        }


        /// <summary>
        ///     Obtiene una lista con los campos de Key, Value concatenados con el conector dado (Ej. [Key, Value] = ["codigo",
        ///     "valor"], connector = " - ", resultado = "codigo - valor")
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="source"></param>
        /// <param name="connector"></param>
        /// <returns></returns>
        public static List<string> ConcatToStringDictionaryKeyValue<TKey, TValue>(Dictionary<TKey, TValue> source,
            string connector)
        {
            var list = source.Select(entry => entry.Key + connector + entry.Value).ToList();

            return list;
        }

        /// <summary>
        ///     Obtiene un listado separado por el separador dado en forma de cadena de texto (Ej: lista{valor1, valor2, valor3} =>
        ///     string = "valor1,valor2,valor3")
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listValues">IEnumerable(T): Arreglo Enumerable para el listado</param>
        /// <param name="separator">
        ///     string: Indica cuál texto/símbolo será usado como separador de lista (Ej: separator = ",";
        ///     stringList = "valor1, valor2, valor3"
        /// </param>
        /// <param name="quotation">
        ///     string: Encierra el valor de la lista con este text (Ej: quotation = "'", valorLista =
        ///     "'valor'"; quotation = "***", valorLista = "***valor***")
        /// </param>
        /// <returns></returns>
        public static string GetListInSeparator<T>(IEnumerable<T> listValues, string separator, string quotation = null)
        {
            if (listValues == null)
                return null;
            var enumerable = listValues as IList<T> ?? listValues.ToList();

            if (!enumerable.Any())
                return null;
            var stringList = enumerable.Aggregate("",
                (current, value) => current + quotation + value + quotation + separator);

            return stringList.Substring(0, stringList.Length - 1);
        }

        /// <summary>
        ///     Obtiene el valor verdadero según el criterio de entrada. Si value es TRUE, VERDADERO, Y, YES, SI, ó 1
        /// </summary>
        /// <param name="value">Object: valor a analizar</param>
        /// <param name="nullable">bool: indica si se asume nulo/vacío como verdadero. True null es true, false null es false</param>
        /// <returns>boolean: true si value es TRUE, VERDADERO, Y, YES, SI ó 1</returns>
        public static bool IsTrue(object value, bool nullable = false)
        {
            try
            {
                if (value == null)
                    return nullable;
                var stringValue = Convert.ToString(value);
                if (string.IsNullOrWhiteSpace(stringValue))
                    return nullable;

                stringValue = stringValue.Trim();
                return stringValue.ToUpper().Equals("TRUE") ||
                       stringValue.ToUpper().Equals("VERDADERO") ||
                       stringValue.ToUpper().Equals("Y") ||
                       stringValue.ToUpper().Equals("YES") ||
                       stringValue.ToUpper().Equals("SI") ||
                       stringValue.ToUpper().Equals("S") ||
                       stringValue.ToUpper().Equals("1");
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        ///     Convierte una cadena de tipo "key - separator - value" en keyValuePair Ej. "23 - Description" ->
        ///     KeyValuePair(string, string){"23", "Description"}
        /// </summary>
        /// <param name="keyValue">string: cadena de tipo llave/descripción (Ej. "23 - Description")</param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns>KeyValuePair(string, string)</returns>
        public static KeyValuePair<string, string> GetCodeKeyValue(string keyValue, string separator = " - ")
        {
            //return nulo
            if (keyValue == null)
                return new KeyValuePair<string, string>();
            //return empty
            if (keyValue.Equals(""))
                return new KeyValuePair<string, string>("", "");
            //return key,value

            if (keyValue.Contains(separator))
                return new KeyValuePair<string, string>(
                    keyValue.Substring(0, keyValue.IndexOf(separator, StringComparison.Ordinal)),
                    keyValue.Substring(keyValue.IndexOf(separator, StringComparison.Ordinal) + separator.Length));

            //return key,empty
            return new KeyValuePair<string, string>(keyValue, "");
        }

        /// <summary>
        ///     Obtiene una cadena con el código/llave a partir de una cadena código-descripción (Ej. Ingresa "03 - Acción" :::
        ///     Obtiene "03")
        /// </summary>
        /// <param name="keyValue"></param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns></returns>
        public static string GetCodeKey(string keyValue, string separator = " - ")
        {
            var codeKeyValue = GetCodeKeyValue(keyValue, separator);
            return codeKeyValue.Key;
        }

        /// <summary>
        ///     Obtiene una cadena con el código/llave a partir de una cadena código-descripción (Ej. Ingresa "03 - Acción" :::
        ///     Obtiene "Acción")
        /// </summary>
        /// <param name="keyValue"></param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns></returns>
        public static string GetCodeValue(string keyValue, string separator = " - ")
        {
            var codeKeyValue = GetCodeKeyValue(keyValue, separator);
            return codeKeyValue.Value;
        }

        /// <summary>
        ///     Obtiene una lista de tipo string a partir de la llave y valor del listado de keyValuePairList
        /// </summary>
        /// <param name="ellipseCodeItemsList">
        ///     List(EllipseCodeItem{string, string}): Listado tipo EllipseCodeItem con los datos de
        ///     llaves y valores
        /// </param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns>string: List{code - description}</returns>
        public static List<string> GetCodeList(List<EllipseCodeItem> ellipseCodeItemsList, string separator = " - ")
        {
            return ellipseCodeItemsList.Select(item => item.code + separator + item.description).ToList();
        }

        /// <summary>
        ///     Obtiene una lista de tipo string a partir de la llave y valor del listado de keyValuePairList
        /// </summary>
        /// <param name="keyValuePairList">
        ///     List(KeyValuePair{string, string}): Listado tipo KeyValuePair con los datos de llaves y
        ///     valores
        /// </param>
        /// <param name="separator">Separador para el Key y el Value (Ej. " - ")</param>
        /// <returns>string: List{key - value}</returns>
        public static List<string> GetCodeList(List<KeyValuePair<string, string>> keyValuePairList,
            string separator = " - ")
        {
            return keyValuePairList.Select(item => item.Key + separator + item.Value).ToList();
        }

        /// <summary>
        ///     Obtiene una lista de tipo string a partir de la llave y valor del listado de Dictionart
        /// </summary>
        /// <param name="dictionaryPair">
        ///     List(Dictionary{string, string}): Listado tipo Dictionary con los datos de llaves y
        ///     valores
        /// </param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns>string: List{key - value}</returns>
        public static List<string> GetCodeList(Dictionary<string, string> dictionaryPair, string separator = " - ")
        {
            return dictionaryPair.Select(item => item.Key + separator + item.Value).ToList();
        }

        public static string ReplaceQueryStringRegexWhiteSpaces(string text, string oldValue, string newValue)
        {
            var newstring = Regex.Replace(text, @"\s+", " ");
            return newstring.Replace(oldValue, newValue);
        }

        public static string CombineUrls(string baseUrl, string relativeUrl)
        {
            var baseUri = new Uri(baseUrl);
            return new Uri(baseUri, relativeUrl).AbsoluteUri;
        }
    }
}
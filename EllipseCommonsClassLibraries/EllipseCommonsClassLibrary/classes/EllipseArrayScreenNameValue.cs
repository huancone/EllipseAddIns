using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Screen = SharedClassLibrary.Ellipse.ScreenService;

namespace SharedClassLibrary.Ellipse
{
    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    public class ArrayScreenNameValue
    {
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        private List<Screen.ScreenNameValueDTO> _arrayFields;
        public int Length;

        public ArrayScreenNameValue()
        {
            _arrayFields = new List<Screen.ScreenNameValueDTO>();
            Length = 0;
        }

        /// <summary>
        ///     Crea un nuevo elemento del tipo ArrayScreenNameValue a partir de un arreglo de tipo [] ScreenFieldDTO
        /// </summary>
        /// <param name="arrayScreenFieldDto">Screen.ScreenFieldDTO</param>
        // ReSharper disable once ParameterTypeCanBeEnumerable.Local
        public ArrayScreenNameValue(Screen.ScreenFieldDTO[] arrayScreenFieldDto)
        {
            _arrayFields = new List<Screen.ScreenNameValueDTO>();
            Length = 0;

            foreach (var screenField in arrayScreenFieldDto)
                Add(screenField);
        }

        /// <summary>
        ///     Crea un nuevo elemento del tipo ArrayScreenNameValue a partir de un arreglo de una lista de ScreenNameValueDTO
        /// </summary>
        /// <param name="arrayFields">Lista Screen.ScreenNameValueDTO Lista para creación del objeto</param>
        public ArrayScreenNameValue(List<Screen.ScreenNameValueDTO> arrayFields)
        {
            _arrayFields = arrayFields;
            Length = arrayFields.Count();
        }

        /// <summary>
        ///     Crea un nuevo elemento del tipo ArrayScreenNameValue a partir de un arreglo de una lista de ScreenNameValueDTO
        /// </summary>
        /// <param name="arrayFields">Lista Screen.ScreenNameValueDTO Lista para creación del objeto</param>
        public ArrayScreenNameValue(Screen.ScreenNameValueDTO[] arrayFields)
        {
            _arrayFields = arrayFields.ToList();
            Length = arrayFields.Count();
        }

        /// <summary>
        ///     Agrega un nuevo elemento a la lista de campos
        /// </summary>
        /// <param name="fieldName">string: Nombre del campo</param>
        /// <param name="fieldValue">string: Valor del campo</param>
        public void Add(string fieldName, string fieldValue)
        {
            var item = new Screen.ScreenNameValueDTO();

            if (fieldName == null)
                return;

            item.fieldName = fieldName;
            item.value = fieldValue;
            Length++;
            _arrayFields.Add(item);
        }

        /// <summary>
        ///     Agrega un nuevo elemento a la lista de campos a partir de un Screen.ScreenFieldDTO
        /// </summary>
        /// <param name="screenField">Screen.ScreenFieldDTO: elemento obtenido de un reply de un screen</param>
        public void Add(Screen.ScreenFieldDTO screenField)
        {
            Add(screenField.fieldName, screenField.value);
        }

        /// <summary>
        ///     Devuelve la primera coincidencia de tipo ScreenNameValueDTO que encuentre que coincida con el nombre del campo
        ///     fieldName
        /// </summary>
        /// <param name="fieldName">string: fieldName del campo ScreenNameValueDTO a buscar</param>
        /// <returns>ScreenNameValueDTO: elemento de tipo ScreenNameValueDTO resultante o null si no hay coincidencia</returns>
        public Screen.ScreenNameValueDTO GetField(string fieldName)
        {
            foreach (var item in _arrayFields)
                if (item.fieldName == fieldName)
                    return item;

            return null;
        }

        /// <summary>
        ///     Establece el valor a la primera coincidencia de tipo ScreenNameValueDTO que encuentre que coincida con el nombre
        ///     del campo fieldName. No hace cambios si no encuentra nada
        /// </summary>
        /// <param name="fieldName">string: fieldName del campo ScreenNameValueDTO a modificar</param>
        /// <param name="fieldValue">string: valor del campo ScreenNameValueDTO a modificar</param>
        /// <returns></returns>
        public void SetField(string fieldName, string fieldValue)
        {
            foreach (var item in _arrayFields.Where(item => item.fieldName == fieldName))
            {
                item.value = fieldValue;
                return;
            }
        }

        /// <summary>
        ///     Devuelve la lista en forma de arreglo para el envío de parámetros al screen service
        /// </summary>
        /// <returns>ScreenNameValueDTO[]: arreglo de campos (fieldName, fieldValue) para el screen service</returns>
        public Screen.ScreenNameValueDTO[] ToArray()
        {
            return _arrayFields.ToArray();
        }
    }
}
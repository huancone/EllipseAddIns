using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using System.Text.RegularExpressions;


namespace EllipseCommonsClassLibrary
{
    public class Utils
    {
        /// <summary>
        /// Obtiene una cadena con el nombre de una variable dada
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item">Variable a obtener el nombre</param>
        /// <returns>string: nombre de una variable (Ej. int numeroEntero = 3; //output: numeroEntero)</returns>
        public static string GetVarName<T>(T item) where T : class
        {
            return typeof(T).GetProperties()[0].Name;
        }

        /// <summary>
        /// Divide el text ingresado en un arreglo string[] teniendo en cuenta los saltos de línea y la longitud de línea máxima deseada
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
                textArray.Add(text);
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
        /// Obtiene una lista con los campos de Key, Value concatenados con el conector dado (Ej. [Key, Value] = ["codigo", "valor"], connector = " - ", resultado = "codigo - valor")
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="source"></param>
        /// <param name="connector"></param>
        /// <returns></returns>
        public static List<string> ConcatToStringDictionaryKeyValue<TKey, TValue>(Dictionary<TKey, TValue> source, string connector)
        {
            var list = source.Select(entry => entry.Key + connector + entry.Value).ToList();

            return list;
        }

        /// <summary>
        /// Obtiene un listado separado por el separador dado en forma de cadena de texto (Ej: lista{valor1, valor2, valor3} => string = "valor1,valor2,valor3")
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listValues">IEnumerable(T): Arreglo Enumerable para el listado</param>
        /// <param name="separator">string: Indica cuál texto/símbolo será usado como separador de lista (Ej: separator = ","; stringList = "valor1, valor2, valor3"</param>
        /// <param name="quotation">string: Encierra el valor de la lista con este text (Ej: quotation = "'", valorLista = "'valor'"; quotation = "***", valorLista = "***valor***")</param>
        /// <returns></returns>
        public static string GetListInSeparator<T>(IEnumerable<T> listValues, string separator, string quotation = null)
        {
            var stringList = listValues.Aggregate("", (current, value) => current + quotation + value + quotation + separator);

            return stringList.Substring(0, stringList.Length - 1);
        }

        /// <summary>
        /// Obtiene el valor verdadero según el criterio de entrada. Si value es TRUE, VERDADERO, Y, YES, SI, ó 1
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
        /// Convierte una cadena de tipo "key - separator - value" en keyValuePair Ej. "23 - Description" -> KeyValuePair(string, string){"23", "Description"}
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
        /// Obtiene una cadena con el código/llave a partir de una cadena código-descripción (Ej. Ingresa "03 - Acción" ::: Obtiene "03")
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
        /// Obtiene una cadena con el código/llave a partir de una cadena código-descripción (Ej. Ingresa "03 - Acción" ::: Obtiene "Acción")
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
        /// Obtiene una lista de tipo string a partir de la llave y valor del listado de keyValuePairList
        /// </summary>
        /// <param name="ellipseCodeItemsList">List(EllipseCodeItem{string, string}): Listado tipo EllipseCodeItem con los datos de llaves y valores</param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns>string: List{code - description}</returns>
        public static List<string> GetCodeList(List<EllipseCodeItem> ellipseCodeItemsList, string separator = " - ")
        {
            return ellipseCodeItemsList.Select(item => item.code + separator + item.description).ToList();
        }
        /// <summary>
        /// Obtiene una lista de tipo string a partir de la llave y valor del listado de keyValuePairList
        /// </summary>
        /// <param name="keyValuePairList">List(KeyValuePair{string, string}): Listado tipo KeyValuePair con los datos de llaves y valores</param>
        /// <param name="separator">Separador para el Key y el Value (Ej. " - ")</param>
        /// <returns>string: List{key - value}</returns>
        public static List<string> GetCodeList(List<KeyValuePair<string, string>> keyValuePairList, string separator = " - ")
        {
            return keyValuePairList.Select(item => item.Key + separator + item.Value).ToList();
        }

        /// <summary>
        /// Obtiene una lista de tipo string a partir de la llave y valor del listado de Dictionart
        /// </summary>
        /// <param name="dictionaryPair">List(Dictionary{string, string}): Listado tipo Dictionary con los datos de llaves y valores</param>
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
    }

    public static class TimeOperations
    {
        public static List<ShiftSlot> GetSlots(ShiftSlot[] shifts, DateTime startEvent, DateTime endEvent)
        {
            if (endEvent < startEvent)
                throw new ArgumentException("La fecha final no puede ser menor a la fecha inicial");

            //to establish the date part of the datetime according to the starttime of the shift day / desplazamiento horario según inicio del día por turno
            startEvent = startEvent.AddTicks(-shifts[0].GetStartDateTime().TimeOfDay.Ticks).Date + startEvent.TimeOfDay;
            endEvent = endEvent.AddTicks(-shifts[0].GetStartDateTime().TimeOfDay.Ticks).Date + endEvent.TimeOfDay;

            var i = 0;
            //To find the starting shift for the event / para encontrar el turno inicial del evento
            var shiftStartLessShiftEnd = shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;
            var eventStartLessShiftStart = startEvent.TimeOfDay.TotalSeconds < shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds;
            var eventStartLessShiftEnd = startEvent.TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;

            while (shiftStartLessShiftEnd && (eventStartLessShiftStart || (!eventStartLessShiftStart && !eventStartLessShiftEnd)) ||
                !shiftStartLessShiftEnd && (eventStartLessShiftStart && !eventStartLessShiftEnd)
                )
            {
                i++;
                shiftStartLessShiftEnd = shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;
                eventStartLessShiftStart = startEvent.TimeOfDay.TotalSeconds < shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds;
                eventStartLessShiftEnd = startEvent.TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;

            }
            var slotStart = startEvent;
            var slotEnd = slotStart.Date + shifts[i].GetEndDateTime().TimeOfDay;
            slotEnd = slotEnd.AddDays((shifts[i].GetEndDateTime().Date - shifts[i].GetStartDateTime().Date).TotalDays);//adiciona si hay un cambio en el día

            var slotList = new List<ShiftSlot>();

            //while (endEvent.Date > slotEnd.Date ||
            //    (endEvent.Date == slotEnd.Date && !(
            //        shifts[i].StartHour.TotalSeconds <= shifts[i].EndHour.TotalSeconds
            //            ? (endEvent.TimeOfDay.TotalSeconds >= shifts[i].StartHour.TotalSeconds && endEvent.TimeOfDay.TotalSeconds < shifts[i].EndHour.TotalSeconds)
            //            : (endEvent.TimeOfDay.TotalSeconds >= shifts[i].StartHour.TotalSeconds || endEvent.TimeOfDay.TotalSeconds < shifts[i].EndHour.TotalSeconds)
            //       )))
            while(endEvent.Date > slotEnd.Date || (endEvent.Date == slotEnd.Date && (slotStart.TimeOfDay.TotalSeconds <= slotEnd.TimeOfDay.TotalSeconds  
                ? (endEvent.TimeOfDay.TotalSeconds > slotStart.TimeOfDay.TotalSeconds && endEvent.TimeOfDay.TotalSeconds > slotEnd.TimeOfDay.TotalSeconds)
                : (endEvent.TimeOfDay.TotalSeconds < slotStart.TimeOfDay.TotalSeconds && endEvent.TimeOfDay.TotalSeconds > slotEnd.TimeOfDay.TotalSeconds)
                )))
            {
                var newSlot = new ShiftSlot();
                newSlot.SetDate(slotStart.Date);
                newSlot.SetStartTime(slotStart.TimeOfDay);
                newSlot.SetEndTime(shifts[i].GetEndDateTime().TimeOfDay);
                newSlot.ShiftCode = shifts[i].ShiftCode;
                slotList.Add(newSlot);
                i++;

                if (i >= shifts.Length)
                {
                    i = 0;
                    slotStart = slotStart.Date.AddDays(1);
                }

                slotStart = slotStart.Date + shifts[i].GetStartDateTime().TimeOfDay;
                slotEnd = slotStart.Date + shifts[i].GetEndDateTime().TimeOfDay;
                slotEnd = slotEnd.AddDays((shifts[i].GetEndDateTime().Date - shifts[i].GetStartDateTime().Date).TotalDays);//adiciona si hay un cambio en el día

            }

            var lastSlot = new ShiftSlot();
            lastSlot.SetDate(slotStart);
            lastSlot.SetStartTime(slotStart.TimeOfDay);
            lastSlot.SetEndTime(endEvent.TimeOfDay);
            lastSlot.ShiftCode = shifts[i].ShiftCode;
            slotList.Add(lastSlot);

            return slotList;
        }

        /// <summary>
        /// Convierte una hora en formato de número (3, 1.6, 36.3) a formato HHMM (03:00, 01:36, 12:18). Si la hora ingresada excede el valor de 24 horas, esta es truncada al día.
        /// </summary>
        /// <param name="hourTime">Hora de forma numérica (Ej: 11, 8.4, 3.1, 28.4) </param>
        /// <param name="separator"></param>
        // ReSharper disable once InconsistentNaming
        public static string ConvertDecimalHourToHHMM(string hourTime, string separator = null)
        {
            if (separator == null)
                separator = ":";
            if (string.IsNullOrWhiteSpace(hourTime))
                return "00" + separator + "00";

            var hh = Convert.ToDecimal(hourTime) % 24;
            var mm = hh - Math.Truncate(hh);
                                        
            mm = Math.Abs(Math.Truncate(mm * 60));
            hh = Math.Truncate(hh);

            var newHour = hh + separator + (Convert.ToInt32(mm)).ToString("D2");
            return newHour;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="hourTime">Hora de forma numérica (Ej: 11, 8.4, 3.1, 28.4) </param>
        /// <param name="separator"></param>
        // ReSharper disable once InconsistentNaming
        public static string ConvertDecimalHourToHHMM(float hourTime, char separator)
        {
            var hh = Convert.ToDecimal(hourTime) % 24;
            var mm = hh - Math.Truncate(hh);

            mm = Math.Abs(Math.Truncate(mm * 60));
            hh = Math.Truncate(hh);

            var newHour = hh + separator + (Convert.ToInt32(mm)).ToString("D2");
            return newHour;
        }

        public static string FormatDateToString(DateTime date, string format, string dateSeparator = "")
        {
            if (format.Equals(DateTimeFormats.DateYYMMDD))
                return "" + date.Year.ToString("0000").Substring(2) + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00");
            if (format.Equals(DateTimeFormats.DateYYDDMM))
                return "" + date.Year.ToString("0000").Substring(2) + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00");
            if (format.Equals(DateTimeFormats.DateMMDDYY))
                return "" + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Year.ToString("0000").Substring(2);
            if (format.Equals(DateTimeFormats.DateDDMMYY))
                return "" + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Year.ToString("0000").Substring(2);

            if (format.Equals(DateTimeFormats.DateYYYYMMDD))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00");
            if (format.Equals(DateTimeFormats.DateYYYYDDMM))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00");
            if (format.Equals(DateTimeFormats.DateMMDDYYYY))
                return "" + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Year.ToString("0000");
            if (format.Equals(DateTimeFormats.DateDDMMYYYY))
                return "" + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Year.ToString("0000");

            throw new ArgumentException("Not a valid Date format", "format");
        }
        public static string FormatDateTimeToString(DateTime date, string format, string dateSeparator = null, string timeSeparator = null, string splitSeparator = null)
        {
            dateSeparator = dateSeparator ?? "";
            timeSeparator = timeSeparator ?? "";
            splitSeparator = splitSeparator ?? " ";

            if (format.Equals(DateTimeFormats.DateTimeYYYYMMDD_HHMM))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00") + splitSeparator + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00");
            if (format.Equals(DateTimeFormats.DateTimeYYYYMMDD_HHMMSS))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Month.ToString("00") + dateSeparator + date.Day.ToString("00") + splitSeparator + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00") + timeSeparator + date.Second.ToString("00");
            if (format.Equals(DateTimeFormats.DateTimeYYYYDDMM_HHMM))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00") + splitSeparator + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00");
            if (format.Equals(DateTimeFormats.DateTimeYYYYDDMM_HHMMSS))
                return "" + date.Year.ToString("0000") + dateSeparator + date.Day.ToString("00") + dateSeparator + date.Month.ToString("00") + splitSeparator + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00") + timeSeparator + date.Second.ToString("00");

            throw new ArgumentException("Not a valid DateTime format", "format");
        }

        public static string FormatDateTimeToString(TimeSpan time, string format, string dateSeparator = null, string timeSeparator = null, string splitSeparator = null)
        {
            var date = new DateTime();
            date = date.Add(time);

            return FormatDateTimeToString(date, format, dateSeparator, timeSeparator, splitSeparator);
        }
        public static string FormatTimeToString(TimeSpan time, string format, string timeSeparator)
        {
            var date = new DateTime();
            date = date.Add(time);
            
            if (format.Equals(DateTimeFormats.TimeHHMMSS))
                return "" + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00") + timeSeparator + date.Second.ToString("00");
            if (format.Equals(DateTimeFormats.TimeHHMM))
                return "" + date.Hour.ToString("00") + timeSeparator + date.Minute.ToString("00");

            throw new ArgumentException("Not a valid Time format", "format");
        }
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public static class DateTimeFormats
        {
            public static string DateYYMMDD = "YY-MM-DD";
            public static string DateYYDDMM = "YY-DD-MM";
            public static string DateMMDDYY = "MM-DD-YY";
            public static string DateDDMMYY = "DD-MM-YY";

            public static string DateYYYYMMDD = "YYYY-MM-DD";
            public static string DateYYYYDDMM = "YYYY-DD-MM";
            public static string DateMMDDYYYY = "MM-DD-YYYY";
            public static string DateDDMMYYYY = "DD-MM-YYYY";

            public static string DateTimeYYYYMMDD_HHMM = "YYYY-MM-DD_HH-MM";
            public static string DateTimeYYYYMMDD_HHMMSS = "YYYY-MM-DD_HH-MM-SS";
            public static string DateTimeYYYYDDMM_HHMM = "YYYY-DD-MM_HH-MM";
            public static string DateTimeYYYYDDMM_HHMMSS = "YYYY-DD-MM_HH-MM-SS";

            public static string TimeHHMM = "HH-MM";
            public static string TimeHHMMSS = "HH-MM-SS";
        }

        /// <summary>
        /// Valida si la fecha YYYYMMDD ha superado el tiempo máximo permitido en días
        /// </summary>
        /// <param name="date">string: fecha en formato yyyyMMdd</param>
        /// <param name="daysLimit">int: número de días permitidos</param>
        /// <returns></returns>
        public static bool ValidateUserStatus(string date, int daysLimit)
        {
            var datetime = DateTime.ParseExact(date, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            return DateTime.Today.Subtract(datetime).TotalDays <= daysLimit;
        }
    }
    public class ShiftSlot
    {
        private DateTime _startDate;
        private DateTime _endDate;
        public string ShiftCode;

        /// <summary>
        /// Inicializa el objecto
        /// </summary>
        public ShiftSlot()
        {
            _startDate = new DateTime();
            _endDate = new DateTime();
        }
        /// <summary>
        /// Inicializa el objeto
        /// </summary>
        /// <param name="shiftCode"></param>
        public ShiftSlot(string shiftCode)
        {
            _startDate = new DateTime();
            _endDate = new DateTime();
            ShiftCode = shiftCode;
        }

        /// <summary>
        /// Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="date">DateTime: fecha para el objeto. La establece como fecha de inicio y fecha final</param>
        /// <param name="startHour">TimeSpan: tiempo inicial del objeto</param>
        /// <param name="endHour">TimeSpan: tiempo final del objeto</param>
        /// <param name="shiftCode">string: código del objeto</param>
        public ShiftSlot(DateTime date, TimeSpan startHour, TimeSpan endHour, string shiftCode = null)
        {
            _startDate = new DateTime(date.Year, date.Month, date.Day);
            _startDate = _startDate.Add(startHour);

            _endDate = new DateTime(date.Year, date.Month, date.Day);
            _endDate = _endDate.Add(endHour);

            ShiftCode = shiftCode;
        }
        /// <summary>
        /// Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="startDate">DateTime: fecha inicial para el objeto</param>
        /// <param name="startHour">TimeSpan: tiempo inicial del objeto</param>
        /// <param name="endDate">DateTime: fecha final para el objeto</param>
        /// <param name="endHour">TimeSpan: tiempo final del objeto</param>
        /// <param name="shifCode">string: código del objeto</param>
        public ShiftSlot(DateTime startDate, TimeSpan startHour, DateTime endDate, TimeSpan endHour, string shifCode = null)
        {
            _startDate = new DateTime(startDate.Year, startDate.Month, startDate.Day);
            _startDate = _startDate.Add(startHour);

            _endDate = new DateTime(endDate.Year, endDate.Month, endDate.Day);
            _endDate = _endDate.Add(endHour);

            ShiftCode = shifCode;
        }
        /// <summary>
        /// /// Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="startDateTime">DateTime: Fecha y Tiempo inicial para el objeto</param>
        /// <param name="endDateTime">DateTime: Fecha y Tiempo final para el objeto</param>
        /// <param name="shiftCode">string: código del objeto</param>
        public ShiftSlot(DateTime startDateTime, DateTime endDateTime, string shiftCode = null)
        {
            _startDate = startDateTime;
            _endDate = endDateTime;
            ShiftCode = shiftCode;
        }

        /// <summary>
        /// Obtiene la fecha actual del objeto. Corresponde a la fecha inicial
        /// </summary>
        /// <returns>DateTime: Fecha del objeto</returns>
        public DateTime GetDate()
        {
            return _startDate.Date;
        }

        /// <summary>
        /// Obtiene la Fecha y Hora inicial del objeto
        /// </summary>
        /// <returns>DateTime</returns>
        public DateTime GetStartDateTime()
        {
            return _startDate;
        }
        /// <summary>
        /// Obtiene la Fecha y Hora final del objeto
        /// </summary>
        /// <returns>DateTime</returns>
        public DateTime GetEndDateTime()
        {
            return _endDate;
        }

        /// <summary>
        /// Establece la fecha del objeto. Establece fecha inicial y fecha final con el parámetro ingresado
        /// </summary>
        /// <param name="date"></param>
        public void SetDate(DateTime date)
        {
            var startTime = _startDate.TimeOfDay;
            _startDate = new DateTime(date.Year, date.Month, date.Day);
            _startDate = _startDate.Add(startTime);

            var endTime = _endDate.TimeOfDay;
            _endDate = new DateTime(date.Year, date.Month, date.Day);
            _endDate = _endDate.Add(endTime);
        }
        /// <summary>
        /// Establece Fecha y Hora inicial del objeto
        /// </summary>
        /// <param name="startDateTime"></param>
        public void SetStartDateTime(DateTime startDateTime)
        {
            _startDate = startDateTime;
        }
        /// <summary>
        /// Establece Hora inicial del objeto conservando su fecha
        /// </summary>
        /// <param name="startTime"></param>
        public void SetStartTime(TimeSpan startTime)
        {
            var dateOnly = _startDate.Date;
            dateOnly = dateOnly.Add(startTime);
            _startDate = dateOnly;
        }
        /// <summary>
        /// Establece Fecha y Hora final del objeto
        /// </summary>
        /// <param name="endDateTime"></param>
        public void SetEndDateTime(DateTime endDateTime)
        {
            _endDate = endDateTime;
        }
        /// <summary>
        /// Establece Hora final del objeto conservando su fecha
        /// </summary>
        /// <param name="endTime"></param>
        public void SetEndTime(TimeSpan endTime)
        {
            var dateOnly = _endDate.Date;
            dateOnly = dateOnly.Add(endTime);
            _endDate = dateOnly;
        }
    }

    public static class ShiftConstants
    {
        public static class ShiftCodes
        {
            public static string HourToHourCode = "HH";
            public static string DailyZeroCode = "A24";
            public static string DailyMorningCode = "D66";
            public static string DayNightCode = "DN";
            public static string HourToHourDescription = "Hour to Hour";
            public static string DailyZeroDescription = "Day start at 00:00";
            public static string DailyMorningDescription = "Day start at 06:00";
            public static string DayNightDescription = "Day 06:00-18:00 Night 18:00-06:00";
        }
        public static class ShiftPeriods
        {
            public static ShiftSlot [] GetHourToHourShiftSlots()
            {
                var slotSlist = new List<ShiftSlot>();

                for(var i=0; i<24; i++)
                {
                    var slot = new ShiftSlot ((i + 1).ToString("00"));
                    slot.SetStartTime(new TimeSpan(i, 00, 00));
                    slot.SetEndTime(new TimeSpan(i+1, 00, 00));

                    slotSlist.Add(slot);
                }

                return slotSlist.ToArray();
            }
            public static ShiftSlot[] GetDailyNightShiftSlots()
            {
                var slotArray = new ShiftSlot[2];

                slotArray[0] = new ShiftSlot("D");
                slotArray[0].SetStartTime(new TimeSpan(06, 00, 00));
                slotArray[0].SetEndTime(new TimeSpan(18, 00, 00));
                slotArray[1] = new ShiftSlot("N");
                slotArray[1].SetStartTime(new TimeSpan(18, 00, 00));
                slotArray[1].SetEndTime(new TimeSpan(01, 06, 00, 00));//dia siguiente

                return slotArray;
            }
            public static ShiftSlot[] GetDailyZeroSlots()
            {
                var slotArray = new ShiftSlot[1];

                slotArray[0] = new ShiftSlot("A");
                slotArray[0].SetStartTime(new TimeSpan(00, 00, 00));
                slotArray[0].SetEndTime(new TimeSpan(24, 00, 00));

                return slotArray;
            }
            public static ShiftSlot[] GetDailyMorningSlots()
            {
                var slotArray = new ShiftSlot[1];

                slotArray[0] = new ShiftSlot("DY");
                slotArray[0].SetStartTime(new TimeSpan(06, 00, 00));
                slotArray[0].SetEndTime(new TimeSpan(01, 06, 00, 00));//dia siguiente

                return slotArray;
            }
        }
    }

    public static class FileWriter
    {
        public static void WriteTextToFile(string text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) || urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;
            
            if (urlPath == null)
                urlPath = "";
            File.WriteAllText(Path.Combine(urlPath, filename), text);
        }
        public static void WriteTextToFile(string [] text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) ||
            //      urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;

            if (urlPath == null)
                urlPath = "";
            File.WriteAllLines(Path.Combine(urlPath, filename), text);
        }
        public static void AppendTextToFile(string text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) ||
            //      urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;

            if (urlPath == null)
                urlPath = "";
            using (var file = new StreamWriter(Path.Combine(urlPath, filename), true))
            {
                file.WriteLine(text);
            }
        }

        public static void CreateDirectory(string directoryPath)
        {
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(directoryPath))
                    return;

                // Try to create the directory.
                Directory.CreateDirectory(directoryPath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("FileWriter:CreateDirectory::" + directoryPath, ex.Message);
                throw;
            }

        }
        public static void DeleteDirectory(string directoryPath)
        {
            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(directoryPath))
                    return;

                // Try to delete the directory.
                var di = new DirectoryInfo(directoryPath);
                di.Delete();
            }
            catch (Exception ex)
            {
                Debugger.LogError("FileWriter:DeleteDirectory::" + directoryPath, ex.Message);
                throw;
            }
        }
        public static void DeleteFile(string directoryPath, string fileName)
        {
            DeleteFile(Path.Combine(directoryPath, fileName));
        }
        public static void DeleteFile(string urlFileName)
        {
            try
            {
                // Determine whether the file exists.
                if (!File.Exists(urlFileName))
                    return;

                // Try to delete the file.
                var fi = new FileInfo(urlFileName);
                fi.Delete();
            }
            catch (Exception ex)
            {
                Debugger.LogError("FileWriter:DeleteFile::" + urlFileName, ex.Message);
                throw;
            }
        }
        public static bool CheckDirectoryExist(string directoryPath)
        {
            try
            {
                // Determine whether the directory exists.
                return Directory.Exists(directoryPath);
            }
            catch (Exception ex)
            {
                Debugger.LogError("FileWriter:CheckDirectoryExist::" + directoryPath, ex.Message);
                throw;
            }
        }

        public static void CopyFileToDirectory(string fileName, string sourcePath, string targetPath, bool overwrite = true)
        {
            try
            {
                var sourceFile = Path.Combine(sourcePath, fileName);
                var destFile = Path.Combine(targetPath, fileName);

                File.Copy(sourceFile, destFile, overwrite);
            }
            catch (Exception ex)
            {
                Debugger.LogError("FileWriter:CopyFileToDirectory", ex.Message);
                throw;
            }
        }

    }

    public static class MathUtil
    {
        /// <summary>
        /// Indica si un número está dentro del rango aceptado con respecto a otro. Si compareNumber está dentro de los valores de baseNumber +/- (baseNumber * threshold)
        /// </summary>
        /// <param name="baseNumber">object: Número base de operación </param>
        /// <param name="compareNumber">object: Número a comparar</param>
        /// <param name="threshold">object: Rango de comparación (Ej.: 0.1, 0.3, 1, 1.3, 2)</param>
        /// <returns>bool: true si compareNumber está dentro del rango threshold de baseNumber (Ej. baseNumber: 6, compareNumber: 4. Si threshold es 0.5 es true, si threshold es 0.3 es false</returns>
        public static bool InThreshold(object baseNumber, object compareNumber, object threshold)
        {
            if (baseNumber == null || compareNumber == null)
                throw new NullReferenceException("No se puede realizar la operación. Ingreso debe ser diferente de nulo");

            var num1 = Convert.ToDecimal(baseNumber);
            var num2 = Convert.ToDecimal(compareNumber);
            var th = Convert.ToDecimal(threshold);
            var abs = Math.Abs(num1 - num2);

            return (num1*th) >= abs;
        }

        public static string ToOrdinal(this long number)
        {
            if (number < 0) return number.ToString();
            long rem = number % 100;
            if (rem >= 11 && rem <= 13) return number + "th";

            switch (number % 10)
            {
                case 1:
                    return number + "st";
                case 2:
                    return number + "nd";
                case 3:
                    return number + "rd";
                default:
                    return number + "th";
            }
        }

        public static string ToOrdinal(this int number)
        {
            return ((long)number).ToOrdinal();
        }

        public static string ToOrdinal(this string number)
        {
            if (string.IsNullOrEmpty(number)) return number;

            var dict = new Dictionary<string, string>
            {
                {"zero", "zeroth"},
                {"nought", "noughth"},
                {"one", "first"},
                {"two", "second"},
                {"three", "third"},
                {"four", "fourth"},
                {"five", "fifth"},
                {"six", "sixth"},
                {"seven", "seventh"},
                {"eight", "eighth"},
                {"nine", "ninth"},
                {"ten", "tenth"},
                {"eleven", "eleventh"},
                {"twelve", "twelfth"},
                {"thirteen", "thirteenth"},
                {"fourteen", "fourteenth"},
                {"fifteen", "fifteenth"},
                {"sixteen", "sixteenth"},
                {"seventeen", "seventeenth"},
                {"eighteen", "eighteenth"},
                {"nineteen", "nineteenth"},
                {"twenty", "twentieth"},
                {"thirty", "thirtieth"},
                {"forty", "fortieth"},
                {"fifty", "fiftieth"},
                {"sixty", "sixtieth"},
                {"seventy", "seventieth"},
                {"eighty", "eightieth"},
                {"ninety", "ninetieth"},
                {"hundred", "hundredth"},
                {"thousand", "thousandth"},
                {"million", "millionth"},
                {"billion", "billionth"},
                {"trillion", "trillionth"},
                {"quadrillion", "quadrillionth"},
                {"quintillion", "quintillionth"}
            };


            // rough check whether it's a valid number
            var temp = number.ToLower().Trim().Replace(" and ", " ");
            var words = temp.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);

            if (words.Any(word => !dict.ContainsKey(word)))
                return number;

            // extract last word
            number = number.TrimEnd().TrimEnd('-');
            var index = number.LastIndexOfAny(new[] { ' ', '-' });
            var last = number.Substring(index + 1);

            // make replacement and maintain original capitalization
            if (last == last.ToLower())
                last = dict[last];
            else if (last == last.ToUpper())
                last = dict[last.ToLower()].ToUpper();
            else
                last = last.ToLower();
                last = char.ToUpper(dict[last][0]) + dict[last].Substring(1);

            return number.Substring(0, index + 1) + last;
        }
    }
}

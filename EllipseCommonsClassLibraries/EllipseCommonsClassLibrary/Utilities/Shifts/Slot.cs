using System;

namespace CommonsClassLibrary.Utilities.Shifts
{
    public class Slot
    {
        private System.DateTime _endDate;
        private System.DateTime _startDate;
        public string ShiftCode;

        /// <summary>
        ///     Inicializa el objecto
        /// </summary>
        public Slot()
        {
            _startDate = new System.DateTime();
            _endDate = new System.DateTime();
        }

        /// <summary>
        ///     Inicializa el objeto
        /// </summary>
        /// <param name="shiftCode"></param>
        public Slot(string shiftCode)
        {
            _startDate = new System.DateTime();
            _endDate = new System.DateTime();
            ShiftCode = shiftCode;
        }

        /// <summary>
        ///     Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="date">System.DateTime: fecha para el objeto. La establece como fecha de inicio y fecha final</param>
        /// <param name="startHour">TimeSpan: tiempo inicial del objeto</param>
        /// <param name="endHour">TimeSpan: tiempo final del objeto</param>
        /// <param name="shiftCode">string: código del objeto</param>
        public Slot(System.DateTime date, TimeSpan startHour, TimeSpan endHour, string shiftCode = null)
        {
            _startDate = new System.DateTime(date.Year, date.Month, date.Day);
            _startDate = _startDate.Add(startHour);

            _endDate = new System.DateTime(date.Year, date.Month, date.Day);
            _endDate = _endDate.Add(endHour);

            ShiftCode = shiftCode;
        }

        /// <summary>
        ///     Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="startDate">System.DateTime: fecha inicial para el objeto</param>
        /// <param name="startHour">TimeSpan: tiempo inicial del objeto</param>
        /// <param name="endDate">System.DateTime: fecha final para el objeto</param>
        /// <param name="endHour">TimeSpan: tiempo final del objeto</param>
        /// <param name="shifCode">string: código del objeto</param>
        public Slot(System.DateTime startDate, TimeSpan startHour, System.DateTime endDate, TimeSpan endHour, string shifCode = null)
        {
            _startDate = new System.DateTime(startDate.Year, startDate.Month, startDate.Day);
            _startDate = _startDate.Add(startHour);

            _endDate = new System.DateTime(endDate.Year, endDate.Month, endDate.Day);
            _endDate = _endDate.Add(endHour);

            ShiftCode = shifCode;
        }

        /// <summary>
        ///     /// Inicializa el objeto con los parámetros indicados
        /// </summary>
        /// <param name="startSystem.DateTime">System.DateTime: Fecha y Tiempo inicial para el objeto</param>
        /// <param name="endSystem.DateTime">System.DateTime: Fecha y Tiempo final para el objeto</param>
        /// <param name="shiftCode">string: código del objeto</param>
        public Slot(System.DateTime startDateTime, System.DateTime endDateTime, string shiftCode = null)
        {
            _startDate = startDateTime;
            _endDate = endDateTime;
            ShiftCode = shiftCode;
        }

        /// <summary>
        ///     Obtiene la fecha actual del objeto. Corresponde a la fecha inicial
        /// </summary>
        /// <returns>System.DateTime: Fecha del objeto</returns>
        public System.DateTime GetDate()
        {
            return _startDate.Date;
        }

        /// <summary>
        ///     Obtiene la Fecha y Hora inicial del objeto
        /// </summary>
        /// <returns>System.DateTime</returns>
        public System.DateTime GetStartDateTime()
        {
            return _startDate;
        }

        /// <summary>
        ///     Obtiene la Fecha y Hora final del objeto
        /// </summary>
        /// <returns>System.DateTime</returns>
        public System.DateTime GetEndDateTime()
        {
            return _endDate;
        }

        /// <summary>
        ///     Establece la fecha del objeto. Establece fecha inicial y fecha final con el parámetro ingresado
        /// </summary>
        /// <param name="date"></param>
        public void SetDate(System.DateTime date)
        {
            var startTime = _startDate.TimeOfDay;
            _startDate = new System.DateTime(date.Year, date.Month, date.Day);
            _startDate = _startDate.Add(startTime);

            var endTime = _endDate.TimeOfDay;
            _endDate = new System.DateTime(date.Year, date.Month, date.Day);
            _endDate = _endDate.Add(endTime);
        }

        /// <summary>
        ///     Establece Fecha y Hora inicial del objeto
        /// </summary>
        /// <param name="startSystem.DateTime"></param>
        public void SetStartDateTime(System.DateTime startDateTime)
        {
            _startDate = startDateTime;
        }

        /// <summary>
        ///     Establece Hora inicial del objeto conservando su fecha
        /// </summary>
        /// <param name="startTime"></param>
        public void SetStartTime(TimeSpan startTime)
        {
            var dateOnly = _startDate.Date;
            dateOnly = dateOnly.Add(startTime);
            _startDate = dateOnly;
        }

        /// <summary>
        ///     Establece Fecha y Hora final del objeto
        /// </summary>
        /// <param name="endSystem.DateTime"></param>
        public void SetEndDateTime(System.DateTime endDateTime)
        {
            _endDate = endDateTime;
        }

        /// <summary>
        ///     Establece Hora final del objeto conservando su fecha
        /// </summary>
        /// <param name="endTime"></param>
        public void SetEndTime(TimeSpan endTime)
        {
            var dateOnly = _endDate.Date;
            dateOnly = dateOnly.Add(endTime);
            _endDate = dateOnly;
        }
    }
}
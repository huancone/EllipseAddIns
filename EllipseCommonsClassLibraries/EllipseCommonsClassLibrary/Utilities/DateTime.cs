using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using CommonsClassLibrary.Utilities.Shifts;

namespace CommonsClassLibrary.Utilities
{
    public partial class MyUtilities
    {
        public class DateTime
        {
            public static string DateDefaultFormat = "yyyyMMdd";
            public static string DateTimeDefaultFormat = "yyyyMMdd hhmmss";
            public static string TimeDefaultFormat = "hhmmss";
            public static string DateYYMMDD = "yy-MM-dd";
            public static string DateYYDDMM = "yy-dd-MM";
            public static string DateMMDDYY = "MM-dd-yy";
            public static string DateDDMMYY = "dd-MM-yy";

            public static string DateYYYYMMDD = "yyyy-MM-dd";
            public static string DateYYYYDDMM = "yyyy-dd-MM";
            public static string DateMMDDYYYY = "MM-dd-yyyy";
            public static string DateDDMMYYYY = "dd-MM-yyyy";

            public static string DateTimeYYYYMMDD_HHMM = "yyyy-MM-dd_hh-mm";
            public static string DateTimeYYYYMMDD_HHMMSS = "yyyy-MM-dd_hh-mm-ss";
            public static string DateTimeYYYYDDMM_HHMM = "yyyy-dd-MM_hh-mm";
            public static string DateTimeYYYYDDMM_HHMMSS = "yyyy-dd-MM_hh-mm-ss";

            public static string TimeHHMM = "hh-mm";
            public static string TimeHHMMSS = "hh-mm-ss";

            public static List<Slot> GetSlots(Slot[] shifts, System.DateTime startEvent, System.DateTime endEvent)
            {
                if (endEvent < startEvent)
                    throw new ArgumentException("La fecha final no puede ser menor a la fecha inicial");

                //to establish the date part of the datetime according to the starttime of the shift day / desplazamiento horario según inicio del día por turno
                startEvent = startEvent.AddTicks(-shifts[0].GetStartDateTime().TimeOfDay.Ticks).Date + startEvent.TimeOfDay;
                endEvent = endEvent.AddTicks(-shifts[0].GetStartDateTime().TimeOfDay.Ticks).Date + endEvent.TimeOfDay;

                var i = 0;
                //To find the starting shift for the event / para encontrar el turno inicial del evento
                var shiftStartLessShiftEnd = shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds <
                                             shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;
                var eventStartLessShiftStart =
                    startEvent.TimeOfDay.TotalSeconds < shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds;
                var eventStartLessShiftEnd =
                    startEvent.TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;

                while (shiftStartLessShiftEnd &&
                       (eventStartLessShiftStart || !eventStartLessShiftStart && !eventStartLessShiftEnd) ||
                       !shiftStartLessShiftEnd && eventStartLessShiftStart && !eventStartLessShiftEnd
                )
                {
                    i++;
                    shiftStartLessShiftEnd = shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds <
                                             shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;
                    eventStartLessShiftStart = startEvent.TimeOfDay.TotalSeconds <
                                               shifts[i].GetStartDateTime().TimeOfDay.TotalSeconds;
                    eventStartLessShiftEnd =
                        startEvent.TimeOfDay.TotalSeconds < shifts[i].GetEndDateTime().TimeOfDay.TotalSeconds;
                }

                var slotStart = startEvent;
                var slotEnd = slotStart.Date + shifts[i].GetEndDateTime().TimeOfDay;
                slotEnd = slotEnd.AddDays((shifts[i].GetEndDateTime().Date - shifts[i].GetStartDateTime().Date)
                    .TotalDays); //adiciona si hay un cambio en el día

                var slotList = new List<Slot>();

                //while (endEvent.Date > slotEnd.Date ||
                //    (endEvent.Date == slotEnd.Date && !(
                //        shifts[i].StartHour.TotalSeconds <= shifts[i].EndHour.TotalSeconds
                //            ? (endEvent.TimeOfDay.TotalSeconds >= shifts[i].StartHour.TotalSeconds && endEvent.TimeOfDay.TotalSeconds < shifts[i].EndHour.TotalSeconds)
                //            : (endEvent.TimeOfDay.TotalSeconds >= shifts[i].StartHour.TotalSeconds || endEvent.TimeOfDay.TotalSeconds < shifts[i].EndHour.TotalSeconds)
                //       )))
                while (endEvent.Date > slotEnd.Date || endEvent.Date == slotEnd.Date &&
                    (slotStart.TimeOfDay.TotalSeconds <= slotEnd.TimeOfDay.TotalSeconds
                        ? endEvent.TimeOfDay.TotalSeconds > slotStart.TimeOfDay.TotalSeconds &&
                          endEvent.TimeOfDay.TotalSeconds > slotEnd.TimeOfDay.TotalSeconds
                        : endEvent.TimeOfDay.TotalSeconds < slotStart.TimeOfDay.TotalSeconds &&
                          endEvent.TimeOfDay.TotalSeconds > slotEnd.TimeOfDay.TotalSeconds
                    ))
                {
                    var newSlot = new Slot();
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
                    slotEnd = slotEnd.AddDays((shifts[i].GetEndDateTime().Date - shifts[i].GetStartDateTime().Date)
                        .TotalDays); //adiciona si hay un cambio en el día
                }

                var lastSlot = new Slot();
                lastSlot.SetDate(slotStart);
                lastSlot.SetStartTime(slotStart.TimeOfDay);
                lastSlot.SetEndTime(endEvent.TimeOfDay);
                lastSlot.ShiftCode = shifts[i].ShiftCode;
                slotList.Add(lastSlot);

                return slotList;
            }


            /// <summary>
            ///     Valida si la fecha YYYYMMDD ha superado el tiempo máximo permitido en días
            /// </summary>
            /// <param name="date">string: fecha en formato indicado por el parámetro dateFormat</param>
            /// <param name="daysLimit">int: número de días permitidos</param>
            /// <returns></returns>
            public static bool IsDateValid(string date, int daysLimit)
            {
                var cultureInfo = CultureInfo.CurrentCulture;
                var dateFormat = DateDefaultFormat;
                return IsDateValid(date, dateFormat, daysLimit, cultureInfo);
            }

            /// <summary>
            ///     Valida si la fecha YYYYMMDD ha superado el tiempo máximo permitido en días
            /// </summary>
            /// <param name="date">string: fecha en formato indicado por el parámetro dateFormat</param>
            /// <param name="dateFormat">string: formato de  fecha del valor date (Ej. yyyyMMdd)</param>
            /// <param name="daysLimit">int: número de días permitidos</param>
            /// <returns></returns>
            public static bool IsDateValid(string date, string dateFormat, int daysLimit)
            {
                var cultureInfo = CultureInfo.CurrentCulture;
                return IsDateValid(date, dateFormat, daysLimit, cultureInfo);
            }

            public static bool IsDateValid(string date, string dateFormat, int daysLimit, CultureInfo cultureInfo)
            {
                try
                {
                    var dateTime = System.DateTime.ParseExact(date, dateFormat, cultureInfo);
                    return IsDateValid(dateTime, daysLimit);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Unable to convert value to a date. Please be sure to provide a date value with format " + dateFormat + ". " + ex.Message);
                }
            }

            public static bool IsDateValid(System.DateTime date, int daysLimit)
            {
                return System.DateTime.Today.Subtract(date).TotalDays <= daysLimit;
            }


            /// <summary>
            ///     Convierte una hora en formato de número (3, 1.6, 36.3) a formato HHMM (03:00, 01:36, 12:18). Si la hora ingresada
            ///     excede el valor de 24 horas, esta es truncada al día.
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
                var mm = hh - System.Math.Truncate(hh);

                mm = System.Math.Abs(System.Math.Truncate(mm * 60));
                hh = System.Math.Truncate(hh);

                var newHour = Convert.ToInt32(hh).ToString("D2") + separator + Convert.ToInt32(mm).ToString("D2");
                return newHour;
            }

            /// <summary>
            /// </summary>
            /// <param name="hourTime">Hora de forma numérica (Ej: 11, 8.4, 3.1, 28.4) </param>
            /// <param name="separator"></param>
            // ReSharper disable once InconsistentNaming
            public static string ConvertDecimalHourToHHMM(float hourTime, char separator)
            {
                var hh = Convert.ToDecimal(hourTime) % 24;
                var mm = hh - System.Math.Truncate(hh);

                mm = System.Math.Abs(System.Math.Truncate(mm * 60));
                hh = System.Math.Truncate(hh);

                var newHour = Convert.ToInt32(hh).ToString("D2") + separator + Convert.ToInt32(mm).ToString("D2");
                return newHour;
            }
        }
    }
}
using System;
using System.Diagnostics.CodeAnalysis;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Utilities.MyDateTime;

namespace EllipseReferenceCodesClassLibrary
{
    public class ReferenceCodeEntity
    {
        public string EntityType;
        public string RefNo;
        public string RepeatCount;
        public string FieldType;
        public string ShortName;
        public string ScreenLiteral;
        public int Length;
        public bool StdTextFlag;

    }
    public class ReferenceCodeItem
    {
        public string EntityType;//WKO Work Order, WRQ WorkRequest
        public string EntityValue;
        public string RefNo;
        public string SeqNum;
        public string RefCode;
        public string FieldType;
        public string ShortName;
        public string ScreenLiteral;
        public bool StdTextFlag;//Flag para indicar si va a cambiarse
        public string StdtxtId;//El Id con el que se crea el StdText
        public string StdText;//Para el texto

        public ReferenceCodeItem(string entityType, string entityValue, string refNo, string seqNum, string refCode)
        {
            EntityType = entityType;
            EntityValue = entityValue;
            RefNo = refNo;
            SeqNum = seqNum;
            RefCode = refCode;
        }
        public ReferenceCodeItem(string entityType, string entityValue, string refNo, string seqNum, string refCode, string stdTextId, string stdText)
        {
            EntityType = entityType;
            EntityValue = entityValue;
            RefNo = refNo;
            SeqNum = seqNum;
            RefCode = refCode;
            StdTextFlag = true;
            StdtxtId = stdTextId;
            StdText = stdText;
        }


        public ReferenceCodeItem()
        { }
    }

    public class StandardTextItem
    {
        public string StdTextKey;//RC Reference Code

    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public static class ReferenceCodeItems
    {
        public class ReferenceDate : ReferenceCodeItem
        {
            private DateTime _date;
            private new string RefCode;
            private new readonly string FieldType;
            public ReferenceDate()
            {
                SetDate(1900, 01, 01);
                FieldType = "D";
            }
            public ReferenceCodeItem ToItem()
            {
                var item = new ReferenceCodeItem()
                {
                    EntityType = EntityType,
                    EntityValue = EntityValue,
                    RefNo = RefNo,
                    SeqNum = SeqNum,
                    RefCode = RefCode,
                    FieldType = FieldType
                };
                return item;
            }

            public void SetYear(object year)
            {
                int intYear;
                if (year == null)
                    throw new ArgumentNullException("year", "year can't be a null value");
                try
                {
                    intYear = Convert.ToInt16(year);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not a valid year. " + ex.Message, "year");
                }
                if (intYear < 1900 || intYear > 9999)
                    throw new ArgumentOutOfRangeException("year", "Year can't be less than 1900 or more than 9999");
                
                _date = new DateTime(intYear, _date.Month, _date.Day);
                UpdateRefCode();
            }
            public void SetMonth(object month)
            {
                int intMonth;
                if (month == null)
                    throw new ArgumentNullException("month", "month can't be a null value");
                try
                {
                    intMonth = Convert.ToInt16(month);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not a valid month." + ex.Message, "month");
                }
                if (intMonth < 1 || intMonth > 12)
                    throw new ArgumentOutOfRangeException("month", "month has to be between 1 and 12");
                
                _date = new DateTime(_date.Year, intMonth, _date.Day);
                UpdateRefCode();
            }
            public void SetDay(object day)
            {
                int intDay;
                if (day == null)
                    throw new ArgumentNullException("day", "month can't be a null value");
                try
                {
                    intDay = Convert.ToInt16(day);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not a valid day." + ex.Message, "day");
                }
                if (intDay < 1 || intDay > 31)
                    throw new ArgumentOutOfRangeException("day", "day has to be between 1 and 31");

                _date = new DateTime(_date.Year, _date.Month, intDay);
                UpdateRefCode();
            }

            public void SetDate(object date)
            {
                try
                {
                    var sdate = Convert.ToString(date);
                    if(sdate.Length < 8 || sdate.Length > 8)
                        throw new Exception("Value is not a valid date YYYYMMDD");
                    SetDate(sdate.Substring(0, 4), sdate.Substring(4, 2), sdate.Substring(6, 2));
                }
                catch (Exception)
                {
                    throw new Exception("Value is not a valid date YYYYMMDD");
                }
            }
            public void SetDate(object year, object month, object day)
            {
                SetYear(year);
                SetMonth(month);
                SetDay(day);
                UpdateRefCode();
            }

            public void SetDate(DateTime date)
            {
                SetDate(date.Year, date.Month, date.Day);
            }
            private void UpdateRefCode()
            {
                RefCode = MyUtilities.ToString(_date);
            }

            public string GetRefCode()
            {
                return RefCode;
            }

            public DateTime GetDate()
            {
                return _date;
            }

        }

        public class ReferenceTime : ReferenceCodeItem//HH:MM
        {
            private TimeSpan _time;
            private new string RefCode;
            private new readonly string FieldType;

            public ReferenceTime()
            {
                SetTime(00, 00);
                FieldType = "T";
            }
            public ReferenceCodeItem ToItem()
            {
                var item = new ReferenceCodeItem()
                {
                    EntityType = EntityType,
                    EntityValue = EntityValue,
                    RefNo = RefNo,
                    SeqNum = SeqNum,
                    RefCode = RefCode,
                    FieldType = FieldType
                };
                return item;
            }

            public void SetHour(object hour)
            {
                int intHour;
                if (hour == null)
                    throw new ArgumentNullException("hour", "hour can't be a null value");
                try
                {
                    intHour = Convert.ToInt16(hour);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not a valid hour." + ex.Message, "hour");
                }
                if (intHour < 0 || intHour > 24)
                    throw new ArgumentOutOfRangeException("hour", "hour can't be less than 0 or more than 24");

                _time = new TimeSpan(intHour, _time.Minutes, _time.Seconds);
                UpdateRefCode();
            }
            public void SetMinutes(object minutes)
            {
                int intMinutes;
                if (minutes == null)
                    throw new ArgumentNullException("minutes", "minutes can't be a null value");
                try
                {
                    intMinutes = Convert.ToInt16(minutes);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not a valid minutes." + ex.Message, "minutes");
                }
                if (intMinutes < 0 || intMinutes > 59)
                    throw new ArgumentOutOfRangeException("minutes", "minutes can't be less than 0 or more than 59");

                _time = new TimeSpan(_time.Hours, intMinutes, _time.Seconds);
                UpdateRefCode();
            }
            public void SetSeconds(object seconds)
            {
                int intSeconds;
                if (seconds == null)
                    throw new ArgumentNullException("seconds", "minutes can't be a null value");
                try
                {
                    intSeconds = Convert.ToInt16(seconds);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Value is not valid for seconds. " + ex.Message, "seconds");
                }
                if (intSeconds < 0 || intSeconds > 59)
                    throw new ArgumentOutOfRangeException("seconds", "minutes can't be less than 0 or more than 59");

                _time = new TimeSpan(_time.Hours, _time.Minutes, intSeconds);
                UpdateRefCode();
            }
            public void SetTime(object hours, object minutes, object seconds)
            {
                SetHour(hours);
                SetMinutes(minutes);
                SetSeconds(seconds);
                UpdateRefCode();
            }
            public void SetTime(object hours, object minutes)
            {
                SetHour(hours);
                SetMinutes(minutes);
                SetSeconds(00);
                UpdateRefCode();
            }

            public void SetTime(TimeSpan time)
            {
                SetTime(time.Hours, time.Minutes, time.Seconds);
            }
            private void UpdateRefCode()
            {
                RefCode = MyUtilities.ToString(_time, "hh:mm");
            }
            public string GetRefCode()
            {
                return RefCode;
            }

            public TimeSpan GetTime()
            {
                return _time;
            }
        }

        public class ReferenceText : ReferenceCodeItem
        {
            private new string RefCode;
            private new readonly string FieldType;

            public ReferenceText()
            {
                RefCode = "";
                FieldType = null;
            }
            public ReferenceText(string value)
            {
                SetValue(value);
                FieldType = null;
            }
            public ReferenceCodeItem ToItem()
            {
                var item = new ReferenceCodeItem()
                {
                    EntityType = EntityType,
                    EntityValue = EntityValue,
                    RefNo = RefNo,
                    SeqNum = SeqNum,
                    RefCode = RefCode,
                    FieldType = FieldType
                };
                return item;
            }

            public void SetValue(string value)
            {
                RefCode = value;
            }

            public string GetValue()
            {
                return GetRefCode();
            }

            public string GetRefCode()
            {
                return RefCode;
            }
        }

        public class ReferenceBoolean : ReferenceCodeItem
        {
            private new string RefCode;
            private new readonly string FieldType;

            public ReferenceBoolean()
            {
                FieldType = null;
                RefCode = "";
            }
            public ReferenceCodeItem ToItem()
            {
                var item = new ReferenceCodeItem()
                {
                    EntityType = EntityType,
                    EntityValue = EntityValue,
                    RefNo = RefNo,
                    SeqNum = SeqNum,
                    RefCode = RefCode,
                    FieldType = FieldType
                };
                return item;
            }

            public void SetValue(object value)
            {
                RefCode = MyUtilities.IsTrue(value) ? "Y" : "N";
            }

            public bool GetValue()
            {
                return MyUtilities.IsTrue(RefCode);
            }

            public string GetRefCode()
            {
                return RefCode;
            }
        }

        public class ReferenceNumeric : ReferenceCodeItem
        {
            private new string RefCode;
            private new readonly string FieldType;

            public ReferenceNumeric()
            {
                RefCode = "";
                FieldType = "N";
            }

            public ReferenceNumeric(object value)
            {
                FieldType = "N";
                SetValue(value);
            }

            public void SetValue(object value)
            {
                if (value == null)
                    RefCode = null;
                else if(Convert.ToString(value).Equals(""))
                    RefCode = "";
                else
                    RefCode = "" + Convert.ToInt32(value);
            }
            public ReferenceCodeItem ToItem()
            {
                var item = new ReferenceCodeItem()
                {
                    EntityType = EntityType,
                    EntityValue = EntityValue,
                    RefNo = RefNo,
                    SeqNum = SeqNum,
                    RefCode = RefCode,
                    FieldType = FieldType
                };
                return item;
            }
            public int GetValue()
            {
                return Convert.ToInt32(GetRefCode());
            }

            public string GetRefCode()
            {
                return RefCode;
            }
        }

    }
}

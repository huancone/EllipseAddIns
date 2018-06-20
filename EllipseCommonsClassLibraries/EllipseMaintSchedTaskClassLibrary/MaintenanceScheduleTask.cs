using System.Collections.Generic;

namespace EllipseMaintSchedTaskClassLibrary
{
    public class MaintenanceScheduleTask
    {
        public string RecType;
        public string EquipmentGrpId;
        public string EquipmentNo;
        public string EquipmentDescription;
        public string CompCode;
        public string CompModCode;
        public string MaintenanceSchTask;

        public string ConAstSegFr;
        public string ConAstSegTo;

        public string SchedDescription1;
        public string SchedDescription2;
        public string WorkGroup;
        public string AssignPerson;
        public string JobDescCode;
        public string StdJobNo;
        public string DistrictCode;
        public string AutoRequisitionInd;
        public string MsHistFlag;

        public string SchedInd;

        public string StatType1;
        public string LastSchedStat1;
        public string SchedFreq1;
        public string LastPerfStat1;

        public string StatType2;
        public string LastSchedStat2;
        public string SchedFreq2;
        public string LastPerfStat2;

        public string LastSchedDate;
        public string LastPerfDate;

        public string NextSchedDate;
        public string NextSchedStat;
        public string NextSchedValue;

        public string ShutdownType;
        public string ShutdownEquip;
        public string ShutdownNo;
        public string CondMonPos;
        public string CondMonType;

        public string StatutoryFlg;

        public string OccurrenceType;
        public string DayOfWeek;
        public string DayOfMonth;
        public string StartMonth;
        public string StartYear;
        public string AllowMultiple;
    }
    public static class MstType
    {
        public static string Egi = "GS";
        public static string Equipment = "ES";
    }
    public static class MstIndicatorList
    {
        public static string LastSchedDate = "LAST SCHED DATE";
        public static string LastSchedDateCode = "1";
        public static string LastSchedStat = "LAST SCHED STAT";
        public static string LastSchedStatCode = "2";
        public static string LastPerfDate = "LAST PERF DATE";
        public static string LastPerfDateCode = "3";
        public static string LastPerfStat = "LAST PERF STAT";
        public static string LastPerfStatCode = "4";
        public static string DualLastSched = "DUAL LAST SCHED";
        public static string DualLastSchedCode = "5";
        public static string DualLastPerf = "DUAL LAST PERF";
        public static string DualLastPerfCode = "6";
        public static string FixedDate = "FIXED DATE";
        public static string FixedDateCode = "7";
        public static string FixedDay = "FIXED DAY";
        public static string FixedDayCode = "8";
        public static string Inactive = "INACTIVE TASK";
        public static string InactiveCode = "9";

        public static string Active = "ACTIVE";

        public static string GetIndicatorCode(string statusName)
        {
            if (statusName == LastSchedDate)
                return LastSchedDateCode;
            if (statusName == LastSchedStat)
                return LastSchedStatCode;
            if (statusName == LastPerfDate)
                return LastPerfDateCode;
            if (statusName == LastPerfStat)
                return LastPerfStatCode;
            if (statusName == DualLastSched)
                return DualLastSchedCode;
            if (statusName == DualLastPerf)
                return DualLastPerfCode;
            if (statusName == FixedDate)
                return FixedDateCode;
            if (statusName == FixedDay)
                return FixedDayCode;
            if (statusName == Inactive)
                return InactiveCode;
            return null;
        }

        public static string GetIndicatorName(string statusCode)
        {
            if (statusCode == LastSchedDateCode)
                return LastSchedDate;
            if (statusCode == LastSchedStatCode)
                return LastSchedStat;
            if (statusCode == LastPerfDateCode)
                return LastPerfDate;
            if (statusCode == LastPerfStatCode)
                return LastPerfStat;
            if (statusCode == DualLastSchedCode)
                return DualLastSched;
            if (statusCode == DualLastPerfCode)
                return DualLastPerf;
            if (statusCode == FixedDateCode)
                return FixedDate;
            if (statusCode == FixedDayCode)
                return FixedDay;
            if (statusCode == InactiveCode)
                return Inactive;
            return null;
        }

        public static List<string> GetIndicatorNames()
        {
            var list = new List<string> { LastSchedDate, LastSchedStat, LastPerfDate, LastPerfStat, DualLastSched, DualLastPerf, FixedDate, FixedDay, Inactive };
            return list;
        }
        public static List<string> GetIndicatorCodes()
        {
            var list = new List<string> { LastSchedDateCode, LastSchedStatCode, LastPerfDateCode, LastPerfStatCode, DualLastSchedCode, DualLastPerfCode, FixedDateCode, FixedDayCode, InactiveCode };
            return list;
        }
        public static List<string> GetActiveIndicatorNames()
        {
            var list = new List<string> { LastSchedDate, LastSchedStat, LastPerfDate, LastPerfStat, DualLastSched, DualLastPerf, FixedDate, FixedDay };
            return list;
        }
        public static List<string> GetActiveIndicatorCodes()
        {
            var list = new List<string> { LastSchedDateCode, LastSchedStatCode, LastPerfDateCode, LastPerfStatCode, DualLastSchedCode, DualLastPerfCode, FixedDateCode, FixedDayCode };
            return list;
        }

        public static List<string> GetIndicatorsList(string separator = " - ")
        {
            var list = new List<string> { LastSchedDateCode + separator + LastSchedDate, LastSchedStatCode + separator + LastSchedStat, LastPerfDateCode + separator + LastPerfDate, LastPerfStatCode + separator + LastPerfStat, DualLastSchedCode + separator + DualLastSched, DualLastPerfCode + separator + DualLastPerf, FixedDateCode + separator + FixedDate, FixedDayCode + separator + FixedDay, InactiveCode + separator + Inactive };
            return list;
        }


    }
    

}

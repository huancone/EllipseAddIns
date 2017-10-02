using System.Diagnostics.CodeAnalysis;

namespace EllipseCommonsClassLibrary.Utilities.MyDateTime
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public static class Formats
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
}

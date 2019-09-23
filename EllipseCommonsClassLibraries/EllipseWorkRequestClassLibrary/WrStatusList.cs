using System.Collections.Generic;

namespace EllipseWorkRequestClassLibrary
{
    public static class WrStatusList
    {
        public static string Open = "OPEN";
        public static string OpenCode = "O";
        public static string Closed = "CLOSED";
        public static string ClosedCode = "C";
        public static string Cancelled = "CANCELLED";
        public static string CancelledCode = "L";
        public static string InWork = "IN_WORK";
        public static string InWorkCode = "W";
        public static string Estimated = "ESTIMATED";
        public static string EstimatedCode = "E";

        public static string Uncompleted = "UNCOMPLETED";

        /// <summary>
        /// Obtiene el código del estado a partir del nombre (Ej. Parámetro OPEN, resultado O)
        /// </summary>
        /// <param name="statusName"></param>
        /// <returns></returns>
        public static string GetStatusCode(string statusName)
        {
            if (statusName == Open)
                return OpenCode;
            if (statusName == Closed)
                return ClosedCode;
            if (statusName == Cancelled)
                return CancelledCode;
            if (statusName == InWork)
                return InWorkCode;
            if (statusName == Estimated)
                return EstimatedCode;
            return null;
        }
        /// <summary>
        /// Obtiene el nombre de un estado a partir del código (Ej. Parámetro O, resultado OPEN)
        /// </summary>
        /// <param name="statusCode"></param>
        /// <returns></returns>
        public static string GetStatusName(string statusCode)
        {
            if (statusCode == OpenCode)
                return Open;
            if (statusCode == ClosedCode)
                return Closed;
            if (statusCode == CancelledCode)
                return Cancelled;
            if (statusCode == InWorkCode)
                return InWork;
            if (statusCode == EstimatedCode)
                return Estimated;
            return null;
        }

        public static List<string> GetStatusNames()
        {
            var list = new List<string> { Open, Closed, Cancelled, InWork, Estimated };
            return list;
        }
        public static List<string> GetStatusCodes()
        {
            var list = new List<string> { OpenCode, ClosedCode, CancelledCode, InWorkCode, EstimatedCode };
            return list;
        }
        public static List<string> GetUncompletedStatusNames()
        {
            var list = new List<string> { Open, InWork, Estimated };
            return list;
        }
        public static List<string> GetUncompletedStatusCodes()
        {
            var list = new List<string> { OpenCode, InWorkCode, EstimatedCode };
            return list;
        }


    }

}

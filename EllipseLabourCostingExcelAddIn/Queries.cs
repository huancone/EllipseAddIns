using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Utilities;

namespace EllipseLabourCostingExcelAddIn
{
    internal static class Queries
    {
        public static string GetGroupEmployeesQuery(string workGroup, string dbReference, string dbLink)
        {
            var query = "WITH EMPLOYEES AS (" +
                        "  SELECT " +
                        "       (CASE WHEN TRIM(EMP.PREF_NAME) IS NULL THEN EMP.EMPLOYEE_ID WHEN REGEXP_LIKE(EMP.PREF_NAME, '^[^a-zA-Z]*$') THEN EMP.PREF_NAME ELSE EMP.EMPLOYEE_ID END) CEDULA, " +
                        "       EMP.EMPLOYEE_ID, " +
                        "       EMP.PREF_NAME, " +
                        "       EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE " +
                        "   FROM " +
                        "       " + dbReference + ".MSF723" + dbLink + " WE " +
                        "       INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP " +
                        "       ON EMP.EMPLOYEE_ID   = WE.EMPLOYEE_ID " +
                        "       OR TRIM(EMP.PREF_NAME)   = TRIM(WE.EMPLOYEE_ID) " +
                        "   WHERE " +
                        "       WE.WORK_GROUP = '" + workGroup + "' " +
                        "       AND ( WE.STOP_DT_REVSD   = '00000000' " +
                        "             OR ( 99999999 - WE.STOP_DT_REVSD ) >= TO_CHAR( SYSDATE, 'YYYYMMDD' ) ) " +
                        "       AND WE.REC_723_TYPE    = 'W' " +
                        "), " +
                        "EMPORDERED AS(" +
                        "  SELECT EMP.*, ROW_NUMBER() OVER(PARTITION BY CEDULA ORDER BY EMPLOYEE_ID) RK FROM EMPLOYEES EMP" +
                        " ) " +
                        "SELECT * FROM EMPORDERED WHERE RK = 1";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

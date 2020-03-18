using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseLabourCostingExcelAddIn
{
    public static class Queries
    {
        public static string GetGroupEmployeesQuery(string workGroup, string dbReference, string dbLink)
        {
            var query = "SELECT " +
                        "     EMP.EMPLOYEE_ID   CEDULA, " +
                        "     EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE " +
                        " FROM " +
                        "     " + dbReference + ".MSF723" + dbLink + " WE " +
                        "     INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP " +
                        "     ON EMP.EMPLOYEE_ID   = WE.EMPLOYEE_ID " +
                        "     OR TRIM(EMP.PREF_NAME)   = TRIM(WE.EMPLOYEE_ID) " +
                        " WHERE " +
                        "     WE.WORK_GROUP = '" + workGroup + "' " +
                        "     AND ( WE.STOP_DT_REVSD   = '00000000' " +
                        "           OR ( 99999999 - WE.STOP_DT_REVSD ) >= TO_CHAR( SYSDATE, 'YYYYMMDD' ) ) " +
                        "     AND WE.REC_723_TYPE    = 'W'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}

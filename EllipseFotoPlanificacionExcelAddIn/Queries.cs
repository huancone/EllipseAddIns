using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;
using System.Threading;
using System.Windows.Forms.VisualStyles;

namespace EllipseFotoPlanificacionExcelAddIn
{
    public static partial class Queries
    {
        public static string GetFetchSigmanPhotoQuery(string dbReference, string dbLink, string districtCode, string monitoringPeriod, string workGroup)
        {
            var workGroupParam = "";
            if (!string.IsNullOrWhiteSpace(workGroup))
                workGroupParam = " AND PL.WORK_GROUP = '" + workGroup + "' ";
            //escribimos el query
            var query = "" +
                        " SELECT" +
                        " PL.PERIODO_MONITOREO, PL.WORK_GROUP, PL.EQUIP_NO, PL.COMPONENT_CODE, PL.MODIFIED_CODE, PL.MAINT_SCH_TASK, PL.NEXT_SCH_DATE, PL.LAST_PERF_DATE, PL.FECHA_FOTO " +
                        " FROM CUMPLIMIENTO_MST PL" +
                        " WHERE PL.PERIODO_MONITOREO = '" + monitoringPeriod + "'"
                        + workGroupParam;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}
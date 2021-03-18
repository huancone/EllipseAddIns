using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Utilities;

namespace PlaneacionFerrocarril.PlanHistory
{
    internal class Queries
    {
        internal static IQueryParamCollection GetReviewPlanHistoryQuery(string startDate, string finishDate, string workGroup, string idConcepto)
        {
            var paramDate = string.IsNullOrWhiteSpace(finishDate) ? " AND FECHA = :" + nameof(startDate) : " AND FECHA BETWEEN :" + nameof(startDate) + " AND :" + nameof(finishDate);
            var paramWorkGroup = string.IsNullOrWhiteSpace(workGroup) ? "" : " AND GRUPO = :" + nameof(workGroup);
            var paramIdConcepto = string.IsNullOrWhiteSpace(idConcepto) ? "" : " AND ID_CONCEPTO = :" + nameof(idConcepto);



            var query = " SELECT FECHA, GRUPO, ID_CONCEPTO, CONCEPTO, VALOR1, VALOR2  FROM SIGMDC.HISTORIAL_PROGRAMACION WHERE" +
                        paramDate + paramWorkGroup + paramIdConcepto;

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(startDate), startDate));
            if(!string.IsNullOrWhiteSpace(finishDate))
                qpCollection.AddParam(new OracleParameter(nameof(finishDate), finishDate));
            if (!string.IsNullOrWhiteSpace(workGroup))
                qpCollection.AddParam(new OracleParameter(nameof(workGroup), workGroup));
            if (!string.IsNullOrWhiteSpace(idConcepto))
                qpCollection.AddParam(new OracleParameter(nameof(idConcepto), idConcepto));

            return qpCollection;
        }

        internal static IQueryParamCollection GetUpdatePlanHistoryItemQuery(string fecha, string grupo, string idConcepto, string concepto, string valor1, string valor2)
        {
            var query = "MERGE INTO SIGMDC.HISTORIAL_PROGRAMACION PH USING " +
                        " (SELECT " +
                        "  :" + nameof(fecha) + " FECHA, " +
                        "  :" + nameof(grupo) + " GRUPO, " +
                        "  :" + nameof(idConcepto) + " ID_CONCEPTO, " +
                        "  :" + nameof(concepto) + " CONCEPTO, " +
                        "  :" + nameof(valor1) + " VALOR1, " +
                        "  :" + nameof(valor2) + " VALOR2 " +
                        "  FROM DUAL) PHI ON ( " +
                        "  PH.FECHA = PHI.FECHA AND PH.GRUPO = PHI.GRUPO AND PH.ID_CONCEPTO = PHI.ID_CONCEPTO " +
                        " ) " +
                        " WHEN MATCHED THEN UPDATE SET " +
                        "   PH.CONCEPTO = PHI.CONCEPTO," +
                        "   PH.VALOR1 = PHI.VALOR1," +
                        "   PH.VALOR2 = PHI.VALOR2" +
                        " WHEN NOT MATCHED THEN INSERT(" +
                        "   FECHA, " +
                        "   GRUPO, " +
                        "   ID_CONCEPTO, " +
                        "   CONCEPTO, " +
                        "   VALOR1, " +
                        "   VALOR2 " +
                        " ) " +
                        " VALUES(" +
                        "   PHI.FECHA, " +
                        "   PHI.GRUPO, " +
                        "   PHI.ID_CONCEPTO, " +
                        "   PHI.CONCEPTO, " +
                        "   PHI.VALOR1, " +
                        "   PHI.VALOR2 " +
                        " ) ";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var qpCollection = new OracleQueryParamCollection(query);
            qpCollection.AddParam(new OracleParameter(nameof(fecha), fecha));
            qpCollection.AddParam(new OracleParameter(nameof(grupo), grupo));
            qpCollection.AddParam(new OracleParameter(nameof(idConcepto), idConcepto));
            qpCollection.AddParam(new OracleParameter(nameof(concepto), concepto));
            qpCollection.AddParam(new OracleParameter(nameof(valor1), valor1));
            qpCollection.AddParam(new OracleParameter(nameof(valor2), valor2));

            return qpCollection;
        }
    }
}

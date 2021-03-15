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
        public static IQueryParamCollection ReviewPlanHistoryQuery(string startDate, string finishDate, string workGroup, string idConcepto)
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
    }
}

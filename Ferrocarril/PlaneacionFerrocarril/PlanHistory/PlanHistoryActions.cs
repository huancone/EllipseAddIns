using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Ellipse;

namespace PlaneacionFerrocarril.PlanHistory
{
    public class PlanHistoryActions
    {
        public static List<PlanHistoryItem> ReviewPlanHistory(EllipseFunctions ef, string startDate, string finishDate, string workGroup, string idConcepto)
        {
            var list = new List<PlanHistoryItem>();
            var sqlQuery = Queries.ReviewPlanHistoryQuery(startDate, finishDate, workGroup, idConcepto);
            var dr = ef.GetQueryResult(sqlQuery);
            if (dr == null || dr.IsClosed)
                return list;
            while (dr.Read())
            {
                list.Add(new PlanHistoryItem(dr));
            }

            return list;
        }
    }
}

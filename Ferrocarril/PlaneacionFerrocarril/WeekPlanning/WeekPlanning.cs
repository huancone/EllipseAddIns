using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using SharedClassLibrary.Ellipse;

namespace PlaneacionFerrocarril
{
    public partial class WeekPlanning
    {
        public static List<TaskItem> GetWorkGroupTaskItems(EllipseFunctions eFunctions, string workGroup, string startDate, string finishDate, string additional)
        {
            var list = new List<TaskItem>();

            var sqlQuery = Queries.GetTaskResourcesQuery(workGroup, startDate, finishDate, additional);


            var dr = eFunctions.GetQueryResult(sqlQuery);
            if (dr == null || dr.IsClosed)
                return list;

            while (dr.Read())
                list.Add(new TaskItem(dr));

            return list;
        }
        public static List<WorkGroupResource> GetWorkGroupAvailableResources(EllipseFunctions eFunctions, string workGroup)
        {
            var list = new List<WorkGroupResource>();
            var sqlQuery = Queries.GetWorkGroupResourcesQuery(workGroup);

            var dr = eFunctions.GetQueryResult(sqlQuery);
            if (dr == null || dr.IsClosed)
                return list;

            while (dr.Read())
                list.Add(new WorkGroupResource(dr));

            return list;
        }
    }
}

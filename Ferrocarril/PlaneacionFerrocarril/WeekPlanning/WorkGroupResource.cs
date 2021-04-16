using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using SharedClassLibrary.Utilities;

namespace PlaneacionFerrocarril
{
    public class WorkGroupResource
    {
        public string Type;
        public string Description;
        public decimal ActualHours;
        public decimal EstimatedHours;

        public WorkGroupResource()
        {

        }
        public WorkGroupResource(IDataRecord dr)
        {
            Type = dr["RESOURCE_TYPE"].ToString().Trim();
            Description = dr["RES_DESC"].ToString().Trim();

            EstimatedHours = MyUtilities.ToDecimal(dr["EST_HRS"].ToString(), 2);
            ActualHours = MyUtilities.ToDecimal(dr["ACT_HRS"].ToString(), 2);

        }
    }
}

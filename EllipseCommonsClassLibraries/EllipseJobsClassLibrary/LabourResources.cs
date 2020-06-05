using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class LabourResources
    {
        public string WorkGroup { get; set; }
        public string ResourceCode { get; set; }
        public string Date { get; set; }
        public double Quantity { get; set; }
        public double AvailableLabourHours { get; set; }
        public double EstimatedLabourHours { get; set; }
        public double RealLabourHours { get; set; }
        public string EmployeeId { get; set; }
        public string EmployeeName { get; set; }
    }
}

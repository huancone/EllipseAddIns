using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class TaskSearchParam
    {
        public bool AdditionalInformation;
        public bool IncludeMst;
        public bool OverlappingDates;

        public TaskSearchParam()
        {
            AdditionalInformation = true;
            IncludeMst = true;
            OverlappingDates = false;
        }
    }
}

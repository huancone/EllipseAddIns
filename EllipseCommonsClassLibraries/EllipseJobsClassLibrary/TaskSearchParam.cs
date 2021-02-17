using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseJobsClassLibrary
{
    public class TaskSearchParam
    {

        public string StartDate;
        public string FinishDate;
        public bool AdditionalInformation;
        public bool IncludeMst;
        public bool OverlappingDates;
        public string DateInclude;
        public string District;
        public string SearchEntity;

        public List<string> WorkGroups;

        public TaskSearchParam()
        {
            AdditionalInformation = true;
            IncludeMst = true;
            OverlappingDates = false;
            WorkGroups = new List<string>();
        }
    }
}

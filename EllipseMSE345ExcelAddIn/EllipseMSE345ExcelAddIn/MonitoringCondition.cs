using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSE345ExcelAddIn
{
    public class MonitoringCondition
    {
        public string CautionLow;
        public string CautionUpper;
        public string ComponentCode;
        public string ComponentDescription;
        public string DangerLow;
        public string DangerUpper;
        public string Egi;
        public string Equipment;
        public string MeassureCode;
        public string MeassureDescription;
        public string ModifierCode;
        public string ModifierDescription;
        public string PositionCode;
        public string PositionDescription;
        public string Type;
        public string TypeDescription { get; set; }
    }
}

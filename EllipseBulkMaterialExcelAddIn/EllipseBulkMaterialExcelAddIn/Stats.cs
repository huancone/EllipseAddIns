using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseBulkMaterialExcelAddIn
{
    public class Stats
    {
        public string EquipNo { get; set; }
        public string StatType { get; set; }
        public decimal MeterValue { get; set; }
        public string StatDate { get; set; }

        public string Error { get; set; }
    }
}

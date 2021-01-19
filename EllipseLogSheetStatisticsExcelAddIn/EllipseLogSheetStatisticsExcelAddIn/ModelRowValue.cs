using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseLogSheetStatisticsExcelAddIn
{
    public class ModelRowValue
    {
        public string Code;
        public string EquipReference;
        public ObjectFlagValue Operator;
        public ObjectFlagValue Account;
        public ObjectFlagValue WorkOrder;
        public ObjectFlagValue Source;
        public ObjectFlagValue Destination;
        public ObjectFlagValue Material;

        public ObjectFlagValue[] Inputs;

        public ModelRowValue()
        {
            Code = "";
            Operator = new ObjectFlagValue();
            Account = new ObjectFlagValue();
            WorkOrder = new ObjectFlagValue();
            Source = new ObjectFlagValue();
            Destination = new ObjectFlagValue();
            Material = new ObjectFlagValue();
            Inputs = new ObjectFlagValue[10];
            for (var i = 0; i < 10; i++)
                Inputs[i] = new ObjectFlagValue();
        }


    }
}

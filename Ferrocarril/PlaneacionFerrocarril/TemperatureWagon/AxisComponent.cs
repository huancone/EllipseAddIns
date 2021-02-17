using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PlaneacionFerrocarril.TemperatureWagon
{
    public class AxisComponent
    {
        public string Axis;
        public string Differential;
        public string BearingLeft;
        public string BearingRight;
        public string WheelLeft;
        public string WheelRight;

        public static AxisComponent[] GetDefaultAxisComponents()
        {
            var list = new List<AxisComponent>();

            for (var i = 0; i < 6; i++)
            {
                var item = new AxisComponent();
                item.Axis = "" + (i + 1);
                item.Differential = "TEJE"+(i+1)+"DIF";
                item.BearingLeft = "TEJE" + (i + 1) + "ROI";
                item.BearingRight = "TEJE" + (i + 1) + "ROD";
                item.WheelLeft = "TEJE" + (i + 1) + "RUI";
                item.WheelRight = "TEJE" + (i + 1) + "RUD";

                list.Add(item);
            }

            return list.ToArray();
        }
    }
}

using System;
using System.Collections.Generic;

namespace CommonsClassLibrary.Utilities.Shifts
{
    public class ShiftPeriods
    {
        public static Slot[] GetHourToHourSlots()
        {
            var slotSlist = new List<Slot>();

            for (var i = 0; i < 24; i++)
            {
                var slot = new Slot((i + 1).ToString("00"));
                slot.SetStartTime(new TimeSpan(i, 00, 00));
                slot.SetEndTime(new TimeSpan(i + 1, 00, 00));

                slotSlist.Add(slot);
            }

            return slotSlist.ToArray();
        }

        public static Slot[] GetDailyNightSlots()
        {
            var slotArray = new Slot[2];

            slotArray[0] = new Slot("D");
            slotArray[0].SetStartTime(new TimeSpan(06, 00, 00));
            slotArray[0].SetEndTime(new TimeSpan(18, 00, 00));
            slotArray[1] = new Slot("N");
            slotArray[1].SetStartTime(new TimeSpan(18, 00, 00));
            slotArray[1].SetEndTime(new TimeSpan(01, 06, 00, 00)); //dia siguiente

            return slotArray;
        }

        public static Slot[] GetDailyZeroSlots()
        {
            var slotArray = new Slot[1];

            slotArray[0] = new Slot("A");
            slotArray[0].SetStartTime(new TimeSpan(00, 00, 00));
            slotArray[0].SetEndTime(new TimeSpan(24, 00, 00));

            return slotArray;
        }

        public static Slot[] GetDailyMorningSlots()
        {
            var slotArray = new Slot[1];

            slotArray[0] = new Slot("DY");
            slotArray[0].SetStartTime(new TimeSpan(06, 00, 00));
            slotArray[0].SetEndTime(new TimeSpan(01, 06, 00, 00)); //dia siguiente

            return slotArray;
        }
    }
}
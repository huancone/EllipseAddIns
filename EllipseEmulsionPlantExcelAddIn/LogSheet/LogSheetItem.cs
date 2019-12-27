using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Classes;

namespace EllipseEmulsionPlantExcelAddIn.LogSheet
{
    public class LogSheetItem
    {
        public string ModelName;
        public string Date;
        public string ShiftCode;

        public List<LogSheetHeaderItem> ModelHeader;
        public List<LogSheetEquipmentModelItem> ModelItems;
        public List<LogSheetEquipmentInputItem> InputItems;

        public LogSheetItem(string modelName, string date, string shiftCode)
        {
            ModelName = modelName;
            Date = date;
            ShiftCode = shiftCode;
        }

        public LogSheetItem(string modelName, string date, string shiftCode, List<LogSheetEquipmentInputItem> inputItems)
        {
            ModelName = modelName;
            Date = date;
            ShiftCode = shiftCode;
            InputItems = inputItems;
        }

        public LogSheetItem(string modelName, string date, string shiftCode, List<LogSheetEquipmentInputItem> inputItems, List<LogSheetHeaderItem> modelHeader, List<LogSheetEquipmentModelItem> modelItems)
        {
            ModelName = modelName;
            Date = date;
            ShiftCode = shiftCode;
            InputItems = inputItems;
            ModelHeader = modelHeader;
            ModelItems = modelItems;
        }
    }
}

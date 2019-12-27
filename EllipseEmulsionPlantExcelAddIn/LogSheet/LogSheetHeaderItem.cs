using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseEmulsionPlantExcelAddIn.LogSheet
{
    public class LogSheetHeaderItem
    {
        public string ModelCode;
        public int Index;
        public string HeaderName;
        public string ValueType;

        public LogSheetHeaderItem()
        {

        }

        public LogSheetHeaderItem(string modelCode, int index, string headerName, string valueType)
        {
            ModelCode = modelCode;
            Index = index;
            HeaderName = headerName;
            ValueType = valueType;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Utilities;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseEmulsionPlantExcelAddIn.LogSheet
{
    public class LogSheetEquipmentInputItem
    {
        public string Action;

        public MyKeyValuePair<string, string> PlantNo;
        public MyKeyValuePair<string, string> Operator;
        public MyKeyValuePair<string, string> AccountCode;
        public MyKeyValuePair<string, string> WorkOrder;
               
        public MyKeyValuePair<string, string> Input1;
        public MyKeyValuePair<string, string> Input2;
        public MyKeyValuePair<string, string> Input3;
        public MyKeyValuePair<string, string> Input4;
        public MyKeyValuePair<string, string> Input5;
        public MyKeyValuePair<string, string> Input6;
        public MyKeyValuePair<string, string> Input7;
        public MyKeyValuePair<string, string> Input8;
        public MyKeyValuePair<string, string> Input9;
        //public MyKeyValuePair<string, string> Input10;

        public LogSheetEquipmentInputItem()
        {
            Action = ""; //S ???, I Insert
            PlantNo = new MyKeyValuePair<string, string>("PLANT_NO2I", null);
            Operator = new MyKeyValuePair<string, string>("OPERATOR2I", null);
            AccountCode = new MyKeyValuePair<string, string>("ACCOUNT_CODE2I", null);
            WorkOrder = new MyKeyValuePair<string, string>("WORK_ORDER2I", null);
            Input1 = new MyKeyValuePair<string, string>("INPUT_12I", null);
            Input2 = new MyKeyValuePair<string, string>("INPUT_22I", null);
            Input3 = new MyKeyValuePair<string, string>("INPUT_32I", null);
            Input4 = new MyKeyValuePair<string, string>("INPUT_42I", null);
            Input5 = new MyKeyValuePair<string, string>("INPUT_52I", null);
            Input6 = new MyKeyValuePair<string, string>("INPUT_62I", null);
            Input7 = new MyKeyValuePair<string, string>("INPUT_72I", null);
            Input8 = new MyKeyValuePair<string, string>("INPUT_82I", null);
            Input9 = new MyKeyValuePair<string, string>("INPUT_92I", null);
            //Input10 = new MyKeyValuePair<string, string>("INPUT_102I", null);

        }
    }
}

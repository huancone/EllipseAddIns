using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSE140DeleteExcelAddIn
{
    public class Requisition
    {
        public string District { get; set; }
        public string Requesition { get; set; }
        public string Item { get; set; }
        public string Warehouse { get; set; }
        public string ReqType { get; set; }
        public string ItemType { get; set; }
        public string StockCode { get; set; }
        public string QtyReq { get; set; }
        public string UnitOfMeasure { get; set; }

        public Requisition(System.Array MyValues)
        {
            this.District = GetValueToCell(MyValues, 1, 1);
            this.Requesition = GetValueToCell(MyValues, 1, 2);
            this.Item = GetValueToCell(MyValues, 1, 3);
            this.Warehouse = GetValueToCell(MyValues, 1, 4);
            this.ReqType = GetValueToCell(MyValues, 1, 5);
            this.ItemType = GetValueToCell(MyValues, 1, 6);
            this.StockCode = GetValueToCell(MyValues, 1, 7);
            this.QtyReq = GetValueToCell(MyValues, 1, 8);
            this.UnitOfMeasure = GetValueToCell(MyValues, 1, 9);
        }

        private static string GetValueToCell(System.Array MyValues, int x, int y)
        {
            if (MyValues.GetValue(x, y) != null)
            {
                return MyValues.GetValue(x, y).ToString();
            }
            return "";
        }
    }
}

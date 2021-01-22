using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseStockCodesExcelAddIn
{
    public static class PurchaseOrderActions
    {
        public static class OrderStatus
        {
            public static string UnprintedCode = "0";
            public static string Unprinted = "UNPRINTED";
            public static string PrintedCode = "1";
            public static string Printed = "PRINTED";
            public static string ModifiedCode = "2";
            public static string Modified = "MODIFIED";
            public static string CancelledCode = "3";
            public static string Cancelled = "CANCELLED";
            public static string CompletedCode = "9";
            public static string Completed = "COMPLETED";
            public static string UncompletedCode = "U";
            public static string Uncompleted = "UNCOMPLETED";

            public static Dictionary<string, string> GetStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {UnprintedCode, Unprinted},
                    {PrintedCode, Printed},
                    {ModifiedCode, Modified},
                    {CancelledCode, Cancelled},
                    {CompletedCode, Completed},
                    {UncompletedCode, Uncompleted}
                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
    }
}

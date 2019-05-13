using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.SearchConstants
{
    public static class RequisitionStatus
    {
        public static KeyValuePair<string, string> Pending = new KeyValuePair<string, string>("P", "PENDING");
        public static KeyValuePair<string, string> Unauthorized = new KeyValuePair<string, string>("U", "UNAUTHORIZED");
        public static KeyValuePair<string, string> Authorized = new KeyValuePair<string, string>("A", "AUTHORIZED");
        public static KeyValuePair<string, string> Awaiting = new KeyValuePair<string, string>("2", "AWAITING RECEIPT");
        public static KeyValuePair<string, string> Completed = new KeyValuePair<string, string>("9", "COMPLETED");
        public static KeyValuePair<string, string> Uncompleted = new KeyValuePair<string, string>("U", "UNCOMPLETED");


        public static List<string> GetStatusNames(bool uncompletedCustom = false)
        {
            var list = new List<string>
            {
                Pending.Value,
                Unauthorized.Value,
                Authorized.Value,
                Awaiting.Value,
                Completed.Value,
            };
            if (uncompletedCustom)
                list.Add(Uncompleted.Value);
            return list;
        }

        public static List<KeyValuePair<string, string>> GetStatusList(bool uncompletedCustom = false)
        {
            var list = new List<KeyValuePair<string, string>>
            {
                Pending,
                Unauthorized,
                Authorized,
                Awaiting,
                Completed
            };
            if (uncompletedCustom)
                list.Add(Uncompleted);
            return list;
        }
        private static List<KeyValuePair<string, string>> GetUncompletedStatusList()
        {
            var list = new List<KeyValuePair<string, string>>
            {
                Pending,
                Unauthorized,
                Authorized,
                Awaiting,
            };
            return list;
        }

        public static List<string> GetUncompletedStatusNames()
        {
            var list = new List<string>();
            foreach (var item in GetUncompletedStatusList())
                list.Add(item.Value);
            return list;
        }

        public static List<string> GetUncompletedStatusCodes()
        {
            var list = new List<string>();
            foreach (var item in GetUncompletedStatusList())
                list.Add(item.Key);
            return list;
        }

        public static string GetStatusCode(string statusName)
        {
            foreach(var item in GetStatusList(true))
                if (item.Value.Equals(statusName))
                    return item.Key;
            return null;
        }

        public static string GetStatusName(string statusCode)
        {
            foreach (var item in GetStatusList(true))
                if (item.Key.Equals(statusCode))
                    return item.Value;
            return null;
        }
    }

    public static class RequisitionHdrStatus
    {
        public static KeyValuePair<string, string> Pending = new KeyValuePair<string, string>("P", "PENDING");
        public static KeyValuePair<string, string> NotPrinted = new KeyValuePair<string, string>("0", "NOT PRINTED");
        public static KeyValuePair<string, string> PrintRequested = new KeyValuePair<string, string>("1", "PRINT REQUESTED");
        public static KeyValuePair<string, string> PartiallyAcquitted = new KeyValuePair<string, string>("2", "PARTIALLY ACQUITTED");
        public static KeyValuePair<string, string> IdrCompleted = new KeyValuePair<string, string>("3", "IDR COMPLETED");
        public static KeyValuePair<string, string> Complete = new KeyValuePair<string, string>("9", "COMPLETE INDICATOR");
        public static KeyValuePair<string, string> Uncompleted = new KeyValuePair<string, string>("U", "UNCOMPLETED");

        public static List<string> GetStatusNames(bool uncompletedCustom = false)
        {
            var list = new List<string>
            {
                Pending.Value,
                NotPrinted.Value,
                PrintRequested.Value,
                PartiallyAcquitted.Value,
                IdrCompleted.Value,
                Complete.Value,
            };
            if (uncompletedCustom)
                list.Add(Uncompleted.Value);
            return list;
        }

        private static List<KeyValuePair<string, string>> GetStatusList(bool uncompletedCustom = false)
        {
            var list = new List<KeyValuePair<string, string>>
            {
                Pending,
                NotPrinted,
                PrintRequested,
                PartiallyAcquitted,
                IdrCompleted,
                Complete
            };
            if (uncompletedCustom)
                list.Add(Uncompleted);
            return list;
        }
        private static List<KeyValuePair<string, string>> GetUncompletedStatusList()
        {
            var list = new List<KeyValuePair<string, string>>
            {
                Pending,
                NotPrinted,
                PrintRequested,
                PartiallyAcquitted,
                IdrCompleted,
            };
            return list;
        }

        public static List<string> GetUncompletedStatusNames()
        {
            var list = new List<string>();
            foreach (var item in GetUncompletedStatusList())
                list.Add(item.Value);
            return list;
        }

        public static List<string> GetUncompletedStatusCodes()
        {
            var list = new List<string>();
            foreach (var item in GetUncompletedStatusList())
                list.Add(item.Key);
            return list;
        }

        public static string GetStatusCode(string statusName)
        {
            foreach (var item in GetStatusList(true))
                if (item.Value.Equals(statusName))
                    return item.Key;
            return null;
        }

        public static string GetStatusName(string statusCode)
        {
            foreach (var item in GetStatusList(true))
                if (item.Key.Equals(statusCode))
                    return item.Value;
            return null;
        }
    }
}


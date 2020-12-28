using System.Collections.Generic;
using System.Linq;
using SharedClassLibrary.Ellipse;

namespace EllipseRequisitionClassLibrary
{
    public static class Requisition
    {
        public static class AuthorizationStatus
        {
            public static string UnauthorizedCode = "U";
            public static string Unauthorized = "UNAUTHORIZED";
            public static string AuthorizedCode = "A";
            public static string Authorized = "AUTHORIZED";
            public static string ApprovalNotifyCode = "N";
            public static string ApprovalNotify = "NOTIFY OF APPROVAL";

            public static Dictionary<string, string> GetAuthorizationStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {UnauthorizedCode, Unauthorized},
                    {AuthorizedCode, Authorized},
                    {ApprovalNotifyCode, ApprovalNotify}
                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetAuthorizationStatusList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetAuthorizationStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }

        public static class TransactionType
        {
            public static Dictionary<string, string> GetTransactionTypeList(EllipseFunctions ellipseFunctions)
            {
                var codeList = ellipseFunctions.GetItemCodes("IT");
                return codeList.ToDictionary(code => code.Code, code => code.Description);
            }

            public static string GetTypeCode(string statusName, EllipseFunctions ellipseFunctions)
            {
                var statusDictionary = GetTransactionTypeList(ellipseFunctions);
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetTypeName(string statusCode, EllipseFunctions ellipseFunctions)
            {
                var statusDictionary = GetTransactionTypeList(ellipseFunctions);
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
        public static class PriorityCodes
        {
            public static Dictionary<string, string> GetPrioriyCodesList(EllipseFunctions ellipseFunctions)
            {
                var codeList = ellipseFunctions.GetItemCodes("PI");
                return codeList.ToDictionary(code => code.Code, code => code.Description);
            }

            public static string GetPriorityCode(string statusName, EllipseFunctions ellipseFunctions)
            {
                var statusDictionary = GetPrioriyCodesList(ellipseFunctions);
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetPriorityName(string statusCode, EllipseFunctions ellipseFunctions)
            {
                var statusDictionary = GetPrioriyCodesList(ellipseFunctions);
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
        public static class RequisitionType
        {
            public static string WarehouseTransferCode = "WH";
            public static string WarehouseTransfer = "Warehouse Transfer";
            public static string DistrictTransferCode = "TR";
            public static string DistrictTransfer = "District Transfer";
            public static string NormalRequisitionCode = "NI";
            public static string NormalRequisition = "Normal Requisition";
            public static string ReturnToSupplierCode = "RS";
            public static string ReturnToSupplier = "Return to Supplier";
            public static string RecallRequisitionCode = "RC";
            public static string RecallRequisition = "Recall Requisition";
            public static string LoanRequisitionCode = "LN";
            public static string LoanRequisition = "Loan Requisition";
            public static string LoanReturnCode = "LR";
            public static string LoanReturn = "Loan Return";
            public static string DisposalRequisitionCode = "DI";
            public static string DisposalRequisition = "Disposal Requisition";
            public static string CreditRequisitionCode = "CR";
            public static string CreditRequisition = "Credit/Return Requisition";
            public static string CashSaleCode = "SC";
            public static string CashSale = "Cash/Stores Sale";
            public static string CreditSaleCode = "SR";
            public static string CreditSale = "Credit Sale";
            public static string RotationCreditRequisitionCode = "TC";
            public static string RotationCreditRequisition = "Rotation Cred. Requisition";
            public static string RotationRecallCode = "TL";
            public static string RotationRecall = "Rotation Recall";
            public static string RotationIssueCode = "TI";
            public static string RotationIssue = "Rotation Issue";
            public static string ShortFormCode = "NS";
            public static string ShortForm = "Short Form";
            public static string ManualRequisitionCode = "NM";
            public static string ManualRequisition = "Manual Requisition";
            public static string ReferredDemandCode = "RD";
            public static string ReferredDemand = "Referred Demand";

            public static Dictionary<string, string> GetRequisitionTypeList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {WarehouseTransferCode, WarehouseTransfer},
                    {DistrictTransferCode, DistrictTransfer},
                    {NormalRequisitionCode, NormalRequisition},
                    {ReturnToSupplierCode, ReturnToSupplier},
                    {RecallRequisitionCode, RecallRequisition},
                    {LoanRequisitionCode, LoanRequisition},
                    {LoanReturnCode, LoanReturn},
                    {DisposalRequisitionCode, DisposalRequisition},
                    {CreditRequisitionCode, CreditRequisition},
                    {CashSaleCode, CashSale},
                    {CreditSaleCode, CreditSale},
                    {RotationCreditRequisitionCode, RotationCreditRequisition},
                    {RotationRecallCode, RotationRecall},
                    {RotationIssueCode, RotationIssue},
                    {ShortFormCode, ShortForm},
                    {ManualRequisitionCode, ManualRequisition},
                    {ReferredDemandCode, ReferredDemand}
                };

                return statusDictionary;
            }

            public static string GetTypeCode(string statusName)
            {
                var statusDictionary = GetRequisitionTypeList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetTypeName(string statusCode)
            {
                var statusDictionary = GetRequisitionTypeList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
        public static class RequisitionStatus
        {
            public static string PendingCode = "P";
            public static string Pending = "PENDING";
            public static string NotPrintedCode = "0";
            public static string NotPrinted = "NOT PRINTED";
            public static string PrintRequestedCode = "1";
            public static string PrintRequested= "PRINT REQUESTED";
            public static string PartiallyAcquittedCode = "2";
            public static string PartiallyAcquitted = "PARTIALLY ACQUITTED";
            public static string IdrCompletedCode = "3";
            public static string IdrCompleted = "IDR COMPLETED";
            public static string CompleteCode = "9";
            public static string Complete = "COMPLETE";

            public static Dictionary<string, string> GetRequisitionStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {PendingCode, Pending},
                    {NotPrintedCode, NotPrinted},
                    {PrintRequestedCode, PrintRequested},
                    {PartiallyAcquittedCode, PartiallyAcquitted},
                    {IdrCompletedCode, IdrCompleted},
                    {CompleteCode, Complete}
                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetRequisitionStatusList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }
 
            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetRequisitionStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }

        public class ItemStatus
        {
            public static string OutSideLeadTimeCode = "0";
            public static string OutsideLeadTime = "OUTSIDE LEAD TIME";
            public static string InsideLeadTimeCode = "1";
            public static string InsideLeadTime = "INSIDE LEAD TIME";
            public static string WarehouseServicingCode = "2";
            public static string WarehouseServicing = "WAREHOUSE SERVICING";
            public static string StockUnavailableCode = "3";
            public static string StockUnavailable = "STOCK UNAVAILABLE";
            public static string CreditIssueRequisitionCode = "4";
            public static string CreditIssueRequisition = "CREDIT ISSUE REQUISITION";
            public static string UnauthorizedCode = "5";
            public static string Unauthorized = "UNAUTHORIZED";
            public static string IssuedNotCompleteCode = "6";
            public static string IssuedNotComplete = "ISSUED NOT COMPLETE";
            public static string CompleteCode = "9";
            public static string Complete = "COMPLETE";

            public static Dictionary<string, string> GetItemStatusList()
            {
                var statusDictionary = new Dictionary<string, string>
                {
                    {OutSideLeadTimeCode, OutsideLeadTime},
                    {InsideLeadTimeCode, InsideLeadTime},
                    {WarehouseServicingCode, WarehouseServicing},
                    {StockUnavailableCode, StockUnavailable},
                    {CreditIssueRequisitionCode, CreditIssueRequisition},
                    {UnauthorizedCode, Unauthorized},
                    {IssuedNotCompleteCode, IssuedNotComplete},
                    {CompleteCode, Complete}

                };

                return statusDictionary;
            }

            public static string GetStatusCode(string statusName)
            {
                var statusDictionary = GetItemStatusList();
                return statusDictionary.ContainsValue(statusName) ? statusDictionary.FirstOrDefault(x => x.Value == statusName).Key : null;
            }

            public static string GetStatusName(string statusCode)
            {
                var statusDictionary = GetItemStatusList();
                return statusDictionary.ContainsKey(statusCode) ? statusDictionary[statusCode] : null;
            }
        }
    }
}

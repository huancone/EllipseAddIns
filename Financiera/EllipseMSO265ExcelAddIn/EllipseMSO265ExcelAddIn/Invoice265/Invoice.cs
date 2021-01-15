using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
    public class Invoice
    {
        public string District;
        public string SupplierNo;
        public string SupplierMnemonic;
        public string GovernmentId;
        public string InvoiceNo;
        public string InvoiceAmount;
        public decimal TaxAmount;

        public string Accountant;
        public string OriginalInvoiceNo;
        public string Currency;
        public string HandlingCode;
        public string ControlAccountGroupCode;

        public string InvoiceDate;
        public string InvoiceReceivedDate;
        public string DueDate;

        public string SettlementDiscount;
        public string DiscountDate;

        public string BankBranchCode;
        public string BankAccountNo;

        public string Ref;

        public bool Equals(Invoice invoice)
        {
            if (District != invoice.District) return false;
            if (SupplierNo != invoice.SupplierNo) return false;

            if (SupplierMnemonic != invoice.SupplierMnemonic) return false;
            if (GovernmentId != invoice.GovernmentId) return false;
            if (InvoiceNo != invoice.InvoiceNo) return false;
            if (InvoiceAmount != invoice.InvoiceAmount) return false;
            //if (TaxAmount != invoice.TaxAmount) return false;

            if (Accountant != invoice.Accountant) return false;
            if (OriginalInvoiceNo != invoice.OriginalInvoiceNo) return false;
            if (Currency != invoice.Currency) return false;
            if (HandlingCode != invoice.HandlingCode) return false;
            if (ControlAccountGroupCode != invoice.ControlAccountGroupCode) return false;

            if (InvoiceDate != invoice.InvoiceDate) return false;
            if (InvoiceReceivedDate != invoice.InvoiceReceivedDate) return false;
            if (DueDate != invoice.DueDate) return false;

            if (SettlementDiscount != invoice.SettlementDiscount) return false;
            if (DiscountDate != invoice.DiscountDate) return false;

            if (BankBranchCode != invoice.BankBranchCode) return false;
            if (BankAccountNo != invoice.BankAccountNo) return false;

            return true;
        }
    }

}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
    public class NominaParameters
    {
        [CsvColumn(FieldIndex = 1)] public string BranchCode;

        [CsvColumn(FieldIndex = 2)] public string BankAccount;

        [CsvColumn(FieldIndex = 3)] public string Accountant;

        [CsvColumn(FieldIndex = 4)] public string SupplierNo;

        [CsvColumn(FieldIndex = 5)] public string SupplierMnemonic;

        [CsvColumn(FieldIndex = 6)] public string Currency;

        [CsvColumn(FieldIndex = 7)] public string InvoiceNo;

        [CsvColumn(FieldIndex = 8)] public string InvoiceDate;

        [CsvColumn(FieldIndex = 9)] public string DueDate;

        [CsvColumn(FieldIndex = 10)] public string InvoiceAmount;

        [CsvColumn(FieldIndex = 11)] public string Description;

        [CsvColumn(FieldIndex = 12)] public string Ref;

        [CsvColumn(FieldIndex = 13)] public string ItemValue;

        [CsvColumn(FieldIndex = 14)] public string Account;

        [CsvColumn(FieldIndex = 15)] public string AuthorizedBy;

        [CsvColumn(FieldIndex = 16)] public string Value01;

        [CsvColumn(FieldIndex = 17)] public string Value02;

        [CsvColumn(FieldIndex = 18)] public string Value03;

        [CsvColumn(FieldIndex = 19)] public string Value04;

        [CsvColumn(FieldIndex = 20)] public string Value05;

        [CsvColumn(FieldIndex = 21)] public string Value06;

        [CsvColumn(FieldIndex = 22)] public string Value07;

        [CsvColumn(FieldIndex = 23)] public string Value08;

        [CsvColumn(FieldIndex = 24)] public string Value09;

        [CsvColumn(FieldIndex = 25)] public string Value10;
    }

}

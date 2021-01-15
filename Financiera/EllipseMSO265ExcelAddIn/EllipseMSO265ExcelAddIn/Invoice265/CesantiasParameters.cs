using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
    public class CesantiasParameters
    {
        [CsvColumn(FieldIndex = 1)] public string SupplierMnemonic;

        [CsvColumn(FieldIndex = 2)] public string SupplierName;

        [CsvColumn(FieldIndex = 3)] public string Reference;

        [CsvColumn(FieldIndex = 4)] public string Description;

        [CsvColumn(FieldIndex = 5)] public string InvoiceDate;

        [CsvColumn(FieldIndex = 6)] public string DueDate;

        [CsvColumn(FieldIndex = 7)] public string Account;

        [CsvColumn(FieldIndex = 8)] public string Currency;

        [CsvColumn(FieldIndex = 9)] public string ItemValue;

        [CsvColumn(FieldIndex = 10)] public string InvoiceAmount;

        [CsvColumn(FieldIndex = 11)] public string AuthorizedBy;

        [CsvColumn(FieldIndex = 12)] public string BranchCode;

        [CsvColumn(FieldIndex = 13)] public string BankAccount;
    }

}

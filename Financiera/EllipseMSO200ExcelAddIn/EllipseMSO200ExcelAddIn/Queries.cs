using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO200ExcelAddIn
{
    public static class Queries
    {
        public static string GetSupplierInvoiceInfo(string districtCode, string cedula, string dbReference, string dbLink)
        {
            var sqlQuery = "SELECT " +
                           "  TRIM(BI.BANK_ACCT_NO) BANK_ACCT_NO, " +
                           "  TRIM(BI.TAX_FILE_NO) TAX_FILE_NO, " +
                           "  TRIM(BI.SUPPLIER_NO) SUPPLIER_NO, " +
                           "  COUNT(BI.SUPPLIER_NO) OVER(PARTITION BY BI.TAX_FILE_NO) CANTIDAD_REGISTROS " +
                           "FROM " +
                           "  " + dbReference + ".MSF203" + dbLink + " BI " +
                           "WHERE " +
                           "  BI.TAX_FILE_NO = '" + cedula + "' " +
                           "  AND BI.DSTRCT_CODE = '" + districtCode + "' ";
            return sqlQuery;
        }
    }
}

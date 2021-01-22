using System;
using System.Collections.Generic;
using System.Linq;
using SharedClassLibrary.Utilities;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
    internal static class Queries
    {
        public static string GetSupplierInvoiceInfoQuery(string districtCode, string supplierNo, string supplierTaxFileNo, string dbReference, string dbLink)
        {
            var paramDistrict = districtCode;
            if (string.IsNullOrWhiteSpace(paramDistrict))
                paramDistrict = "ICOR";

            var paramSupplierNo = supplierNo;
            if (!string.IsNullOrWhiteSpace(paramSupplierNo))
                paramSupplierNo = " AND SUP.SUPPLIER_NO = '" + paramSupplierNo + "'";

            var paramSupplierTaxFileNo = supplierTaxFileNo;
            if (!string.IsNullOrWhiteSpace(paramSupplierTaxFileNo))
                paramSupplierTaxFileNo = " AND BNK.TAX_FILE_NO = '" + paramSupplierTaxFileNo + "'";

            var sqlQuery = "SELECT " +
                           "   BNK.DSTRCT_CODE," +
                           "   TRIM(SUP.SUPPLIER_NO) SUPPLIER_NO," +
                           "   TRIM(BNK.TAX_FILE_NO) TAX_FILE_NO," +
                           "   TRIM(SUP.SUP_STATUS) ST_ADRESS," +
                           "   TRIM(BNK.SUP_STATUS) ST_BUSINESS," +
                           "   TRIM(SUP.SUPPLIER_NAME) SUPPLIER_NAME," +
                           "   TRIM(SUP.CURRENCY_TYPE) CURRENCY_TYPE," +
                           "   TRIM(BNK.BANK_ACCT_NAME) BANK_ACCT_NAME," +
                           "   TRIM(BNK.BANK_ACCT_NO) BANK_ACCT_NO," +
                           "   BNK.SUP_STATUS," +
                           "   TRIM(BNK.DEF_BRANCH_CODE) DEF_BRANCH_CODE," +
                           "   TRIM(BNK.DEF_BANK_ACCT_NO) DEF_BANK_ACCT_NO," +
                           "   COUNT(BNK.SUPPLIER_NO) OVER(PARTITION BY BNK.TAX_FILE_NO) CANTIDAD_REGISTROS" +
                           " FROM ELLIPSE.MSF200 SUP" +
                           " INNER JOIN ELLIPSE.MSF203 BNK" +
                           " ON SUP.SUPPLIER_NO  = BNK.SUPPLIER_NO" +
                           " WHERE" +
                           " BNK.DSTRCT_CODE = '" + paramDistrict + "'" +
                           paramSupplierNo +
                           paramSupplierTaxFileNo +
                           " AND BNK.SUP_STATUS <> 9";

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");
            return sqlQuery;
        }

        public static string GetTaxCodeListQuery(List<string> taxCodesParamList, string taxGroupCode)
        {
            var paramTaxes = "";
            if (taxCodesParamList != null && taxCodesParamList.Any())
                paramTaxes = " AND TXC.ATAX_CODE IN (" + MyUtilities.GetListInSeparator(taxCodesParamList, ",", "'") + ")";


            const string paramGroupIndicator = " AND (TRIM(GRP_LEVEL_IND) IS NULL OR TRIM(GRP_LEVEL_IND) = 'N')";

            var conditionalGroup = "";
            var paramGroupCode = "";
            if (!string.IsNullOrWhiteSpace(taxGroupCode))
            {
                conditionalGroup = " JOIN ELLIPSE.MSF014 TXG ON TXG.REL_ATAX_CODE = TXC.ATAX_CODE";
                paramGroupCode = " AND TXG.ATAX_CODE = '" + taxGroupCode + "'";
            }

            var sqlQuery = "SELECT TC.TABLE_CODE, TC.TABLE_DESC, TXC.ATAX_CODE, TXC.DESCRIPTION, TXC.TAX_REF, TXC.ATAX_RATE_9, TXC.DEFAULTED_IND, TXC.DEDUCT_SW" +
                           " FROM ELLIPSE.MSF010 TC JOIN ELLIPSE.MSF013 TXC ON TC.TABLE_CODE = TXC.ATAX_CODE" + conditionalGroup +
                           " WHERE TC.TABLE_TYPE = '+ADD' " +
                           paramGroupIndicator +
                           paramTaxes +
                           paramGroupCode;

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }

        public static string GetTaxGroupCodeListQuery(List<string> taxGroupCodeParamList)
        {
            var paramTaxes = "";
            if (taxGroupCodeParamList != null && taxGroupCodeParamList.Count > 0)
                paramTaxes = " AND TXC.ATAX_CODE IN (" + MyUtilities.GetListInSeparator(taxGroupCodeParamList, ",", "'") + ")";

            var paramGroupIndicator = " AND TRIM(GRP_LEVEL_IND) = 'Y'";

            var sqlQuery = "SELECT TXC.ATAX_CODE, TXC.DESCRIPTION, TXC.TAX_REF, TXC.ATAX_RATE_9, TXC.DEFAULTED_IND, TXC.DEDUCT_SW" +
                           " FROM ELLIPSE.MSF013 TXC " +
                           " WHERE " +
                           paramGroupIndicator +
                           paramTaxes;

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }
    }

}

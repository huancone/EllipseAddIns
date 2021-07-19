
namespace EllipseSAO900AddIn
{
    /// <summary>
    ///     Consultas SQL a las bases de datos de ELLIPSE
    /// </summary>
    internal static class Queries
    {
        public static string GetEmployeeName(string employeeId, string dbReference, string dbLink)
        {
            var sqlQuery = " " +
                           "SELECT DISTINCT " +
                           "  EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE " +
                           "FROM " +
                           "  " + dbReference + ".MSF870" + dbLink + " POS " +
                           "INNER JOIN " + dbReference + ".MSF878" + dbLink + " EMPOS " +
                           "ON" +
                           "  EMPOS.POSITION_ID = POS.POSITION_ID " +
                           "AND " +
                           "  (" +
                           "    EMPOS.POS_STOP_DATE > TO_CHAR ( SYSDATE, 'YYYYMMDD' ) " +
                           "  OR EMPOS.POS_STOP_DATE = '00000000' " +
                           "  ) " +
                           "INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP " +
                           "ON " +
                           "  EMPOS.EMPLOYEE_ID = EMP.EMPLOYEE_ID " +
                           "WHERE " +
                           "EMPOS.EMPLOYEE_ID = '" + employeeId + "' ";
            return sqlQuery;
        }

        public static string GetTransactionInfo(string districtCode, string numTransaction, string dbReference,
            string dbLink)
        {
            var processDate = numTransaction.Substring(0, 8);
            var transNo = numTransaction.Substring(8, 11);
            var userNo = numTransaction.Substring(19, 4);
            var recType = numTransaction.Substring(23, 1);
            var sqlQuery = " " +
                           "SELECT " +
                           "  TR.FULL_PERIOD, " +
                           "  TR.ACCOUNT_CODE, " +
                           "  DECODE ( TRIM ( TR.PROJECT_NO ), NULL, DECODE ( TRIM ( TR.WORK_ORDER ), NULL, ' ', TR.WORK_ORDER ), TR.PROJECT_NO ) PROJECT_NO, " +
                           "  DECODE ( TRIM ( TR.PROJECT_NO ), NULL, DECODE ( TRIM ( TR.WORK_ORDER ), NULL, ' ', 'W' ), 'P' ) IND," +
                           "  TR.TRAN_AMOUNT, " +
                           "  TR.TRAN_AMOUNT_S " +
                           "FROM " +
                           "  " + dbReference + ".MSF900" + dbLink + " TR " +
                           "WHERE " +
                           "  TR.DSTRCT_CODE = '" + districtCode + "' " +
                           "AND TR.PROCESS_DATE = '" + processDate + "' " +
                           "AND TR.TRANS_NO = '" + transNo + "' " +
                           "AND TR.USERNO = '" + userNo + "' " +
                           "AND TR.REC900_TYPE = '" + recType + "' ";

            return sqlQuery;
        }

        public static string GetAccountCodeInfo(string districtCode, string accountCode, string dbReference,
            string dbLink)
        {
            var sqlQuery = " " +
                           "SELECT " +
                           "  CC.ACTIVE_STATUS, " +
                           "  CC.ACCOUNT_CODE, " +
                           "  CC.PROJ_ENTRY_IND, " +
                           "  CC.WO_ENTRY_IND, " +
                           "  CC.SUBLEDGER_IND " +
                           "FROM " +
                           "  " + dbReference + ".MSF966" + dbLink + " CC " +
                           "WHERE " +
                           "  CC.DSTRCT_CODE = '" + districtCode + "' " +
                           "AND CC.ACCOUNT_CODE = '" + accountCode + "'";
            return sqlQuery;
        }

        public static string GetSupplierName(string districtCode, string supplierId, string dbReference,
            string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT " +
                           "  SUP.SUPPLIER_NO, " +
                           "  SUP.SUPPLIER_NAME " +
                           "FROM " +
                           "  " + dbReference + ".MSF200 SUP" + dbLink + " " +
                           "INNER JOIN " + dbReference + ".MSF203" + dbLink + " SD " +
                           "ON " +
                           "  SD.SUPPLIER_NO = SUP.SUPPLIER_NO " +
                           "WHERE " +
                           "  SUP.SUPPLIER_NO = '" + supplierId + "' " +
                           "  AND SD.DSTRCT_CODE = '" + districtCode + "'";
            return sqlQuery;
        }

        public static string GetContractNameDesc(string document, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT " +
                           "  CONTRACT_DESC " +
                           "FROM " +
                           "  " + dbReference + ".MSF384" + dbLink + " " +
                           "WHERE " +
                           "  CONTRACT_NO = '" + document + "'";
            return sqlQuery;
        }

        public static string GetPurchaseOrder(string document, string supplierNo, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT " +
                           "  PO_NO " +
                           "FROM " +
                           "  " + dbReference + ".MSF220" + dbLink + " " +
                           "WHERE " +
                           "  PO_NO = '" + document + "' " +
                           "AND SUPPLIER_NO = '" + supplierNo + "'";
            return sqlQuery;
        }
    }


}

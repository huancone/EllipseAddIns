using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO905ExcelAddIn
{
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

        public static string GetTransactionInfo(string districtCode, string numTransaction, string dbReference, string dbLink)
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

        public static string GetAccountCodeInfo(string districtCode, string accountCode, string dbReference, string dbLink)
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

        public static string GetSupplierName(string districtCode, string supplierId, string dbReference, string dbLink)
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

        public static string GetSupplierMnemonic(string mnemonic, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT " +
                           "  GL_COLLOQ_CD, " +
                           "  COUNT ( * ) OVER ( ) CANTIDAD " +
                           "FROM " +
                           "  " + dbReference + ".MSF922" + dbLink + " " +
                           "WHERE " +
                           "  GL_COLLOQ_TY = '7' " +
                           "AND COLLOQ_NAME = '" + mnemonic + "'";
            return sqlQuery;
        }

        public static string GetJournalInfo(string districtCode, string journal, string dbReference, string dbLink)
        {
            var sqlQuery = " " +
                "SELECT " +
                "  TR.DSTRCT_CODE, " +
                "  TR.TRAN_GROUP_KEY, " +
                "  ( TR.PROCESS_DATE || TR.TRANSACTION_NO || TR.USERNO || TR.REC900_TYPE ) NUMTXT, " +
                "  TR.FULL_PERIOD, " +
                "  TR.REC900_TYPE, " +
                "  TR.TRAN_TYPE, " +
                "  TR.POSTED_STATUS, " +
                "  TR.MANJNL_VCHR, " +
                "  TR.ACCOUNTANT, " +
                "  TR.ACCOUNT_CODE, " +
                "  TR.PROJECT_NO, " +
                "  TR.TRAN_AMOUNT, " +
                "  TR.TRAN_AMOUNT_S, " +
                "  TR.CURRENCY_TYPE, " +
                "  TR.CREATION_DATE, " +
                "  TR.CREATION_TIME, " +
                "  TR.CREATION_USER, " +
                "  TR.MIMS_SL_KEY, " +
                "  TR.JOURNAL_DESC, " +
                "  TR.DOCUMENT_REF, " +
                "  TR.AUTO_JNL_FLG, " +
                "  TR.JOURNAL_TYPE " +
                "FROM " +
                "  " + dbReference + ".MSF900" + dbLink + " TR " +
                "INNER JOIN " + dbReference + ".MSFX90 X90" + dbLink + " " +
                "ON " +
                "  TR.TRAN_GROUP_KEY = X90.DSTRCT_CODE || X90.PROCESS_DATE || X90.TRANSACTION_NO || X90.USERNO || X90.REC900_TYPE " +
                "WHERE " +
                "  X90.DSTRCT_CODE = '" + districtCode + "' " +
                "AND X90.JOURNAL_NO = '" + journal + "' " +
                "ORDER BY " +
                "  1, " +
                "  2, " +
                "  3 ";

            return sqlQuery;
        }

        public static string GetSupplierNit(string nit, string dbReference, string dbLink)
        {
            var sqlQuery = " " +
                           "SELECT " +
                           " DISTINCT NIT " +
                           "FROM " +
                           "  ( " +
                           "    SELECT " +
                           "      trim(TAX_FILE_NO) NIT " +
                           "    FROM " +
                           "      " + dbReference + ".MSF203" + dbLink + " " +
                           "    WHERE " +
                           "      TRIM (TAX_FILE_NO) = '" + nit + "' " +
                           "    UNION " +
                           "    SELECT " +
                           "      TRIM(GOVT_ID_NO) NIT " +
                           "    FROM " +
                           "      " + dbReference + ".MSF503" + dbLink + " " +
                           "    WHERE " +
                           "      TRIM (GOVT_ID_NO) = '" + nit + "' " +
                           "    UNION " +
                           "    SELECT " +
                           "      TRIM(TABLE_CODE) NIT " +
                           "    FROM " +
                           "      " + dbReference + ".MSF010" + dbLink + " " +
                           "    WHERE " +
                           "      TABLE_TYPE          = '+NIT' " +
                           "    AND TRIM (TABLE_CODE) = '" + nit + "' " +
                           "  )  ";

            return sqlQuery;
        }
    }

}

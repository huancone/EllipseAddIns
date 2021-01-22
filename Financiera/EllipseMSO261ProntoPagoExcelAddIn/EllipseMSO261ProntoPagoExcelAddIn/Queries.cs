using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO261ProntoPagoExcelAddIn
{
    internal static class Queries
    {
        public static string GetSupplierInvoiceInfo(string districtCode, string supplierNo, string invoiceNo, string dbReference, string dbLink)
        {
            var sqlQuery = "SELECT " +
                           "  INV.SUPPLIER_NO, " +
                           "  INV.EXT_INV_NO, " +
                           "  INV.SD_DATE FEC_MOD_PAGO, " +
                           "  INV.DUE_DATE, " +
                           "  INV.PMT_STATUS, " +
                           "  SUP.SUPPLIER_NAME NOM_SUPPLIER, " +
                           "  TRIM ( INV.BRANCH_CODE ) || '-' || TRIM ( INV.BANK_ACCT_NO ) ORIG_BANK, " +
                           "  SBI.NO_OF_DAYS_PAY, " +
                           "  DECODE ( INV.CURRENCY_TYPE, 'USD ', DECODE ( INV.LOC_INV_AMD, '0', INV.LOC_INV_ORIG, INV.LOC_INV_AMD ), DECODE ( INV.FOR_INV_AMD, '0', INV.FOR_INV_ORIG, INV.FOR_INV_AMD ) ) VLR_FACTURA, " +
                           "  DECODE ( SUBSTR ( INVOICE_LINE_ITEM.INV_ITEM_DESC, 1, 4 ), 'CNT:', INVOICE_LINE_ITEM.FOR_VAL_INVD, DECODE ( INV.CURRENCY_TYPE, 'USD ', DECODE ( INV.LOC_INV_AMD, '0', INV.LOC_INV_ORIG, INV.LOC_INV_AMD ), DECODE ( INV.FOR_INV_AMD, '0', INV.FOR_INV_ORIG, INV.FOR_INV_AMD ) ) - DECODE ( INV.CURRENCY_TYPE, 'PES ', NVL ( 0, INV.ATAX_AMT_FOR ), 0 ) ) VRBASE, " +
                           "  SBI.NO_OF_DAYS_PAY - ( TO_DATE ( DECODE ( TRIM ( INV.SD_DATE ), NULL, INV.INV_RCPT_DATE, INV.SD_DATE ), 'YYYYMMDD' ) - TO_DATE ( INV.INV_RCPT_DATE, 'YYYYMMDD' ) ) DIREFENCIA, " +
                           "  INV.SD_AMOUNT VR_OTROS_DESCTS " +
                           "FROM " +
                           "  ELLIPSE.MSF260 INV " +
                           "LEFT JOIN ELLIPSE.MSF203 SBI " +
                           "ON " +
                           "  INV.DSTRCT_CODE = SBI.DSTRCT_CODE " +
                           "AND INV.SUPPLIER_NO = SBI.SUPPLIER_NO " +
                           "LEFT JOIN ELLIPSE.MSF200 SUP " +
                           "ON " +
                           "  SBI.SUPPLIER_NO = SUP.SUPPLIER_NO " +
                           "INNER JOIN ELLIPSE.MSF26A INVOICE_LINE_ITEM " +
                           "ON " +
                           "  INVOICE_LINE_ITEM.SUPPLIER_NO = INV.SUPPLIER_NO " +
                           "AND INVOICE_LINE_ITEM.INV_NO = INV.INV_NO " +
                           "AND INVOICE_LINE_ITEM.DSTRCT_CODE = INV.DSTRCT_CODE " +
                           "AND INVOICE_LINE_ITEM.INV_ITEM_NO = '001' " +
                            "WHERE " +
                            "   INV.DSTRCT_CODE = '" + districtCode + "' " +
                            "AND INV.SUPPLIER_NO = '" + supplierNo + "' " +
                            "AND INV.INV_NO = '" + invoiceNo + "' ";

            return sqlQuery;
        }
    }

}

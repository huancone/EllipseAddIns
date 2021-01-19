using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Ellipse;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
    public class Supplier
    {
        public Supplier()
        {
        }

        public Supplier(EllipseFunctions eFunctions, string supplierNo, string supplierTaxFileNo, string districtCode = "ICOR")
        {
            var sqlQuery = Queries.GetSupplierInvoiceInfoQuery(districtCode, supplierNo, supplierTaxFileNo, eFunctions.DbReference, eFunctions.DbLink);

            var drSupplierInfo = eFunctions.GetQueryResult(sqlQuery);

            if (!drSupplierInfo.Read())
                throw new Exception("No se han econtrado datos para el Supplier Ingresado");

            var cant = Convert.ToInt16(drSupplierInfo["CANTIDAD_REGISTROS"].ToString());
            if (cant > 1)
                throw new Exception("Se ha encontrado más de un registro activo para el Supplier Ingresado");

            SupplierNo = drSupplierInfo["SUPPLIER_NO"].ToString();
            TaxFileNo = drSupplierInfo["TAX_FILE_NO"].ToString();
            StAdress = drSupplierInfo["ST_ADRESS"].ToString();
            StBusiness = drSupplierInfo["ST_BUSINESS"].ToString();
            SupplierName = drSupplierInfo["SUPPLIER_NAME"].ToString();
            CurrencyType = drSupplierInfo["CURRENCY_TYPE"].ToString();
            AccountName = drSupplierInfo["BANK_ACCT_NAME"].ToString();
            AccountNo = drSupplierInfo["BANK_ACCT_NO"].ToString();
            BankBranchCode = drSupplierInfo["DEF_BRANCH_CODE"].ToString();
            BankBranchAccountNo = drSupplierInfo["DEF_BANK_ACCT_NO"].ToString();
            Status = drSupplierInfo["SUP_STATUS"].ToString();
        }

        public string SupplierNo;
        public string TaxFileNo;
        public string StAdress;
        public string StBusiness;
        public string SupplierName;
        public string CurrencyType;
        public string AccountName;
        public string AccountNo;
        public string Status;
        public string BankBranchCode;
        public string BankBranchAccountNo;
    }

}

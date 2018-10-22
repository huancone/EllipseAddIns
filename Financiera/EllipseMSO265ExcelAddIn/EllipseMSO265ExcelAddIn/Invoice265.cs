using System;
using System.Collections.Generic;
using System.Linq;
using LINQtoCSV;
using System.Web.Services.Ellipse.Post;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;
using Util = System.Web.Services.Ellipse.Post.Util;

namespace EllipseMSO265ExcelAddIn
{
    public static class Invoice265
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

        public class InvoiceItem
        {
            public string Description;
            public decimal ItemValue;
            public decimal TaxValue;
            public string Account;
            public string AuthorizedBy;
            public string WorkOrderProjectNo;
            public string WorkOrderProjectIndicator;
            public string ItemDistrict;
            public string EquipNo;
            public List<TaxCodeItem> TaxList;
            public decimal FirstTaxAdjustment;//Usado para calcular las diferencias manuales (override) sobre el primer tax del item
        }
        public class TaxCodeItem
        {
            public string TaxCode;
            public string TaxDescription;
            public string TaxReference;
            public decimal TaxRatePerc;
            public string DefaultToInvoiceItem;
            public string Deduct;
        }

        public static List<TaxCodeItem> GetTaxCodeList(EllipseFunctions eFunctions)
        {
            return GetTaxCodeList(eFunctions, null, null);
        }

        public static List<TaxCodeItem> GetTaxCodeList(EllipseFunctions eFunctions, List<string> taxCodeParamList)
        {
            if (taxCodeParamList == null || !taxCodeParamList.Any())
                return new List<TaxCodeItem>();
            return GetTaxCodeList(eFunctions, taxCodeParamList, null);
        }

        public static List<TaxCodeItem> GetTaxCodeList(EllipseFunctions eFunctions, string taxGroupCode)
        {
            return string.IsNullOrWhiteSpace(taxGroupCode) ? null : GetTaxCodeList(eFunctions, null, taxGroupCode);
        }

        public static List<TaxCodeItem> GetTaxGroupCodeList(EllipseFunctions eFunctions, List<string> taxGroupCodeParamList = null)
        {
            var taxList = new List<TaxCodeItem>();

            
            var dataReader = eFunctions.GetQueryResult(Queries.GetTaxGroupCodeListQuery(taxGroupCodeParamList));

            if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
            {
                eFunctions.CloseConnection();
                return taxList;
            }

            while (dataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage);//!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            eFunctions.CloseConnection();
            return taxList;
        }
        
        private static List<TaxCodeItem> GetTaxCodeList(EllipseFunctions eFunctions, List<string> taxCodesParamList, string taxGroupCode)
        {
            var taxList = new List<TaxCodeItem>();

            var dataReader = eFunctions.GetQueryResult(Queries.GetTaxCodeListQuery(taxCodesParamList, taxGroupCode));

            if (dataReader == null || dataReader.IsClosed || !dataReader.HasRows)
            {
                eFunctions.CloseConnection();
                return taxList;
            }

            while (dataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage);//!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            eFunctions.CloseConnection();
            return taxList;
        }
        public class Supplier
        {
            public Supplier()
            {
            }

            public Supplier(EllipseFunctions eFunctions, string supplierNo, string supplierTaxFileNo, string districtCode = "ICOR")
            {
                var sqlQuery = Queries.GetSupplierInvoiceInfoQuery(districtCode, supplierNo, supplierTaxFileNo, eFunctions.dbReference, eFunctions.dbLink);

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

        public class CesantiasParameters
        {
            [CsvColumn(FieldIndex = 1)]
            public string SupplierMnemonic;

            [CsvColumn(FieldIndex = 2)]
            public string SupplierName;

            [CsvColumn(FieldIndex = 3)]
            public string Reference;

            [CsvColumn(FieldIndex = 4)]
            public string Description;

            [CsvColumn(FieldIndex = 5)]
            public string InvoiceDate;

            [CsvColumn(FieldIndex = 6)]
            public string DueDate;

            [CsvColumn(FieldIndex = 7)]
            public string Account;

            [CsvColumn(FieldIndex = 8)]
            public string Currency;

            [CsvColumn(FieldIndex = 9)]
            public string ItemValue;

            [CsvColumn(FieldIndex = 10)]
            public string InvoiceAmount;

            [CsvColumn(FieldIndex = 11)]
            public string AuthorizedBy;

            [CsvColumn(FieldIndex = 12)]
            public string BranchCode;

            [CsvColumn(FieldIndex = 13)]
            public string BankAccount;
        }

        public class NominaParameters
        {
            [CsvColumn(FieldIndex = 1)]
            public string BranchCode;

            [CsvColumn(FieldIndex = 2)]
            public string BankAccount;

            [CsvColumn(FieldIndex = 3)]
            public string Accountant;

            [CsvColumn(FieldIndex = 4)]
            public string SupplierNo;

            [CsvColumn(FieldIndex = 5)]
            public string SupplierMnemonic;

            [CsvColumn(FieldIndex = 6)]
            public string Currency;

            [CsvColumn(FieldIndex = 7)]
            public string InvoiceNo;

            [CsvColumn(FieldIndex = 8)]
            public string InvoiceDate;

            [CsvColumn(FieldIndex = 9)]
            public string DueDate;

            [CsvColumn(FieldIndex = 10)]
            public string InvoiceAmount;

            [CsvColumn(FieldIndex = 11)]
            public string Description;

            [CsvColumn(FieldIndex = 12)]
            public string Ref;

            [CsvColumn(FieldIndex = 13)]
            public string ItemValue;

            [CsvColumn(FieldIndex = 14)]
            public string Account;

            [CsvColumn(FieldIndex = 15)]
            public string AuthorizedBy;

            [CsvColumn(FieldIndex = 16)]
            public string Value01;

            [CsvColumn(FieldIndex = 17)]
            public string Value02;

            [CsvColumn(FieldIndex = 18)]
            public string Value03;

            [CsvColumn(FieldIndex = 19)]
            public string Value04;

            [CsvColumn(FieldIndex = 20)]
            public string Value05;

            [CsvColumn(FieldIndex = 21)]
            public string Value06;

            [CsvColumn(FieldIndex = 22)]
            public string Value07;

            [CsvColumn(FieldIndex = 23)]
            public string Value08;

            [CsvColumn(FieldIndex = 24)]
            public string Value09;

            [CsvColumn(FieldIndex = 25)]
            public string Value10;
        }

        public static class InvoiceActions
        {
            public static decimal GetItemTaxAdjustment(decimal itemValue, decimal calculatedTaxValue, decimal manualTaxValue, List<TaxCodeItem> taxList )
            {
                if (taxList == null || taxList.Count <= 0)
                    return 0;
                decimal taxDifference = (int)Math.Abs(calculatedTaxValue - manualTaxValue);
                var taxAdjustment = 0;
                if (taxDifference > 0)
                {
                    decimal taxValueItem = itemValue * (taxList[0].TaxRatePerc / 100);
                    if (MyUtilities.IsTrue(taxList[0].Deduct))
                        taxValueItem = taxValueItem * -1;
                    if (calculatedTaxValue > manualTaxValue)
                        taxAdjustment = (int)Math.Round(taxValueItem - taxDifference, 0);
                    else
                        taxAdjustment = (int)Math.Round(taxValueItem + taxDifference, 0);
                }

                return taxAdjustment;
            }
            public static List<TaxCodeItem> GetItemTaxList(EllipseFunctions eFunctions, string taxGroupCode, string additionalTaxCodes)
            {
                var listTaxes = new List<string>();
                //group list
                var groupTaxCodeList = GetTaxCodeList(eFunctions, taxGroupCode);
                //additional codes list
                if (!string.IsNullOrWhiteSpace(additionalTaxCodes) && additionalTaxCodes.Contains(";"))
                {
                    var splitArray = additionalTaxCodes.Split(';');
                    foreach (var item in splitArray)
                        listTaxes.Add(item);
                }
                else if (!string.IsNullOrWhiteSpace(additionalTaxCodes))
                {
                    additionalTaxCodes = MyUtilities.GetCodeKey(additionalTaxCodes);
                    listTaxes.Add(additionalTaxCodes);
                }
                
                if (listTaxes.Count != listTaxes.Distinct().Count())
                    throw new Exception("Impuesto Duplicado");

                if (groupTaxCodeList != null && groupTaxCodeList.Count > 0)
                    foreach (var taxItem in groupTaxCodeList)
                        listTaxes.Add(taxItem.TaxCode);


                return GetTaxCodeList(eFunctions, listTaxes);
            }
            public static decimal GetCalculatedItemTaxValue(decimal itemValue, List<TaxCodeItem> taxList)
            {
                decimal calculatedTaxValue = 0;
                foreach (var tax in taxList)
                {

                    decimal taxValueItem = itemValue * (tax.TaxRatePerc / 100);
                    if (MyUtilities.IsTrue(tax.Deduct))
                        taxValueItem = taxValueItem * -1;

                    calculatedTaxValue += (int)Math.Round(taxValueItem, 0);
                }

                return (int)Math.Round(calculatedTaxValue, 0);
            }
            public static ResponseDTO LoadNonInvoice(EllipseFunctions eFunctions, Invoice invoice, List<InvoiceItem> invoiceItemList)
            {
                //Default Values
                if (string.IsNullOrWhiteSpace(invoice.District))
                    invoice.District = "ICOR";
                if (string.IsNullOrWhiteSpace(invoice.Currency))
                    invoice.Currency = "PES";
                if (string.IsNullOrWhiteSpace(invoice.HandlingCode))
                    invoice.HandlingCode = "PN";
                
                if (string.IsNullOrWhiteSpace(invoice.BankBranchCode) || string.IsNullOrWhiteSpace(invoice.BankAccountNo))
                {
                    var supplier = new Supplier(eFunctions, invoice.SupplierNo, invoice.SupplierMnemonic, invoice.District);

                    invoice.BankBranchCode = supplier.BankBranchCode;
                    invoice.BankAccountNo = supplier.BankBranchAccountNo;
                    if(string.IsNullOrWhiteSpace(invoice.Currency))
                        invoice.Currency = supplier.CurrencyType;
                }

                foreach (var invoiceItem in invoiceItemList)
                {
                    if (string.IsNullOrWhiteSpace(invoiceItem.ItemDistrict))
                        invoiceItem.ItemDistrict = invoice.District;
                }
                //

                //Total Tax Amount
                //TO DO
                //

                var responseDto = eFunctions.InitiatePostConnection();

                if (responseDto.GotErrorMessages())
                    throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                //Abrimos la pantalla
                var requestXml = "<interaction>" +
                                    "   <actions>" +
                                    "       <action>" +
                                    "           <name>initialScreen</name>" +
                                    "           <data>" +
                                    "               <screenName>MSO265</screenName>" +
                                    "           </data>" +
                                    "           <id>" + Util.GetNewOperationId() + "</id>" +
                                    "           </action>" +
                                    "   </actions>" +
                                    "   <connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId>" +
                                    "   <application>ServiceInteraction</application>" +
                                    "   <applicationPage>unknown</applicationPage>" +
                                    "</interaction>";

                responseDto = eFunctions.ExecutePostRequest(requestXml);

                if (responseDto.GotErrorMessages())
                    throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));



                //Ingresamos la información principal
                if (!responseDto.ResponseString.Contains("MSM265A"))
                    throw new Exception("No se ha podido ingresar al programa MSO265");
                requestXml = "<interaction>                                                     ";
                requestXml = requestXml + "	<actions>";
                requestXml = requestXml + "		<action>";
                requestXml = requestXml + "			<name>submitScreen</name>";
                requestXml = requestXml + "			<data>";
                requestXml = requestXml + "				<inputs>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>DSTRCT_CODE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.District + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>SUPPLIER_NO1I</name>";
                requestXml = requestXml + "						<value>" + invoice.SupplierNo + "</value>";
                requestXml = requestXml + "					</screenField>";
                if (string.IsNullOrWhiteSpace(invoice.SupplierNo) && !string.IsNullOrWhiteSpace(invoice.SupplierMnemonic))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>MNEMONIC1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.SupplierMnemonic + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                if (!string.IsNullOrWhiteSpace(invoice.GovernmentId))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>GOVT_ID1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.GovernmentId + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }

                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>INV_NO1I</name>";
                requestXml = requestXml + "						<value>" + invoice.InvoiceNo + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "                 <screenField>";
                requestXml = requestXml + "                 	<name>INV_AMT1I</name>";
                requestXml = requestXml + "                 	<value>" + invoice.InvoiceAmount + "</value>";
                requestXml = requestXml + "                 </screenField>";
                if (invoice.TaxAmount != 0)
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ADD_TAX_AMOUNT1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.TaxAmount + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }

                if (!string.IsNullOrWhiteSpace(invoice.Accountant))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNTANT1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.Accountant + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                if (!string.IsNullOrWhiteSpace(invoice.OriginalInvoiceNo))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ORG_INV_NO1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.OriginalInvoiceNo + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>CURRENCY_TYPE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.Currency + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>HANDLE_CDE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.HandlingCode + "</value>";
                requestXml = requestXml + "					</screenField>";
                if (!string.IsNullOrWhiteSpace(invoice.ControlAccountGroupCode))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCT_GRP_CODE1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.ControlAccountGroupCode + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>INV_DATE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.InvoiceDate + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>INV_RCPT_DATE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.InvoiceReceivedDate + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>DUE_DATE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.DueDate + "</value>";
                requestXml = requestXml + "					</screenField>";
                if (!string.IsNullOrWhiteSpace(invoice.SettlementDiscount))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>SD_AMOUNT1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.SettlementDiscount + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                if (!string.IsNullOrWhiteSpace(invoice.DiscountDate))
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>SD_DATE1I</name>";
                    requestXml = requestXml + "						<value>" + invoice.DiscountDate + "</value>";
                    requestXml = requestXml + "					</screenField>";
                }
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>BRANCH_CODE1I</name>";
                requestXml = requestXml + "						<value>" + invoice.BankBranchCode + "</value>";
                requestXml = requestXml + "					</screenField>";
                requestXml = requestXml + "					<screenField>";
                requestXml = requestXml + "						<name>BANK_ACCT_NO1I</name>";
                requestXml = requestXml + "						<value>" + invoice.BankAccountNo + "</value>";
                requestXml = requestXml + "					</screenField>";

                var itemIndex = 1;
                foreach (var invoiceItem in invoiceItemList)
                {
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_ITEM_DESC1I" + itemIndex + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItem.Description + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I" + itemIndex + "</name>";
                    requestXml = requestXml + "					    <value>" + invoiceItem.ItemValue + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I" + itemIndex + "</name>";
                    requestXml = requestXml + "					   	<value>" + invoiceItem.ItemDistrict + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>AUTH_BY1I" + itemIndex + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItem.AuthorizedBy + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNT1I" + itemIndex + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItem.Account + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (!string.IsNullOrWhiteSpace(invoiceItem.WorkOrderProjectNo))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>WORK_ORDER1I" + itemIndex + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItem.WorkOrderProjectNo + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (!string.IsNullOrWhiteSpace(invoiceItem.WorkOrderProjectIndicator))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>WORK_PROJ_IND" + itemIndex + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItem.WorkOrderProjectIndicator + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (!string.IsNullOrWhiteSpace(invoiceItem.EquipNo))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>PLANT_NO1I" + itemIndex + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItem.EquipNo + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    if (invoiceItem.TaxList != null && invoiceItem.TaxList.Count > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACTION1I" + itemIndex + "</name>";
                        requestXml = requestXml + "						<value>T</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
                    itemIndex++;
                }
                requestXml = requestXml + "				</inputs>";
                requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                requestXml = requestXml + "			</data>";
                requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                requestXml = requestXml + "		</action>";
                requestXml = requestXml + "	</actions>                                                       ";
                requestXml = requestXml + "	<chains/>                                                        ";
                requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                requestXml = requestXml + "	<application>ServiceInteraction</application>                    ";
                requestXml = requestXml + "	<applicationPage>unknown</applicationPage>                       ";
                requestXml = requestXml + "</interaction>                                                    ";

                requestXml = requestXml.Replace("&", "&amp;");
                responseDto = eFunctions.ExecutePostRequest(requestXml);

                if (responseDto.GotErrorMessages())
                    throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                //Pantalla de información del proveedor a la que ingresa internamente por el MNEMONIC / Cedula
                if (string.IsNullOrWhiteSpace(invoice.SupplierNo) && !string.IsNullOrWhiteSpace(invoice.SupplierMnemonic))
                {
                    if (!responseDto.ResponseString.Contains("MSM202A"))
                        throw new Exception("Se ha producido un error al intentar validar la información del Supplier");
                    requestXml = "<interaction> ";
                    requestXml = requestXml + "	<actions> ";
                    requestXml = requestXml + "		<action> ";
                    requestXml = requestXml + "			<name>submitScreen</name> ";
                    requestXml = requestXml + "			<data> ";
                    requestXml = requestXml + "				<inputs> ";
                    requestXml = requestXml + "					<screenField> ";
                    requestXml = requestXml + "						<name>SUP_MNEMONIC1I</name> ";
                    requestXml = requestXml + "						<value>" + invoice.SupplierMnemonic + "</value> ";
                    requestXml = requestXml + "					</screenField> ";
                    requestXml = requestXml + "					<screenField> ";
                    requestXml = requestXml + "						<name>SUP_STATUS_IND1I</name> ";
                    requestXml = requestXml + "						<value>A</value> ";
                    requestXml = requestXml + "					</screenField> ";
                    requestXml = requestXml + "				</inputs> ";
                    requestXml = requestXml + "				<screenName>MSM202A</screenName> ";
                    requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                    requestXml = requestXml + "			</data> ";
                    requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                    requestXml = requestXml + "		</action> ";
                    requestXml = requestXml + "	</actions> ";
                    requestXml = requestXml + "	<chains/> ";
                    requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                    requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                    requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                    requestXml = requestXml + "</interaction> ";

                    requestXml = requestXml.Replace("&", "&amp;");
                    responseDto = eFunctions.ExecutePostRequest(requestXml);

                    if (responseDto.GotErrorMessages())
                        throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                }
                // - supplier selection

                //Pantalla de Impuestos
                foreach (var invoiceItem in invoiceItemList)
                {
                    if (invoiceItem.TaxList != null && invoiceItem.TaxList.Count > 0)
                    {
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<inputs> ";
                        var taxIndex = 1;
                        foreach (var tax in invoiceItem.TaxList)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value>" + tax.TaxCode + "</value> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }
                        while (taxIndex <= 12)
                        {
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>ATAX_CODE1I" + taxIndex + "</name> ";
                            requestXml = requestXml + "						<value/> ";
                            requestXml = requestXml + "					</screenField> ";
                            taxIndex++;
                        }
                        requestXml = requestXml + "				</inputs> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                        //confirmación impuestos
                        if (!responseDto.ResponseString.Contains("MSM26JA"))
                            throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");
                        requestXml = "<interaction> ";
                        requestXml = requestXml + "	<actions> ";
                        requestXml = requestXml + "		<action> ";
                        requestXml = requestXml + "			<name>submitScreen</name> ";
                        requestXml = requestXml + "			<data> ";
                        requestXml = requestXml + "				<screenName>MSM26JA</screenName> ";
                        requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction> ";
                        if (invoiceItem.FirstTaxAdjustment != 0)
                        {
                            requestXml = requestXml + "				<inputs> ";
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>TAX_VALUE1I1</name> ";
                            requestXml = requestXml + "						<value>" + Math.Abs(invoiceItem.FirstTaxAdjustment) + "</value> ";
                            requestXml = requestXml + "					</screenField> ";
                            requestXml = requestXml + "				</inputs> ";
                        }
                        requestXml = requestXml + "			</data> ";
                        requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id> ";
                        requestXml = requestXml + "		</action> ";
                        requestXml = requestXml + "	</actions> ";
                        requestXml = requestXml + "	<chains/> ";
                        requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId> ";
                        requestXml = requestXml + "	<application>ServiceInteraction</application> ";
                        requestXml = requestXml + "	<applicationPage>unknown</applicationPage> ";
                        requestXml = requestXml + "</interaction> ";

                        requestXml = requestXml.Replace("&", "&amp;");
                        responseDto = eFunctions.ExecutePostRequest(requestXml);

                        if (responseDto.GotErrorMessages())
                            throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                        if (invoiceItem.FirstTaxAdjustment != 0)
                        {
                            responseDto = eFunctions.ExecutePostRequest(requestXml);

                            if (responseDto.GotErrorMessages())
                                throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                        }
                    }
                }
                //

                //Pantalla de confirmación inicial
                if (!responseDto.ResponseString.Contains("MSM265A"))
                    throw new Exception("Se ha producido un error al intentar completar el proceso");
                requestXml = "<interaction>";
                requestXml = requestXml + "	<actions>";
                requestXml = requestXml + "		<action>";
                requestXml = requestXml + "			<name>submitScreen</name>";
                requestXml = requestXml + "			<data>";
                requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                requestXml = requestXml + "			</data>";
                requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                requestXml = requestXml + "		</action>";
                requestXml = requestXml + "	</actions>";
                requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                requestXml = requestXml + "	<application>ServiceInteraction</application>";
                requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                requestXml = requestXml + "</interaction>";

                responseDto = eFunctions.ExecutePostRequest(requestXml);
                if (responseDto.GotErrorMessages())
                    throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                //Pantalla de confirmación final
                if (!responseDto.ResponseString.Contains("MSM265A"))
                    throw new Exception("Se ha producido un error al intentar completar el proceso");
                requestXml = "<interaction>";
                requestXml = requestXml + "	<actions>";
                requestXml = requestXml + "		<action>";
                requestXml = requestXml + "			<name>submitScreen</name>";
                requestXml = requestXml + "			<data>";
                requestXml = requestXml + "				<screenName>MSM265A</screenName>";
                requestXml = requestXml + "				<screenAction>TRANSMIT</screenAction>";
                requestXml = requestXml + "			</data>";
                requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
                requestXml = requestXml + "		</action>";
                requestXml = requestXml + "	</actions>";
                requestXml = requestXml + "	<connectionId>" + eFunctions.PostServiceProxy.ConnectionId + "</connectionId>";
                requestXml = requestXml + "	<application>ServiceInteraction</application>";
                requestXml = requestXml + "	<applicationPage>unknown</applicationPage>";
                requestXml = requestXml + "</interaction>";

                responseDto = eFunctions.ExecutePostRequest(requestXml);
                if (responseDto.GotErrorMessages())
                    throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));

                return responseDto;
            }
        }

        public static class Queries
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
                    paramSupplierNo = " AND BNK.TAX_FILE_NO = '" + paramSupplierTaxFileNo + "'";

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
                               " AND BNK.SUP_STATUS <> 9";
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
}

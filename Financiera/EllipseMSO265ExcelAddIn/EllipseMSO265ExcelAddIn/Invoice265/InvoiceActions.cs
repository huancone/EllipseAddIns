using System;
using System.Collections.Generic;
using System.Linq;
using LINQtoCSV;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Utilities;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using Math = System.Math;


namespace EllipseMSO265ExcelAddIn.Invoice265
{
    public static class InvoiceActions
    {
        public static decimal GetItemTaxAdjustment(decimal itemValue, decimal calculatedTaxValue, decimal manualTaxValue, List<TaxCodeItem> taxList)
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
                    taxAdjustment = (int)Math.Round(taxValueItem - taxDifference, MidpointRounding.AwayFromZero);
                else
                    taxAdjustment = (int)Math.Round(taxValueItem + taxDifference, MidpointRounding.AwayFromZero);
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
                taxValueItem = (int)Math.Round(taxValueItem, MidpointRounding.AwayFromZero);

                if (MyUtilities.IsTrue(tax.Deduct))
                    taxValueItem = taxValueItem * -1;

                calculatedTaxValue += taxValueItem;
            }

            return (int)Math.Round(calculatedTaxValue, MidpointRounding.AwayFromZero);
        }

        public static Screen.ScreenDTO LoadNonInvoice(EllipseFunctions ef, string urlService, Screen.OperationContext opContext, Invoice invoice, List<InvoiceItem> invoiceItemList)
        {
                var screenService = new Screen.ScreenService();
                screenService.Url = urlService + "/ScreenService";

                ef.RevertOperation(opContext, screenService);
                //Default Values
                if (string.IsNullOrWhiteSpace(invoice.Currency))
                    invoice.Currency = "PES";
                if (string.IsNullOrWhiteSpace(invoice.HandlingCode))
                    invoice.HandlingCode = "PN";
                
                if (string.IsNullOrWhiteSpace(invoice.BankBranchCode) || string.IsNullOrWhiteSpace(invoice.BankAccountNo))
                {
                    var supplier = new Supplier(ef, invoice.SupplierNo, invoice.SupplierMnemonic, invoice.District);

                    invoice.BankBranchCode = supplier.BankBranchCode;
                    invoice.BankAccountNo = supplier.BankBranchAccountNo;
                    if (string.IsNullOrWhiteSpace(invoice.Currency))
                        invoice.Currency = supplier.CurrencyType;
                }

                foreach (var invoiceItem in invoiceItemList)
                {
                    if (string.IsNullOrWhiteSpace(invoiceItem.ItemDistrict))
                        invoiceItem.ItemDistrict = invoice.District;
                }
                //

                //ejecutamos el programa
                var reply = screenService.executeScreen(opContext, "MSO265");

                //Validamos el ingreso
                if (reply.mapName != "MSM265A")
                    throw new Exception("No se ha podido ingresar al programa MSO265");



                //La pantalla tiene un límite de registro de 3 ítems. Por lo que se debe procesar en la primera pantalla 3 ítems y posteriormente repetir el ejercicio
                var currentItemLimit = 3;
                var currentItemIndex = 0;


                var arrayFields = new ArrayScreenNameValue();
                while (currentItemIndex < invoiceItemList.Count)
                {
                    if (currentItemIndex < 3)
                    {

                        //Ingresamos la información principal
                        //se adicionan los valores a los campos

                        arrayFields.Add("DSTRCT_CODE1I", invoice.District);
                        arrayFields.Add("SUPPLIER_NO1I", invoice.SupplierNo);
                        if (string.IsNullOrWhiteSpace(invoice.SupplierNo) && !string.IsNullOrWhiteSpace(invoice.SupplierMnemonic))
                            arrayFields.Add("MNEMONIC1I", invoice.SupplierMnemonic);
                        if (!string.IsNullOrWhiteSpace(invoice.GovernmentId))
                            arrayFields.Add("GOVT_ID1I", invoice.GovernmentId);
                        arrayFields.Add("INV_NO1I", invoice.InvoiceNo);
                        arrayFields.Add("INV_AMT1I", invoice.InvoiceAmount);
                        if (invoice.TaxAmount != 0)
                            arrayFields.Add("ADD_TAX_AMOUNT1I", "" + invoice.TaxAmount);
                        if (!string.IsNullOrWhiteSpace(invoice.Accountant))
                            arrayFields.Add("ACCOUNTANT1I", invoice.Accountant);
                        if (!string.IsNullOrWhiteSpace(invoice.OriginalInvoiceNo))
                            arrayFields.Add("ORG_INV_NO1I", invoice.OriginalInvoiceNo);
                        arrayFields.Add("CURRENCY_TYPE1I", invoice.Currency);
                        arrayFields.Add("HANDLE_CDE1I", invoice.HandlingCode);
                        if (!string.IsNullOrWhiteSpace(invoice.ControlAccountGroupCode))
                            arrayFields.Add("ACCT_GRP_CODE1I", invoice.ControlAccountGroupCode);
                        arrayFields.Add("INV_DATE1I", invoice.InvoiceDate);
                        arrayFields.Add("INV_RCPT_DATE1I", invoice.InvoiceReceivedDate);
                        arrayFields.Add("DUE_DATE1I", invoice.DueDate);
                        if (!string.IsNullOrWhiteSpace(invoice.SettlementDiscount))
                            arrayFields.Add("SD_AMOUNT1I", invoice.SettlementDiscount);
                        if (!string.IsNullOrWhiteSpace(invoice.DiscountDate))
                            arrayFields.Add("SD_DATE1I", invoice.DiscountDate);
                        arrayFields.Add("BRANCH_CODE1I", invoice.BankBranchCode);
                        arrayFields.Add("BANK_ACCT_NO1I", invoice.BankAccountNo);

                    }

                    for (var i = currentItemLimit - 3; i < currentItemLimit; i++)
                    {
                        currentItemIndex++;
                        var iItem = currentItemIndex - (currentItemLimit - 3);
                        if (invoiceItemList == null || currentItemIndex > invoiceItemList.Count)
                            break;

                        arrayFields.Add("INV_ITEM_DESC1I" + iItem, invoiceItemList[i].Description);
                        arrayFields.Add("INV_ITEM_VALUE1I" + iItem, "" + invoiceItemList[i].ItemValue);
                        arrayFields.Add("ACCT_DSTRCT1I" + iItem, invoiceItemList[i].ItemDistrict);
                        arrayFields.Add("AUTH_BY1I" + iItem, invoiceItemList[i].AuthorizedBy);
                        arrayFields.Add("ACCOUNT1I" + iItem, invoiceItemList[i].Account);

                        if (!string.IsNullOrWhiteSpace(invoiceItemList[i].WorkOrderProjectNo))
                            arrayFields.Add("WORK_ORDER1I" + iItem, invoiceItemList[i].WorkOrderProjectNo);

                        if (!string.IsNullOrWhiteSpace(invoiceItemList[i].WorkOrderProjectIndicator))
                            arrayFields.Add("PROJECT_IND1I" + iItem, invoiceItemList[i].WorkOrderProjectIndicator);

                        if (!string.IsNullOrWhiteSpace(invoiceItemList[i].EquipNo))
                            arrayFields.Add("PLANT_NO1I" + iItem, invoiceItemList[i].EquipNo);

                        if (invoiceItemList[i].TaxList != null && invoiceItemList[i].TaxList.Count > 0)
                            arrayFields.Add("ACTION1I" + iItem, "T");
                    }

                    var request = new Screen.ScreenSubmitRequestDTO
                    {
                        screenFields = arrayFields.ToArray(),
                        screenKey = "1"
                    };
                    reply = screenService.submit(opContext, request);
                    if (ef.CheckReplyError(reply))
                    {
                        if(reply != null && !string.IsNullOrWhiteSpace(reply.message) && reply.message.Contains("X2:0011 - INPUT REQUIRED"))
                            throw new Exception(reply.message + " " + reply.currentCursorFieldName);
                        throw new Exception(reply.message);
                    }

                    //Pantalla de información del proveedor a la que ingresa internamente por el MNEMONIC / Cedula. Aquí es <4 porque el contador es válido para los tres primeros ítems
                    if (currentItemIndex < 4 && string.IsNullOrWhiteSpace(invoice.SupplierNo) && !string.IsNullOrWhiteSpace(invoice.SupplierMnemonic))
                    {
                        if (reply.mapName != "MSM202A")
                            throw new Exception("Se ha producido un error al intentar validar la información del Supplier");
                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("SUP_MNEMONIC1I", invoice.SupplierMnemonic);
                        arrayFields.Add("SUP_STATUS_IND1I", "A");

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };

                        reply = screenService.submit(opContext, request);
                        if (ef.CheckReplyError(reply))
                            throw new Exception(reply.message);

                    }

                    // - supplier selection
                    //Pantalla de Impuestos
                    for (var i = currentItemLimit - 3; i < currentItemLimit; i++)
                    {
                        if (invoiceItemList == null || i >= invoiceItemList.Count)
                            break;
                        if (invoiceItemList[i].TaxList != null && invoiceItemList[i].TaxList.Count > 0)
                        {
                            if (reply.mapName != "MSM26JA")
                                throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");


                            arrayFields = new ArrayScreenNameValue();
                            var taxIndex = 1;
                            foreach (var tax in invoiceItemList[i].TaxList)
                            {
                                arrayFields.Add("ATAX_CODE1I" + taxIndex, tax.TaxCode);
                                taxIndex++;
                            }

                            while (taxIndex <= 12)
                            {
                                arrayFields.Add("ATAX_CODE1I" + taxIndex, "");
                                taxIndex++;
                            }

                            request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenFields = arrayFields.ToArray(),
                                screenKey = "1"
                            };

                            reply = screenService.submit(opContext, request);
                            if (ef.CheckReplyError(reply))
                                throw new Exception(reply.message);

                            //confirmación impuestos
                            if (reply.mapName != "MSM26JA")
                                throw new Exception("Se ha producido un error al intentar añadir los códigos de Impuestos");

                            arrayFields = new ArrayScreenNameValue();

                            if (invoiceItemList[i].FirstTaxAdjustment != 0)
                                arrayFields.Add("TAX_VALUE1I1", "" + Math.Abs(invoiceItemList[i].FirstTaxAdjustment));

                            request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenFields = arrayFields.ToArray(),
                                screenKey = "1"
                            };

                            reply = screenService.submit(opContext, request);
                            if (ef.CheckReplyError(reply))
                                throw new Exception(reply.message);

                            if (invoiceItemList[i].FirstTaxAdjustment != 0)
                            {
                                request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };

                                reply = screenService.submit(opContext, request);
                                if (ef.CheckReplyError(reply))
                                    throw new Exception(reply.message);
                            }
                        }
                    }

                    //Pantalla de confirmación por ítems
                    do
                    {
                        if (reply.mapName != "MSM265A")
                            throw new Exception("Se ha producido un error al intentar completar el proceso");

                        arrayFields = new ArrayScreenNameValue();
                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };

                        reply = screenService.submit(opContext, request);
                        if (ef.CheckReplyError(reply))
                            throw new Exception(reply.message);

                    } while (ef.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.message.Contains("4493: ALL AMOUNTS ARE IN THE CURRENCY DISPLAYED ABOVE"));

                    currentItemLimit = currentItemLimit + 3;
                }
                //Si es múltiplo de 3 ítems debe volver a confirmar porque presenta una pantalla de nuevos ítems vacía y no repetirá el ciclo
                if (currentItemIndex == invoiceItemList.Count && (currentItemIndex % 3) == 0)
                {
                    arrayFields = new ArrayScreenNameValue();
                    var request = new Screen.ScreenSubmitRequestDTO
                    {
                        screenFields = arrayFields.ToArray(),
                        screenKey = "1"
                    };

                    reply = screenService.submit(opContext, request);
                    if (ef.CheckReplyError(reply))
                        throw new Exception(reply.message);
                }

            return reply;
        }

        /*
        public static ResponseDTO LoadNonInvoicePost(EllipseFunctions eFunctions, Invoice invoice, List<InvoiceItem> invoiceItemList)
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
                if (string.IsNullOrWhiteSpace(invoice.Currency))
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

            //La pantalla tiene un límite de registro de 3 ítems. Por lo que se debe procesar en la primera pantalla 3 ítems y posteriormente repetir el ejercicio
            var currentItemLimit = 3;
            var currentItemIndex = 0;

            while (currentItemIndex < invoiceItemList.Count)
            {
                if (!responseDto.ResponseString.Contains("MSM265A"))
                    throw new Exception("No se ha podido ingresar al programa MSO265");
                requestXml = "<interaction>                                                     ";
                requestXml = requestXml + "	<actions>";
                requestXml = requestXml + "		<action>";
                requestXml = requestXml + "			<name>submitScreen</name>";
                requestXml = requestXml + "			<data>";
                requestXml = requestXml + "				<inputs>";

                //Solo se envían estos datos la primera vez
                if (currentItemIndex < 3)
                {
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
                }

                for (var i = currentItemLimit - 3; i < currentItemLimit; i++)
                {
                    currentItemIndex++;
                    var iItem = currentItemIndex - (currentItemLimit - 3);
                    if (invoiceItemList == null || currentItemIndex > invoiceItemList.Count)
                        break;
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>INV_ITEM_DESC1I" + iItem + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItemList[i].Description + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					    <name>INV_ITEM_VALUE1I" + iItem + "</name>";
                    requestXml = requestXml + "					    <value>" + invoiceItemList[i].ItemValue + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "					 	<name>ACCT_DSTRCT1I" + iItem + "</name>";
                    requestXml = requestXml + "					   	<value>" + invoiceItemList[i].ItemDistrict + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>AUTH_BY1I" + iItem + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItemList[i].AuthorizedBy + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    requestXml = requestXml + "					<screenField>";
                    requestXml = requestXml + "						<name>ACCOUNT1I" + iItem + "</name>";
                    requestXml = requestXml + "						<value>" + invoiceItemList[i].Account + "</value>";
                    requestXml = requestXml + "					</screenField>";
                    if (!string.IsNullOrWhiteSpace(invoiceItemList[i].WorkOrderProjectNo))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>WORK_ORDER1I" + iItem + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItemList[i].WorkOrderProjectNo + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }

                    if (!string.IsNullOrWhiteSpace(invoiceItemList[i].WorkOrderProjectIndicator))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>WORK_PROJ_IND" + iItem + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItemList[i].WorkOrderProjectIndicator + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }

                    if (!string.IsNullOrWhiteSpace(invoiceItemList[i].EquipNo))
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>PLANT_NO1I" + iItem + "</name>";
                        requestXml = requestXml + "						<value>" + invoiceItemList[i].EquipNo + "</value>";
                        requestXml = requestXml + "					</screenField>";
                    }

                    if (invoiceItemList[i].TaxList != null && invoiceItemList[i].TaxList.Count > 0)
                    {
                        requestXml = requestXml + "					<screenField>";
                        requestXml = requestXml + "						<name>ACTION1I" + iItem + "</name>";
                        requestXml = requestXml + "						<value>T</value>";
                        requestXml = requestXml + "					</screenField>";
                    }
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

                //Pantalla de información del proveedor a la que ingresa internamente por el MNEMONIC / Cedula. Aquí es <4 porque el contador es válido para los tres primeros ítems
                if (currentItemIndex < 4 && string.IsNullOrWhiteSpace(invoice.SupplierNo) && !string.IsNullOrWhiteSpace(invoice.SupplierMnemonic))
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
                for (var i = currentItemLimit - 3; i < currentItemLimit; i++)
                {
                    if (invoiceItemList == null || i >= invoiceItemList.Count)
                        break;
                    if (invoiceItemList[i].TaxList != null && invoiceItemList[i].TaxList.Count > 0)
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
                        foreach (var tax in invoiceItemList[i].TaxList)
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
                        if (invoiceItemList[i].FirstTaxAdjustment != 0)
                        {
                            requestXml = requestXml + "				<inputs> ";
                            requestXml = requestXml + "					<screenField> ";
                            requestXml = requestXml + "						<name>TAX_VALUE1I1</name> ";
                            requestXml = requestXml + "						<value>" + Math.Abs(invoiceItemList[i].FirstTaxAdjustment) + "</value> ";
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

                        if (invoiceItemList[i].FirstTaxAdjustment != 0)
                        {
                            responseDto = eFunctions.ExecutePostRequest(requestXml);

                            if (responseDto.GotErrorMessages())
                                throw new Exception(responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text)));
                        }
                    }
                }
                //

                //Pantalla de confirmación por ítems
                do
                {
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
                } while (responseDto.GotWarningMessages() || responseDto.ResponseString.Contains("4493: ALL AMOUNTS ARE IN THE CURRENCY DISPLAYED ABOVE"));

                currentItemLimit = currentItemLimit + 3;
            }

            //Si es múltiplo de 3 ítems debe volver a confirmar porque presenta una pantalla de nuevos ítems vacía y no repetirá el ciclo
            if (currentItemIndex == invoiceItemList.Count && (currentItemIndex % 3) == 0)
            {
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
            }

            return responseDto;
        }*/

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

            if (dataReader == null || dataReader.IsClosed)
                return taxList;

            while (dataReader.Read())
            {

                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage); //!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            return taxList;
        }

        private static List<TaxCodeItem> GetTaxCodeList(EllipseFunctions eFunctions, List<string> taxCodesParamList, string taxGroupCode)
        {
            var taxList = new List<TaxCodeItem>();

            var dataReader = eFunctions.GetQueryResult(Queries.GetTaxCodeListQuery(taxCodesParamList, taxGroupCode));

            if (dataReader == null || dataReader.IsClosed)
                return taxList;

            while (dataReader.Read())
            {
                // ReSharper disable once UseObjectOrCollectionInitializer
                var tax = new TaxCodeItem();
                tax.TaxCode = dataReader["ATAX_CODE"].ToString().Trim();
                tax.TaxDescription = dataReader["DESCRIPTION"].ToString().Trim();
                tax.TaxReference = dataReader["TAX_REF"].ToString().Trim();
                var taxPercentage = dataReader["ATAX_RATE_9"];
                tax.TaxRatePerc = Convert.ToDecimal(taxPercentage); //!string.IsNullOrWhiteSpace(taxPercentage) ? Convert.ToDecimal(taxPercentage) : default(decimal);
                tax.DefaultToInvoiceItem = dataReader["DEFAULTED_IND"].ToString().Trim();
                tax.Deduct = dataReader["DEDUCT_SW"].ToString().Trim();

                taxList.Add(tax);
            }

            return taxList;
        }
    }
}

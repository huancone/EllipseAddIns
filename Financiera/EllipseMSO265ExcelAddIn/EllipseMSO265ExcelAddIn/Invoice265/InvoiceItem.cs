using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSO265ExcelAddIn.Invoice265
{
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
        public decimal FirstTaxAdjustment; //Usado para calcular las diferencias manuales (override) sobre el primer tax del item
    }
}

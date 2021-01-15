using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Ellipse;

namespace EllipseMSO261ProntoPagoExcelAddIn
{
    public class SupplierInvoiceInfo
    {

        public SupplierInvoiceInfo()
        {

        }

        public static SupplierInvoiceInfo GetSupplierInvoiceInfo(string supplier, string factura, EllipseFunctions ef)
        {

            var invoice = new SupplierInvoiceInfo();
            var sqlQuery = Queries.GetSupplierInvoiceInfo("ICOR", supplier, factura, ef.DbReference, ef.DbLink);

            var dr = ef.GetQueryResult(sqlQuery);

            if (dr == null || dr.IsClosed)
                return null;
            
            
            invoice.Supplier = dr["SUPPLIER_NO"].ToString();
            invoice.Factura = dr["EXT_INV_NO"].ToString();
            invoice.Fechapagosolicitada = null;
            invoice.Fechapagooriginal = dr["DUE_DATE"].ToString();
            invoice.PmtStatus = dr["PMT_STATUS"].ToString();
            invoice.Proveedor = dr["NOM_SUPPLIER"].ToString();
            invoice.CodigoBancoOriginal = dr["ORIG_BANK"].ToString();
            invoice.St = Convert.ToDouble(dr["NO_OF_DAYS_PAY"].ToString());
            invoice.Vrtotalfactura = Convert.ToDouble(dr["VLR_FACTURA"].ToString());
            invoice.VrBasedeDescuento = Convert.ToDouble(dr["VRBASE"].ToString());
            invoice.Diferencia = Convert.ToDouble(dr["DIREFENCIA"].ToString());
            invoice.Descuentocalculado = Convert.ToDouble(dr["VR_OTROS_DESCTS"].ToString());
            invoice.Vrdescuentoaplicado = Convert.ToDouble(dr["VR_OTROS_DESCTS"].ToString());
            invoice.Fechadepagomodificada = dr["FEC_MOD_PAGO"].ToString();
            invoice.BancodePagoModificado = dr["ORIG_BANK"].ToString();
            invoice.Error = "Success";

            return invoice;
        }
        public string Supplier { get; set; }
        public string Factura { get; set; }
        public string Fechapagosolicitada { get; set; }
        public string Fechapagooriginal { get; set; }
        public string PmtStatus { get; set; }
        public string Proveedor { get; set; }
        public string CodigoBancoOriginal { get; set; }
        public double St { get; set; }
        public double Vrtotalfactura { get; set; }
        public double VrBasedeDescuento { get; set; }
        public double Diferencia { get; set; }
        public double Descuentocalculado { get; set; }
        public double Vrdescuentoaplicado { get; set; }
        public string Fechadepagomodificada { get; set; }
        public string BancodePagoModificado { get; set; }
        public string Error { get; set; }
    }

}

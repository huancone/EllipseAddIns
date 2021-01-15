using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Ellipse;

namespace EllipseMSO905ExcelAddIn
{
    public class EllipseNit
    {
        public string Nit { get; set; }
        public string Error { get; set; }

        public EllipseNit(EllipseFunctions ef, string nit)
        {
            try
            {
                if (string.IsNullOrEmpty(nit))
                {
                    Error = "Nit Invalido";
                    return;
                }

                var ellipseNitQuery = Queries.GetSupplierNit(nit, ef.DbReference, ef.DbLink);
                var drEllipseNit = ef.GetQueryResult(ellipseNitQuery);
                if (drEllipseNit != null && !drEllipseNit.IsClosed)
                {
                    drEllipseNit.Read();
                    Nit = drEllipseNit["NIT"].ToString();
                    Error = null;
                }
                else
                {
                    Nit = null;
                    Error = "Nit no registrado";
                }
                if (drEllipseNit != null) drEllipseNit.Close();
            }
            catch (Exception error)
            {
                Error = error.Message;
            }
        }
    }
}

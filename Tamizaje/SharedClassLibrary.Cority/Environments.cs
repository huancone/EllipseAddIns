using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedClassLibrary.Cority
{
    public static class Environments
    {
        public static string GetServiceUrl(string serviceName)
        {
            switch (serviceName.ToUpper())
            {
                case "PRODUCTIVO":
                    return @"https://cerrejon.cority.com/webservice/MGIPService.svc?singleWsdl";
                case "TEST":
                    return @"https://cerrejon.maspcl1.medgate.com/gx2test/webservice/MGIPService.svc?singleWsdl";
                default:
                    throw new Exception("No se ha podido hallar el servidor " + serviceName + ". Seleccione un servidor válido e intente nuevamente");
            }
        }
    }
}

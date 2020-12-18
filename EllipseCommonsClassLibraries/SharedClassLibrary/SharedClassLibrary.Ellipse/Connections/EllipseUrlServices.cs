using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharedClassLibrary.Ellipse.Connections
{
    public class WebService
    {
        public static string Productivo = @"/ellipse/webservice/ellprod"; //XPath
        public static string Contingencia = @"/ellipse/webservice/ellcont"; //XPath
        public static string Desarrollo = @"/ellipse/webservice/elldesa"; //XPath
        public static string Test = @"/ellipse/webservice/elltest"; //XPath

        public static string UrlProductivo = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlContingencia = "http://ews-el8prod.lmnerp02.cerrejon.com/ews/services";
        public static string UrlDesarrollo = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services";
        public static string UrlTest = "http://ews-el8test.lmnerp03.cerrejon.com/ews/services/";
    }

    public class UrlPost
    {
        public static string Productivo = @"/ellipse/url/ellprod"; //XPath
        public static string Contingencia = @"/ellipse/url/ellcont"; //XPath
        public static string Desarrollo = @"/ellipse/url/elldesa"; //XPath
        public static string Test = @"/ellipse/url/elltest"; //XPath

        public static string UrlProductivo =
            "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.9.17_84/bind?app=";

        public static string UrlContingencia =
            "http://ellipse-el8prod.lmnerp02.cerrejon.com/ria-Ellipse-8.9.8_446/bind?app=";

        public static string UrlDesarrollo =
            "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.17_84/bind?app=";

        public static string UrlTest =
            "http://ellipse-el8test.lmnerp03.cerrejon.com/ria-Ellipse-8.9.17_84/bind?app=";
    }
}

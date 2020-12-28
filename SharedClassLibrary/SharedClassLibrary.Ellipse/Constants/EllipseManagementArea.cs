using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharedClassLibrary.Ellipse.Constants
{
    public class ManagementArea
    {
        public static List<KeyValuePair<string, string>> GetManagementArea()
        {
            var itemList = new List<KeyValuePair<string, string>> {ManejoDeCarbon, Mantenimiento, SoporteOperacion};
            return itemList;
        }
        public static KeyValuePair<string, string> ManejoDeCarbon = new KeyValuePair<string, string>("MDC", "MANEJO DE CARBON");
        public static KeyValuePair<string, string> Mantenimiento = new KeyValuePair<string, string>("MNTTO", "MANTENIMIENTO");
        public static KeyValuePair<string, string> SoporteOperacion = new KeyValuePair<string, string>("SOP", "SOPORTE A LA OPERACIÓN");
    }
}

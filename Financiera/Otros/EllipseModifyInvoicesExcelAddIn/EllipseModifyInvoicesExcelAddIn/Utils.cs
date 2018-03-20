using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseModifyInvoicesExcelAddIn
{
    public class Utils
    {
        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static String formatearCeldaACadena(String celda) {
            if(celda != null)
            {
                return celda.Trim();
            }
            else
            {
                return "";
            }
        }

        /*
         * VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA, ELIMINA LOS ESPACIOS y OBTIENE UNA SUBCADENA
          return: String
         */
        public static String formatearCeldaACadenaYDividir(String celda, int inicio, int fin)
        {
            if (celda != null)
            {
                if(celda.Length > fin)
                    return celda.Substring(inicio-1, fin).Trim();
                else
                    return celda.Trim();
            }
            else
            {
                return "";
            }
        }

        public static String formatearCeldaANumero(String celda,String formato)
        {
            if (celda != null)
            {
                return String.Format("{0:" + formato + "}", celda.Trim());
            }
            else
            {
                return "";
            }
        }

        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static String formatearCeldaACadenaSufijo(String celda, String sufijo)
        {
            if (celda != null)
            {
                if (celda.Trim().Contains(sufijo))
                {
                    return celda.Trim();
                }
                else
                {
                    return celda.Trim() + "%";
                }
            }
            else
            {
                return "";
            }
        }

        public static String formatearCeldaACadenaPadLeft(String celda,int longitud)
        {
            if (celda != null)
            {
                celda=celda.Trim();
                if (celda.Length >= longitud)
                    return celda;
                else
                    return celda.PadLeft(longitud, '0');

            }
            else
            {
                return "";
            }
        }
    }
}

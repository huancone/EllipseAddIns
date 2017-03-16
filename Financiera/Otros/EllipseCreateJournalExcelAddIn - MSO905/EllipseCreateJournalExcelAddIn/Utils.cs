using System;

namespace EllipseCreateJournalExcelAddIn
{
    public class Utils
    {
        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static string formatearCeldaACadena(string celda)
        {
            if (celda != null)
            {
                return celda.Trim().ToUpper();
            }
            return "";
        }

        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static double formatearCeldaADouble(string celda)
        {
            if (celda != null)
            {
                try
                {
                    return double.Parse(celda);
                }
                catch (Exception exc)
                {
                    return 0L;
                }
            }
            return 0L;
        }

        /*
         * VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA, ELIMINA LOS ESPACIOS y OBTIENE UNA SUBCADENA
          return: String
         */

        public static string formatearCeldaACadenaYDividir(string celda, int inicio, int fin)
        {
            if (celda != null)
            {
                if (celda.Length > fin)
                    return celda.Substring(inicio - 1, fin).Trim().ToUpper();
                return celda.Trim().ToUpper();
            }
            return "";
        }

        public static string formatearCeldaANumero(string celda, string formato)
        {
            if (celda != null)
            {
                return string.Format("{0:" + formato + "}", celda.Trim().ToUpper());
            }
            return "";
        }

        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static string formatearCeldaACadenaSufijo(string celda, string sufijo)
        {
            if (celda != null)
            {
                if (celda.Trim().Contains(sufijo))
                {
                    return celda.Trim().ToUpper();
                }
                return celda.Trim().ToUpper() + "%";
            }
            return "";
        }

        public static string formatearCeldaACadenaPadLeft(string celda, int longitud)
        {
            if (celda != null)
            {
                celda = celda.Trim().ToUpper();
                if (celda.Length >= longitud)
                    return celda;
                return celda.PadLeft(longitud, '0');
            }
            return "";
        }
    }
}
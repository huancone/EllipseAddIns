namespace EllipseMSO685Opc3ModifyExcelAddIn
{
    public class Utils
    {
        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static string FormatearCeldaACadena(string celda) {
            if(celda != null)
            {
                return celda.Trim().ToUpper();
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
        public static string FormatearCeldaACadenaYDividir(string celda, int inicio, int fin)
        {
            if (celda != null)
            {
                if(celda.Length > fin)
                    return celda.Substring(inicio - 1, fin).Trim().ToUpper();
                else
                    return celda.Trim().ToUpper();
            }
            else
            {
                return "";
            }
        }

        public static string FormatearCeldaANumero(string celda,string formato)
        {
            if (celda != null)
            {
                return string.Format("{0:" + formato + "}", celda.Trim().ToUpper());
            }
            else
            {
                return "";
            }
        }

        //VALIDA SI EL OBJETO ES NULO, SI ES NULO LO CONVIERTE A CADENA Y ELIMINA LOS ESPACIOS
        public static string FormatearCeldaACadenaSufijo(string celda, string sufijo)
        {
            if (celda != null)
            {
                if (celda.Trim().Contains(sufijo))
                {
                    return celda.Trim().ToUpper();
                }
                else
                {
                    return celda.Trim().ToUpper() + "%";
                }
            }
            else
            {
                return "";
            }
        }

        public static string FormatearCeldaACadenaPadLeft(string celda,int longitud)
        {
            if (celda != null)
            {
                celda = celda.Trim().ToUpper();
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

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace PlaneacionFerrocarril.PlanHistory
{
    public class PlanHistoryItem
    {
        public string Fecha;
        public string Grupo;
        public string IdConcepto;
        public string Concepto;
        public string Valor1;
        public string Valor2;

        public PlanHistoryItem()
        {
        }

        public PlanHistoryItem(IDataRecord dr)
        {
            Fecha = dr["FECHA"].ToString();
            Grupo = dr["GRUPO"].ToString();
            IdConcepto = dr["ID_CONCEPTO"].ToString();
            Concepto = dr["CONCEPTO"].ToString();
            Valor1 = dr["VALOR1"].ToString();
            Valor2 = dr["VALOR2"].ToString();
        }
    }
}

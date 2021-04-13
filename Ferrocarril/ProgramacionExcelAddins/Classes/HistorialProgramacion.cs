using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProgramacionExcelAddins.Classes
{
    public class HistorialProgramacion
    {
      private string _fecha;
        private string _grupo;
        private string _idConcepto;
        private string _concepto;
        private decimal _valor1;
        private decimal _valor2;    
        public void SetFecha ( string fecha)
        {
            this._fecha = fecha;
        }
        public string GetFecha()
        {
            return this._fecha;
        }
        public void SetGrupo(string grupo)
        {
            this._grupo = grupo;
        }
        public string GetGrupo()
        {
            return this._grupo;
        }
        public void SetIdConcepto(string idConcepto)
        {
            this._idConcepto = idConcepto;
        }
        public string GetIdConcepto()
        {
            return this._idConcepto;
        }
        public void SetConcepto(string concepto)
        {
            this._concepto = concepto;
        }
        public string GetConcepto()
        {
            return this._concepto;
        }
        public void SetValor1(decimal valor1)
        {
            this._valor1 = valor1;
        }
        public decimal GetValor1()
        {
            return this._valor1;
        }
        public void SetValor2(decimal valor2)
        {
            this._valor2 = valor2;
        }
        public decimal GetValor2()
        {
            return this._valor2;
        }
    }
}

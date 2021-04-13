using System;
using System.Collections.Generic;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Ellipse.Connections;

namespace BonoDeTopeados
{
    public class DatePeriod
    {
        public int Year;
        public int Period;
        private static List<string> _plOperacionesCedulas;
        private static string _periodYear;
        private static string _periodNumber;

        public static void SetPeriod(int year, int number)
        {
            SetPeriod(year.ToString(), number.ToString());
        }

        public static void SetPeriod(string year, string number)
        {
            _periodYear = year;
            _periodNumber = number;
        }
        public DatePeriod()
        {

        }

        public DatePeriod(int year, int period)
        {
            Year = year;
            Period = period;
        }
        public static DatePeriod GetPeriod(string cedula, int codDependencia, int year, int month, string modoPeriodo = "NORMAL")
        {
            if (modoPeriodo.Equals("MES CORRIDO"))
                return GetLagPeriod(year, month);
            if (modoPeriodo.Equals("MES FIJO"))
                return GetNormalPeriod(year, month);

            //Los grupos de operaciones utilizan un periodo normal (tres meses calendarios) mientras que mantenimiento utiliza un periodo lag (mes anterior)


            //Dependencias Operaciones
            var fcOperacionesDependencias = new List<int> { 102571, 102572, 102573, 102574 };
            var pbvOperacionesDependencias = new List<int> { 104262, 104261 };
            var pbvOperacionesMarinasDependencias = new List<int> { 104482, 104483, 104484, 104485, 104490, 104491, 105014, 105015, 105016, 105017 };
            //estas cédulas se deben tomar siempre del archivo TBL_KPI_SILOS
            if (_plOperacionesCedulas == null)
            {
                if (_periodYear == null)
                    throw new Exception("Año y Periodo de la clase DatePeriod no establecida. Se requiere para los valores de Plantas de Carbón Operaciones");
                //plOperacionesCedulas = new List<string> {"8791029", "84032997", "84009505", "5172132", "5165706", "5164429", "84009735", "17957448", "84080914", "15186148", "1120746035", "17971385", "84093114", "84008323"};
                
                _plOperacionesCedulas = new List<string>();

                using (var dbConn = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo)))
                {
                    var sqlQuery = Queries.GetPlantOperationsEmployees(_periodYear, _periodNumber);
                    var dReader = dbConn.GetQueryResult(sqlQuery);
                    if (dReader != null && !dReader.IsClosed)
                    {
                        while (dReader.Read())
                            _plOperacionesCedulas.Add("" + dReader["CEDULA"].ToString().Trim());
                    }
                }
            }

            foreach (var d in fcOperacionesDependencias)
            {
                if (d == codDependencia)
                    return GetNormalPeriod(year, month);
            }
            foreach (var d in pbvOperacionesDependencias)
            {
                if (d == codDependencia)
                    return GetNormalPeriod(year,month);
            }
            foreach (var d in pbvOperacionesMarinasDependencias)
            {
                if (d == codDependencia)
                    return GetNormalPeriod(year, month);
            }
            foreach (var c in _plOperacionesCedulas)
            {
                if (c.Equals(cedula))
                    return GetNormalPeriod(year, month);
            }
            //NORMAL PERIOD - OPERACIONES
            return GetLagPeriod(year, month);
        }

        public static DatePeriod GetLagPeriod(int year, int month)
        {
            
            //LAGPERIOD
            if (month >= 12)
                return new DatePeriod(year + 1, 1);
            if (month < 3)
                return new DatePeriod(year, 1);
            if (month < 6)
                return new DatePeriod(year, 2);
            if (month < 9)
                return new DatePeriod(year, 3);
            return new DatePeriod(year, 4);
        }
        public static DatePeriod GetNormalPeriod(int year, int month)
        {
            if (month <= 3)
                return new DatePeriod(year, 1);
            if (month <= 6)
                return new DatePeriod(year, 2);
            if (month <= 9)
                return new DatePeriod(year, 3);
            return new DatePeriod(year, 4);
        }
    }
}

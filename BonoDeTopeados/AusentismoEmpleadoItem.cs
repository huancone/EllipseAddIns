using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;


namespace BonoDeTopeados
{
    public class AusentismoEmpleadoItem
    {
        [CsvColumn(Name = "EMPLID", FieldIndex = 1)] public string EmployeeId;
        [CsvColumn(Name = "NAME", FieldIndex = 2)] public string Nombre;
        [CsvColumn(Name = "ANO_PROCS", FieldIndex = 3)] public int Anho;
        [CsvColumn(Name = "MES_PROCS", FieldIndex = 4)] public int Mes;
        [CsvColumn(Name = "HORAS_PROGR", FieldIndex = 5)] public double HrProg;
        [CsvColumn(Name = "HORAS_AUS_PDI", FieldIndex = 6)] public double HrAusPdi;
        [CsvColumn(Name = "HORAS_AUS_BONO", FieldIndex = 7)] public double HrAusBono;
        [CsvColumn(Name = "HORAS_VACACIONES", FieldIndex = 8)] public double HrVacaciones;
        [CsvColumn(Name = "HORAS_886", FieldIndex = 9)] public double Hr886;
    }
}

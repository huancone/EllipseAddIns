using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;


namespace BonoDeTopeados
{
    public class ListadoEmpleadoItem
    {
        [CsvColumn(Name = "Emplid", FieldIndex = 1)] public string EmployeeId;
        [CsvColumn(Name = "Nombre", FieldIndex = 2)] public string Nombre;
        [CsvColumn(Name = "Rol", FieldIndex = 3)] public int Rol;
        [CsvColumn(Name = "Cargo", FieldIndex = 4)] public string Cargo;
        [CsvColumn(Name = "Nivel_Empleado", FieldIndex = 5)] public int NivelEmpleado;
        [CsvColumn(Name = "Nivel_Cargo", FieldIndex = 6)] public int NivelCargo;
        [CsvColumn(Name = "Cod_Superintendencia", FieldIndex = 7)] public int CodigoSuperintendencia;
        [CsvColumn(Name = "Des_Superintendencia", FieldIndex = 8)] public string DescSuperintendencia;
        [CsvColumn(Name = "Cod_Cuadrilla", FieldIndex = 9)] public int CodCuadrilla;
        [CsvColumn(Name = "Des_Cuadrilla", FieldIndex = 10)] public string DescCuadrilla;
        [CsvColumn(Name = "Fecha", FieldIndex = 11)] public string Fecha;
        [CsvColumn(Name = "Turno", FieldIndex = 12)] public string Turno;
        [CsvColumn(Name = "Supervisor", FieldIndex = 13)] public string Supervisor;
    }
}

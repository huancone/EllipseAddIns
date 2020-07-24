using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Utilities.Shifts;

namespace BonoDeTopeados
{
    public class EmployeeTurns
    {
        public string Cedula;
        public int Anho;
        public int Periodo;
        public string Nombre;
        public string Cargo;
        public int CodSuperintendencia;
        public string DescSuperintendencia;
        public int CodDependencia;
        public string DescDependencia;
        public string Supervisor;
        public int NivelEmpleado;
        public int NivelCargo;
        public int Rol;
        public string Estado;
        public int TurnoD;
        public int TurnoL;
        public int TurnoI;
        public int TurnoM;
        public int TurnoN;
        public int TurnoP;
        public int TurnoT;
        public int TurnoV;
        public double TurnoOtro;

        public EmployeeTurns(ListadoEmpleadoItem empleadoItem, string modoPeriodo)
        {
            Cedula = empleadoItem.EmployeeId;
            var date = MyUtilities.ToDateTime(empleadoItem.Fecha, "yyyy-mm-dd");
            Anho = MyUtilities.ToInteger(MyUtilities.ToString(date, "yyyy"));
            var month = MyUtilities.ToInteger(MyUtilities.ToString(date, "mm"));
            Periodo = Period(empleadoItem.EmployeeId, empleadoItem.CodCuadrilla, month, modoPeriodo);
            Nombre = empleadoItem.Nombre;
            Cargo = empleadoItem.Cargo;

            CodSuperintendencia = empleadoItem.CodigoSuperintendencia;
            DescSuperintendencia = empleadoItem.DescSuperintendencia;
            CodDependencia = empleadoItem.CodCuadrilla;
            DescDependencia = empleadoItem.DescCuadrilla;
            Supervisor = empleadoItem.Supervisor;
            NivelEmpleado = empleadoItem.NivelEmpleado;
            NivelCargo = empleadoItem.NivelCargo;
            Rol = empleadoItem.Rol;
            if (empleadoItem.NivelCargo > empleadoItem.NivelEmpleado)
                Estado = "Pio";
            else if (empleadoItem.NivelCargo == empleadoItem.NivelEmpleado)
                Estado = "Topeado";
            else
                Estado = "NA";

            if (empleadoItem.Turno.Equals("D"))
                TurnoD++;
            else if (empleadoItem.Turno.Equals("L"))
                TurnoL++;
            else if (empleadoItem.Turno.Equals("I"))
                TurnoI++;
            else if (empleadoItem.Turno.Equals("M"))
                TurnoM++;
            else if (empleadoItem.Turno.Equals("N"))
                TurnoN++;
            else if (empleadoItem.Turno.Equals("P"))
                TurnoP++;
            else if (empleadoItem.Turno.Equals("T"))
                TurnoT++;
            else if (empleadoItem.Turno.Equals("V"))
                TurnoV++;
        }

        public EmployeeTurns()
        {

        }

        public int Period(string cedula, int codDependencia, int month, string modoPeriodo = "NORMAL")
        {
            if (modoPeriodo.Equals("MES CORRIDO"))
                return LagPeriod(month);
            if(modoPeriodo.Equals("MES FIJO"))
                return NormalPeriod(month);
            //Dependencias Operaciones
            var fcOperacionesDependencias = new List<int> {102571, 102572, 102573, 102574};
            var pbvOperacionesDependencias = new List<int> {104262,104261 };
            var pbvOperacionesMarinasDependencias = new List<int> { 104482, 104483, 104484, 104485, 104490, 104491, 105014 , 105015, 105016, 105017 };
            var plOperacionesCedulas = new List<string> { "8791029", "84032997", "84009505", "5172132", "5165706", "5164429", "84009735", "17957448", "84080914", "15186148", "1120746035", "17971385", "84093114", "84008323"};

            foreach(var d in fcOperacionesDependencias)
            {
                if (d == codDependencia)
                    return NormalPeriod(month);
            }
            foreach (var d in pbvOperacionesDependencias)
            {
                if (d == codDependencia)
                    return NormalPeriod(month);
            }
            foreach (var d in pbvOperacionesMarinasDependencias)
            {
                if (d == codDependencia)
                    return NormalPeriod(month);
            }
            foreach (var c in plOperacionesCedulas)
            {
                if (c.Equals(cedula))
                    return NormalPeriod(month);
            }
            //NORMAL PERIOD - OPERACIONES
            return LagPeriod(month);
        }

        private int LagPeriod(int month)
        {
            //LAGPERIOD
            if (month >= 12)
                return 1;
            else if (month < 3)
                return 1;
            else if (month < 6)
                return 2;
            else if (month < 9)
                return 3;
            else return 4;
        }
        private int NormalPeriod(int month)
        {
            if (month <= 3)
                return 1;
            else if (month <= 6)
                return 2;
            else if (month <= 9)
                return 3;
            else return 4;
        }
        public bool Equals(EmployeeTurns employeeTurn, bool ignoreTurns = false)
        {
            //Cedula
            if (string.IsNullOrWhiteSpace(Cedula) && !string.IsNullOrWhiteSpace(employeeTurn.Cedula))
                return false;
            if (!string.IsNullOrWhiteSpace(Cedula) && string.IsNullOrWhiteSpace(employeeTurn.Cedula))
                return false;
            if (!string.IsNullOrWhiteSpace(Cedula) && !string.IsNullOrWhiteSpace(employeeTurn.Cedula))
                if (!Cedula.Trim().PadLeft(15, '0').Equals(employeeTurn.Cedula.Trim().PadLeft(15, '0')))
                    return false;
            //
            if (Anho != employeeTurn.Anho)
                return false;
            if (Periodo != employeeTurn.Periodo)
                return false;
            if (!Nombre.Equals(employeeTurn.Nombre))
                return false;
            if (!Cargo.Equals(employeeTurn.Cargo))
                return false;
            if (CodSuperintendencia != employeeTurn.CodSuperintendencia)
                return false;
            if (!DescSuperintendencia.Equals(employeeTurn.DescSuperintendencia))
                return false;
            if (CodDependencia != employeeTurn.CodDependencia)
                return false;
            if (!DescDependencia.Equals(employeeTurn.DescDependencia))
                return false;
            if (!Supervisor.Equals(employeeTurn.Supervisor))
                return false;
            if (NivelEmpleado != employeeTurn.NivelEmpleado)
                return false;
            if (NivelCargo != employeeTurn.NivelCargo)
                return false;
            if (Rol != employeeTurn.Rol)
                return false;
            if (!Estado.Equals(employeeTurn.Estado))
                return false;

            if (!ignoreTurns)
            {
                if (TurnoD != employeeTurn.TurnoD)
                    return false;
                if (TurnoL != employeeTurn.TurnoL)
                    return false;
                if (TurnoI != employeeTurn.TurnoI)
                    return false;
                if (TurnoM != employeeTurn.TurnoM)
                    return false;
                if (TurnoN != employeeTurn.TurnoN)
                    return false;
                if (TurnoP != employeeTurn.TurnoP)
                    return false;
                if (TurnoT != employeeTurn.TurnoT)
                    return false;
                if (TurnoV != employeeTurn.TurnoV)
                    return false;
                if (TurnoOtro != employeeTurn.TurnoOtro)
                    return false;
            }

            return true;
        }

        public void SumTurns(EmployeeTurns employeeTurn)
        {
            TurnoD += employeeTurn.TurnoD;
            TurnoL += employeeTurn.TurnoL;
            TurnoI += employeeTurn.TurnoI;
            TurnoM += employeeTurn.TurnoM;
            TurnoN += employeeTurn.TurnoN;
            TurnoP += employeeTurn.TurnoP;
            TurnoT += employeeTurn.TurnoT;
            TurnoV += employeeTurn.TurnoV;
            TurnoOtro += employeeTurn.TurnoOtro;
        }
    }
}

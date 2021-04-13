using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BonoDeTopeados
{
    internal class Queries
    {
        public static string InsertEmployeeTurnType(EmployeeTurns empTurn)
        {
            var query = "MERGE INTO SIGMDC.TBL_RH_EMPLOYEE_TURNOTYPE T USING " +
                        "(SELECT " +
                        " '" + empTurn.Anho + "' ANO, " +
                        " '" + empTurn.Periodo + "' PERIODO, " +
                        " '" + empTurn.Cedula + "' CEDULA, " +
                        " '" + empTurn.Nombre + "' NOMBRE, " +
                        " '" + empTurn.Cargo + "' CARGO, " +
                        " '" + empTurn.CodSuperintendencia + "' COD_SUPERINTENDENCIA, " +
                        " '" + empTurn.DescSuperintendencia + "' SUPERINTENDENCIA, " +
                        " '" + empTurn.CodDependencia + "' COD_DEPENDENCIA, " +
                        " '" + empTurn.DescDependencia + "' DESC_DEPENDENCIA, " +
                        " '" + empTurn.Supervisor + "' SUPERVISOR, " +
                        " '" + empTurn.NivelEmpleado + "' NVL_EMP, " +
                        " '" + empTurn.NivelCargo + "' NVL_CARGO, " +
                        " '" + empTurn.Rol + "' ROL, " +
                        " '" + empTurn.Estado + "' ESTADO, " +
                        " '" + empTurn.TurnoD + "' TURNO_D, " +
                        " '" + empTurn.TurnoL + "' TURNO_L, " +
                        " '" + empTurn.TurnoM + "' TURNO_M, " +
                        " '" + empTurn.TurnoI + "' TURNO_I, " +
                        " '" + empTurn.TurnoN + "' TURNO_N, " +
                        " '" + empTurn.TurnoP + "' TURNO_P, " +
                        " '" + empTurn.TurnoT + "' TURNO_T, " +
                        " '" + empTurn.TurnoV + "' TURNO_V, " +
                        " '" + empTurn.TurnoOtro + "' TURNO_OTRO " +
                        " FROM DUAL) E ON ( " +
                        " T.ANO = E.ANO " +
                        " AND T.PERIODO = E.PERIODO " +
                        " AND T.CEDULA = E.CEDULA " +
                        ") " +
                        "WHEN MATCHED THEN UPDATE SET " + 
                        "  T.CARGO = E.CARGO, " +
                        "  T.COD_SUPERINTENDENCIA = E.COD_SUPERINTENDENCIA, " +
                        "  T.SUPERINTENDENCIA = E.SUPERINTENDENCIA, " +
                        "  T.COD_DEPENDENCIA = E.COD_DEPENDENCIA, " +
                        "  T.DESC_DEPENDENCIA = E.DESC_DEPENDENCIA, " +
                        "  T.SUPERVISOR = E.SUPERVISOR, " +
                        "  T.NVL_EMP = E.NVL_EMP, " +
                        "  T.NVL_CARGO = E.NVL_CARGO, " +
                        "  T.ROL = E.ROL, " +
                        "  T.ESTADO = E.ESTADO, " +
                        "  T.TURNO_D = E.TURNO_D, " +
                        "  T.TURNO_L = E.TURNO_L, " +
                        "  T.TURNO_M = E.TURNO_M, " +
                        "  T.TURNO_I = E.TURNO_I, " +
                        "  T.TURNO_N = E.TURNO_N, " +
                        "  T.TURNO_P = E.TURNO_P, " +
                        "  T.TURNO_T = E.TURNO_T, " +
                        "  T.TURNO_V = E.TURNO_V, " +
                        "  T.TURNO_OTRO = E.TURNO_OTRO " +
                        "WHEN NOT MATCHED THEN INSERT(" +
                        "  ANO, " +
                        "  PERIODO, " +
                        "  CEDULA, " +
                        "  NOMBRE, " +
                        "  CARGO, " +
                        "  COD_SUPERINTENDENCIA, " +
                        "  SUPERINTENDENCIA, " +
                        "  COD_DEPENDENCIA, " +
                        "  DESC_DEPENDENCIA, " +
                        "  SUPERVISOR, " +
                        "  NVL_EMP, " +
                        "  NVL_CARGO, " +
                        "  ROL, " +
                        "  ESTADO, " +
                        "  TURNO_D, " +
                        "  TURNO_L, " +
                        "  TURNO_M, " +
                        "  TURNO_I, " +
                        "  TURNO_N, " +
                        "  TURNO_P, " +
                        "  TURNO_T, " +
                        "  TURNO_V, " +
                        "  TURNO_OTRO " +
                        ") " +
                        "VALUES(" +
                        "  E.ANO, " +
                        "  E.PERIODO, " +
                        "  E.CEDULA, " +
                        "  E.NOMBRE, " +
                        "  E.CARGO, " +
                        "  E.COD_SUPERINTENDENCIA, " +
                        "  E.SUPERINTENDENCIA, " +
                        "  E.COD_DEPENDENCIA, " +
                        "  E.DESC_DEPENDENCIA, " +
                        "  E.SUPERVISOR, " +
                        "  E.NVL_EMP, " +
                        "  E.NVL_CARGO, " +
                        "  E.ROL, " +
                        "  E.ESTADO, " +
                        "  E.TURNO_D, " +
                        "  E.TURNO_L, " +
                        "  E.TURNO_M, " +
                        "  E.TURNO_I, " +
                        "  E.TURNO_N, " +
                        "  E.TURNO_P, " +
                        "  E.TURNO_T, " +
                        "  E.TURNO_V, " +
                        "  E.TURNO_OTRO " +
                        ") ";

            return query;
        }


        public static string InsertEmployeeTurn886(EmployeeTurns empTurn)
        {
            var query = "MERGE INTO SIGMDC.TBL_RH_EMPLOYEE_TURNOTYPE T USING " +
                        "(SELECT " +
                        " '" + empTurn.Anho + "' ANO, " +
                        " '" + empTurn.Periodo + "' PERIODO, " +
                        " '" + empTurn.Cedula + "' CEDULA, " +
                        " '" + empTurn.Nombre + "' NOMBRE, " +
                        " '" + empTurn.Cargo + "' CARGO, " +
                        " '" + empTurn.CodSuperintendencia + "' COD_SUPERINTENDENCIA, " +
                        " '" + empTurn.DescSuperintendencia + "' SUPERINTENDENCIA, " +
                        " '" + empTurn.CodDependencia + "' COD_DEPENDENCIA, " +
                        " '" + empTurn.DescDependencia + "' DESC_DEPENDENCIA, " +
                        " '" + empTurn.Supervisor + "' SUPERVISOR, " +
                        " '" + empTurn.NivelEmpleado + "' NVL_EMP, " +
                        " '" + empTurn.NivelCargo + "' NVL_CARGO, " +
                        " '" + empTurn.Rol + "' ROL, " +
                        " '" + empTurn.Estado + "' ESTADO, " +
                        " '" + empTurn.TurnoD + "' TURNO_D, " +
                        " '" + empTurn.TurnoL + "' TURNO_L, " +
                        " '" + empTurn.TurnoM + "' TURNO_M, " +
                        " '" + empTurn.TurnoI + "' TURNO_I, " +
                        " '" + empTurn.TurnoN + "' TURNO_N, " +
                        " '" + empTurn.TurnoP + "' TURNO_P, " +
                        " '" + empTurn.TurnoT + "' TURNO_T, " +
                        " '" + empTurn.TurnoV + "' TURNO_V, " +
                        " '" + empTurn.TurnoOtro + "' TURNO_OTRO " +
                        " FROM DUAL) E ON ( " +
                        " T.ANO = E.ANO " +
                        " AND T.PERIODO = E.PERIODO " +
                        " AND T.CEDULA = E.CEDULA " +
                        ") " +
                        "WHEN MATCHED THEN UPDATE SET " +
                        //"  T.CARGO = E.CARGO, " +
                        //"  T.COD_SUPERINTENDENCIA = E.COD_SUPERINTENDENCIA, " +
                        //"  T.SUPERINTENDENCIA = E.SUPERINTENDENCIA, " +
                        //"  T.COD_DEPENDENCIA = E.COD_DEPENDENCIA, " +
                        //"  T.DESC_DEPENDENCIA = E.DESC_DEPENDENCIA, " +
                        //"  T.SUPERVISOR = E.SUPERVISOR, " +
                        //"  T.NVL_EMP = E.NVL_EMP, " +
                        //"  T.NVL_CARGO = E.NVL_CARGO, " +
                        //"  T.ROL = E.ROL, " +
                        //"  T.ESTADO = E.ESTADO, " +
                        //"  T.TURNO_D = E.TURNO_D, " +
                        //"  T.TURNO_L = E.TURNO_L, " +
                        //"  T.TURNO_M = E.TURNO_M, " +
                        //"  T.TURNO_I = E.TURNO_I, " +
                        //"  T.TURNO_N = E.TURNO_N, " +
                        //"  T.TURNO_P = E.TURNO_P, " +
                        //"  T.TURNO_T = E.TURNO_T, " +
                        //"  T.TURNO_V = E.TURNO_V, " +
                        "  T.TURNO_OTRO = E.TURNO_OTRO " +
                        "WHEN NOT MATCHED THEN INSERT(" +
                        "  ANO, " +
                        "  PERIODO, " +
                        "  CEDULA, " +
                        "  NOMBRE, " +
                        "  CARGO, " +
                        "  COD_SUPERINTENDENCIA, " +
                        "  SUPERINTENDENCIA, " +
                        "  COD_DEPENDENCIA, " +
                        "  DESC_DEPENDENCIA, " +
                        "  SUPERVISOR, " +
                        "  NVL_EMP, " +
                        "  NVL_CARGO, " +
                        "  ROL, " +
                        "  ESTADO, " +
                        "  TURNO_D, " +
                        "  TURNO_L, " +
                        "  TURNO_M, " +
                        "  TURNO_I, " +
                        "  TURNO_N, " +
                        "  TURNO_P, " +
                        "  TURNO_T, " +
                        "  TURNO_V, " +
                        "  TURNO_OTRO " +
                        ") " +
                        "VALUES(" +
                        "  E.ANO, " +
                        "  E.PERIODO, " +
                        "  E.CEDULA, " +
                        "  E.NOMBRE, " +
                        "  E.CARGO, " +
                        "  E.COD_SUPERINTENDENCIA, " +
                        "  E.SUPERINTENDENCIA, " +
                        "  E.COD_DEPENDENCIA, " +
                        "  E.DESC_DEPENDENCIA, " +
                        "  E.SUPERVISOR, " +
                        "  E.NVL_EMP, " +
                        "  E.NVL_CARGO, " +
                        "  E.ROL, " +
                        "  E.ESTADO, " +
                        "  E.TURNO_D, " +
                        "  E.TURNO_L, " +
                        "  E.TURNO_M, " +
                        "  E.TURNO_I, " +
                        "  E.TURNO_N, " +
                        "  E.TURNO_P, " +
                        "  E.TURNO_T, " +
                        "  E.TURNO_V, " +
                        "  E.TURNO_OTRO " +
                        ") ";

            return query;
        }

        public static string GetEmployeeTurnType(string cedula, int anho, int periodo)
        {
            var query = "SELECT " +
                        "  E.ANO, " +
                        "  E.PERIODO, " +
                        "  E.CEDULA, " +
                        "  E.NOMBRE, " +
                        "  E.CARGO, " +
                        "  E.COD_SUPERINTENDENCIA, " +
                        "  E.SUPERINTENDENCIA, " +
                        "  E.COD_DEPENDENCIA, " +
                        "  E.DESC_DEPENDENCIA, " +
                        "  E.SUPERVISOR, " +
                        "  E.NVL_EMP, " +
                        "  E.NVL_CARGO, " +
                        "  E.ROL, " +
                        "  E.ESTADO, " +
                        "  E.TURNO_D, " +
                        "  E.TURNO_L, " +
                        "  E.TURNO_M, " +
                        "  E.TURNO_I, " +
                        "  E.TURNO_N, " +
                        "  E.TURNO_P, " +
                        "  E.TURNO_T, " +
                        "  E.TURNO_V, " +
                        "  E.TURNO_OTRO " +
                        " FROM SIGMDC.TBL_RH_EMPLOYEE_TURNOTYPE E" +
                        " WHERE E.CEDULA = '" + cedula + "' AND E.ANO = '" + anho + "' AND E.PERIODO = '" + periodo + "'";

            return query;
        }

        public static string GetPlantOperationsEmployees(string anho, string periodo)
        {
            var query = "SELECT EMPS.CEDULA" +
            " FROM SIGMDC.TBL_KPIS_SILOS EMPS" +
                " WHERE EMPS.ANO = '" + anho + "' AND EMPS.PERIODO = '" + periodo + "'";
            return query;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SharedClassLibrary.Utilities;
using System.Data;
using SharedClassLibrary;

namespace InvestigacionesIcam
{
    public static class Queries
    {

        public static string GetAccidentsQuery(List <KeyValuePair<int?, string>> searchParameters)
        {
            //Parámetros de Accidentes
            string paramCodAccidente = null, paramFechaInicialAccidente = null, paramFechaFinalAccidente = null, 
                paramCodSuperIntendenciaAccidente = null, paramEstadoAccidente = null, 
                paramIdResponsableAccidente = null, paramPotencialAccidente = null;


            foreach (var searchParam in searchParameters)
            {
                if (string.IsNullOrWhiteSpace(searchParam.Value))
                    continue;

                switch (searchParam.Key)
                {
                    //Parámetros de Accidentes
                    case (int)IxSearchParameters.CodigoAccidente:
                        paramCodAccidente = " AND INF.COD_ACCIDENTE = '" + searchParam.Value + "'";
                        break;

                    case (int)IxSearchParameters.FechaInicialAccidente:
                        paramFechaInicialAccidente = " AND INF.FEC_CASO >= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaFinalAccidente:
                        paramFechaFinalAccidente = " AND INF.FEC_CASO <= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.CodigoSuperIntendenciaAccidente:
                        paramCodSuperIntendenciaAccidente = " AND INF.COD_SUPERINTENDENCIA = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.EstadoAccidente:
                        if (searchParam.Value.ToUpper().Equals("CANCELADO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO = '9' ";
                        else if (searchParam.Value.ToUpper().Equals("ACTIVO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO <> '9' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableAccidente:
                        paramIdResponsableAccidente = " AND INF.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.PotencialAccidente:
                        if(searchParam.Value.ToUpper().Equals("BAJO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'BAJO' ";
                        if (searchParam.Value.ToUpper().Equals("MEDIO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'MEDIO' ";
                        if (searchParam.Value.ToUpper().Equals("ALTO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'ALTO' ";
                        if (searchParam.Value.ToUpper().Equals("ICAM"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL IN ('MEDIO', 'ALTO') ";
                        break;

                    default:
                        break;
                }
            }

            var sqlQuery = @"SELECT 
                      INF.COD_ACCIDENTE, INF.COD_ESTADO,
                      (CASE WHEN INF.COD_ESTADO= 9 THEN 'Cancelado'  ELSE 'Activo' END) DESC_ESTADO,
                      TO_CHAR(INF.FEC_CASO, 'YYYYMMDD') FEC_CASO, INF.HOR_CASO, 
                      TO_CHAR(INF.FEC_TURNO, 'YYYYMMDD') FEC_TURNO, INF.COD_TURNO,
                      INF.NOM_REPORTA, INF.NRO_ID_REPORTA,
                      INF.USR_GENERA, TO_CHAR(INF.FEC_GENERA, 'YYYYMMDD') FEC_GENERA, INF.NRO_ID_RESPONSABLE, AEMP.NAME NOMBRE_RESPONSABLE,
                      INF.USR_MODIFICA, TO_CHAR(INF.FEC_MODIFICA, 'YYYYMMDD') FEC_MODIFICA,
                      INF.DES_CARGO, INF.COD_EMPRESA_REP, INF.COD_DEPARTAMENTO_REP, 
                      INF.DES_INFORME, INF.DES_COMENTARIOS, 
                      INF.IND_LESION, INF.IND_DANO, INF.IND_AMBIENTAL, INF.IND_COMUNIDAD, INF.IND_INCENDIO,
                      INF.COD_POSIBLE_OCURRENCIA, INF.COD_POSIBLE_LESION, INF.COD_PELIGRO_ASOCIADO,
                      INF.DES_POTENCIAL, 
                      INF.COD_TAREA_REALIZABA, 
                      INF.COD_AREA, AAREA.DES_AREA, INF.COD_SITIO, ASIT.DES_TIPO DESC_SITIO, INF.COD_LOCALIZACION, ALOC.DES_TIPO DESC_LOCALIZACION, INF.COD_SUPERINTENDENCIA, ASUP.DES_SUPERINTENDENCIA,
                      INF.COD_SECCION, TO_CHAR(INF.FEC_ENTREGA_INVESTIGACION, 'YYYYMMDD') FEC_ENTREGA_INVESTIGACION, INF.IND_CONFIDENCIAL, TO_CHAR(INF.FEC_COMITE, 'YYYYMMDD') FEC_COMITE, 
                      INF.NRO_ID_APRUEBA_INV, INF.IND_GRAVE, INF.COD_TIPO_FRCP, INF.IND_CERO_BARRERAS, INF.IND_ACC_REPETITIVO, INF.IND_EXISTE_ESTANDAR_CONTROL, INF.IND_REPORTE_DIARIO
                    FROM ADMINSIIO.AC_INFORMES INF 
                      LEFT JOIN ADMINSIIO.SYS_SUPERINTENDENCIA ASUP ON INF.COD_SUPERINTENDENCIA = ASUP.COD_SUPERINTENDENCIA 
                      LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS AEMP ON INF.NRO_ID_RESPONSABLE = AEMP.EMPLID
                      LEFT JOIN ADMINSIIO.SYS_AREAS AAREA ON INF.COD_AREA = AAREA.COD_AREA 
                      LEFT JOIN ADMINSIIO.SYS_TIPO ALOC ON INF.COD_LOCALIZACION = ALOC.COD_TIPO AND ALOC.COD_CLASE = 'LO'
                      LEFT JOIN ADMINSIIO.SYS_TIPO ASIT ON INF.COD_LOCALIZACION = ASIT.COD_TIPO AND ASIT.COD_CLASE = 'ST'
                    WHERE 
                    ";
            if (!string.IsNullOrWhiteSpace(paramCodAccidente))
                sqlQuery += paramCodAccidente;
            else
            {
                //Parámetros de Accidente
                if (!string.IsNullOrWhiteSpace(paramFechaInicialAccidente))
                    sqlQuery += paramFechaInicialAccidente;
                if (!string.IsNullOrWhiteSpace(paramFechaFinalAccidente))
                    sqlQuery += paramFechaFinalAccidente;
                if (!string.IsNullOrWhiteSpace(paramCodSuperIntendenciaAccidente))
                    sqlQuery += paramCodSuperIntendenciaAccidente;
                if (!string.IsNullOrWhiteSpace(paramEstadoAccidente))
                    sqlQuery += paramEstadoAccidente;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableAccidente))
                    sqlQuery += paramIdResponsableAccidente;
                if (!string.IsNullOrWhiteSpace(paramPotencialAccidente))
                    sqlQuery += paramPotencialAccidente;

            }

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }


        public static string GetRecomendationsQuery(List<KeyValuePair<int?, string>> searchParameters)
        {
            //Parámetros de Accidentes
            string paramCodAccidente = null, paramFechaInicialAccidente = null, paramFechaFinalAccidente = null,
                paramCodSuperIntendenciaAccidente = null, paramEstadoAccidente = null,
                paramIdResponsableAccidente = null, paramPotencialAccidente = null;
            //Parámetros de Recomendaciones
            string paramCodEstadoRecomendacion = null, paramIdResponsableRecomendacion = null, paramCodDepartamentoResponsableRecomendacion = null, 
                paramFechaCierreRecomendacion = null, paramFechaPlaneadaRecomendacion = null;

            foreach (var searchParam in searchParameters)
            {
                if (string.IsNullOrWhiteSpace(searchParam.Value))
                    continue;

                switch (searchParam.Key)
                {
                    //Parámetros de Accidentes
                    case (int)IxSearchParameters.CodigoAccidente:
                        paramCodAccidente = " AND INF.COD_ACCIDENTE = '" + searchParam.Value + "'";
                        break;

                    case (int)IxSearchParameters.FechaInicialAccidente:
                        paramFechaInicialAccidente = " AND INF.FEC_CASO >= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaFinalAccidente:
                        paramFechaFinalAccidente = " AND INF.FEC_CASO <= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.CodigoSuperIntendenciaAccidente:
                        paramCodSuperIntendenciaAccidente = " AND INF.COD_SUPERINTENDENCIA = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.EstadoAccidente:
                        if (searchParam.Value.ToUpper().Equals("CANCELADO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO = '9' ";
                        else if (searchParam.Value.ToUpper().Equals("ACTIVO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO <> '9' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableAccidente:
                        paramIdResponsableAccidente = " AND INF.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.PotencialAccidente:
                        if (searchParam.Value.ToUpper().Equals("BAJO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'BAJO' ";
                        if (searchParam.Value.ToUpper().Equals("MEDIO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'MEDIO' ";
                        if (searchParam.Value.ToUpper().Equals("ALTO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'ALTO' ";
                        if (searchParam.Value.ToUpper().Equals("ICAM"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL IN ('MEDIO', 'ALTO') ";
                        break;

                    //Parámetros de Recomendaciones
                    case (int)IxSearchParameters.CodigoEstadoRecomendacion:
                        paramCodEstadoRecomendacion = " AND REC.COD_ESTADO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableRecomendacion:
                        paramIdResponsableRecomendacion = " AND REC.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.CodigoDepartamentoRecomendacion:
                        paramCodDepartamentoResponsableRecomendacion = " AND REC.COD_DEPARTAMENTO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;
                    case (int)IxSearchParameters.FechaCreacionRecomendacion:
                        paramFechaCierreRecomendacion = " AND REC.FEC_CIERRE = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;
                    case (int)IxSearchParameters.FechaPlaneadaRecomendacion:
                        paramFechaPlaneadaRecomendacion = " AND REC.FEC_PLANEADA = = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;
                    default:
                        break;
                }
            }

            var sqlQuery = @" 
                SELECT 
                  (REC.COD_SISTEMA||'-'||REC.COD_EVENTO||'-'||REC.COD_RECOMENDACION) ID_RECOMENDACION, REC.COD_SISTEMA,
                  REC.COD_EVENTO, 
                  REC.COD_RECOMENDACION, REC.DES_RECOMENDACION, REC.COD_ESTADO, RCODEST.DES_ESTADO, 
                  TO_CHAR(REC.FEC_CREACION, 'YYYYMMDD') FEC_CREACION, REC.USR_CREACION, TO_CHAR(REC.FEC_PLANEADA, 'YYYYMMDD') FEC_PLANEADA, TO_CHAR(REC.FEC_CIERRE, 'YYYYMMDD') FEC_CIERRE,
                  REC.NRO_ID_RESPONSABLE, REMP.NAME NOMBRE_RESPONSABLE, REC.COD_DEPARTAMENTO, RDEP.DES_LARGA DESC_DEPARTAMENTO, REC.COD_TIPO_ERTC,
                  REC.COD_DEPTID
                FROM 
                  ADMINSIIO.AC_INFORMES INF JOIN ADMINSIIO.TR_RECOMENDACIONES REC ON INF.COD_ACCIDENTE = REC.COD_EVENTO AND REC.COD_SISTEMA = '001'
                  LEFT JOIN ADMINSIIO.TR_ESTADOS RCODEST ON REC.COD_ESTADO = RCODEST.COD_ESTADO AND RCODEST.COD_SISTEMA = '011'
                  LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS REMP ON REC.NRO_ID_RESPONSABLE = REMP.EMPLID
                  LEFT JOIN ADMINSIIO.SYS_DEPARTAMENTOS RDEP ON REC.COD_DEPARTAMENTO = RDEP.COD_DEPARTAMENTO 
                WHERE 
                    ";
            if (!string.IsNullOrWhiteSpace(paramCodAccidente))
                sqlQuery += paramCodAccidente;
            else
            {
                //Parámetros de Accidente
                if (!string.IsNullOrWhiteSpace(paramFechaInicialAccidente))
                    sqlQuery += paramFechaInicialAccidente;
                if (!string.IsNullOrWhiteSpace(paramFechaFinalAccidente))
                    sqlQuery += paramFechaFinalAccidente;
                if (!string.IsNullOrWhiteSpace(paramCodSuperIntendenciaAccidente))
                    sqlQuery += paramCodSuperIntendenciaAccidente;
                if (!string.IsNullOrWhiteSpace(paramEstadoAccidente))
                    sqlQuery += paramEstadoAccidente;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableAccidente))
                    sqlQuery += paramIdResponsableAccidente;
                if (!string.IsNullOrWhiteSpace(paramPotencialAccidente))
                    sqlQuery += paramPotencialAccidente;
                //Parámetros de Recomendaciones
                if (!string.IsNullOrWhiteSpace(paramCodEstadoRecomendacion))
                    sqlQuery += paramCodEstadoRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableRecomendacion))
                    sqlQuery += paramIdResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramCodDepartamentoResponsableRecomendacion))
                    sqlQuery += paramCodDepartamentoResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaCierreRecomendacion))
                    sqlQuery += paramFechaCierreRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaPlaneadaRecomendacion))
                    sqlQuery += paramFechaPlaneadaRecomendacion;

            }
        

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }




        public static string GetPlansQuery(List<KeyValuePair<int?, string>> searchParameters)
        {
            //Parámetros de Accidentes
            string paramCodAccidente = null, paramFechaInicialAccidente = null, paramFechaFinalAccidente = null,
                paramCodSuperIntendenciaAccidente = null, paramEstadoAccidente = null,
                paramIdResponsableAccidente = null, paramPotencialAccidente = null;
            //Parámetros de Recomendaciones
            string paramCodEstadoRecomendacion = null, paramIdResponsableRecomendacion = null, paramCodDepartamentoResponsableRecomendacion = null,
                paramFechaCierreRecomendacion = null, paramFechaPlaneadaRecomendacion = null;
            //Parámetros de Planes de Acción
            string paramIdResponsablePlan = null, paramCodEstadoPlan = null, paramFechaPlaneadaPlan = null;

            foreach (var searchParam in searchParameters)
            {
                if (string.IsNullOrWhiteSpace(searchParam.Value))
                    continue;

                switch (searchParam.Key)
                {                   
                    //Parámetros de Accidentes
                    case (int)IxSearchParameters.CodigoAccidente:
                        paramCodAccidente = " AND INF.COD_ACCIDENTE = '" + searchParam.Value + "'";
                        break;

                    case (int)IxSearchParameters.FechaInicialAccidente:
                        paramFechaInicialAccidente = " AND INF.FEC_CASO >= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaFinalAccidente:
                        paramFechaFinalAccidente = " AND INF.FEC_CASO <= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.CodigoSuperIntendenciaAccidente:
                        paramCodSuperIntendenciaAccidente = " AND INF.COD_SUPERINTENDENCIA = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.EstadoAccidente:
                        if (searchParam.Value.ToUpper().Equals("CANCELADO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO = '9' ";
                        else if (searchParam.Value.ToUpper().Equals("ACTIVO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO <> '9' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableAccidente:
                        paramIdResponsableAccidente = " AND INF.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.PotencialAccidente:
                        if (searchParam.Value.ToUpper().Equals("BAJO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'BAJO' ";
                        if (searchParam.Value.ToUpper().Equals("MEDIO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'MEDIO' ";
                        if (searchParam.Value.ToUpper().Equals("ALTO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'ALTO' ";
                        if (searchParam.Value.ToUpper().Equals("ICAM"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL IN ('MEDIO', 'ALTO') ";
                        break;

                    //Parámetros de Recomendaciones
                    case (int)IxSearchParameters.CodigoEstadoRecomendacion:
                        paramCodEstadoRecomendacion = " AND REC.COD_ESTADO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableRecomendacion:
                        paramIdResponsableRecomendacion = " AND REC.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.CodigoDepartamentoRecomendacion:
                        paramCodDepartamentoResponsableRecomendacion = " AND REC.COD_DEPARTAMENTO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.FechaCreacionRecomendacion:
                        paramFechaCierreRecomendacion = " AND REC.FEC_CIERRE = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaPlaneadaRecomendacion:
                        paramFechaPlaneadaRecomendacion = " AND REC.FEC_PLANEADA = = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    //Parámetros de Planes de Acción
                    case (int)IxSearchParameters.CodigoEstadoPlan:
                        paramCodEstadoPlan = " AND PAC.COD_ESTADO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.IdResponsablePlan:
                        paramIdResponsablePlan = " AND PAC.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.FechaPlaneadaPlan:
                        paramFechaPlaneadaPlan = " AND PAC.FEC_PLANEADA = = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    default:
                        break;
                }
            }

            var sqlQuery = @" 
                SELECT 
                  (REC.COD_SISTEMA||'-'||REC.COD_EVENTO||'-'||REC.COD_RECOMENDACION) ID_RECOMENDACION,
                  PAC.COD_SISTEMA, PAC.COD_EVENTO, PAC.COD_RECOMENDACION, 
                  PAC.COD_PLAN, 
                  ADMINSIIO.LTOVC_TR_PLANES_ACC_DES_PLAN(PAC.ROWID) AS DES_PLAN,
                  PAC.USR_CREACION, TO_CHAR(PAC.FEC_PLANEADA, 'YYYYMMDD') FEC_PLANEADA, PAC.NRO_ID_RESPONSABLE, PEMP.NAME NOMBRE_RESPONSABLE,
                  PAC.COD_ESTADO, PCODEST.DES_ESTADO, PAC.POR_AVANCE,
                  ADMINSIIO.LTOVC_TR_AVANCES_DES_AVANCE(AVA.ROWID) AS DES_AVANCE
                FROM 
                  ADMINSIIO.AC_INFORMES INF JOIN ADMINSIIO.TR_RECOMENDACIONES REC ON INF.COD_ACCIDENTE = REC.COD_EVENTO AND REC.COD_SISTEMA = '001'
                  JOIN ADMINSIIO.TR_PLANES_ACCION PAC ON (REC.COD_SISTEMA = PAC.COD_SISTEMA AND REC.COD_EVENTO = PAC.COD_EVENTO AND REC.COD_RECOMENDACION = PAC.COD_RECOMENDACION)
                  LEFT JOIN ADMINSIIO.TR_ESTADOS PCODEST ON PAC.COD_ESTADO = PCODEST.COD_ESTADO AND PCODEST.COD_SISTEMA = '011'
                  LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS PEMP ON PAC.NRO_ID_RESPONSABLE = PEMP.EMPLID
                  LEFT JOIN ADMINSIIO.TR_AVANCES AVA ON PAC.COD_EVENTO = AVA.COD_EVENTO AND PAC.COD_RECOMENDACION = AVA.COD_RECOMENDACION AND PAC.COD_SISTEMA = AVA.COD_SISTEMA AND PAC.COD_PLAN = AVA.COD_PLAN
                WHERE 
                ";
            if (!string.IsNullOrWhiteSpace(paramCodAccidente))
                sqlQuery += paramCodAccidente;
            else
            {
                //Parámetros de Accidente
                if (!string.IsNullOrWhiteSpace(paramFechaInicialAccidente))
                    sqlQuery += paramFechaInicialAccidente;
                if (!string.IsNullOrWhiteSpace(paramFechaFinalAccidente))
                    sqlQuery += paramFechaFinalAccidente;
                if (!string.IsNullOrWhiteSpace(paramCodSuperIntendenciaAccidente))
                    sqlQuery += paramCodSuperIntendenciaAccidente;
                if (!string.IsNullOrWhiteSpace(paramEstadoAccidente))
                    sqlQuery += paramEstadoAccidente;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableAccidente))
                    sqlQuery += paramIdResponsableAccidente;
                if (!string.IsNullOrWhiteSpace(paramPotencialAccidente))
                    sqlQuery += paramPotencialAccidente;
                //Parámetros de Recomendaciones
                if (!string.IsNullOrWhiteSpace(paramCodEstadoRecomendacion))
                    sqlQuery += paramCodEstadoRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableRecomendacion))
                    sqlQuery += paramIdResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramCodDepartamentoResponsableRecomendacion))
                    sqlQuery += paramCodDepartamentoResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaCierreRecomendacion))
                    sqlQuery += paramFechaCierreRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaPlaneadaRecomendacion))
                    sqlQuery += paramFechaPlaneadaRecomendacion;
                //Parámetros de Planes de Acción
                if (!string.IsNullOrWhiteSpace(paramCodEstadoPlan))
                    sqlQuery += paramCodEstadoPlan;
                if (!string.IsNullOrWhiteSpace(paramIdResponsablePlan))
                    sqlQuery += paramIdResponsablePlan;
                if (!string.IsNullOrWhiteSpace(paramFechaPlaneadaPlan))
                    sqlQuery += paramFechaPlaneadaPlan;
            }

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }

        
    
        public static string GetAccidentsAllQuery(List<KeyValuePair<int?, string>> searchParameters)
        {
            //Parámetros de Accidentes
            string paramCodAccidente = null, paramFechaInicialAccidente = null, paramFechaFinalAccidente = null,
                paramCodSuperIntendenciaAccidente = null, paramEstadoAccidente = null,
                paramIdResponsableAccidente = null, paramPotencialAccidente = null;
            //Parámetros de Recomendaciones
            string paramCodEstadoRecomendacion = null, paramIdResponsableRecomendacion = null, paramCodDepartamentoResponsableRecomendacion = null,
                paramFechaCierreRecomendacion = null, paramFechaPlaneadaRecomendacion = null;
            //Parámetros de Planes de Acción
            string paramIdResponsablePlan = null, paramCodEstadoPlan = null, paramFechaPlaneadaPlan = null;

            foreach (var searchParam in searchParameters)
            {
                if (string.IsNullOrWhiteSpace(searchParam.Value))
                    continue;
                switch (searchParam.Key)
                {
                    //Parámetros de Accidentes
                    case (int)IxSearchParameters.CodigoAccidente:
                        paramCodAccidente = " AND INF.COD_ACCIDENTE = '" + searchParam.Value + "'";
                        break;

                    case (int)IxSearchParameters.FechaInicialAccidente:
                        paramFechaInicialAccidente = " AND INF.FEC_CASO >= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaFinalAccidente:
                        paramFechaFinalAccidente = " AND INF.FEC_CASO <= TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.CodigoSuperIntendenciaAccidente:
                        paramCodSuperIntendenciaAccidente = " AND INF.COD_SUPERINTENDENCIA = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.EstadoAccidente:
                        if (searchParam.Value.ToUpper().Equals("CANCELADO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO = '9' ";
                        else if (searchParam.Value.ToUpper().Equals("ACTIVO"))
                            paramCodSuperIntendenciaAccidente = " AND INF.COD_ESTADO <> '9' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableAccidente:
                        paramIdResponsableAccidente = " AND INF.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.PotencialAccidente:
                        if (searchParam.Value.ToUpper().Equals("BAJO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'BAJO' ";
                        if (searchParam.Value.ToUpper().Equals("MEDIO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'MEDIO' ";
                        if (searchParam.Value.ToUpper().Equals("ALTO"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL = 'ALTO' ";
                        if (searchParam.Value.ToUpper().Equals("ICAM"))
                            paramPotencialAccidente = " AND INF.DES_POTENCIAL IN ('MEDIO', 'ALTO') ";
                        break;

                    //Parámetros de Recomendaciones
                    case (int)IxSearchParameters.CodigoEstadoRecomendacion:
                        paramCodEstadoRecomendacion = " AND REC.COD_ESTADO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.IdResponsableRecomendacion:
                        paramIdResponsableRecomendacion = " AND REC.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.CodigoDepartamentoRecomendacion:
                        paramCodDepartamentoResponsableRecomendacion = " AND REC.COD_DEPARTAMENTO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.FechaCreacionRecomendacion:
                        paramFechaCierreRecomendacion = " AND REC.FEC_CIERRE = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    case (int)IxSearchParameters.FechaPlaneadaRecomendacion:
                        paramFechaPlaneadaRecomendacion = " AND REC.FEC_PLANEADA = = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    //Parámetros de Planes de Acción
                    case (int)IxSearchParameters.CodigoEstadoPlan:
                        paramCodEstadoPlan = " AND PAC.COD_ESTADO = '" + MyUtilities.GetCodeKey(searchParam.Value) + "' ";
                        break;

                    case (int)IxSearchParameters.IdResponsablePlan:
                        paramIdResponsablePlan = " AND PAC.NRO_ID_RESPONSABLE = '" + searchParam.Value + "' ";
                        break;

                    case (int)IxSearchParameters.FechaPlaneadaPlan:
                        paramFechaPlaneadaPlan = " AND PAC.FEC_PLANEADA = = TO_DATE('" + searchParam.Value + "', 'YYYYMMDD') ";
                        break;

                    default:
                        break;
                }
            }

            var sqlQuery = @" 
                SELECT 
                  INF.COD_ACCIDENTE Codigo_Acc, REC.COD_RECOMENDACION Codigo_Recomend, PAC.COD_PLAN Codigo_Plan, 
                  INF.COD_ESTADO Codigo_Estado_Acc,
                  (CASE WHEN INF.COD_ESTADO= 9 THEN 'Cancelado'  ELSE 'Activo' END) Desc_Estado_Acc,
                  TO_CHAR(INF.FEC_CASO, 'YYYYMMDD')  Fecha_Caso_Acc, INF.HOR_CASO Hora_Caso_Acc, 
                  TO_CHAR(INF.FEC_TURNO, 'YYYYMMDD')  Fecha_Turno_Acc, INF.COD_TURNO Codigo_Turno_Acc,
                  INF.NOM_REPORTA Nombre_Reporta_Acc, INF.NRO_ID_REPORTA Id_Reporta_Acc,
                  INF.USR_GENERA Usuario_Generacion_Acc, TO_CHAR(INF.FEC_GENERA, 'YYYYMMDD')  Fecha_Generacion_Acc, INF.NRO_ID_RESPONSABLE Id_Responsable_Acc, AEMP.NAME Nombre_Responsable_Acc,
                  INF.USR_MODIFICA Usuario_Modifica_Acc, TO_CHAR(INF.FEC_MODIFICA, 'YYYYMMDD')  Fecha_Modifica_Acc,
                  INF.DES_CARGO Desc_Cargo_Acc, INF.COD_EMPRESA_REP Codigo_Empresa_Rep_Acc, INF.COD_DEPARTAMENTO_REP Codigo_Departamento_Rep_Acc, 
                  INF.DES_INFORME Desc_Informe_Acc, INF.DES_COMENTARIOS Desc_Comentarios_Acc, 
                  INF.IND_LESION Ind_Lesion_Acc, INF.IND_DANO Ind_Dano_Acc, INF.IND_AMBIENTAL Ind_Ambiental_Acc, INF.IND_COMUNIDAD Ind_Communidad_Acc, INF.IND_INCENDIO Ind_Incendio_Acc,
                  INF.COD_POSIBLE_OCURRENCIA Codigo_Posible_Ocurrencia_Acc, INF.COD_POSIBLE_LESION Codigo_Posible_Lesion_Acc, INF.COD_PELIGRO_ASOCIADO Codigo_Peligro_Asociado_Acc,
                  INF.DES_POTENCIAL Desc_Potencial_Acc, 
                  INF.COD_TAREA_REALIZABA Codigo_Tarea_Realizaba_Acc, 
                  INF.COD_AREA Codigo_Area_Acc, AAREA.DES_AREA Desc_Area_Acc, INF.COD_SITIO Codigo_Sitio_Acc, ASIT.DES_TIPO Desc_Sitio_Acc, INF.COD_LOCALIZACION Codigo_Localizacion_Acc, ALOC.DES_TIPO Desc_Localizacion_Acc, INF.COD_SUPERINTENDENCIA Codigo_SuperIntendencia_Acc, ASUP.DES_SUPERINTENDENCIA Desc_SuperIntendencia_Acc,
                  INF.COD_SECCION Codigo_Seccion_Acc, TO_CHAR(INF.FEC_ENTREGA_INVESTIGACION, 'YYYYMMDD')  Fecha_Entrega_Inv_Acc, INF.IND_CONFIDENCIAL Ind_Confidencial_Acc, TO_CHAR(INF.FEC_COMITE, 'YYYYMMDD')  Fecha_Comite_Acc, 
                  INF.NRO_ID_APRUEBA_INV Id_Aprueba_Inv_Acc, INF.IND_GRAVE Ind_Grave_Acc, INF.COD_TIPO_FRCP COdigo_Tipo_Frcp_Acc, INF.IND_CERO_BARRERAS Ind_Cero_Barreras_Acc, INF.IND_ACC_REPETITIVO Ind_Accidente_Repetitivo_Acc, INF.IND_EXISTE_ESTANDAR_CONTROL Ind_Existe_Ecc_Acc, INF.IND_REPORTE_DIARIO Ind_Reporte_Diario_Acc,
                  (REC.COD_SISTEMA||'-'||REC.COD_EVENTO||'-'||REC.COD_RECOMENDACION) Id_Recomend, REC.COD_SISTEMA Codigo_Sistema,
                  REC.DES_RECOMENDACION Desc_Recomend, REC.COD_ESTADO Cod_Estado_Recomend, RCODEST.DES_ESTADO Des_Estado_Recomend, 
                  TO_CHAR(REC.FEC_CREACION, 'YYYYMMDD')  Fecha_Creacion_Recomend, REC.USR_CREACION Usuario_Creacion_Recomend, TO_CHAR(REC.FEC_PLANEADA, 'YYYYMMDD')  Fecha_Plan_Recomend, TO_CHAR(REC.FEC_CIERRE, 'YYYYMMDD')  Fecha_Cierre_Recomend,
                  REC.NRO_ID_RESPONSABLE Id_Resp_Recomend, REMP.NAME Nombre_Resp_Recomend, REC.COD_DEPARTAMENTO Cod_Departamento_Recomend, RDEP.DES_LARGA Desc_Despartamento_Recomend, REC.COD_TIPO_ERTC Codigo_Tipo_Ertc_Recomend,
                  REC.COD_DEPTID Codigo_DeptId_Recomend,
                  ADMINSIIO.LTOVC_TR_PLANES_ACC_DES_PLAN(PAC.ROWID) AS Desc_Plan,
                  PAC.USR_CREACION Usuario_Creacion_Plan, TO_CHAR(PAC.FEC_PLANEADA, 'YYYYMMDD')  Fecha_Planeada_Plan, PAC.NRO_ID_RESPONSABLE Id_Responsable_Plan, PEMP.NAME Nombre_Responsable_Plan,
                  PAC.COD_ESTADO Codigo_Estado_Plan, PCODEST.DES_ESTADO Desc_Estado_Plan, PAC.POR_AVANCE Porcentaje_Avance_Plan,
                  ADMINSIIO.LTOVC_TR_AVANCES_DES_AVANCE(AVA.ROWID) AS Desc_Avance_Plan
                FROM 
                  ADMINSIIO.AC_INFORMES INF 
                  LEFT JOIN ADMINSIIO.SYS_SUPERINTENDENCIA ASUP ON INF.COD_SUPERINTENDENCIA = ASUP.COD_SUPERINTENDENCIA 
                  LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS AEMP ON INF.NRO_ID_RESPONSABLE = AEMP.EMPLID
                  LEFT JOIN ADMINSIIO.SYS_AREAS AAREA ON INF.COD_AREA = AAREA.COD_AREA 
                  LEFT JOIN ADMINSIIO.SYS_TIPO ALOC ON INF.COD_LOCALIZACION = ALOC.COD_TIPO AND ALOC.COD_CLASE = 'LO'
                  LEFT JOIN ADMINSIIO.SYS_TIPO ASIT ON INF.COD_LOCALIZACION = ASIT.COD_TIPO AND ASIT.COD_CLASE = 'ST'
                  LEFT JOIN ADMINSIIO.TR_RECOMENDACIONES REC ON INF.COD_ACCIDENTE = REC.COD_EVENTO AND REC.COD_SISTEMA = '001'
                  LEFT JOIN ADMINSIIO.TR_ESTADOS RCODEST ON REC.COD_ESTADO = RCODEST.COD_ESTADO AND RCODEST.COD_SISTEMA = '011'
                  LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS REMP ON REC.NRO_ID_RESPONSABLE = REMP.EMPLID
                  LEFT JOIN ADMINSIIO.SYS_DEPARTAMENTOS RDEP ON REC.COD_DEPARTAMENTO = RDEP.COD_DEPARTAMENTO 
                  LEFT JOIN ADMINSIIO.TR_PLANES_ACCION PAC ON (REC.COD_SISTEMA = PAC.COD_SISTEMA AND REC.COD_EVENTO = PAC.COD_EVENTO AND REC.COD_RECOMENDACION = PAC.COD_RECOMENDACION)
                  LEFT JOIN ADMINSIIO.TR_ESTADOS PCODEST ON PAC.COD_ESTADO = PCODEST.COD_ESTADO AND PCODEST.COD_SISTEMA = '011'
                  LEFT JOIN ADMINSIIO.RH_V_EMPLEADOS PEMP ON PAC.NRO_ID_RESPONSABLE = PEMP.EMPLID
                  LEFT JOIN ADMINSIIO.TR_AVANCES AVA ON PAC.COD_EVENTO = AVA.COD_EVENTO AND PAC.COD_RECOMENDACION = AVA.COD_RECOMENDACION AND PAC.COD_SISTEMA = AVA.COD_SISTEMA AND PAC.COD_PLAN = AVA.COD_PLAN
                WHERE
                ";
            if (!string.IsNullOrWhiteSpace(paramCodAccidente))
                sqlQuery += paramCodAccidente;
            else
            {
                //Parámetros de Accidente
                if (!string.IsNullOrWhiteSpace(paramFechaInicialAccidente))
                    sqlQuery += paramFechaInicialAccidente;
                if (!string.IsNullOrWhiteSpace(paramFechaFinalAccidente))
                    sqlQuery += paramFechaFinalAccidente;
                if (!string.IsNullOrWhiteSpace(paramCodSuperIntendenciaAccidente))
                    sqlQuery += paramCodSuperIntendenciaAccidente;
                if (!string.IsNullOrWhiteSpace(paramEstadoAccidente))
                    sqlQuery += paramEstadoAccidente;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableAccidente))
                    sqlQuery += paramIdResponsableAccidente;
                if (!string.IsNullOrWhiteSpace(paramPotencialAccidente))
                    sqlQuery += paramPotencialAccidente;
                //Parámetros de Recomendaciones
                if (!string.IsNullOrWhiteSpace(paramCodEstadoRecomendacion))
                    sqlQuery += paramCodEstadoRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramIdResponsableRecomendacion))
                    sqlQuery += paramIdResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramCodDepartamentoResponsableRecomendacion))
                    sqlQuery += paramCodDepartamentoResponsableRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaCierreRecomendacion))
                    sqlQuery += paramFechaCierreRecomendacion;
                if (!string.IsNullOrWhiteSpace(paramFechaPlaneadaRecomendacion))
                    sqlQuery += paramFechaPlaneadaRecomendacion;
                //Parámetros de Planes de Acción
                if (!string.IsNullOrWhiteSpace(paramCodEstadoPlan))
                    sqlQuery += paramCodEstadoPlan;
                if (!string.IsNullOrWhiteSpace(paramIdResponsablePlan))
                    sqlQuery += paramIdResponsablePlan;
                if (!string.IsNullOrWhiteSpace(paramFechaPlaneadaPlan))
                    sqlQuery += paramFechaPlaneadaPlan;
            }

            sqlQuery = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(sqlQuery, "WHERE AND", "WHERE ");

            return sqlQuery;
        }
    }
}

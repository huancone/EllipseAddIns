using System.Collections.Generic;
using System.Linq;

namespace EllipseCommonsClassLibrary.Constants
{
    using System.Diagnostics.CodeAnalysis;

    public static class WoTypeMtType
    {
        /// <summary>
        /// Obtiene listado de objeto de Tipo de Orden vs Tipo de mantenimiento (MT Type, MT Desc, OT Type, OT Desc)
        /// </summary>
        /// <returns></returns>
        public static List<WoTypeMtTypeCode> GetWoTypeMtTypeList()
        {
            var typeList = new List<WoTypeMtTypeCode>
                {
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "CA", "CALIBRACION"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "CO", "CAMBIO DE COMPONENTE MAYOR"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "IP", "SERVICIOS E INSPECCIONES (SEIS)"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "IS", "INSPECCIONES"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "LA", "LAVADO"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "LU", "LUBRICACION"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "OH", "OVERHAUL"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "PB", "PRECIO BASE"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "RM", "REPARACION/CAMBIO DE COMPONENTE MENOR"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "RP", "REPARACIONES PROGRAMADAS"),
                    new WoTypeMtTypeCode("PE", "PREVENTIVO", "SN", "SERVICIO NO CONFORME"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AA", "ANALISIS DE ACEITES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AC", "ANALISIS DE COMBUSTIBLE"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "AV", "ANALISIS DE VIBRACIONES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "BC", "BASADA EN CONDICION"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "BK", "PRUEBA BAKER"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "DU", "DETECCION ULTRASONICA"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "ET", "CORRIENTES DE EDDY"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "IE", "INSPECCION ESTRUCTURAL"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "IR", "INSPECCION TERMOGRAFICA"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "ME", "MEDICIONES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "MT", "PARTICULAS MAGNETICAS"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "P0", "ANÁLISIS REFRIGERANTE"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "PT", "TINTAS PENETRANTES"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "UT", "ULTRASONIDO"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "VI", "VIDEO"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "VT", "INSPECCION VISUAL"),
                    new WoTypeMtTypeCode("PD", "PREDICTIVO", "WR", "WINDROCK"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "RE", "REPARACION"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "A", "ACCIDENTE"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "AT", "ATENTADO"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "CO", "CAMBIO DE COMPONENTE MAYOR"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "DO", "DAÑO OPERACIONAL"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "EV", "EVENTO DE BASEMAN"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "G", "GARANTIA"),
                    new WoTypeMtTypeCode("CO", "A LA FALLA - CORRECTIVO", "SE", "SERVICIO IMIS"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "ET", "ESTUDIO TECNICO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "FA", "FABRICACION"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "MC", "CAMBIO DE EQUIPO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "MN", "MONTAJE NUEVO"),
                    new WoTypeMtTypeCode("PT", "PROACTIVO", "RD", "REDISEÑO O MODIFICACIONES"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AF", "ANALISIS DE FALLAS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AR", "ANALISIS DE RESULTADOS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "AS", "ACTIVIDADES  DE SIO & MEDIO AMBIENTE"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "FA", "FABRICACION"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "MP", "MOVILIZACIÓN DE COMPONENTES"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "OP", "OPERACION DE EQUIPOS"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "SL", "SIN LABOR"),
                    new WoTypeMtTypeCode("NM", "NO MANTENIMIENTO", "SM", "SOPORTE AL MANTENIMIENTO")

                };
            return typeList;
        }

        public static Dictionary<string, string> GetPriorityCodeList()
        {
            var dictionaryList = new Dictionary<string, string>
            {
                {"P0", "CRÍTICA - Detener equipo inmediatamente"},
                {"P1", "URGENTE - Programar en ventana en curso"},
                {"P2", "PRIORITARIA - Programar más tardar en ventana siguiente"},
                {"P3", "RUTINA - Programar en el próximo PM"},
                {"P4", "RUTINA - Programar según oportunidad"},
                {"BE", "INST/IMIS - EMERGENCIA - Atención 1h Cierre 7 días"},
                {"B1", "INST/IMIS - ALTA - Atención 48h Cierre 7 días"},
                {"B2", "INST/IMIS - NORMAL - Atención 6 días Cierre 15 días"},
                {"B3", "INST/IMIS - BAJA - Atención 9 días cierre 30 días"},
            };

            return dictionaryList;
        }
        /// <summary>
        /// Obtiene arreglo Dictionary{key, value} con listado de los códigos de Tipo de Orden admitidos {codigo, descripcion}
        /// </summary>
        /// <returns></returns>
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here.")]
        public static Dictionary<string, string> GetWoTypeList()
        {
            var listType = GetWoTypeMtTypeList();

            var woTypeList = new Dictionary<string, string>();
            foreach (var type in listType.Where(type => !woTypeList.ContainsKey(type.WoTypeCode)))
            {
                woTypeList.Add(type.WoTypeCode, type.WoTypeDesc);
            }
            return woTypeList;
        }


        /// <summary>
        /// Obtiene arreglo Dictionary{key, value} con listado de los códigos de Tipo de Mantenimiento admitidos {codigo, descripcion}
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetMtTypeList()
        {
            var listType = GetWoTypeMtTypeList();

            var mtTypeList = new Dictionary<string, string>();
            foreach (var type in listType.Where(type => !mtTypeList.ContainsKey(type.MtTypeCode)))
            {
                mtTypeList.Add(type.MtTypeCode, type.MtTypeDesc);
            }
            return mtTypeList;
        }
        /// <summary>
        /// Valida la prioridad de una orden/std establecida para MDC
        /// </summary>
        /// <param name="priority">string: código de prioridad</param>
        /// <param name="district">string: distrito al que pertenece la orden-std</param>
        /// <param name="workGroup">string: grupo de trabajo</param>
        /// <returns>true si la prioridad es válida, false si no es válida</returns>
        public static bool ValidatePriority(string priority, string district = null, string workGroup = null)
        {
            if (priority == null)
                return false;

            priority = priority.Trim();

            if (district == null || district.Trim().Equals("ICOR"))
            {
                if (priority.Equals("P0") || priority.Equals("P1") || priority.Equals("P2") || priority.Equals("P3") || priority.Equals("P4"))
                    return true;
            }
            else if (district.Trim().Equals("INST"))
            {
                if (workGroup != null && (workGroup.Trim().Equals("AAPREV") && (priority.Equals("P0") || priority.Equals("P1") || priority.Equals("P2") || priority.Equals("P3") || priority.Equals("P4"))))
                    return true;
                if (priority.Equals("B1") || priority.Equals("B2") || priority.Equals("B3"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Valida la relación de Tipo de Orden vs Tipo de Mantenimiento de una orden/std establecida para MDC
        /// </summary>
        /// <param name="woType">string: Tipo de Orden</param>
        /// <param name="mtType">string: Tipo de Mantenimiento</param>
        /// <returns>true si la relación es válida, false si no es válida</returns>
        public static bool ValidateWoMtTypeCode(string woType, string mtType)
        {
            if (woType == null || mtType == null)
                return false;

            woType = woType.Trim();
            mtType = mtType.Trim();

            var typeList = GetWoTypeMtTypeList();

            return typeList.Any(type => woType == type.WoTypeCode && mtType == type.MtTypeCode);
        }

        public class WoTypeMtTypeCode
        {
            public string MtTypeCode;
            public string MtTypeDesc;
            public string WoTypeCode;
            public string WoTypeDesc;

            public WoTypeMtTypeCode(string mtTypeCode, string mtTypeDesc, string woTypeCode, string woTypeDesc)
            {
                MtTypeCode = mtTypeCode;
                MtTypeDesc = mtTypeDesc;
                WoTypeCode = woTypeCode;
                WoTypeDesc = woTypeDesc;
            }
        }


    }
    
}

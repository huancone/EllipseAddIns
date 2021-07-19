using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InvestigacionesIcam
{
    public class AccRecPlan
    {
        public string CodigoAccidente;
        public string CodigoEstadoAccidente;
        public string EstadoAccidente;
        public string FechaCasoAccidente;
        public string HoraCasoAccidente;
        public string FechaTurnoAccidente;
        public string CodigoTurnoAccidente;
        public string NombreReportaAccidente;
        public string IdReportaAccidente;
        public string UsuarioCreacionAccidente;
        public string FechaCreacionAccidente;
        public string IdResponsableAccidente;
        public string NombreResponsableAccidente;
        public string UsuarioModificadorAccidente;
        public string FechaModificadorAccidente;
        public string DescripcionCargoAccidente;
        public string CodigoEmpresaRepresentanteAccidente;
        public string CodigoDepartamentoRepresentanteAccidente;
        public string DescripcionInformeAccidente;
        public string ComentariosAccidente;
        public string IndicadorLesionAccidente;
        public string IndicadorDañoAccidente;
        public string IndicadorAmbientalAccidente;
        public string IndicadorComunidadAccidente;
        public string IndicadorIncendioAccidente;
        public string IndicadorCeroBarrerasAccidente;
        public string IndicadorAccidenteRepetitivoAccidente;
        public string IndicadorExisteEccAccidente;
        public string IndicadorReporteDiarioAccidente;
        public string IndicadorGraveAccidente;
        public string IndicadorConfindencialAccidente;
        public string CodigoPosibleOcurrenciaAccidente;
        public string CodigoPosibleLesionAccidente;
        public string CodigoPeligroAsociadoAccidente;
        public string PotencialAccidente;
        public string CodigoTareaRealizadaAccidente;
        public string CodigoAreaAccidente;
        public string AreaAccidente;
        public string CodigoSitioAccidente;
        public string SitioAccidente;
        public string CodigoLocalizacionAccidente;
        public string LocalizacionAccidente;
        public string CodigoSuperIntendenciaAccidente;
        public string SuperIntendenciaAccidente;
        public string CodigoSeccionAccidente;
        public string FechaEntregaInvestigaciónAccidente;
        public string FechaComiteAccidente;
        public string IdApruebaInvestigacionAccidente;
        public string CodigoTipoFrcpAccidente;

        public string IdRecomendacion;
        public string CodigoSistema;
        //public string CodigoAccidente;
        public string CodigoRecomendacion;
        public string DescripcionRecomendacion;
        public string CodigoEstadoRecomendacion;
        public string EstadoRecomendacion;
        public string FechaCreacionRecomendacion;
        public string UsuarioCreacionRecomendacion;
        public string FechaPlaneadaRecomendacion;
        public string FechaCierreRecomendacion;
        public string IdResponsableRecomendacion;
        public string NombreResponsableRecomendacion;
        public string CodigoDepartamentoRecomendacion;
        public string DepartamentoRecomendacion;
        public string CodigoTipoErtcRecomendacion;
        public string CodigoDepIdRecomendacion;

        //public string IdRecomendacion;
        //public string CodigoSistema;
        //public string CodigoAccidente;
        //public string CodigoRecomendacion;
        public string CodigoPlan;
        public string DescripcionPlan;
        public string UsuarioCreacionPlan;
        public string FechaPlaneadaPlan;
        public string IdResponsablePlan;
        public string NombreResponsablePlan;
        public string CodigoEstadoPlan;
        public string EstadoPlan;
        public string PorcentajeAvancePlan;
        public string AvancePlan;

        public AccRecPlan()
        {
        }

        public AccRecPlan(IDataRecord dr)
        {

            CodigoAccidente = dr["Codigo_Acc"].ToString().Trim();
            CodigoEstadoAccidente = dr["Codigo_Estado_Acc"].ToString().Trim();
            EstadoAccidente = dr["Desc_Estado_Acc"].ToString().Trim();
            FechaCasoAccidente = dr["Fecha_Caso_Acc"].ToString().Trim();
            HoraCasoAccidente = dr["Hora_Caso_Acc"].ToString().Trim();
            FechaTurnoAccidente = dr["Fecha_Turno_Acc"].ToString().Trim();
            CodigoTurnoAccidente = dr["Codigo_Turno_Acc"].ToString().Trim();
            NombreReportaAccidente = dr["Nombre_Reporta_Acc"].ToString().Trim();
            IdReportaAccidente = dr["Id_Reporta_Acc"].ToString().Trim();
            UsuarioCreacionAccidente = dr["Usuario_Generacion_Acc"].ToString().Trim();
            FechaCreacionAccidente = dr["Fecha_Generacion_Acc"].ToString().Trim();
            IdResponsableAccidente = dr["Id_Responsable_Acc"].ToString().Trim();
            NombreResponsableAccidente = dr["Nombre_Responsable_Acc"].ToString().Trim();
            UsuarioModificadorAccidente = dr["Usuario_Modifica_Acc"].ToString().Trim();
            FechaModificadorAccidente = dr["Fecha_Modifica_Acc"].ToString().Trim();
            DescripcionCargoAccidente = dr["Desc_Cargo_Acc"].ToString().Trim();
            CodigoEmpresaRepresentanteAccidente = dr["Codigo_Empresa_Rep_Acc"].ToString().Trim();
            CodigoDepartamentoRepresentanteAccidente = dr["Codigo_Departamento_Rep_Acc"].ToString().Trim();
            DescripcionInformeAccidente = dr["Desc_Informe_Acc"].ToString().Trim();
            ComentariosAccidente = dr["Desc_Comentarios_Acc"].ToString().Trim();
            IndicadorLesionAccidente = dr["Ind_Lesion_Acc"].ToString().Trim();
            IndicadorDañoAccidente = dr["Ind_Dano_Acc"].ToString().Trim();
            IndicadorAmbientalAccidente = dr["Ind_Ambiental_Acc"].ToString().Trim();
            IndicadorComunidadAccidente = dr["Ind_Communidad_Acc"].ToString().Trim();
            IndicadorIncendioAccidente = dr["Ind_Incendio_Acc"].ToString().Trim();

            IndicadorCeroBarrerasAccidente = dr["Ind_Cero_Barreras_Acc"].ToString().Trim();
            IndicadorAccidenteRepetitivoAccidente = dr["Ind_Accidente_Repetitivo_Acc"].ToString().Trim();
            IndicadorExisteEccAccidente = dr["Ind_Existe_Ecc_Acc"].ToString().Trim();
            IndicadorReporteDiarioAccidente = dr["Ind_Reporte_Diario_Acc"].ToString().Trim();
            IndicadorGraveAccidente = dr["Ind_Grave_Acc"].ToString().Trim();
            IndicadorConfindencialAccidente = dr["Ind_Confidencial_Acc"].ToString().Trim();
            CodigoPosibleOcurrenciaAccidente = dr["Codigo_Posible_Ocurrencia_Acc"].ToString().Trim();

            CodigoPosibleLesionAccidente = dr["Codigo_Posible_Lesion_Acc"].ToString().Trim();
            CodigoPeligroAsociadoAccidente = dr["Codigo_Peligro_Asociado_Acc"].ToString().Trim();
            PotencialAccidente = dr["Desc_Potencial_Acc"].ToString().Trim();
            CodigoTareaRealizadaAccidente = dr["Codigo_Tarea_Realizaba_Acc"].ToString().Trim();
            CodigoAreaAccidente = dr["Codigo_Area_Acc"].ToString().Trim();
            AreaAccidente = dr["Desc_Area_Acc"].ToString().Trim();
            CodigoSitioAccidente = dr["Codigo_Sitio_Acc"].ToString().Trim();
            SitioAccidente = dr["Desc_Sitio_Acc"].ToString().Trim();
            CodigoLocalizacionAccidente = dr["Codigo_Localizacion_Acc"].ToString().Trim();
            LocalizacionAccidente = dr["Desc_Localizacion_Acc"].ToString().Trim();
            CodigoSuperIntendenciaAccidente = dr["Codigo_SuperIntendencia_Acc"].ToString().Trim();
            SuperIntendenciaAccidente = dr["Desc_SuperIntendencia_Acc"].ToString().Trim();
            CodigoSeccionAccidente = dr["Codigo_Seccion_Acc"].ToString().Trim();
            FechaEntregaInvestigaciónAccidente = dr["Fecha_Entrega_Inv_Acc"].ToString().Trim();
            FechaComiteAccidente = dr["Fecha_Comite_Acc"].ToString().Trim();
            IdApruebaInvestigacionAccidente = dr["Id_Aprueba_Inv_Acc"].ToString().Trim();
            CodigoTipoFrcpAccidente = dr["COdigo_Tipo_Frcp_Acc"].ToString().Trim();

            IdRecomendacion = dr["Id_Recomend"].ToString().Trim();
            CodigoSistema = dr["Codigo_Sistema"].ToString().Trim();
            CodigoRecomendacion = dr["Codigo_Recomend"].ToString().Trim();
            DescripcionRecomendacion = dr["Desc_Recomend"].ToString().Trim();
            CodigoEstadoRecomendacion = dr["Cod_Estado_Recomend"].ToString().Trim();
            EstadoRecomendacion = dr["Des_Estado_Recomend"].ToString().Trim();
            FechaCreacionRecomendacion = dr["Fecha_Creacion_Recomend"].ToString().Trim();
            UsuarioCreacionRecomendacion = dr["Usuario_Creacion_Recomend"].ToString().Trim();
            FechaPlaneadaRecomendacion = dr["Fecha_Plan_Recomend"].ToString().Trim();
            FechaCierreRecomendacion = dr["Fecha_Cierre_Recomend"].ToString().Trim();
            IdResponsableRecomendacion = dr["Id_Resp_Recomend"].ToString().Trim();
            NombreResponsableRecomendacion = dr["Nombre_Resp_Recomend"].ToString().Trim();
            CodigoDepartamentoRecomendacion = dr["Cod_Departamento_Recomend"].ToString().Trim();
            DepartamentoRecomendacion = dr["Desc_Despartamento_Recomend"].ToString().Trim();
            CodigoTipoErtcRecomendacion = dr["Codigo_Tipo_Ertc_Recomend"].ToString().Trim();
            CodigoDepIdRecomendacion = dr["Codigo_DeptId_Recomend"].ToString().Trim();

            CodigoPlan = dr["Codigo_Plan"].ToString().Trim();
            DescripcionPlan = dr["Desc_Plan"].ToString().Trim();
            UsuarioCreacionPlan = dr["Usuario_Creacion_Plan"].ToString().Trim();
            FechaPlaneadaPlan = dr["Fecha_Planeada_Plan"].ToString().Trim();
            IdResponsablePlan = dr["Id_Responsable_Plan"].ToString().Trim();
            NombreResponsablePlan = dr["Nombre_Responsable_Plan"].ToString().Trim();
            CodigoEstadoPlan = dr["Codigo_Estado_Plan"].ToString().Trim();
            EstadoPlan = dr["Desc_Estado_Plan"].ToString().Trim();
            PorcentajeAvancePlan = dr["Porcentaje_Avance_Plan"].ToString().Trim();
            AvancePlan = dr["Desc_Avance_Plan"].ToString().Trim();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SharedClassLibrary.Utilities;

namespace InvestigacionesIcam
{
    public class Accidente
    {
        public string CodigoAccidente;
        public string CodigoEstado;
        public string Estado;
        public string FechaCaso;
        public string HoraCaso;
        public string FechaTurno;
        public string CodigoTurno;
        public string NombreReporta;
        public string IdReporta;
        public string UsuarioCreacion;
        public string FechaCreacion;
        public string IdResponsable;
        public string NombreResponsable;
        public string UsuarioModificador;
        public string FechaModificador;
        public string DescripcionCargo;
        public string CodigoEmpresaRepresentante;
        public string CodigoDepartamentoRepresentante;
        public string DescripcionInforme;
        public string Comentarios;
        public string IndicadorLesion;
        public string IndicadorDaño;
        public string IndicadorAmbiental;
        public string IndicadorComunidad;
        public string IndicadorIncendio;
        public string CodigoPosibleOcurrencia;
        public string CodigoPosibleLesion;
        public string CodigoPeligroAsociado;
        public string Potencial;
        public string CodigoTareaRealizada;
        public string CodigoArea;
        public string Area;
        public string CodigoSitio;
        public string Sitio;
        public string CodigoLocalizacion;
        public string Localizacion;
        public string CodigoSuperIntendencia;
        public string SuperIntendencia;
        public string CodigoSeccion;
        public string FechaEntregaInvestigación;
        public string IndicadorConfindencial;
        public string FechaComite;
        public string IdApruebaInvestigacion;
        public string IndicadorGrave;
        public string CodigoTipoFrcp;
        public string IndicadorCeroBarreras;
        public string IndicadorAccidenteRepetitivo;
        public string IndicadorExisteEcc;
        public string IndicadorReporteDiario;

        public Accidente()
        {
        }
        public Accidente(IDataRecord dr)
        {
            CodigoAccidente = dr["COD_ACCIDENTE"].ToString().Trim();
            CodigoEstado = dr["COD_ESTADO"].ToString().Trim();
            Estado = dr["DESC_ESTADO"].ToString().Trim();
            FechaCaso = dr["FEC_CASO"].ToString().Trim();
            HoraCaso = dr["HOR_CASO"].ToString().Trim();
            FechaTurno = dr["FEC_TURNO"].ToString().Trim();
            CodigoTurno = dr["COD_TURNO"].ToString().Trim();
            NombreReporta = dr["NOM_REPORTA"].ToString().Trim();
            IdReporta = dr["NRO_ID_REPORTA"].ToString().Trim();
            UsuarioCreacion = dr["USR_GENERA"].ToString().Trim();
            FechaCreacion = dr["FEC_GENERA"].ToString().Trim();
            IdResponsable = dr["NRO_ID_RESPONSABLE"].ToString().Trim();
            NombreResponsable = dr["NOMBRE_RESPONSABLE"].ToString().Trim();
            UsuarioModificador = dr["USR_MODIFICA"].ToString().Trim();
            FechaModificador = dr["FEC_MODIFICA"].ToString().Trim();
            DescripcionCargo = dr["DES_CARGO"].ToString().Trim();
            CodigoEmpresaRepresentante = dr["COD_EMPRESA_REP"].ToString().Trim();
            CodigoDepartamentoRepresentante = dr["COD_DEPARTAMENTO_REP"].ToString().Trim();
            DescripcionInforme = dr["DES_INFORME"].ToString().Trim();
            Comentarios = dr["DES_COMENTARIOS"].ToString().Trim();
            IndicadorLesion = dr["IND_LESION"].ToString().Trim();
            IndicadorDaño = dr["IND_DANO"].ToString().Trim();
            IndicadorAmbiental = dr["IND_AMBIENTAL"].ToString().Trim();
            IndicadorComunidad = dr["IND_COMUNIDAD"].ToString().Trim();
            IndicadorIncendio = dr["IND_INCENDIO"].ToString().Trim();
            CodigoPosibleOcurrencia = dr["COD_POSIBLE_OCURRENCIA"].ToString().Trim();
            CodigoPosibleLesion = dr["COD_POSIBLE_LESION"].ToString().Trim();
            CodigoPeligroAsociado = dr["COD_PELIGRO_ASOCIADO"].ToString().Trim();
            Potencial = dr["DES_POTENCIAL"].ToString().Trim();
            CodigoTareaRealizada = dr["COD_TAREA_REALIZABA"].ToString().Trim();
            CodigoArea = dr["COD_AREA"].ToString().Trim();
            Area = dr["DES_AREA"].ToString().Trim();
            CodigoSitio = dr["COD_SITIO"].ToString().Trim();
            Sitio = dr["DESC_SITIO"].ToString().Trim();
            CodigoLocalizacion = dr["COD_LOCALIZACION"].ToString().Trim();
            Localizacion = dr["DESC_LOCALIZACION"].ToString().Trim();
            CodigoSuperIntendencia = dr["COD_SUPERINTENDENCIA"].ToString().Trim();
            SuperIntendencia = dr["DES_SUPERINTENDENCIA"].ToString().Trim();
            CodigoSeccion = dr["COD_SECCION"].ToString().Trim();
            FechaEntregaInvestigación = dr["FEC_ENTREGA_INVESTIGACION"].ToString().Trim();
            IndicadorConfindencial = dr["IND_CONFIDENCIAL"].ToString().Trim();
            FechaComite = dr["FEC_COMITE"].ToString().Trim();
            IdApruebaInvestigacion = dr["NRO_ID_APRUEBA_INV"].ToString().Trim();
            IndicadorGrave = dr["IND_GRAVE"].ToString().Trim();
            CodigoTipoFrcp = dr["COD_TIPO_FRCP"].ToString().Trim();
            IndicadorCeroBarreras = dr["IND_CERO_BARRERAS"].ToString().Trim();
            IndicadorAccidenteRepetitivo = dr["IND_ACC_REPETITIVO"].ToString().Trim();
            IndicadorExisteEcc = dr["IND_EXISTE_ESTANDAR_CONTROL"].ToString().Trim();
            IndicadorReporteDiario = dr["IND_REPORTE_DIARIO"].ToString().Trim();
        }
    }
}

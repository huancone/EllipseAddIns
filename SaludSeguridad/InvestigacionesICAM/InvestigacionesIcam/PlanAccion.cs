using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InvestigacionesIcam
{
    public class PlanAccion
    {
        public string IdRecomendacion;
        public string CodigoSistema;
        public string CodigoAccidente;
        public string CodigoRecomendacion;
        public string CodigoPlan;
        public string Descripcion;
        public string UsuarioCreacion;
        public string FechaPlaneada;
        public string IdResponsable;
        public string NombreResponsable;
        public string CodigoEstado;
        public string Estado;
        public string PorcentajeAvance;
        public string Avance;

        public PlanAccion()
        {

        }
        public PlanAccion(IDataRecord dr)
        {
            IdRecomendacion = dr["ID_RECOMENDACION"].ToString().Trim();
            CodigoSistema = dr["COD_SISTEMA"].ToString().Trim();
            CodigoAccidente = dr["COD_EVENTO"].ToString().Trim();
            CodigoRecomendacion = dr["COD_RECOMENDACION"].ToString().Trim();
            CodigoPlan = dr["COD_PLAN"].ToString().Trim();
            Descripcion = dr["DES_PLAN"].ToString().Trim();
            UsuarioCreacion = dr["USR_CREACION"].ToString().Trim();
            FechaPlaneada = dr["FEC_PLANEADA"].ToString().Trim();
            IdResponsable = dr["NRO_ID_RESPONSABLE"].ToString().Trim();
            NombreResponsable = dr["NOMBRE_RESPONSABLE"].ToString().Trim();
            CodigoEstado = dr["COD_ESTADO"].ToString().Trim();
            Estado = dr["DES_ESTADO"].ToString().Trim();
            PorcentajeAvance = dr["POR_AVANCE"].ToString().Trim();
            Avance = dr["DES_AVANCE"].ToString().Trim();
        }
    }
}

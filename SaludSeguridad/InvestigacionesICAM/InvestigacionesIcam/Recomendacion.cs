using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InvestigacionesIcam
{
    public class Recomendacion
    {
        public string IdRecomendacion;
        public string CodigoSistema;
        public string CodigoAccidente;
        public string CodigoRecomendacion;
        public string Descripcion;
        public string CodigoEstado;
        public string Estado;
        public string FechaCreacion;
        public string UsuarioCreacion;
        public string FechaPlaneada;
        public string FechaCierre;
        public string IdResponsable;
        public string NombreResponsable;
        public string CodigoDepartamento;
        public string Departamento;
        public string CodigoTipoErtc;
        public string CodigoDepId;

        public Recomendacion()
        {

        }
        public Recomendacion(IDataRecord dr)
        {
            IdRecomendacion = dr["ID_RECOMENDACION"].ToString().Trim();
            CodigoSistema = dr["COD_SISTEMA"].ToString().Trim();
            CodigoAccidente = dr["COD_EVENTO"].ToString().Trim();
            CodigoRecomendacion = dr["COD_RECOMENDACION"].ToString().Trim();
            Descripcion = dr["DES_RECOMENDACION"].ToString().Trim();
            CodigoEstado = dr["COD_ESTADO"].ToString().Trim();
            Estado = dr["DES_ESTADO"].ToString().Trim();
            FechaCreacion = dr["FEC_CREACION"].ToString().Trim();
            UsuarioCreacion = dr["USR_CREACION"].ToString().Trim();
            FechaPlaneada = dr["FEC_PLANEADA"].ToString().Trim();
            FechaCierre = dr["FEC_CIERRE"].ToString().Trim();
            IdResponsable = dr["NRO_ID_RESPONSABLE"].ToString().Trim();
            NombreResponsable = dr["NOMBRE_RESPONSABLE"].ToString().Trim();
            CodigoDepartamento = dr["COD_DEPARTAMENTO"].ToString().Trim();
            Departamento = dr["DESC_DEPARTAMENTO"].ToString().Trim();
            CodigoTipoErtc = dr["COD_TIPO_ERTC"].ToString().Trim();
            CodigoDepId = dr["COD_DEPTID"].ToString().Trim();

        }
    }
}

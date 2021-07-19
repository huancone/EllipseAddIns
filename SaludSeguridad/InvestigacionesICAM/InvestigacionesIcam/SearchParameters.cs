using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvestigacionesIcam
{
    public enum IxSearchParameters
    {
        [Description("Código")] CodigoAccidente,
        [Description("Fecha Inicial Accidente")] FechaInicialAccidente,
        [Description("Fecha Final Accidente")] FechaFinalAccidente,
        [Description("Estado Accidente")] EstadoAccidente,
        //[Description("Id Reporta Accidente")] IdReportaAccidente,
        //[Description("Usuario Creacion Accidente")] UsuarioCreacionAccidente,
        //[Description("Fecha Creacion Accidente")] FechaCreacionAccidente,
        [Description("Id Responsable Accidente")] IdResponsableAccidente,
        [Description("Potencial Accidente")] PotencialAccidente,
        //[Description("Codigo Area Accidente")] CodigoAreaAccidente,
        //[Description("Codigo Sitio Accidente")] CodigoSitioAccidente,
        //[Description("Codigo Localizacion Accidente")] CodigoLocalizacionAccidente,
        [Description("Codigo SuperIntendencia Accidente")] CodigoSuperIntendenciaAccidente,
        //[Description("Codigo Seccion Accidente")] CodigoSeccionAccidente,
        //[Description("Id Aprueba Investigacion")] IdApruebaInvestigacion,
        //[Description("Fecha Comite Accidente")] FechaComiteAccidente,
        //
        [Description("Codigo Estado Recomendacion")] CodigoEstadoRecomendacion,
        [Description("Fecha Creacion Recomendacion")] FechaCreacionRecomendacion,
        //[Description("Usuario Creacion Recomendacion")] UsuarioCreacionRecomendacion,
        [Description("Fecha Planeada Recomendacion")] FechaPlaneadaRecomendacion,
        //[Description("Fecha Cierre Recomendacion")] FechaCierreRecomendacion,
        [Description("Id Responsable Recomendacion")] IdResponsableRecomendacion,
        [Description("Codigo Departamento Recomendacion")] CodigoDepartamentoRecomendacion,
        //
        //[Description("Usuario Creacion Plan")]UsuarioCreacionPlan,
        [Description("Fecha Planeada Plan")] FechaPlaneadaPlan,
        [Description("Id Responsable Plan")] IdResponsablePlan,
        [Description("Codigo Estado Plan")] CodigoEstadoPlan,


    }

    public static class SearchParameters
    {
        public static string GetDescription(this Enum value)
        {
            var field = value.GetType().GetField(value.ToString());

            DescriptionAttribute attribute
                    = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute))
                        as DescriptionAttribute;

            return attribute == null ? value.ToString() : attribute.Description;
        }

        public static object GetValueFromDescription(string description, bool ignoreError = true)
        {
            foreach (var value in typeof(IxSearchParameters).GetEnumValues())
            {
                var field = value.GetType().GetField(value.ToString());

                DescriptionAttribute attribute
                        = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute))
                            as DescriptionAttribute;

                if (attribute.Description.ToUpper().Equals(description.ToUpper()))
                    return value;
            }

            if (ignoreError)
                return null;
            else
                throw new ArgumentException("Not found: ", nameof(description));
        }
    }
}

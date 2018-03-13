using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseCommonsClassLibrary.Constants
{
    public static class Groups
    {
        public static List<WorkGroup> GetWorkGroupList()
        {
            var groupList = new List<WorkGroup>
            {
                new WorkGroup("AAPREV", "Mantenimiento Aire Acondicionado", "INST", "MINA", "INST"),
                new WorkGroup("BASE9", "OPERACION Y ATENCION BASE 9", ManagementArea.SoporteOperacion.Key, "ENERGIA", "ICOR"),
                new WorkGroup("CALLCEN", "Call Center", "INST", "IMIS", "INST"),
                new WorkGroup("CARGUE2", "Cargue 2", ManagementArea.Mantenimiento.Key, QuarterMasters.AcarreoElectrico.Key, "ICOR"),
                new WorkGroup("CAT2401", "U.A.S. CAMIONES CAT240", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("CAT789C", "Camion 190 ton cat789C", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("CTC", "INSPECCIONES VIAS Y MANTTO. DEL CTC", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR"),
                new WorkGroup("EH320", "U.A.S. CAMIONES DE 320 MINA NORTE", ManagementArea.Mantenimiento.Key, QuarterMasters.AcarreoElectrico.Key, "ICOR"),
                new WorkGroup("ELIVIA1", "GRUPO DE TRABAJO DE LIVIANOS", ManagementArea.SoporteOperacion.Key, "LIVIANOS", "ICOR"),
                new WorkGroup("EMEDIA1", "GRUPO DE TRABAJO MEDIANOS", ManagementArea.SoporteOperacion.Key, "MEDIANOS", "ICOR"),
                new WorkGroup("EQAUXV", "MTTO.EQUIPO VIAS FFCC", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR"),
                new WorkGroup("GI&T", "Grupo de Inspección y Tecnología", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("GRUAS", "UAS GRUAS Y MANEJADORES DE LLANTA", ManagementArea.SoporteOperacion.Key, "GRUAS", "ICOR"),
                new WorkGroup("IALIAL1", "GRUPO DE SEIS Transformadores y Distrib.", ManagementArea.SoporteOperacion.Key, "ENERGIA", "ICOR"),
                new WorkGroup("IAPTAL1", "GRUPO SEIS Taller & soporte SER", ManagementArea.SoporteOperacion.Key, "ENERGIA", "ICOR"),
                new WorkGroup("IBOMBA1", "SEIS DE BOMBAS Super. de Servicio", ManagementArea.SoporteOperacion.Key, "ENERGIA", "ICOR"),
                new WorkGroup("ICARROS", "MTTO.VAGONES", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR"),
                new WorkGroup("LLANTAS", "TALLER DE LLANTAS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("LUBRICA", "LABORES TALLER DE LUBRICACION", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("L1350", "CARGADORES LETORNEAU L1350", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("MCARGA", "UAS MANEJO DE CARGA - OPERADORES", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("MTIL17", "MANTENIMIENTO INDUSTRIAL&PLANTA-AGUA", "INST", "MINA", "ICOR"),
                new WorkGroup("MTOLOC", "MTTO. LOCOMOTORAS", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR"),
                new WorkGroup("MTTOSOP", "UAS EQUIPO DE SOPORTE", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("ORUGAS", "TRACTORES DE ORUGAS D9L Y D11N", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("PCSERVI", "PLANTA DE CARBON", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PlantasDeCarbon.Key, "ICOR"),
                new WorkGroup("PHIDCAS", "PALAS HIDRAULICAS MINA", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("PHS", "UAS PALAS ELECTRICAS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("PPELCOP", "TALLER ELECTRICO/ELECTRONICO", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOAA", "MANTENIMIENTO DE AIRES PBV", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "INST"),
                new WorkGroup("PTOBAND", "MANTTO.BANDAS TRANSPORTADORAS CARBON", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOCAR", "MANTTO MECANICO EQUIPOS DE MANCARB", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOCE", "CARGA Y ESTIBA PUERTO BOLIVAR", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOCP8", "GRUPO SEIS CONTRATO REDES Y MONTAJES", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOINS", "GRUPO DE INSPECCIONES ESTRUCTURALES PBV", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOMET", "CONTRATISTA METALISTERIA Y PINTURA PBV", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOMIN", "MANTENIMIENTO INSTALACIONES PBV", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "INST"),
                new WorkGroup("PTOOM1", "GRUPO DE MANTENIMIENTO MOTORES DIESEL", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOPRED", "GRUPO PREDICTIVOS PUERTO BOLIVAR", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOSEG", "GRUPO CONTROLES CRITICOS", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("PTOTM", "TALLER MECANICO/PLANTA AGUA -PBV", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("RDCAMPO", "GRUPO DE MANTENIMIENTO MOTORES EN CAMPO", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RDCOMPO", "REPARACIÓN DE COMPONENTES MENORES DE MOT", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RDIESEL", "RECONSTRUCCION DE MOTORES", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RECHID", "GRUPO REC.HIDRAULICA DE PRONOSTICOS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RHMENOR", "RECONSTRUIR COMP.MENORES HIDRAULICOS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RELECII", "GRUPO PARA REPARACION DE COMPO. PROGRAMA", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("REMAQ", "MAQUINAS HERRAMIENTAS RECONSTRUCCION", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RERODA", "RECONSTRUCCION TREN DE RODAJE", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("RESOLD", "RECONSTRUCCION SOLDADURA", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("T&A", "TALLER MANTTO.ROLDAN -TRAFICO & ADUANA", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.PuertoBolivar.Key, "ICOR"),
                new WorkGroup("TANQ777", "UAS DE TANQUEROS Y TRAILLAS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("TRACLLA", "UAS DE TRACTORES DE LLANTAS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("VIAS", "UAS DE MOTONIVELADORAS", ManagementArea.Mantenimiento.Key, "MINA", "ICOR"),
                new WorkGroup("VIASM", "MANTENIMIENTO VIAS MINA", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR"),
                new WorkGroup("VIASP", "MANTENIMIENTO DE VIAS PUERTO", ManagementArea.ManejoDeCarbon.Key, QuarterMasters.Ferrocarril.Key, "ICOR")
            };
            return groupList;
        }


        public class WorkGroup
        {
            public string Name;
            public string Description;
            public string Area;
            public string Details;
            public string DistrictCode;

            public WorkGroup(string name, string description, string area, string details, string districtCode)
            {
                Name = name;
                Description = description;
                Area = area;
                Details = details;
                DistrictCode = districtCode;
            }
        }


    }
    
}

using LINQtoCSV;

namespace EllipseSAO900AddIn
{
    /// <summary>
    ///     Clase usada para exportar el documento csv
    /// </summary>
    public class Reclasificaciones
    {
        [CsvColumn(Name = "ACTION", FieldIndex = 1)]
        public string Action { get; set; }

        [CsvColumn(Name = "AUTORIZADOR", FieldIndex = 2)]
        public string Autorizador { get; set; }

        [CsvColumn(Name = "DISTRITO", FieldIndex = 3)]
        public string Distrito { get; set; }

        [CsvColumn(Name = "NUM_TRANSACCION", FieldIndex = 4)]
        public string NumTransaccion { get; set; }

        [CsvColumn(Name = "CCOSTOS", FieldIndex = 5)]
        public string Centro { get; set; }

        [CsvColumn(Name = "PROJ/WO", FieldIndex = 6)]
        public string ProyectoOrden { get; set; }

        [CsvColumn(Name = "IND", FieldIndex = 7)]
        public string Indicador { get; set; }

        [CsvColumn(Name = "DOLARES", FieldIndex = 8)]
        public string Dolares { get; set; }

        [CsvColumn(Name = "PESOS", FieldIndex = 9)]
        public string Pesos { get; set; }

        [CsvColumn(Name = "CCOSTOS_DESTINO", FieldIndex = 10)]
        public string CentroDestino { get; set; }

        [CsvColumn(Name = "EQUIPO", FieldIndex = 11)]
        public string Equipo { get; set; }

        [CsvColumn(Name = "PROJ/WO_DESTINO", FieldIndex = 12)]
        public string ProyectoOrdenDestino { get; set; }

        [CsvColumn(Name = "IND_DESTINO", FieldIndex = 13)]
        public string IndicadorDestino { get; set; }

        [CsvColumn(Name = "RAZON DEL CAMBIO", FieldIndex = 14)]
        public string RazonCambio { get; set; }
    }

}

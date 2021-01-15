using LINQtoCSV;

namespace EllipseSAO900AddIn
{
    /// <summary>
    ///     Clase usada para exportar el documento csv
    /// </summary>
    public class Causaciones
    {
        [CsvColumn(Name = "ACCION", FieldIndex = 1)]
        public string Action { get; set; }

        [CsvColumn(Name = "ITEM", FieldIndex = 2)]
        public string Item { get; set; }

        [CsvColumn(Name = "SUPPLIER", FieldIndex = 3)]
        public string Supplier { get; set; }

        [CsvColumn(Name = "TIPO_DE_DOC", FieldIndex = 4)]
        public string TipoDocumento { get; set; }

        [CsvColumn(Name = "NUM_DE_DOC", FieldIndex = 5)]
        public string NumeroDocumento { get; set; }

        [CsvColumn(Name = "FECHA_SOLCT", FieldIndex = 6)]
        public string FechaSolicitud { get; set; }

        [CsvColumn(Name = "MONEDA", FieldIndex = 7)]
        public string Moneda { get; set; }

        [CsvColumn(Name = "VALOR_TOTAL", FieldIndex = 8)]
        public string ValorTotal { get; set; }

        [CsvColumn(Name = "SOLICITADO_POR", FieldIndex = 9)]
        public string SolicitadorPor { get; set; }

        [CsvColumn(Name = "DISTRITO", FieldIndex = 10)]
        public string Distrito { get; set; }

        [CsvColumn(Name = "C_COSTOS_DETALLE", FieldIndex = 11)]
        public string Centro { get; set; }

        [CsvColumn(Name = "EQUIPO", FieldIndex = 12)]
        public string Equipo { get; set; }

        [CsvColumn(Name = "PROYECTO_WO", FieldIndex = 13)]
        public string ProyectoOrden { get; set; }

        [CsvColumn(Name = "P_W", FieldIndex = 14)]
        public string Ind { get; set; }

        [CsvColumn(Name = "VALOR_PES_o_USD", FieldIndex = 15)]
        public string Valor { get; set; }
    }

}

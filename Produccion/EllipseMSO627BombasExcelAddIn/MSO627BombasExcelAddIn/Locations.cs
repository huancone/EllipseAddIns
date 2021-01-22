using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace MSO627BombasExcelAddIn
{
    public class Locations
    {
        [CsvColumn(FieldIndex = 1)]
        public double X { get; set; }

        [CsvColumn(FieldIndex = 2)]
        public double Y { get; set; }

        [CsvColumn(FieldIndex = 3)]
        public double Z { get; set; }

        [CsvColumn(FieldIndex = 4)]
        public string Sitio { get; set; }

        [CsvColumn(FieldIndex = 5)]
        public string TipoSitio { get; set; }

        [CsvColumn(FieldIndex = 6)]
        public string Tajo { get; set; }

        [CsvColumn(FieldIndex = 7)]
        public string Nombre { get; set; }

    }
}

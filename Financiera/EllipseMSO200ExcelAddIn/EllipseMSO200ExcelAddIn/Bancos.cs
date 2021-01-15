using LINQtoCSV;

namespace EllipseMSO200ExcelAddIn
{
    public class Bancos
    {
        [CsvColumn(FieldIndex = 1)]
        public string CodigoNomina { get; set; }

        [CsvColumn(FieldIndex = 2)]
        public string CodigoMims { get; set; }

        [CsvColumn(FieldIndex = 3)]
        public string NombreInstitucion { get; set; }
    }
}
